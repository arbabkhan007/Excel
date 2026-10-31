"""`Expense Tracker` - everything the properties cost you.

Two rules keep this tab honest.

First, **nothing is counted twice**.  Cleaning lives on the Cleaning
Schedule and maintenance lives on the Maintenance Log, because both
need their own workflow.  If you also type a cleaning invoice in here
it is flagged and kept out of the profit figure - unless you tell
Settings you would rather do it the other way round, in which case the
other tabs step back instead.

Second, **a recurring cost is shown as what it really costs a month**.
An annual insurance premium and a weekly laundry bill are not
comparable until they both say "per month", and a host staring at an
expense breakdown is always asking the per-month question.
"""

from __future__ import annotations

from ..config import EXPENSE_CATEGORIES, Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, flag_yes_no_cf, id_col, 
                    is_capital, lookup_booking, money_cf, 
                    monthly_equivalent, month_col, weekday_col, 
                    year_col)

NAME = "Expense Tracker"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "ExpensePrefix", digits=cfg.digits("expense"),
               width=10, header="Expense ID"),
        Col("date", "Date", 12, "date",
            validate=date_check(),
            note="The date on the receipt. Every row on this tab starts "
                 "here.",
            group="The cost"),
        month_col("date"),
        year_col("date"),
        weekday_col("date", "day", "Day", 6, hidden=True),
        Col("property", "Property", 11, "text",
            validate=list_from("PropertyList"),
            note="Which property it belongs to. Leave blank for a cost "
                 "that covers the whole portfolio - accounting fees, "
                 "your software subscription, the camera.",
            group="The cost"),
        Col("property_name", "Property name", 18, "text",
            formula=('=IF({date}{r}="","",IF({property}{r}="",'
                     '"Whole portfolio",'
                     'IFERROR(INDEX([[Property Database.name]],'
                     'MATCH({property}{r},[[Property Database.id]],0)),'
                     '"Not in Property Database")))'),
            group="The cost"),
        Col("category", "Category", 18, "text",
            validate=list_from("ExpenseCategoryList"),
            note="Pick from the list so the breakdowns add up. Edit the "
                 "list on Settings if your accountant wants different "
                 "headings.",
            group="The cost"),
        Col("vendor", "Paid to", 20, "text",
            note="Who you paid. Free text - it does not have to be "
                 "someone on the Staff & Vendors tab.",
            group="The cost"),
        Col("description", "What it was for", 30, "text",
            group="The cost"),
        Col("amount", "Amount", 12, "money",
            validate=number_check(-1000000, 1000000, "Amount"),
            cf=money_cf(),
            note="What you paid, including tax. A refund goes in as a "
                 "negative number.",
            group="The cost"),
        Col("tax_amount", "of which tax", 11, "money",
            validate=number_check(0, 1000000, "Tax"),
            note="Optional. The VAT or sales tax inside the amount, if "
                 "you reclaim it.",
            hidden=True, group="The cost"),
        Col("method", "Paid by", 14, "text",
            validate=list_from("PaymentMethodList"), group="The cost"),
        Col("receipt", "Receipt ref", 13, "text",
            note="Where the paperwork is - a file name, a folder, an "
                 "email subject. The thing you will want in March.",
            group="The cost"),

        # ----------------------------------------------- how it behaves --
        Col("frequency", "How often", 13, "text",
            validate=list_from("FrequencyList"),
            note="One off, or how often it repeats. This is what turns "
                 "the amount into a monthly equivalent.",
            group="How it behaves"),
        Col("monthly_equiv", "Per month", 11, "money",
            formula=blank("date",
                          monthly_equivalent("{amount}{r}",
                                             "{frequency}{r}")),
            note="What this cost really is each month. A yearly bill "
                 "divided by twelve, a weekly one times 52 over 12. One "
                 "off costs show nothing here.",
            group="How it behaves"),
        Col("booking", "Booking", 11, "text",
            validate=list_from("BookingList"),
            note="Optional. For a cost caused by one specific stay - a "
                 "replacement, a late check-out cleaner, a damage "
                 "repair. The true-profit-per-booking figure uses it.",
            group="How it behaves"),
        Col("booking_prop", "From booking", 11, "text",
            formula=lookup_booking("property"), hidden=True,
            group="How it behaves"),

        # ----------------------------------------------- how it counts --
        Col("tax_category", "Tax treatment", 17, "text",
            validate=list_from("TaxCategoryList"),
            note="How your accountant will want it grouped. The Tax "
                 "Summary tab totals these - it does not decide them "
                 "for you, because deductibility depends on where you "
                 "are and how you hold the property.",
            group="How it counts"),
        Col("is_cleaning", "Is cleaning", 11, "flag",
            formula=blank(
                "date",
                'IF({category}{r}="Cleaning","Yes",'
                'IF({category}{r}="Laundry","Yes","No"))'),
            hidden=True, group="How it counts"),
        Col("is_maintenance", "Is maintenance", 12, "flag",
            formula=blank(
                "date",
                'IF({category}{r}="Maintenance","Yes",'
                'IF({category}{r}="Repairs","Yes","No"))'),
            hidden=True, group="How it counts"),
        Col("is_capital", "Is capital", 10, "flag",
            formula=blank("date",
                          f'IF({is_capital("{category}{r}")},"Yes","No")'),
            hidden=True,
            note="Furniture and appliances are an asset, not a running "
                 "cost. They are totalled for tax but kept out of "
                 "operating profit.",
            group="How it counts"),
        Col("in_pl", "Counts in profit", 12, "flag",
            formula=blank(
                "date",
                'IF({is_capital}{r}="Yes","No",'
                'IF({is_cleaning}{r}="Yes",'
                'IF(CleaningInExpenses="Yes","Yes","No"),'
                'IF({is_maintenance}{r}="Yes",'
                'IF(MaintenanceInExpenses="Yes","Yes","No"),"Yes")))'),
            cf=flag_yes_no_cf(),
            note="No means this row is already counted somewhere else, "
                 "so adding it to profit would double count it. Both "
                 "switches live on Settings.",
            group="How it counts"),
        Col("pl_amount", "Counted as", 12, "money",
            formula=blank("date",
                          'IF({in_pl}{r}="Yes",{amount}{r},0)'),
            note="The figure the dashboard, the monthly report and the "
                 "per-property profit actually add up.",
            group="How it counts"),

        Col("alert", "Action", 26, "status",
            formula=blank(
                "date",
                'IF({amount}{r}="","Add - no amount",'
                'IF({category}{r}="","Add - no category",'
                'IF({in_pl}{r}="No",'
                'IF({is_capital}{r}="Yes","OK - capital, not a running '
                'cost","OK - counted on its own tab"),'
                'IF({receipt}{r}="","Check - no receipt reference",'
                'IF({tax_category}{r}="","Check - no tax treatment",'
                '"OK")))))'),
            cf=alert_cf(), group="How it counts"),
        Col("notes", "Notes", 28, "wrap", group="How it counts"),
    ]

    return SheetSpec(
        name=NAME,
        title="Expense Tracker",
        subtitle="Everything the properties cost you, counted once and "
                 "shown per month",
        columns=columns,
        rows=cfg.row_count(NAME),
        freeze_cols=2,
        samples=sample("expenses", cfg),
        help_text="One row per cost. Attach it to a property, or leave "
                  "the property blank for something that covers the "
                  "whole portfolio. Cleaning and maintenance typed in "
                  "here are flagged and excluded from profit, because "
                  "they already have their own tabs - the two switches "
                  "on Settings reverse that if you would rather run it "
                  "the other way.",
        extra_notes=[
            "Furniture and appliances are treated as capital, not as a "
            "running cost. They are still totalled on the Tax Summary, "
            "but they do not drag down the month they were bought in.",
            "Attach a cost to a booking when one particular stay caused "
            "it. That is what makes the true-profit-per-booking figure "
            "on the Monthly Reports tab worth looking at.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Totals, then the full category breakdown."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index

    row = sheet.section(row, "Everything on this tab", last)
    row = sheet.table_head(row, 0, [
        "Rows", "Total spent", "Counted in profit", "Capital items",
        "Already counted elsewhere", "Recurring per month",
        "Tax inside it", "No receipt"],
        [9, 12, 15, 12, 20, 17, 12, 11])
    sheet.calc_row(row, 0, [
        '=COUNT([[Expense Tracker.date]])',
        '=ROUND(SUM([[Expense Tracker.amount]]),2)',
        '=ROUND(SUM([[Expense Tracker.pl_amount]]),2)',
        '=ROUND(SUMIFS([[Expense Tracker.amount]],'
        '[[Expense Tracker.is_capital]],"Yes"),2)',
        '=ROUND(SUMIFS([[Expense Tracker.amount]],'
        '[[Expense Tracker.in_pl]],"No")-'
        'SUMIFS([[Expense Tracker.amount]],'
        '[[Expense Tracker.is_capital]],"Yes"),2)',
        '=ROUND(SUM([[Expense Tracker.monthly_equiv]]),2)',
        '=ROUND(SUM([[Expense Tracker.tax_amount]]),2)',
        '=COUNTIFS([[Expense Tracker.receipt]],"",'
        '[[Expense Tracker.date]],">0")',
    ], ["int", "money", "money", "money", "money", "money", "money",
        "int"])
    row += 2

    row = sheet.section(row, "Where the money goes", last)
    row = sheet.table_head(row, 0, [
        "Category", "Rows", "Total", "Counted in profit", "Per month",
        "Share of spend", "Biggest single"],
        [20, 8, 12, 15, 11, 13, 14])
    total = 'SUM([[Expense Tracker.amount]])'
    for name in EXPENSE_CATEGORIES:
        key = f'"{name}"'
        cat = f'[[Expense Tracker.category]],{key}'
        amt = f'SUMIFS([[Expense Tracker.amount]],{cat})'
        sheet.calc_row(row, 0, [
            name,
            f"=COUNTIFS({cat})",
            f"=ROUND({amt},2)",
            f'=ROUND(SUMIFS([[Expense Tracker.pl_amount]],{cat}),2)',
            f'=ROUND(SUMIFS([[Expense Tracker.monthly_equiv]],{cat}),2)',
            f'=IF({total}<=0,"",{amt}/{total})',
            f'=IF(COUNTIFS({cat})<=0,"",'
            f'SUMPRODUCT(MAX(([[Expense Tracker.category]]={key})*'
            f'IFERROR([[Expense Tracker.amount]]*1,0))))',
        ], ["text", "int", "money", "money", "money", "percent", "money"])
        row += 1
