"""`Expenses` - everything the salon spends, in the order a P&L wants it.

Rent, product, wages, card fees, the coffee for the waiting area.
Recurring costs carry a frequency so the workbook can tell you what
your fixed monthly overhead really is - the number that decides how
many heads you need through the door before you have made anything.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, month_col, rank_desc,
                     status_cf, week_col, yesno)

NAME = "Expenses"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "ExpensePrefix", digits=4, width=11),
        Col("date", "Date", 11, "date", validate=date_check(),
            note="The day the cost belongs to, not necessarily the "
                 "day it leaves the bank."),
        Col("category", "Category", 22, "text",
            validate=list_from("ExpenseCategoryList"),
            note="These are the lines of your profit and loss and the "
                 "rows of your budget. Keep to the list."),
        Col("payee", "Paid to", 20, "text",
            note="Landlord, wholesaler, the council, the card "
                 "processor."),
        Col("description", "What for", 28, "text"),
        Col("amount", "Amount", 11, "money",
            validate=number_check(0, 1000000),
            note="Before tax, if you can reclaim it. Including tax if "
                 "you cannot - just be consistent."),
        Col("tax", "Tax", 10, "money", validate=number_check(0, 1000000),
            note="Sales tax or VAT on the bill, if you track it "
                 "separately."),
        Col("method", "Method", 14, "text",
            validate=list_from("PaymentMethodList")),
        Col("recurring", "Recurring?", 11, "flag", validate=yesno(),
            note="Yes for anything that comes round again - rent, "
                 "insurance, software, the music licence."),
        Col("frequency", "How often", 13, "text",
            validate=list_from("FrequencyList"),
            note="Only used when recurring is Yes. It is what turns a "
                 "yearly insurance bill into a monthly overhead."),
        Col("due", "Due", 11, "date", validate=date_check(),
            note="When it has to be paid by. Leave blank for anything "
                 "already settled."),
        Col("paid", "Paid?", 9, "flag", validate=yesno()),
        Col("paid_on", "Paid on", 11, "date", validate=date_check()),
        Col("reference", "Reference", 16, "text",
            note="Invoice or receipt number - the thing your "
                 "accountant will ask for."),
        Col("notes", "Notes", 22, "wrap"),
        # ---------------------------------------------------- worked out --
        month_col("date"),
        week_col("date"),
        Col("total", "Total", 11, "money",
            formula=blank(
                "date",
                'ROUND(IF({amount}{r}="",0,{amount}{r})+'
                'IF({tax}{r}="",0,{tax}{r}),2)'),
            cf=[{"type": "data_bar", "bar_color": Palette.primary_mid}],
            note="Amount plus tax. The profit and loss uses the "
                 "amount; this column is what actually left the bank."),
        Col("monthly", "Monthly equivalent", 14, "money",
            formula=blank(
                "date",
                'IF({recurring}{r}<>"Yes",0,'
                'ROUND(IF({amount}{r}="",0,{amount}{r})*'
                'IF({frequency}{r}="Weekly",52/12,'
                'IF({frequency}{r}="Fortnightly",26/12,'
                'IF({frequency}{r}="Four weekly",13/12,'
                'IF({frequency}{r}="Monthly",1,'
                'IF({frequency}{r}="Quarterly",1/3,'
                'IF({frequency}{r}="Twice a year",1/6,'
                'IF({frequency}{r}="Yearly",1/12,0))))))),2))'),
            note="A yearly bill spread over twelve months, a weekly "
                 "one multiplied up. Add this column and you have "
                 "your fixed monthly overhead."),
        Col("is_fixed", "Fixed cost?", 11, "flag",
            formula=blank(
                "date",
                'IF({recurring}{r}="Yes","Yes","No")'),
            hidden=True),
        Col("days_late", "Days late", 10, "int",
            formula=blank(
                "date",
                'IF({paid}{r}="Yes","",IF({due}{r}="","",'
                'TodayDate-{due}{r}))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "bad"}],
            hidden=True),
        Col("status", "Status", 12, "status",
            formula=blank(
                "date",
                'IF({paid}{r}="Yes","Paid",'
                'IF({due}{r}="","Pending",'
                'IF({due}{r}<TodayDate,"Overdue",'
                'IF({due}{r}<=TodayDate+7,"Due soon","Pending"))))'),
            cf=status_cf()),
        Col("budget", "Budget this month", 13, "money",
            formula=blank(
                "date",
                'IFERROR(INDEX(BudgetAmount,MATCH({category}{r},'
                'BudgetCategory,0))*1,0)'),
            hidden=True,
            note="The monthly budget for this category, from the "
                 "Goals & Budget tab."),
        rank_desc("total", guard='{date}{r}=""', header="Size rank"),
        Col("alert", "Check", 26, "status",
            formula=blank(
                "date",
                'IF({amount}{r}="","Add - no amount",'
                'IF({category}{r}="","Add - no category, it will miss '
                'the P&L",'
                'IF({status}{r}="Overdue","Overdue - pay it",'
                'IF({status}{r}="Due soon","Due - coming up",'
                'IF(IF({recurring}{r}="Yes",{frequency}{r}="",FALSE),'
                '"Add - how often does it recur?",'
                'IF({paid}{r}="Yes","OK","Pending"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Expenses",
        subtitle="Every cost, categorised the way the profit and loss "
                 "needs it, with the recurring ones spread properly.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).expenses,
        help_text="Date, category, amount. Mark the regulars as "
                  "recurring and give them a frequency - that is what "
                  "makes the overhead figure believable.",
        extra_notes=[
            "Wages do not belong here. They come from the payroll tab "
            "so they are counted once and only once - the profit and "
            "loss picks them up from there.",
            "Stock bought for resale is a purchase, not an expense "
            "while it sits on the shelf. Log it on the Suppliers & "
            "Purchases tab; the profit and loss counts the cost of "
            "what you actually sold.",
            "The monthly equivalent column is what turns a pile of "
            "bills into a break-even number. Rent plus insurance plus "
            "software plus the loan, spread evenly, is what you have "
            "to earn before anything is yours.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Expense summary under the grid."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "What it costs to open the doors", last)
    sheet.table_head(row, 0, [
        "Costs logged", "Total", "This month", "Fixed monthly overhead",
        "Unpaid", "Overdue", "Biggest single cost", "Biggest category"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Expenses.date]],"<>")',
        '=ROUND(SUM([[Expenses.amount]]),2)',
        '=ROUND(SUMIFS([[Expenses.amount]],[[Expenses.month]],'
        'ThisMonthKey),2)',
        '=ROUND(SUM([[Expenses.monthly]]),2)',
        '=ROUND(SUMIFS([[Expenses.total]],[[Expenses.paid]],"<>Yes"),2)',
        '=COUNTIFS([[Expenses.status]],"Overdue")',
        '=ROUND(MAX([[Expenses.amount]]),2)',
        '=IFERROR(INDEX([[Expenses.category]],MATCH(MAX('
        '[[Expenses.total]]),[[Expenses.total]],0)),"")',
    ], ["int", "money", "money", "money", "money", "int", "money",
        "text"])
