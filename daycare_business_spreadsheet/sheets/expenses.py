"""`Expense Tracker` - everything that goes out, categorised and budgeted.

One row per payment.  Categories come from a list so that a month's
spending can actually be added up, recurring costs are converted to a
monthly equivalent so that an annual insurance premium does not look
like a disaster in March, and each row carries the budget for its own
category so the variance is visible where the money was spent.

Wages are the one trap.  Staff costs already come from the Payroll
Tracker, so rows in the wage categories are excluded from the profit
and loss by default.  If you would rather keep everything here, switch
'Payroll sits in expenses' to Yes on Settings - but do one or the
other, never both.
"""

from __future__ import annotations

from ..config import FREQUENCY_MONTHS, PAYROLL_EXPENSE_CATEGORIES, \
    Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, id_col, month_col,
                     status_cf)

NAME = "Expense Tracker"


def spec(cfg: Config) -> SheetSpec:
    wage_test = "+".join(f'({{category}}{{r}}="{c}")'
                         for c in PAYROLL_EXPENSE_CATEGORIES)

    # Each recurring frequency turned into "times per month".
    freq = "0"
    for label, months in FREQUENCY_MONTHS.items():
        freq = f'IF({{frequency}}{{r}}="{label}",{months:.6f},{freq})'

    columns = [
        id_col("date", "ExpensePrefix", digits=4, width=10,
               header="Expense ID"),
        Col("date", "Date", 11, "date", validate=date_check()),
        month_col("date"),
        Col("category", "Category", 22, "text",
            validate=list_from("ExpenseCategoryList"), cf=status_cf(),
            note="Pick from the list. Typing your own text means the "
                 "monthly category reports will not find it."),
        Col("description", "What it was", 30, "text"),
        Col("vendor", "Paid to", 20, "text"),
        Col("amount", "Amount", 12, "money",
            validate=number_check(-100000, 1000000),
            note="Net of tax if you reclaim it, gross if you do not. "
                 "Be consistent - the workbook does not mind which."),
        Col("tax", "Tax", 10, "money",
            validate=number_check(0, 100000), hidden=True),
        Col("total", "Total paid", 12, "money",
            formula=blank(
                "date",
                'ROUND({amount}{r}+IF({tax}{r}="",0,{tax}{r}),2)'),
            cf=bar_cf(Palette.accent)),
        Col("method", "Paid by", 14, "text",
            validate=list_from("PaymentMethodList")),
        Col("receipt", "Receipt?", 9, "flag",
            validate=list_from("YesNoList"),
            cf=[{"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "good"}],
            note="The column your accountant will thank you for."),
        Col("receipt_ref", "Receipt ref", 13, "text", hidden=True,
            note="Where the paper or the photo lives. A folder name "
                 "is enough."),
        Col("frequency", "Recurring", 13, "text",
            validate=list_from("FrequencyList"),
            note="How often this cost comes round. One off is fine - "
                 "it just means nothing is projected forward."),
        Col("monthly", "Monthly equivalent", 13, "money",
            formula=blank(
                "date",
                f'ROUND({{total}}{{r}}*({freq}),2)'),
            note="A yearly insurance premium divided by twelve, a "
                 "weekly food bill multiplied by 4.33. This is the "
                 "column that tells you your real monthly overhead."),
        Col("is_wages", "Payroll category?", 11, "flag",
            formula=blank("category",
                          f'IF(({wage_test})>0,"Yes","No")'),
            hidden=True),
        Col("in_pl", "Counts in profit?", 11, "flag",
            formula=blank(
                "date",
                'IF({is_wages}{r}<>"Yes","Yes",'
                'IF(PayrollInExpenses="Yes","Yes","No"))'),
            cf=[{"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "muted"}],
            note="Wage rows are left out of the profit figures "
                 "because the Payroll Tracker already counts them. "
                 "Change the switch on Settings if you would rather "
                 "do it the other way round."),
        Col("pl_amount", "In the P&L", 12, "money",
            formula=blank(
                "date",
                'IF({in_pl}{r}="Yes",{total}{r},0)'),
            hidden=True),
        # ------------------------------------------------- the budget --
        Col("budget", "Monthly budget", 12, "money",
            formula=blank(
                "category",
                'IFERROR(INDEX(BudgetAmount,MATCH({category}{r},'
                'BudgetCategory,0))*1,0)'),
            hidden=True,
            note="From the budget table on Settings & Lists."),
        Col("cat_month", "Category this month", 13, "money",
            formula=blank(
                "date",
                'ROUND(SUMIFS({total_first}:{total_last},'
                '{category_first}:{category_last},{category}{r},'
                '{month_first}:{month_last},{month}{r}),2)'),
            hidden=True),
        Col("variance", "Budget variance", 13, "money",
            formula=blank(
                "category",
                'IF({budget}{r}<=0,"",'
                'ROUND({cat_month}{r}-{budget}{r},2))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<=", "value": 0,
                 "fmt": "good"}],
            note="Actual minus budget for this category in this "
                 "month. Positive means over budget."),
        Col("child_cost", "Cost per child", 11, "money",
            formula=blank(
                "date",
                'IF(ChildrenEnrolled<=0,"",'
                'ROUND({total}{r}/ChildrenEnrolled,2))'),
            hidden=True),
        Col("notes", "Notes", 24, "wrap"),
        Col("alert", "Action", 28, "status",
            formula=blank(
                "date",
                'IF({category}{r}="","Add - pick a category",'
                'IF({amount}{r}="","Add - no amount",'
                'IF({receipt}{r}<>"Yes","Missing - no receipt kept",'
                'IF(IF({budget}{r}<=0,FALSE,'
                '{cat_month}{r}>{budget}{r}*1.1),'
                '"Check - category over budget this month",'
                'IF({vendor}{r}="","Add - who was it paid to?",'
                'IF({frequency}{r}="","Add - one off or recurring?",'
                '"OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Expense & Supply Spending",
        subtitle="Everything that goes out, categorised, converted to "
                 "a monthly equivalent and measured against budget.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).expenses,
        help_text="One row per payment. Pick a category from the "
                  "list, say how often it recurs, and the monthly "
                  "equivalent and the budget variance follow.",
        extra_notes=[
            "'Monthly equivalent' is the number to use when you are "
            "working out what the setting costs to run. A yearly "
            "premium and a weekly food bill are not comparable until "
            "both are expressed per month.",
            "Wage rows are deliberately excluded from the profit "
            "figures, because the Payroll Tracker already counts "
            "them. The switch is on Settings if you want it the "
            "other way round.",
            "Set a monthly budget per category on Settings & Lists "
            "and the variance column tells you which categories are "
            "running away from you, in the month they do it.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Spending summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Spending summary", last)
    sheet.table_head(row, 0, [
        "Payments logged", "Total paid", "Counted in profit",
        "Excluded (payroll)", "Monthly overhead", "This month",
        "Last month", "Budget a month", "Over / under budget",
        "Missing receipts"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Expense Tracker.date]],"<>")',
        '=ROUND(SUM([[Expense Tracker.total]]),2)',
        '=ROUND(SUM([[Expense Tracker.pl_amount]]),2)',
        '=ROUND(SUM([[Expense Tracker.total]])-'
        'SUM([[Expense Tracker.pl_amount]]),2)',
        '=ROUND(SUMIFS([[Expense Tracker.monthly]],'
        '[[Expense Tracker.frequency]],"<>One off"),2)',
        '=ROUND(SUMIFS([[Expense Tracker.total]],'
        '[[Expense Tracker.month]],ThisMonthKey),2)',
        '=ROUND(SUMIFS([[Expense Tracker.total]],'
        '[[Expense Tracker.month]],'
        'TEXT(EDATE(TodayDate,-1),"yyyy-mm")),2)',
        '=ROUND(BudgetMonthlyTotal,2)',
        '=ROUND(SUMIFS([[Expense Tracker.total]],'
        '[[Expense Tracker.month]],ThisMonthKey)-BudgetMonthlyTotal,2)',
        '=COUNTIFS([[Expense Tracker.receipt]],"<>Yes",'
        '[[Expense Tracker.date]],"<>")',
    ], ["int", "money", "money", "money", "money", "money", "money",
        "money", "money", "int"])
