"""`Expenses` - everything the business spends.

Wages, products, fuel, insurance, advertising, software. One row per
payment out. The dashboard's profit figure is this tab subtracted from
the invoices, so what is missing here is money you think you have.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, flag_yes_no_cf, id_col, month_col

NAME = "Expenses"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "ExpensePrefix", digits=4, width=11),
        Col("date", "Date", 11, "date", validate=date_check()),
        month_col("date"),
        Col("category", "Category", 22, "text",
            validate=list_from("ExpenseCategoryList"),
            note="Keep to the list. Consistent categories are what "
                 "make the yearly comparison worth reading."),
        Col("type", "Type", 13, "text", validate=list_from("ExpenseTypeList"),
            note="Fixed costs happen whether you work or not. "
                 "Variable costs follow the work. Knowing which is "
                 "which is most of financial planning."),
        Col("description", "What it was", 32, "text"),
        Col("supplier", "Paid to", 20, "text"),
        Col("amount", "Amount (net)", 12, "money",
            validate=number_check(0, 10000000),
            note="Before tax, if you can reclaim it. Otherwise just "
                 "put the amount you paid."),
        Col("tax", "Tax", 10, "money", validate=number_check(0, 1000000)),
        Col("method", "Paid by", 14, "text",
            validate=list_from("PaymentMethodList")),
        Col("job", "Job ID", 11, "id", validate=list_from("JobIDs"),
            note="Only for a cost that belongs to one job - a skip, "
                 "a specialist product, a subcontractor."),
        Col("employee", "Who spent it", 16, "text",
            validate=list_from("EmployeeNames")),
        Col("receipt", "Receipt?", 9, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="No receipt usually means no deduction. Photograph "
                 "them in the van and put the link in the notes."),
        Col("recurring", "Recurring?", 9, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf()),
        Col("deductible", "Claimable?", 9, "flag",
            validate=list_from("YesNoNAList"), cf=flag_yes_no_cf(),
            note="Your accountant decides this, not the "
                 "spreadsheet. The column is here so you can mark "
                 "what you have asked about."),
        Col("notes", "Notes", 28, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("total", "Total paid", 12, "money",
            formula=blank("date",
                          'IF({amount}{r}="",0,{amount}{r})+'
                          'IF({tax}{r}="",0,{tax}{r})'),
            cf=[{"type": "data_bar", "bar_color": "#C86B6B"}]),
        Col("month_total", "This category, this month", 12, "money",
            formula='=IF(OR({date}{r}="",{category}{r}=""),"",'
                    'SUMIFS({total_first}:{total_last},'
                    '{category_first}:{category_last},{category}{r},'
                    '{month_first}:{month_last},{month}{r}))',
            hidden=True),
        Col("share", "Share of the month", 10, "percent0",
            formula='=IF(OR({date}{r}="",{month}{r}=""),"",'
                    'IFERROR({month_total}{r}/SUMIFS('
                    '{total_first}:{total_last},'
                    '{month_first}:{month_last},{month}{r}),""))',
            hidden=True),
        Col("client_name", "On behalf of", 18, "text",
            formula='=IF({job}{r}="","",IFERROR(INDEX('
                    '[[Jobs.client_name]],MATCH({job}{r},[[Jobs.id]],0)),'
                    '""))',
            hidden=True),
        Col("alert", "Check", 24, "status",
            formula=blank(
                "date",
                'IF({amount}{r}="","Add - how much?",'
                'IF({category}{r}="","Add - which category?",'
                'IF({receipt}{r}="No","Action - no receipt",'
                'IF(AND({amount}{r}>500,{description}{r}=""),'
                '"Add - describe a large payment","OK"))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Business Expenses",
        subtitle="Everything that goes out, grouped the way your "
                 "accountant wants to see it.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).expenses,
        help_text="One row per payment. Use the same categories every "
                  "time - that is what makes the monthly comparison "
                  "on the Reports tab worth reading.",
        extra_notes=[
            "Wages belong here too, even if you also keep "
            "timesheets. The dashboard's profit is invoices less "
            "this tab, so leaving wages out flatters you badly.",
            "Mileage has its own tab because it is usually claimed "
            "at a rate per mile rather than as a receipt. Do not "
            "enter it twice.",
            "This is a record of what you spent, not tax advice. "
            "What is and is not claimable depends on where you are "
            "and what you do - ask an accountant.",
        ],
    )
