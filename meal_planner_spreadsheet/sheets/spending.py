"""`Grocery Spending` - what you actually spent.

The planned cost on the meal plan is a forecast. This is the receipt.
Log the shop (or each line of it if you are keeping a close eye on
prices) and the budget tabs compare the two, month by month.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, id_col, month_col, status_cf, week_col,
                     weekday_col)

NAME = "Grocery Spending"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "SpendPrefix", digits=4, width=10),
        Col("date", "Date", 11, "date", validate=date_check()),
        weekday_col("date"),
        week_col("date"),
        month_col("date"),
        Col("store", "Store", 18, "text", validate=list_from("StoreNames")),
        Col("type", "Type of spend", 15, "text",
            validate=list_from("SpendTypeList"),
            note="Weekly shop, top-up, takeaway, eating out, treat. "
                 "The budget splits on this."),
        Col("item", "Item or description", 24, "text",
            validate=list_from("IngredientNames"),
            note="Pick an ingredient to track its price over time, or "
                 "just type 'weekly shop'."),
        Col("qty", "Quantity", 9, "number1",
            validate=number_check(0, 100000)),
        Col("unit", "Unit", 8, "text", validate=list_from("UnitList")),
        Col("amount", "Amount", 11, "money",
            validate=number_check(0, 1000000),
            note="What left your account."),
        Col("unit_price", "Price per unit", 11, "money",
            formula='=IF(OR({amount}{r}="",{qty}{r}="",{qty}{r}=0),"",'
                    '{amount}{r}/{qty}{r})',
            hidden=True),
        Col("payer", "Who paid", 14, "text",
            validate=list_from("MemberNames")),
        Col("method", "Paid with", 13, "text",
            validate=list_from("PaymentList")),
        Col("planned", "Planned?", 9, "flag",
            validate=list_from("YesNoList"),
            note="No means it was an impulse buy - the dashboard "
                 "counts them."),
        Col("budget_line", "Budget line", 15, "text",
            validate=list_from("BudgetLineList")),
        Col("waste", "Thrown away?", 10, "flag",
            validate=list_from("YesNoList"),
            note="Set to Yes if this one ended up in the bin - it "
                 "feeds the waste report."),
        Col("receipt", "Receipt link", 18, "link"),
        Col("month_budget", "That month's budget", 11, "money",
            formula="=MonthlyBudget", hidden=True),
        Col("running_month", "Spent this month", 12, "money",
            formula='=IF({date}{r}="","",SUMIFS({amount_first}:'
                    '{amount_last},{month_first}:{month_last},'
                    '{month}{r}))',
            note="Everything logged in the same month as this row."),
        Col("over", "Over budget?", 12, "status",
            formula='=IF({date}{r}="","",IF({running_month}{r}>'
                    'MonthlyBudget,"Check - over the monthly budget",'
                    '"OK"))',
            cf=alert_cf(), hidden=True),
        Col("status", "Status", 12, "status",
            formula='=IF({date}{r}="","",IF({amount}{r}="",'
                    '"Add - no amount",IF({waste}{r}="Yes","Wasted",'
                    'IF({planned}{r}="No","Check - unplanned",'
                    '"Purchased"))))',
            cf=status_cf()),
        Col("notes", "Notes", 24, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Grocery Spending",
        subtitle="Every shop, top-up and takeaway, so the budget tabs "
                 "can compare what you planned with what you paid.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).spending,
        help_text="One line per shop is plenty to track a budget. Line "
                  "by line entry is only worth it for the handful of "
                  "items whose price you want to watch.",
        extra_notes=[
            "Choosing an ingredient from the drop-down lets the price "
            "book show what you have actually been paying - useful "
            "when the shelf price and the real price drift apart.",
            "Takeaways and meals out belong here too, under their own "
            "spend type, otherwise the monthly food total lies to you.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    spec_ = reg[NAME]
    s = sheet
    last = spec_.last_col_index
    row = spec_.last_data_row + 1
    row = s.section(row, "Spending summary", last)
    row = s.table_head(row, 0, ["Shops logged", "Total spent",
                                "This month", "Monthly budget",
                                "Left this month", "Eating out",
                                "Unplanned buys", "Average shop"],
                       [13, 13, 13, 14, 15, 13, 14, 14])
    diff_row = row
    row = s.calc_row(row, 0, [
        '=COUNTIF([[Grocery Spending.date]],">0")',
        '=SUM([[Grocery Spending.amount]])',
        '=SUMIFS([[Grocery Spending.amount]],'
        '[[Grocery Spending.month]],TEXT(TODAY(),"yyyy-mm"))',
        "=MonthlyBudget",
        '=MonthlyBudget-SUMIFS([[Grocery Spending.amount]],'
        '[[Grocery Spending.month]],TEXT(TODAY(),"yyyy-mm"))',
        '=SUMIFS([[Grocery Spending.amount]],'
        '[[Grocery Spending.type]],"Eating out")+'
        'SUMIFS([[Grocery Spending.amount]],'
        '[[Grocery Spending.type]],"Takeaway")',
        '=COUNTIF([[Grocery Spending.planned]],"No")',
        '=IFERROR(AVERAGEIFS([[Grocery Spending.amount]],'
        '[[Grocery Spending.type]],"Weekly shop"),"")',
    ], ["int", "money", "money", "money", "money", "money", "int",
        "money"])
    s.ws.conditional_format(diff_row, 4, diff_row, 4,
                            {"type": "cell", "criteria": "<", "value": 0,
                             "format": fmts.cf("bad")})
    s.ws.conditional_format(diff_row, 4, diff_row, 4,
                            {"type": "cell", "criteria": ">=", "value": 0,
                             "format": fmts.cf("good")})
