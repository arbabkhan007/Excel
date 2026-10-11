"""`Food Waste` - the uncomfortable tab.

Every bag of salad that turned to soup in the drawer, with a price on
it. It is the fastest way to find the £300 a year most households throw
out: log it for a month, read the reasons column, and change one thing.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, base_unit_of, id_col, lookup_ingredient,
                     month_col, to_base, week_col)

NAME = "Food Waste"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", '"W-"', digits=4, width=8),
        Col("date", "Date", 11, "date", validate=date_check()),
        week_col("date"),
        month_col("date"),
        Col("item", "What went in the bin", 24, "text",
            validate=list_from("IngredientNames"),
            note="Pick an ingredient where you can - it prices the "
                 "waste for you."),
        Col("other", "Or describe it", 20, "text",
            note="'Half a takeaway', 'the kids' leftovers'."),
        Col("qty", "How much", 10, "number1",
            validate=number_check(0, 100000)),
        Col("unit", "Unit", 9, "text", validate=list_from("UnitList")),
        Col("base_qty", "In base units", 11, "number1",
            formula=to_base("qty", "unit"), hidden=True),
        Col("base_unit", "Base unit", 9, "text",
            formula=base_unit_of("unit"), hidden=True),
        Col("unit_cost", "Cost per unit", 10, "money",
            formula=lookup_ingredient("unit_cost", key="item"),
            hidden=True),
        Col("value", "What it cost", 11, "money",
            formula='=IF(OR({base_qty}{r}="",{unit_cost}{r}="",'
                    'NOT(ISNUMBER({unit_cost}{r}))),{manual_cost}{r},'
                    '{base_qty}{r}*{unit_cost}{r})',
            note="Worked out from the price book, or your own figure "
                 "in the next column."),
        Col("manual_cost", "Your estimate", 11, "money",
            validate=number_check(0, 10000),
            note="Use this for anything not in the price book."),
        Col("reason", "Why", 18, "text", validate=list_from("WasteReasonList"),
            note="Be honest. The pattern in this column is the whole "
                 "point of the tab."),
        Col("avoidable", "Avoidable?", 10, "flag",
            validate=list_from("YesNoList"),
            note="Peelings and bones are not. A forgotten bag of "
                 "spinach is."),
        Col("where", "Where it was", 13, "text",
            validate=list_from("StorageList")),
        Col("disposal", "How it was thrown", 14, "text",
            validate=list_from("DisposalList"),
            note="Compost and food waste collections are better than "
                 "landfill, but not buying it is better still."),
        Col("prevent", "What would have stopped it", 30, "wrap",
            note="'Freeze half on the day', 'buy the small bag', "
                 "'actually check the fridge first'."),
        Col("month_total", "Wasted this month", 11, "money",
            formula='=IF({date}{r}="","",SUMIFS({value_first}:'
                    '{value_last},{month_first}:{month_last},'
                    '{month}{r}))', hidden=True),
        Col("alert", "Check", 22, "status",
            formula='=IF({date}{r}="","",IF({avoidable}{r}="Yes",'
                    'IF({value}{r}>5,"Check - avoidable and expensive",'
                    '"Check - avoidable"),"OK"))',
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Food Waste Log",
        subtitle="What got thrown away, what it cost and what would "
                 "have stopped it.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent_soft,
        freeze_cols=2,
        samples=sample(cfg).waste,
        help_text="Log it for four weeks. Almost nobody keeps this up "
                  "forever, and almost everybody changes what they buy "
                  "after the first month.",
        extra_notes=[
            "The value is worked out from the price book, so a "
            "quarter of a bag of rocket costs what you paid for it, "
            "not a guess.",
            "The 'what would have stopped it' column is the one worth "
            "re-reading before you plan next week.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    spec_ = reg[NAME]
    s = sheet
    last = spec_.last_col_index
    row = spec_.last_data_row + 1
    row = s.section(row, "What waste is costing you", last)
    row = s.table_head(row, 0, ["Items binned", "Total cost",
                                "This month", "Avoidable share",
                                "A year at this rate", "Target a month",
                                "Against target"],
                       [13, 13, 13, 15, 18, 15, 14])
    diff_row = row
    row = s.calc_row(row, 0, [
        '=COUNTIF([[Food Waste.date]],">0")',
        '=SUM([[Food Waste.value]])',
        '=SUMIFS([[Food Waste.value]],[[Food Waste.month]],'
        'TEXT(TODAY(),"yyyy-mm"))',
        '=IFERROR(SUMIFS([[Food Waste.value]],'
        '[[Food Waste.avoidable]],"Yes")/SUM([[Food Waste.value]]),"")',
        '=SUMIFS([[Food Waste.value]],[[Food Waste.month]],'
        'TEXT(TODAY(),"yyyy-mm"))*12',
        "=WasteCostTarget",
        '=WasteCostTarget-SUMIFS([[Food Waste.value]],'
        '[[Food Waste.month]],TEXT(TODAY(),"yyyy-mm"))',
    ], ["int", "money", "money", "percent0", "money", "money", "money"])
    s.ws.conditional_format(diff_row, 6, diff_row, 6,
                            {"type": "cell", "criteria": "<", "value": 0,
                             "format": fmts.cf("bad")})
    s.ws.conditional_format(diff_row, 6, diff_row, 6,
                            {"type": "cell", "criteria": ">=", "value": 0,
                             "format": fmts.cf("good")})
