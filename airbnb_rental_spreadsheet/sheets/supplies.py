"""`Supplies Inventory` - linen, loo roll and the things that run out.

Running out of toilet paper on a Friday afternoon is a one-star review
and a forty mile round trip.  This tab works out how fast each line is
actually being consumed - from the number of turnovers that property
has had, not from a guess - and tells you how many weeks of cover is
left before it becomes a problem.

Stock can belong to one property or to a shared store cupboard that
serves the lot.  Leave the property blank and it is treated as shared.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, id_col, money_cf, status_cf)

NAME = "Supplies Inventory"


def spec(cfg: Config) -> SheetSpec:
    cl_prop = "[[Cleaning Schedule.property]]"
    cl_date = "[[Cleaning Schedule.clean_date]]"

    columns = [
        id_col("name", "SupplyPrefix", digits=cfg.digits("supply"),
               width=10, header="Item ID"),
        Col("name", "Item", 24, "text",
            note="Bath towels, king sheets, loo roll, dishwasher "
                 "tablets. One row per thing you re-buy.",
            group="The item"),
        Col("category", "Category", 18, "text",
            validate=list_from("SupplyCategoryList"), group="The item"),
        Col("property", "Property", 11, "text",
            validate=list_from("PropertyList"),
            note="Blank means a shared store cupboard that serves every "
                 "property.",
            group="The item"),
        Col("property_name", "Kept for", 18, "text",
            formula=('=IF({name}{r}="","",IF({property}{r}="",'
                     '"Shared store",'
                     'IFERROR(INDEX([[Property Database.name]],'
                     'MATCH({property}{r},[[Property Database.id]],0)),'
                     '"Not in Property Database")))'),
            group="The item"),
        Col("unit", "Unit", 10, "text",
            note="Each, pack, roll, pair, litre. Whatever you count it "
                 "in.",
            group="The item"),
        Col("supplier", "Supplier", 18, "text", group="The item"),
        Col("sku", "Code / SKU", 13, "text", hidden=True, group="The item"),

        # ----------------------------------------------------- the stock --
        Col("on_hand", "In stock", 10, "number1",
            validate=number_check(0, 100000, "Quantity in stock"),
            note="What is there now. Update it when you count, not "
                 "every time you take one.",
            cf=bar_cf(), group="Stock"),
        Col("min_level", "Minimum", 10, "number1",
            validate=number_check(0, 100000, "Minimum level"),
            note="The level at which you want to be told to order. For "
                 "linen, two full changes per bed is a sensible floor.",
            group="Stock"),
        Col("last_count", "Last counted", 12, "date",
            validate=date_check(), group="Stock"),
        Col("last_restock", "Last restocked", 12, "date",
            validate=date_check(), group="Stock"),
        Col("count_age", "Days since count", 13, "int",
            formula=blank("name",
                          'IF({last_count}{r}="","",'
                          'MAX(0,TodayDate-{last_count}{r}))'),
            hidden=True, group="Stock"),

        # ----------------------------------------------------- the usage --
        Col("per_turn", "Used per turnover", 14, "number1",
            validate=number_check(0, 1000, "Used per turnover"),
            note="How many of these a single changeover gets through. "
                 "Four towels, two sheet sets, one roll of bin bags. "
                 "This is what makes the forecast real.",
            group="Usage"),
        Col("turns_month", "Turnovers a month", 14, "number1",
            formula=blank(
                "name",
                f'IF({{property}}{{r}}="",'
                f'ROUND(COUNTIFS({cl_date},">="&(TodayDate-90),'
                f'{cl_date},"<="&TodayDate)/3,1),'
                f'ROUND(COUNTIFS({cl_prop},{{property}}{{r}},'
                f'{cl_date},">="&(TodayDate-90),'
                f'{cl_date},"<="&TodayDate)/3,1))'),
            note="Counted from the Cleaning Schedule over the last "
                 "ninety days, so it follows your real season rather "
                 "than an assumption.",
            group="Usage"),
        Col("usage_month", "Used a month", 12, "number1",
            formula=blank("name",
                          'ROUND({per_turn}{r}*{turns_month}{r},1)'),
            group="Usage"),
        Col("weeks_cover", "Weeks of cover", 13, "number1",
            formula=blank(
                "name",
                'IF({usage_month}{r}<=0,"",'
                'ROUND({on_hand}{r}/({usage_month}{r}/4.33),1))'),
            cf=[{"type": "cell", "criteria": "<", "value": 1,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<",
                 "value": "=RestockWeeks", "fmt": "warn"},
                {"type": "cell", "criteria": ">=",
                 "value": "=RestockWeeks", "fmt": "good"}],
            note="How long the stock on the shelf will last at the rate "
                 "you are actually turning the property over.",
            group="Usage"),

        # ----------------------------------------------------- the money --
        Col("unit_cost", "Unit cost", 11, "money",
            validate=number_check(0, 100000, "Unit cost"), group="Money"),
        Col("value", "Stock value", 12, "money",
            formula=blank("name",
                          'ROUND({on_hand}{r}*{unit_cost}{r},2)'),
            cf=money_cf(), group="Money"),
        Col("reorder_qty", "Order quantity", 13, "number1",
            validate=number_check(0, 100000, "Order quantity"),
            note="How many you buy at a time. Blank falls back to "
                 "enough to get back above the minimum.",
            group="Money"),
        Col("order_qty_used", "To order", 10, "number1",
            formula=blank(
                "name",
                'IF({needs_order}{r}<>"Yes",0,'
                'IF({reorder_qty}{r}<>"",{reorder_qty}{r},'
                'MAX(0,{min_level}{r}*2-{on_hand}{r})))'),
            group="Money"),
        Col("order_cost", "Order cost", 11, "money",
            formula=blank("name",
                          'ROUND({order_qty_used}{r}*{unit_cost}{r},2)'),
            group="Money"),
        Col("cost_per_turn", "Cost a turnover", 13, "money",
            formula=blank("name",
                          'ROUND({per_turn}{r}*{unit_cost}{r},2)'),
            hidden=True,
            note="What this one line adds to every changeover.",
            group="Money"),

        # --------------------------------------------------- the verdict --
        Col("needs_order", "Order now", 10, "flag",
            formula=blank(
                "name",
                'IF({min_level}{r}<=0,"No",'
                'IF({on_hand}{r}<={min_level}{r},"Yes",'
                'IF({weeks_cover}{r}="","No",'
                'IF({weeks_cover}{r}<RestockWeeks,"Yes","No"))))'),
            hidden=True, group="Verdict"),
        Col("status", "Stock status", 13, "status",
            formula=blank(
                "name",
                'IF({on_hand}{r}<=0,"Out of stock",'
                'IF({min_level}{r}<=0,"Not tracked",'
                'IF({on_hand}{r}<={min_level}{r},"Order now",'
                'IF({needs_order}{r}="Yes","Low",'
                'IF({on_hand}{r}>{min_level}{r}*4,"Overstocked",'
                '"In stock")))))'),
            cf=status_cf(), group="Verdict"),
        Col("alert", "Action", 28, "status",
            formula=blank(
                "name",
                'IF({on_hand}{r}<=0,"Out of stock - order today",'
                'IF({min_level}{r}<=0,"Add - set a minimum level",'
                'IF({needs_order}{r}="Yes",'
                'IF({order_cost}{r}>0,"Order - '
                '"&TEXT({order_qty_used}{r},"0")&" "&{unit}{r},'
                '"Order - stock is low"),'
                'IF({count_age}{r}="","Check - never counted",'
                'IF({count_age}{r}>90,"Check - not counted in three '
                'months","OK")))))'),
            cf=alert_cf(), group="Verdict"),
        Col("notes", "Notes", 26, "wrap", group="Verdict"),
    ]

    return SheetSpec(
        name=NAME,
        title="Supplies Inventory",
        subtitle="Linen, consumables and the things that run out on a "
                 "Friday afternoon",
        columns=columns,
        rows=cfg.row_count(NAME),
        freeze_cols=2,
        samples=sample("supplies", cfg),
        help_text="One row per thing you re-buy. Type what is on the "
                  "shelf, the minimum you want to keep and how many a "
                  "single turnover gets through - the weeks of cover, "
                  "the order quantity and the order cost work "
                  "themselves out from the number of changeovers that "
                  "property has actually had.",
        extra_notes=[
            "Turnovers a month is counted from the Cleaning Schedule "
            "over the last ninety days, so a property that is busy in "
            "August and quiet in February gets a different answer in "
            "each.",
            "Leave the property blank for a shared store cupboard. Its "
            "usage is then based on every turnover across the "
            "portfolio.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index

    row = sheet.section(row, "The store cupboard", last)
    row = sheet.table_head(row, 0, [
        "Lines tracked", "In stock", "Low", "Order now", "Out of stock",
        "Stock value", "Cost to restock", "Cost a turnover"],
        [13, 10, 8, 11, 13, 12, 15, 14])
    sheet.calc_row(row, 0, [
        '=COUNTIF([[Supplies Inventory.name]],"?*")',
        '=COUNTIFS([[Supplies Inventory.status]],"In stock")',
        '=COUNTIFS([[Supplies Inventory.status]],"Low")',
        '=COUNTIFS([[Supplies Inventory.status]],"Order now")',
        '=COUNTIFS([[Supplies Inventory.status]],"Out of stock")',
        '=ROUND(SUM([[Supplies Inventory.value]]),2)',
        '=ROUND(SUM([[Supplies Inventory.order_cost]]),2)',
        '=ROUND(SUM([[Supplies Inventory.cost_per_turn]]),2)',
    ], ["int", "int", "int", "int", "int", "money", "money", "money"])
    row += 2

    sheet.note(row, "'Cost a turnover' adds up what every consumable "
                    "line contributes to a single changeover. Compare "
                    "it with the consumables figure on Settings - if "
                    "they disagree, the Settings number is the one "
                    "making your turnover costs wrong.", last)
