"""`Supply Inventory` - nappies, wipes, paint, paper and lunch.

A childcare setting does not run out of toys.  It runs out of nappies
in size 4 on a Tuesday afternoon, and somebody drives to the shop.
This tab exists to stop that: opening stock plus everything in, less
everything used, wasted or written off, equals what is on the shelf -
and how many weeks that will last at the rate you have actually been
using it.

    stock = opening + received - used - waste + adjustments
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, id_col, status_cf)

NAME = "Supply Inventory"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "SupplyPrefix", digits=3, width=10,
               header="Item ID"),
        Col("name", "Item", 26, "text",
            note="Be specific. 'Nappies size 4' and 'Nappies size 5' "
                 "run out on different days."),
        Col("category", "Category", 20, "text",
            validate=list_from("SupplyCategoryList"), cf=status_cf()),
        Col("unit", "Unit", 11, "text",
            note="Pack, box, litre, each. Whatever you count in."),
        Col("pack_size", "Items per unit", 11, "int",
            validate=number_check(0, 10000, decimal=False),
            hidden=True),
        Col("supplier", "Supplier", 20, "text"),
        Col("supplier_phone", "Supplier phone", 15, "text",
            hidden=True),
        Col("lead_days", "Lead time (days)", 11, "int",
            validate=number_check(0, 180, decimal=False), hidden=True,
            note="How long an order takes to arrive. The reorder "
                 "warning allows for it."),
        # -------------------------------------------------- movements --
        Col("opening", "Opening stock", 11, "number1",
            validate=number_check(0, 100000),
            note="What was on the shelf when you started tracking. "
                 "The date is on Settings."),
        Col("received", "Received", 10, "number1",
            validate=number_check(0, 100000),
            note="Everything that has come in since. Add to this "
                 "figure each time a delivery arrives."),
        Col("used", "Used", 10, "number1",
            validate=number_check(0, 100000)),
        Col("waste", "Waste / damaged", 11, "number1",
            validate=number_check(0, 100000)),
        Col("adjust", "Adjustment", 10, "number1",
            validate=number_check(-100000, 100000),
            note="For stock counts that do not agree with the "
                 "arithmetic. Positive adds, negative takes away."),
        Col("stock", "On the shelf", 11, "number1",
            formula=blank(
                "name",
                'ROUND(IF({opening}{r}="",0,{opening}{r})+'
                'IF({received}{r}="",0,{received}{r})-'
                'IF({used}{r}="",0,{used}{r})-'
                'IF({waste}{r}="",0,{waste}{r})+'
                'IF({adjust}{r}="",0,{adjust}{r}),2)'),
            cf=[{"type": "cell", "criteria": "<=", "value": 0,
                 "fmt": "bold_bad"},
                {"type": "data_bar", "bar_color": Palette.primary_mid}],
            note="Opening plus received, less used, less waste, plus "
                 "adjustments. If it goes negative, something was "
                 "used that was never booked in."),
        # ----------------------------------------------------- money ---
        Col("cost", "Cost a unit", 11, "money",
            validate=number_check(0, 100000)),
        Col("value", "Stock value", 11, "money",
            formula=blank(
                "name",
                'ROUND(MAX(0,{stock}{r})*IF({cost}{r}="",0,'
                '{cost}{r}),2)'),
            cf=bar_cf(Palette.accent)),
        Col("spend", "Spent so far", 11, "money",
            formula=blank(
                "name",
                'ROUND((IF({received}{r}="",0,{received}{r}))*'
                'IF({cost}{r}="",0,{cost}{r}),2)'),
            hidden=True),
        # ------------------------------------------------- how long ----
        Col("weekly_use", "Used a week", 11, "number1",
            formula=blank(
                "name",
                'ROUND(IF({used}{r}="",0,{used}{r})/'
                'MAX(1,TrackingWeeks),2)'),
            note="Everything used, divided by the weeks you have been "
                 "tracking. Your real rate, not an estimate."),
        Col("weeks_cover", "Weeks left", 10, "number1",
            formula=blank(
                "name",
                'IF({weekly_use}{r}<=0,"",'
                'ROUND(MAX(0,{stock}{r})/{weekly_use}{r},1))'),
            cf=[{"type": "cell", "criteria": "<=", "value": 1,
                 "fmt": "bold_bad"},
                {"type": "cell", "criteria": "<=",
                 "value": "=StockWarningWeeks", "fmt": "warn"},
                {"type": "cell", "criteria": ">",
                 "value": "=ReorderWeeks", "fmt": "good"}],
            note="How long what is on the shelf will last at the rate "
                 "you have been getting through it."),
        Col("min_level", "Minimum level", 11, "number1",
            validate=number_check(0, 100000),
            note="Never go below this. Leave blank and the warning "
                 "uses the weeks-of-cover figure instead."),
        Col("order_qty", "Order this much", 12, "number1",
            formula=blank(
                "name",
                'IF({weekly_use}{r}<=0,'
                'IF({min_level}{r}="","",'
                'MAX(0,ROUND({min_level}{r}-{stock}{r},0))),'
                'MAX(0,ROUND({weekly_use}{r}*ReorderWeeks-'
                '{stock}{r},0)))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}],
            note="Enough to bring you back up to the cover you asked "
                 "for on Settings. Zero means you are fine."),
        Col("order_cost", "Order value", 11, "money",
            formula=blank(
                "name",
                'IF({order_qty}{r}="","",'
                'ROUND({order_qty}{r}*IF({cost}{r}="",0,{cost}{r}),'
                '2))'),
            hidden=True),
        Col("last_ordered", "Last ordered", 11, "date",
            validate=date_check()),
        Col("last_counted", "Last counted", 11, "date",
            validate=date_check(), hidden=True),
        Col("status", "Stock status", 12, "status",
            formula=blank(
                "name",
                'IF({stock}{r}<=0,"Out of stock",'
                'IF(IF({min_level}{r}="",FALSE,'
                '{stock}{r}<={min_level}{r}),"Low stock",'
                'IF({weeks_cover}{r}="","In stock",'
                'IF({weeks_cover}{r}<=StockWarningWeeks,"Low stock",'
                '"In stock"))))'),
            cf=status_cf()),
        Col("notes", "Notes", 24, "wrap"),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "name",
                'IF({stock}{r}<0,"Check - stock has gone negative",'
                'IF({stock}{r}=0,"Out of stock - order today",'
                'IF({status}{r}="Low stock","Order - "&'
                'IF({order_qty}{r}="","now",{order_qty}{r}&" "&'
                'IF({unit}{r}="","units",LOWER({unit}{r}))),'
                'IF({cost}{r}="","Add - no unit cost, no stock value",'
                'IF({supplier}{r}="","Add - no supplier recorded",'
                'IF(IF({weeks_cover}{r}="",FALSE,'
                '{weeks_cover}{r}>26),'
                '"Check - over six months of stock on the shelf",'
                '"OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Supply Inventory",
        subtitle="What is on the shelf, how fast it goes, and what to "
                 "order before somebody has to drive to the shop.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).supplies,
        help_text="List what you keep in stock. Update the received "
                  "and used figures as you go, and the shelf count, "
                  "the weeks of cover and the order quantity follow.",
        extra_notes=[
            "Stock is opening plus received, less used, less waste, "
            "plus adjustments. A negative figure is not a bug - it "
            "means something left the cupboard that was never booked "
            "in.",
            "'Weeks left' uses your actual consumption since the "
            "tracking start date on Settings, so it gets more "
            "accurate the longer you use it.",
            "The order quantity brings you back up to the weeks of "
            "cover you asked for on Settings - raise it before a "
            "holiday, lower it if you are short of storage.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Store cupboard summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "The store cupboard", last)
    sheet.table_head(row, 0, [
        "Items tracked", "In stock", "Low stock", "Out of stock",
        "Stock value", "Spent on supplies", "To order now",
        "Order value", "Items with no supplier"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Supply Inventory.name]],"<>")',
        '=COUNTIFS([[Supply Inventory.status]],"In stock")',
        '=COUNTIFS([[Supply Inventory.status]],"Low stock")',
        '=COUNTIFS([[Supply Inventory.status]],"Out of stock")',
        '=ROUND(SUM([[Supply Inventory.value]]),2)',
        '=ROUND(SUM([[Supply Inventory.spend]]),2)',
        '=COUNTIFS([[Supply Inventory.order_qty]],">0")',
        '=ROUND(SUM([[Supply Inventory.order_cost]]),2)',
        '=COUNTIF([[Supply Inventory.name]],"<>")-'
        'COUNTIF([[Supply Inventory.supplier]],"<>")',
    ], ["int", "int", "int", "int", "money", "money", "int", "money",
        "int"])
