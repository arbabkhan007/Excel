"""`Supplies` - the cupboard, and what it is costing you.

Count the stock once, log what you use on the Supply Usage tab, and
this tab works out what is left, how fast it is going and when to
reorder. The cost per use is what makes a job costing honest.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, flag_yes_no_cf, id_col, status_cf

NAME = "Supplies"

_USE = "[[Supply Usage.supply]]"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "SupplyPrefix", digits=4, width=10),
        Col("name", "Item", 26, "text",
            note="What you call it on the shopping list. The Supply "
                 "Usage tab picks from these names."),
        Col("category", "Category", 20, "text",
            validate=list_from("SupplyCategoryList")),
        Col("supplier", "Supplier", 18, "text"),
        Col("unit", "Unit", 11, "text",
            validate=list_from("SupplyUnitList"),
            note="The unit you use it in - a bottle, a litre, a "
                 "roll. Buy in packs, use in units."),
        Col("pack_size", "Units in a pack", 10, "number1",
            validate=number_check(0, 100000),
            note="Six bottles in a case, 24 rolls in a pack."),
        Col("pack_cost", "Pack cost", 11, "money",
            validate=number_check(0, 100000)),
        Col("stock", "Counted stock", 11, "number1",
            validate=number_check(0, 1000000),
            note="How many units you had when you last counted. "
                 "Count the cupboard once a month and type it here."),
        Col("count_date", "Counted on", 11, "date", validate=date_check(),
            note="Usage logged after this date is taken off the "
                 "counted stock."),
        Col("reorder", "Reorder at", 10, "number1",
            validate=number_check(0, 1000000),
            note="The level that triggers a 'low stock' flag."),
        Col("order_qty", "Order quantity", 10, "number1",
            validate=number_check(0, 1000000)),
        Col("last_order", "Last ordered", 11, "date",
            validate=date_check()),
        Col("eco", "Eco / low tox?", 10, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="Worth tracking - more and more clients ask."),
        Col("coshh", "Hazard sheet?", 10, "flag",
            validate=list_from("YesNoNAList"), cf=flag_yes_no_cf(),
            note="Any chemical your staff use should have its safety "
                 "data sheet on file. The Safety tab tracks them."),
        Col("notes", "Notes", 28, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("unit_cost", "Cost a unit", 10, "money",
            formula='=IF(OR({name}{r}="",{pack_cost}{r}="",'
                    '{pack_size}{r}="",{pack_size}{r}=0),"",'
                    'ROUND({pack_cost}{r}/{pack_size}{r},4))',
            note="What one use of it costs. This is the figure the "
                 "Supply Usage tab multiplies."),
        Col("used_total", "Used (all time)", 11, "number1",
            formula='=IF({name}{r}="","",SUMIFS('
                    '[[Supply Usage.qty]],' + _USE + ',{name}{r}))'),
        Col("used_since", "Used since the count", 11, "number1",
            formula='=IF(OR({name}{r}="",{count_date}{r}=""),"",'
                    'SUMIFS([[Supply Usage.qty]],' + _USE + ',{name}{r},'
                    '[[Supply Usage.date]],">"&{count_date}{r}))',
            hidden=True),
        Col("on_hand", "On hand", 10, "number1",
            formula=blank(
                "name",
                'IF({stock}{r}="","",{stock}{r}-IF({used_since}{r}="",0,'
                '{used_since}{r}))'),
            cf=[{"type": "cell", "criteria": "<=", "value": 0,
                 "fmt": "bad"}],
            note="Counted stock less everything logged since the "
                 "count."),
        Col("used_28", "Used in 28 days", 10, "number1",
            formula='=IF({name}{r}="","",SUMIFS([[Supply Usage.qty]],' +
                    _USE + ',{name}{r},[[Supply Usage.date]],'
                    '">="&TodayDate-28))',
            hidden=True),
        Col("daily_use", "Used a day", 9, "number1",
            formula='=IF({name}{r}="","",ROUND({used_28}{r}/28,3))',
            hidden=True),
        Col("days_cover", "Days of cover", 10, "int",
            formula='=IF(OR({name}{r}="",{on_hand}{r}="",'
                    '{daily_use}{r}=0),"",ROUND({on_hand}{r}/'
                    '{daily_use}{r},0))',
            cf=[{"type": "cell", "criteria": "<", "value": "=LowStockDays",
                 "fmt": "warn"},
                {"type": "cell", "criteria": "<=", "value": 0,
                 "fmt": "bad"}],
            note="At the rate of the last four weeks. Blank means "
                 "you have not used any recently."),
        Col("stock_value", "Stock value", 11, "money",
            formula='=IF(OR({name}{r}="",{on_hand}{r}="",'
                    '{unit_cost}{r}=""),"",ROUND({on_hand}{r}*'
                    '{unit_cost}{r},2))',
            note="What is sitting in the cupboard. Add the column up "
                 "and most owners get a surprise."),
        Col("spend", "Spent on it", 11, "money",
            formula='=IF({name}{r}="","",SUMIFS('
                    '[[Supply Usage.cost]],' + _USE + ',{name}{r}))'),
        Col("status", "Stock", 13, "status",
            formula=blank(
                "name",
                'IF({on_hand}{r}="","Not counted",'
                'IF({on_hand}{r}<=0,"Out of stock",'
                'IF(AND({reorder}{r}<>"",{on_hand}{r}<={reorder}{r}),'
                '"Low stock",IF(AND({days_cover}{r}<>"",'
                '{days_cover}{r}<LowStockDays),"Low stock",'
                '"In stock"))))'),
            cf=status_cf()),
        Col("alert", "Action", 26, "status",
            formula=blank(
                "name",
                'IF({status}{r}="Out of stock","Order - none left",'
                'IF({status}{r}="Low stock","Order - running low",'
                'IF({status}{r}="Not counted","Add - count the stock",'
                'IF(AND({coshh}{r}="No",{category}{r}='
                '"Chemicals & detergents"),'
                '"Action - no safety data sheet","OK"))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Cleaning Supplies & Inventory",
        subtitle="What you hold, what you use and what it is really "
                 "costing per job.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).supplies,
        help_text="Count the cupboard once and put the date against "
                  "it. From then on, log what you use on the Supply "
                  "Usage tab and the stock level follows by itself.",
        extra_notes=[
            "Cost per unit is the number that matters. A $14 bottle "
            "of degreaser that does eight ovens costs $1.75 a job - "
            "that belongs in your price.",
            "Days of cover beats a reorder level on its own, "
            "because it notices when a new contract doubles your "
            "usage.",
            "Any chemical your staff use should have a safety data "
            "sheet. Keep them with the file - the Safety tab tracks "
            "which ones you hold.",
        ],
    )
