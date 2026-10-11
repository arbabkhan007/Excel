"""`Stores` - where you shop.

Opening hours, delivery slots, loyalty cards and - the useful bit - how
much of your money each shop takes and how its prices compare.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import alert_cf, id_col, status_cf

NAME = "Stores"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", '"ST-"', digits=2, width=8),
        Col("name", "Store", 20, "text",
            note="Supermarkets, the market stall, the butcher, the "
                 "online shop - anywhere food money goes."),
        Col("type", "Type", 15, "text", validate=list_from("StoreTypeList")),
        Col("location", "Where", 20, "text"),
        Col("travel_min", "Minutes away", 10, "int",
            validate=number_check(0, 300, decimal=False)),
        Col("best_for", "Best for", 24, "text",
            note="'Veg and bread', 'cheap store cupboard', 'the only "
                 "place that does the oat milk'."),
        Col("delivery", "Delivers?", 10, "flag",
            validate=list_from("YesNoList")),
        Col("delivery_fee", "Delivery fee", 11, "money",
            validate=number_check(0, 1000)),
        Col("min_order", "Minimum order", 11, "money",
            validate=number_check(0, 100000)),
        Col("slot", "Usual slot", 16, "text",
            note="'Thursday 6-7pm'. Handy next to the shopping day."),
        Col("loyalty", "Loyalty card", 16, "text"),
        Col("open_hours", "Opening hours", 18, "text"),
        Col("price_level", "Price level", 12, "text",
            validate=list_from("PriceLevelList")),
        Col("rating", "Your rating", 10, "rating",
            validate=number_check(1, 5, decimal=False)),
        Col("items_priced", "Items priced here", 11, "int",
            formula='=IF({name}{r}="","",COUNTIFS('
                    '[[Store Prices.store]],{name}{r}))',
            note="How many lines of the price book you have checked "
                 "at this shop."),
        Col("cheapest_on", "Cheapest on", 11, "int",
            formula='=IF({name}{r}="","",COUNTIFS('
                    '[[Ingredients.best_store]],{name}{r}))',
            note="How many ingredients this shop wins on price."),
        Col("spent", "Spent here", 12, "money",
            formula='=IF({name}{r}="","",SUMIFS('
                    '[[Grocery Spending.amount]],'
                    '[[Grocery Spending.store]],{name}{r}))'),
        Col("visits", "Shops logged", 10, "int",
            formula='=IF({name}{r}="","",COUNTIFS('
                    '[[Grocery Spending.store]],{name}{r}))'),
        Col("avg_shop", "Average shop", 11, "money",
            formula='=IF(OR({visits}{r}="",{visits}{r}=0),"",'
                    '{spent}{r}/{visits}{r})'),
        Col("on_list", "On this week's list", 11, "int",
            formula='=IF({name}{r}="","",COUNTIFS('
                    '[[Grocery List.store]],{name}{r}))'),
        Col("status", "Status", 12, "status",
            validate=list_from("ActiveList"), cf=status_cf()),
        Col("alert", "Check", 22, "status",
            formula='=IF({name}{r}="","",IF({items_priced}{r}=0,'
                    '"Add - no prices logged here yet",'
                    'IF(AND({delivery}{r}="Yes",{min_order}{r}=""),'
                    '"Add - minimum order?","OK")))',
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Grocery Stores",
        subtitle="Where you shop, what each one is good for and what "
                 "each one costs you.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_soft,
        freeze_cols=2,
        samples=sample(cfg).stores,
        help_text="Add a store here and it appears in every store "
                  "drop-down in the workbook.",
    )
