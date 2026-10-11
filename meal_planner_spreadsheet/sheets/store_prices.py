"""`Store Prices` - the same thing, priced in three shops.

The honest version of 'shop around'. Log what a pack costs where, and
the workbook works out the price per unit, the cheapest shop, and
whether the drive to the big supermarket is actually worth it.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, id_col, lookup_ingredient,
                     min_number)

NAME = "Store Prices"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("ingredient", '"SP-"', digits=4, width=8),
        Col("ingredient", "Ingredient", 24, "text",
            validate=list_from("IngredientNames")),
        Col("store", "Store", 18, "text", validate=list_from("StoreNames")),
        Col("brand", "Brand", 16, "text",
            note="Own brand against the big name is usually the "
                 "biggest single saving in a shop."),
        Col("pack_size", "Pack size", 10, "number1",
            validate=number_check(0, 1000000),
            note="In the base unit for that ingredient - g, ml or "
                 "each."),
        Col("unit", "Unit", 8, "text",
            formula=lookup_ingredient("base_unit")),
        Col("price", "Pack price", 11, "money",
            validate=number_check(0, 100000)),
        Col("offer", "On offer?", 10, "flag",
            validate=list_from("YesNoList"),
            note="Offer prices are worth logging but do not plan the "
                 "month around them."),
        Col("unit_cost", "Cost per unit", 11, "money",
            formula='=IF(OR({price}{r}="",{pack_size}{r}="",'
                    '{pack_size}{r}=0),"",{price}{r}/{pack_size}{r})',
            hidden=True),
        Col("per_100", "Cost per 100", 11, "money",
            formula='=IF({unit_cost}{r}="","",{unit_cost}{r}*100)',
            note="The comparable number. This is the column to sort."),
        Col("book_cost", "Price book cost", 11, "money",
            formula=lookup_ingredient("unit_cost"), hidden=True),
        Col("vs_book", "Against the price book", 11, "money",
            formula='=IF(OR({unit_cost}{r}="",{book_cost}{r}="",'
                    'NOT(ISNUMBER({book_cost}{r}))),"",'
                    '({unit_cost}{r}-{book_cost}{r})*100)',
            note="Negative means this shop is cheaper per 100 than the "
                 "price you have in the Ingredients tab.",
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "good"},
                {"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}]),
        Col("best", "Cheapest so far?", 11, "flag",
            formula='=IF(OR({ingredient}{r}="",{per_100}{r}=""),"",'
                    'IF({unit_cost}{r}<=' + min_number(
                        '[[Store Prices.unit_cost]]',
                        '[[Store Prices.ingredient]]={ingredient}{r}')
                    + '+0.0000001,"Yes","No"))',
            cf=[{"type": "text", "criteria": "containing", "value": "Yes",
                 "fmt": "good"}],
            note="Yes means no other shop you have logged beats this "
                 "price per unit."),
        Col("checked", "Price checked", 11, "date", validate=date_check()),
        Col("age_days", "How old", 9, "int",
            formula='=IF({checked}{r}="","",TODAY()-{checked}{r})',
            cf=[{"type": "cell", "criteria": ">", "value": 90,
                 "fmt": "warn"}]),
        Col("alert", "Check", 24, "status",
            formula='=IF({ingredient}{r}="","",'
                    'IF({price}{r}="","Add - no price",'
                    'IF({pack_size}{r}="","Add - no pack size",'
                    'IF(AND({age_days}{r}<>"",{age_days}{r}>180),'
                    '"Check - over six months old","OK"))))',
            cf=alert_cf()),
        Col("notes", "Notes", 24, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Store Price Comparison",
        subtitle="What the same ingredient costs in each shop, reduced "
                 "to one comparable number.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent_soft,
        freeze_cols=3,
        samples=sample(cfg).store_prices,
        help_text="You do not need to price everything. The twenty "
                  "things you buy every week are where the money is.",
        extra_notes=[
            "Cost per 100 is the only fair comparison - it is what "
            "the shelf label shows, and it is how the Ingredients tab "
            "picks the cheapest shop for each item.",
            "Prices go stale. Anything you checked more than six "
            "months ago is flagged so you know not to trust it.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    spec_ = reg[NAME]
    s = sheet
    last = spec_.last_col_index
    row = spec_.last_data_row + 1
    row = s.section(row, "What shopping around is worth", last)
    row = s.table_head(row, 0, ["Prices logged", "Ingredients covered",
                                "Stores compared", "Checked this year",
                                "Total saving if you always bought "
                                "cheapest"],
                       [14, 18, 16, 16, 34])
    row = s.calc_row(row, 0, [
        '=COUNTIF([[Store Prices.ingredient]],"<>")',
        '=SUMPRODUCT((COUNTIF([[Store Prices.ingredient]],'
        '[[Ingredients.name]])>0)*([[Ingredients.name]]<>""))',
        '=SUMPRODUCT((COUNTIF([[Store Prices.store]],'
        '[[Stores.name]])>0)*([[Stores.name]]<>""))',
        '=COUNTIFS([[Store Prices.checked]],">="&DATE(YEAR(TODAY()),1,1))',
        '=SUM([[Ingredients.saving]])',
    ], ["int", "int", "int", "int", "money"])
    row += 1
    s.note(row, "The saving on the right is the sum of every 'cheaper "
                "elsewhere' gap per 100 units. It is a ceiling, not a "
                "promise - petrol, time and the fact that nobody "
                "shops in four places all take a bite out of it.", last)
