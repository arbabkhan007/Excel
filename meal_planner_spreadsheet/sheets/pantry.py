"""`Pantry` - what is actually in the house.

Every jar, bag and tin you have, with how much is left and when it goes
off. The grocery list subtracts this from what your meals need, so you
stop buying a fourth bag of rice; the expiry columns tell you what to
cook first.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, base_unit_of, id_col, lookup_ingredient,
                     status_cf, to_base)

NAME = "Pantry"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("ingredient", "PantryPrefix", digits=4, width=10),
        Col("ingredient", "Ingredient", 24, "text",
            validate=list_from("IngredientNames"),
            note="Choose from the Ingredients tab so the grocery list "
                 "can match it. Add anything new there first."),
        Col("category", "Category", 15, "text",
            formula=lookup_ingredient("category")),
        Col("location", "Where it lives", 14, "text",
            validate=list_from("StorageList"),
            note="Cupboard, Fridge, Freezer, Spice rack... The fridge "
                 "and freezer tab filters on this."),
        Col("qty", "How much is left", 11, "number1",
            validate=number_check(0, 100000),
            note="Update it when you notice, not every meal. Even a "
                 "rough number beats no number."),
        Col("unit", "Unit", 10, "text", validate=list_from("UnitList")),
        Col("base_qty", "In base units", 11, "number1",
            formula=to_base("qty", "unit"),
            note="What the grocery list deducts."),
        Col("base_unit", "Base unit", 9, "text",
            formula=base_unit_of("unit"), hidden=True),
        Col("opened", "Opened?", 9, "flag", validate=list_from("YesNoList")),
        Col("bought", "Bought on", 11, "date", validate=date_check()),
        Col("expiry", "Use by", 11, "date", validate=date_check(),
            note="Leave it blank and the workbook estimates one from "
                 "the shelf life on the Ingredients tab."),
        Col("shelf_days", "Keeps for", 9, "int",
            formula=lookup_ingredient("shelf_days"), hidden=True),
        Col("est_expiry", "Best before (worked out)", 12, "date",
            formula='=IF({expiry}{r}<>"",{expiry}{r},'
                    'IF(OR({bought}{r}="",{shelf_days}{r}="",'
                    'NOT(ISNUMBER({shelf_days}{r}))),"",'
                    '{bought}{r}+{shelf_days}{r}))',
            note="Your date if you typed one, otherwise bought date + "
                 "shelf life."),
        Col("days_left", "Days left", 9, "int",
            formula='=IF({est_expiry}{r}="","",{est_expiry}{r}-TODAY())',
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "between", "minimum": 0,
                 "maximum": "=ExpiryWarningDays", "fmt": "warn"},
                {"type": "cell", "criteria": ">",
                 "value": "=ExpiryWarningDays", "fmt": "good"}]),
        Col("unit_cost", "Cost per unit", 10, "money",
            formula=lookup_ingredient("unit_cost"), hidden=True),
        Col("value", "Value", 11, "money",
            formula='=IF(OR({base_qty}{r}="",{unit_cost}{r}="",'
                    'NOT(ISNUMBER({unit_cost}{r}))),"",'
                    '{base_qty}{r}*{unit_cost}{r})',
            note="What is sitting in your cupboards, at price book "
                 "prices."),
        Col("min_qty", "Keep at least", 10, "number1",
            validate=number_check(0, 100000),
            note="Your re-order level. Below this it is flagged as low "
                 "stock - the staples list uses it."),
        Col("needed", "Plan needs", 11, "number1",
            formula='=IF({ingredient}{r}="","",IFERROR(INDEX('
                    '[[Ingredients.need_now]],MATCH({ingredient}{r},'
                    '[[Ingredients.name]],0)),0))', hidden=True),
        Col("after_plan", "Left after the plan", 12, "number1",
            formula='=IF({base_qty}{r}="","",{base_qty}{r}-'
                    'IF({needed}{r}="",0,{needed}{r}))',
            note="What should be left once you have cooked this "
                 "week's meals.",
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "warn"}]),
        Col("staple", "Staple?", 9, "flag", validate=list_from("YesNoList"),
            note="Things you never want to run out of."),
        Col("status", "Status", 14, "status",
            formula='=IF({ingredient}{r}="","",'
                    'IF({base_qty}{r}=0,"Out of stock",'
                    'IF(AND({days_left}{r}<>"",{days_left}{r}<0),'
                    '"Expired",IF(AND({days_left}{r}<>"",'
                    '{days_left}{r}<=ExpiryWarningDays),"Use soon",'
                    'IF(AND({min_qty}{r}<>"",{base_qty}{r}<'
                    '{min_qty}{r}),"Low stock","In stock")))))',
            cf=status_cf()),
        Col("alert", "What to do", 28, "status",
            formula='=IF({ingredient}{r}="","",'
                    'IF({status}{r}="Expired","Expired - bin it and log '
                    'it on the Food Waste tab",'
                    'IF({status}{r}="Use soon","Use soon - plan a meal '
                    'with it",IF({status}{r}="Out of stock",'
                    'IF({staple}{r}="Yes","Missing - staple has run '
                    'out","Add - restock when convenient"),'
                    'IF({status}{r}="Low stock","Check - running low",'
                    'IF({after_plan}{r}<0,"Check - the plan needs more '
                    'than you have","OK"))))))',
            cf=alert_cf()),
        Col("notes", "Notes", 24, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Pantry Inventory",
        subtitle="What you already own, what it is worth and what is "
                 "about to go off - the tab that shrinks your shopping "
                 "list.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).pantry,
        help_text="One row per thing you have in stock. You do not "
                  "have to be perfect - even counting the ten things "
                  "you buy most often saves real money.",
        extra_notes=[
            "The grocery list takes 'In base units' off whatever your "
            "meals need. An empty pantry simply means nothing is "
            "deducted.",
            "Leave 'Use by' blank and the workbook works one out from "
            "the shelf life on the Ingredients tab - a sensible guess, "
            "not a guarantee. Trust the label and your nose first.",
            "'Keep at least' turns this into a re-order list: anything "
            "below it is flagged as low stock on the dashboard.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Stock value and what needs attention."""
    spec_ = reg[NAME]
    s = sheet
    last = spec_.last_col_index
    row = spec_.last_data_row + 1
    row = s.section(row, "Pantry at a glance", last)
    row = s.table_head(row, 0, ["Items in stock", "Value in stock",
                                "Expired", "Use within "
                                f"{cfg.expiry_warning_days} days",
                                "Low stock", "Out of stock",
                                "Staples below par"],
                       [15, 15, 11, 18, 12, 13, 18])
    row = s.calc_row(row, 0, [
        '=COUNTIF([[Pantry.ingredient]],"<>")',
        '=SUM([[Pantry.value]])',
        '=COUNTIF([[Pantry.status]],"Expired")',
        '=COUNTIF([[Pantry.status]],"Use soon")',
        '=COUNTIF([[Pantry.status]],"Low stock")',
        '=COUNTIF([[Pantry.status]],"Out of stock")',
        '=COUNTIFS([[Pantry.staple]],"Yes",[[Pantry.status]],'
        '"<>In stock")',
    ], ["int", "money", "int", "int", "int", "int", "int"])
    row += 1
    s.note(row, "Tip: filter the Status column on 'Use soon' before you "
                "plan next week. Cooking what you already have is the "
                "single biggest saving in this workbook.", last)
