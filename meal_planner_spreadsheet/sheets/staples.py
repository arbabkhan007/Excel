"""`Staples` - the things you never want to run out of.

Salt, olive oil, pasta, nappies for the dog's dinner. Set a par level
once and the workbook tells you when stock has dropped below it, how
long it will last at your usual rate, and roughly when to buy more.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import alert_cf, id_col, lookup_ingredient, status_cf

NAME = "Staples"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("ingredient", '"STP-"', digits=3, width=9),
        Col("ingredient", "Staple", 24, "text",
            validate=list_from("IngredientNames")),
        Col("category", "Category", 15, "text",
            formula=lookup_ingredient("category")),
        Col("unit", "Unit", 8, "text",
            formula=lookup_ingredient("base_unit")),
        Col("par", "Always keep", 11, "number1",
            validate=number_check(0, 100000),
            note="Your par level: the amount you want on the shelf at "
                 "all times, in base units."),
        Col("in_stock", "In stock", 11, "number1",
            formula='=IF({ingredient}{r}="","",IFERROR(INDEX('
                    '[[Ingredients.in_pantry]],MATCH({ingredient}{r},'
                    '[[Ingredients.name]],0)),0))'),
        Col("gap", "Short by", 10, "number1",
            formula='=IF(OR({par}{r}="",{in_stock}{r}=""),"",'
                    'MAX(0,{par}{r}-{in_stock}{r}))',
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}]),
        Col("usage", "Used a week", 10, "number1",
            validate=number_check(0, 100000),
            note="Roughly how much the house gets through in a week. "
                 "A guess is fine - it only drives the countdown."),
        Col("weeks_left", "Weeks left", 10, "number1",
            formula='=IF(OR({usage}{r}="",{usage}{r}=0,'
                    '{in_stock}{r}=""),"",{in_stock}{r}/{usage}{r})',
            cf=[{"type": "cell", "criteria": "<", "value": 1,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "between", "minimum": 1,
                 "maximum": 2, "fmt": "warn"}]),
        Col("buy_by", "Buy before", 11, "date",
            formula='=IF({weeks_left}{r}="","",TODAY()+'
                    'ROUND({weeks_left}{r}*7,0))'),
        Col("pack_size", "Pack size", 10, "number1",
            formula=lookup_ingredient("pack_size"), hidden=True),
        Col("pack_price", "Pack price", 10, "money",
            formula=lookup_ingredient("pack_price"), hidden=True),
        Col("packs", "Packs to buy", 10, "number1",
            formula='=IF(OR({gap}{r}="",{gap}{r}=0,{pack_size}{r}="",'
                    'NOT(ISNUMBER({pack_size}{r})),{pack_size}{r}=0),'
                    '"",ROUNDUP({gap}{r}/{pack_size}{r},0))'),
        Col("cost", "Cost to restock", 11, "money",
            formula='=IF(OR({packs}{r}="",{pack_price}{r}="",'
                    'NOT(ISNUMBER({pack_price}{r}))),"",'
                    '{packs}{r}*{pack_price}{r})'),
        Col("store", "Buy it at", 16, "text",
            formula=lookup_ingredient("store")),
        Col("bulk", "Worth buying big?", 11, "flag",
            validate=list_from("YesNoList"),
            note="Only if you have the space and it does not go off - "
                 "bulk buying something you bin is not a saving."),
        Col("status", "Status", 13, "status",
            formula='=IF({ingredient}{r}="","",IF({in_stock}{r}=0,'
                    '"Out of stock",IF(AND({par}{r}<>"",'
                    '{in_stock}{r}<{par}{r}),"Low stock","In stock")))',
            cf=status_cf()),
        Col("alert", "Check", 24, "status",
            formula='=IF({ingredient}{r}="","",'
                    'IF({status}{r}="Out of stock","Missing - add it '
                    'to the list",IF(AND({weeks_left}{r}<>"",'
                    '{weeks_left}{r}<1),"Due - less than a week left",'
                    'IF({status}{r}="Low stock","Check - below your '
                    'par level","OK"))))',
            cf=alert_cf()),
        Col("notes", "Notes", 24, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Pantry Staples",
        subtitle="Par levels for the things that should always be in "
                 "the cupboard, with a countdown to the day they run "
                 "out.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_soft,
        freeze_cols=2,
        samples=sample(cfg).staples,
        help_text="Twenty staples is plenty. Set the par level to what "
                  "makes you comfortable, not what fits in the "
                  "cupboard.",
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    spec_ = reg[NAME]
    s = sheet
    last = spec_.last_col_index
    row = spec_.last_data_row + 1
    row = s.section(row, "Restock list", last)
    row = s.table_head(row, 0, ["Staples tracked", "Below par",
                                "Out of stock", "Running out this week",
                                "Cost to restock everything"],
                       [15, 12, 13, 20, 26])
    row = s.calc_row(row, 0, [
        '=COUNTIF([[Staples.ingredient]],"<>")',
        '=COUNTIF([[Staples.status]],"Low stock")',
        '=COUNTIF([[Staples.status]],"Out of stock")',
        '=COUNTIFS([[Staples.weeks_left]],"<1",'
        '[[Staples.weeks_left]],">=0")',
        '=SUM([[Staples.cost]])',
    ], ["int", "int", "int", "int", "money"])
