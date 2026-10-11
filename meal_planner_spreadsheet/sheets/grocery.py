"""`Grocery List` - the shopping list that writes itself.

Nothing in the first ten columns is typed. The list looks at every meal
planned between the two dates on the Settings tab, multiplies each
recipe by the number of people you are feeding, adds up the ingredients,
takes off whatever is already in the pantry and tells you what is left
to buy, in whole packs, with a running cost.

You tick things off as you shop, and anything the recipes do not cover -
bin bags, birthday candles - goes in the extras block underneath.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..writer import list_from, number_check
from .common import ROW_OFFSET, lookup_ingredient, status_cf

NAME = "Grocery List"

EXTRA_ROWS = 24


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("rank", "#", 5, "int",
            formula=f"=ROW()-{ROW_OFFSET}", hidden=True),
        Col("ingredient", "Item", 24, "text",
            formula='=IFERROR(INDEX([[Ingredients.name]],MATCH('
                    '{rank}{r},[[Ingredients.need_rank]],0)),"")',
            note="Automatic. Every ingredient your planned meals need, "
                 "in the order they appear in the price book."),
        Col("category", "Category", 15, "text",
            formula=lookup_ingredient("category")),
        Col("aisle", "Aisle", 16, "text",
            formula=lookup_ingredient("aisle"),
            note="Sort by this column to walk the shop once."),
        Col("store", "Store", 16, "text",
            formula=lookup_ingredient("store")),
        Col("needed", "Needed", 10, "number1",
            formula='=IF({ingredient}{r}="","",IFERROR(INDEX('
                    '[[Ingredients.need_now]],MATCH({ingredient}{r},'
                    '[[Ingredients.name]],0)),""))',
            note="Everything the planned meals call for, in base "
                 "units."),
        Col("unit", "Unit", 7, "text",
            formula=lookup_ingredient("base_unit")),
        Col("in_pantry", "Already have", 11, "number1",
            formula='=IF({ingredient}{r}="","",IFERROR(INDEX('
                    '[[Ingredients.in_pantry]],MATCH({ingredient}{r},'
                    '[[Ingredients.name]],0)),0))',
            note="Straight from the Pantry tab - this is the "
                 "deduction that stops you buying a fourth jar of "
                 "cumin."),
        Col("short", "Still short", 10, "number1",
            formula='=IF({needed}{r}="","",MAX(0,{needed}{r}-'
                    'IF({in_pantry}{r}="",0,{in_pantry}{r})))',
            cf=[{"type": "cell", "criteria": "=", "value": 0,
                 "fmt": "muted"}]),
        Col("pack_size", "Pack size", 10, "number1",
            formula=lookup_ingredient("pack_size"), hidden=True),
        Col("pack_price", "Pack price", 10, "money",
            formula=lookup_ingredient("pack_price"), hidden=True),
        Col("unit_cost", "Cost per unit", 10, "money",
            formula=lookup_ingredient("unit_cost"), hidden=True),
        Col("packs", "Packs to buy", 11, "number1",
            formula='=IF(OR({short}{r}="",{short}{r}=0),"",'
                    'IF(OR({pack_size}{r}="",NOT(ISNUMBER('
                    '{pack_size}{r})),{pack_size}{r}=0),"",'
                    'ROUNDUP({short}{r}/{pack_size}{r},0)))',
            note="How many of the pack you normally buy. 1.2 tins "
                 "becomes 2 tins, because shops do not sell 0.2 of a "
                 "tin."),
        Col("buy_qty", "Buy this much", 12, "number1",
            formula='=IF({short}{r}="","",IF(RoundingRule='
                    '"Exact quantity",{short}{r},IF({packs}{r}="",'
                    '{short}{r},{packs}{r}*{pack_size}{r})))',
            note="Follows the rounding rule on the Settings tab."),
        Col("est_cost", "Estimated cost", 12, "money",
            formula='=IF({buy_qty}{r}="","",IF(RoundingRule='
                    '"Exact quantity",IF({unit_cost}{r}="","",'
                    '{buy_qty}{r}*{unit_cost}{r}),'
                    'IF(OR({packs}{r}="",{pack_price}{r}=""),"",'
                    '{packs}{r}*{pack_price}{r})))',
            cf=[{"type": "data_bar", "bar_color": "#4E8A63"}]),
        Col("running", "Running total", 12, "money",
            formula='=IF({ingredient}{r}="","",SUM({est_cost}${first}:'
                    '{est_cost}{r}))',
            note="Watch this against your weekly budget as you plan."),
        Col("to_buy", "To buy", 10, "number1",
            formula='=IF({got}{r}="Yes",0,IF({buy_qty}{r}="",0,'
                    '{buy_qty}{r}))', hidden=True),
        Col("recipes", "For how many meals", 11, "int",
            formula='=IF({ingredient}{r}="","",COUNTIFS('
                    '[[Recipe Ingredients.ingredient]],{ingredient}{r},'
                    '[[Recipe Ingredients.need_base]],">0"))',
            note="How many of the week's recipes want it."),
        Col("optional", "Optional only?", 11, "flag",
            formula='=IF({ingredient}{r}="","",IF(COUNTIFS('
                    '[[Recipe Ingredients.ingredient]],{ingredient}{r},'
                    '[[Recipe Ingredients.need_base]],">0",'
                    '[[Recipe Ingredients.optional]],"No")=0,"Yes",'
                    '"No"))',
            note="Yes means every recipe that wants it treats it as "
                 "optional - skip it if money is tight.",
            cf=[{"type": "text", "criteria": "containing", "value": "Yes",
                 "fmt": "muted"}]),
        Col("allergen", "Allergen", 12, "text",
            formula=lookup_ingredient("allergen"), hidden=True),
        # ----------------------------------------------- what you do
        Col("got", "Got it?", 9, "flag", validate=list_from("YesNoList"),
            note="Tick as you shop. Ticked rows drop out of the "
                 "remaining total below."),
        Col("paid", "Actually paid", 11, "money",
            validate=number_check(0, 100000),
            note="Optional - fill it in at the till and the workbook "
                 "shows you the gap between estimate and reality."),
        Col("swap", "Swapped for", 16, "text",
            note="What you bought instead when the shelf was empty."),
        Col("note", "Note", 20, "wrap"),
        Col("status", "Status", 12, "status",
            formula='=IF({ingredient}{r}="","",IF({got}{r}="Yes",'
                    '"Bought",IF({short}{r}=0,"In stock",'
                    'IF({est_cost}{r}="","Check - no price",'
                    '"On the list"))))',
            cf=status_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Automatic Grocery List",
        subtitle="Built from your meal plan, minus what is already in "
                 "the pantry, grouped by aisle and priced as you go.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=[],
        help_text="Nothing to type on the left hand side - change the "
                  "shopping window on the Settings tab and the whole "
                  "list rebuilds. Tick 'Got it?' as you shop.",
        extra_notes=[
            "The list covers every meal planned between 'Shopping list "
            "covers from' and 'covers to' on the Settings tab, where "
            "'Shop for it?' on the Meal Plan says Yes.",
            "Quantities are in base units - g, ml or each. 'Buy this "
            "much' rounds them up to whole packs unless you change "
            "the rounding rule on the Settings tab.",
            "Ticking 'Got it?' does not delete the row - it keeps the "
            "record of the shop and takes the item out of the "
            "remaining total.",
            "An ingredient only appears here if it is on the "
            "Ingredients tab. Anything flagged 'Missing' on the Recipe "
            "Ingredients tab is quietly left out, so clear those first.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Totals, the budget check, the store split and the extras block."""
    spec_ = reg[NAME]
    s = sheet
    ws = s.ws
    last = spec_.last_col_index
    row = spec_.last_data_row + 1          # 0 based, one row under the grid

    # ------------------------------------------------------------ totals --
    row = s.section(row, "This shop", last)
    row = s.table_head(row, 0, ["Items on the list", "Already in stock",
                                "Still to buy", "Estimated cost",
                                "Picked up so far", "Left to pick up",
                                "Actually paid"],
                       [16, 15, 13, 15, 16, 15, 14])
    row = s.calc_row(row, 0, [
        '=COUNTIF([[Grocery List.ingredient]],"<>")',
        '=COUNTIFS([[Grocery List.short]],0,'
        '[[Grocery List.ingredient]],"<>")',
        '=COUNTIFS([[Grocery List.short]],">0")',
        '=SUM([[Grocery List.est_cost]])',
        '=COUNTIF([[Grocery List.got]],"Yes")',
        '=SUMIFS([[Grocery List.est_cost]],[[Grocery List.got]],"<>Yes")',
        '=SUM([[Grocery List.paid]])',
    ], ["int", "int", "int", "money", "int", "money", "money"])
    row += 1

    # ------------------------------------------------------ budget check --
    row = s.note(row, "The estimate is measured against the weekly "
                      "grocery budget on the Settings tab. Change the "
                      "shopping window there too - this list follows it.",
                 last)
    row = s.table_head(row, 0, ["Weekly budget", "This shop", "Difference",
                                "Shopping window",
                                "Days to the shopping day"],
                       [15, 14, 14, 28, 20])
    diff_row = row
    row = s.calc_row(row, 0, [
        "=WeeklyBudget",
        '=SUM([[Grocery List.est_cost]])',
        '=WeeklyBudget-SUM([[Grocery List.est_cost]])',
        '=TEXT(ShopFrom,"d mmm")&" to "&TEXT(ShopTo,"d mmm yyyy")',
        '=MAX(0,MOD(IFERROR(MATCH(ShoppingDay,WeekdayList,0),6)-'
        'WEEKDAY(TODAY(),2),7))',
    ], ["money", "money", "money", "text", "int"])
    ws.conditional_format(diff_row, 2, diff_row, 2,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(diff_row, 2, diff_row, 2,
                          {"type": "cell", "criteria": ">=", "value": 0,
                           "format": fmts.cf("good")})
    row += 1

    # --------------------------------------------------- store by store --
    row = s.section(row, "Store by store", last)
    row = s.table_head(row, 0, ["Store", "Items", "Estimated cost",
                                "Still to pick up", "Cheaper elsewhere"],
                       [22, 10, 15, 16, 18])
    for i in range(6):
        cell = f"A{row + 1}"
        s.write(row, 0, f'=IFERROR(IF(INDEX(StoreNames,{i + 1})=0,"",'
                        f'INDEX(StoreNames,{i + 1})),"")',
                fmts.calc("text"))
        s.calc_row(row, 1, [
            f'=IF({cell}="","",COUNTIFS([[Grocery List.store]],{cell}))',
            f'=IF({cell}="","",SUMIFS([[Grocery List.est_cost]],'
            f'[[Grocery List.store]],{cell}))',
            f'=IF({cell}="","",SUMIFS([[Grocery List.est_cost]],'
            f'[[Grocery List.store]],{cell},[[Grocery List.got]],'
            f'"<>Yes"))',
            f'=IF({cell}="","",SUMIFS([[Ingredients.saving]],'
            f'[[Ingredients.store]],{cell}))',
        ], ["int", "money", "money", "money"])
        row += 1
    row += 1

    # ----------------------------------------------------- extras block --
    row = s.section(row, "Extras - things no recipe asked for", last)
    row = s.note(row, "Bin bags, washing up liquid, a birthday cake. "
                      "These rows are yours to type in and they are "
                      "added to the grand total at the bottom.", last)
    row = s.table_head(row, 0, ["Item", "Category", "How many", "Unit",
                                "Estimated cost", "Store", "Got it?",
                                "Note"],
                       [24, 15, 10, 8, 14, 16, 9, 26])
    first_extra = row
    styles = ["text", "text", "number1", "text", "money", "text", "flag",
              "wrap"]
    for i in range(EXTRA_ROWS):
        r = first_extra + i
        for col_i, style in enumerate(styles):
            s.write(r, col_i, None, fmts.input(style))
        ws.data_validation(r, 1, r, 1, list_from("GroceryCategoryList"))
        ws.data_validation(r, 3, r, 3, list_from("UnitList"))
        ws.data_validation(r, 5, r, 5, list_from("StoreNames"))
        ws.data_validation(r, 6, r, 6, list_from("YesNoList"))
    last_extra = first_extra + EXTRA_ROWS - 1
    row = last_extra + 1
    s.merge(row, 0, 3, "Extras total", fmts.chrome("total_label"))
    extras_sum = f"=SUM(E{first_extra + 1}:E{last_extra + 1})"
    ws.write_formula(row, 4, extras_sum,
                     fmts.chrome("total_value",
                                 num_format=cfg.money_format))
    s.merge(row, 5, min(last, 9),
            "Typed in by you - no recipe needed.",
            fmts.chrome("total_value", align="left"))
    row += 2

    # ------------------------------------------------------ grand total --
    s.merge(row, 0, 3, "Everything in the trolley",
            fmts.chrome("total_label"))
    ws.write_formula(
        row, 4,
        s.reg.render('=SUM([[Grocery List.est_cost]])') +
        f"+SUM(E{first_extra + 1}:E{last_extra + 1})",
        fmts.chrome("total_value", num_format=cfg.money_format))
    s.merge(row, 5, min(last, 9),
            "Recipes plus extras. Compare it with the weekly "
            "budget above.",
            fmts.chrome("total_value", align="left"))
