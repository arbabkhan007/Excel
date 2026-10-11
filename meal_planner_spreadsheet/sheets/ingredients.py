"""`Ingredients` - the price book every cost and shopping line reads.

One row per ingredient you buy. It holds the category and aisle (so the
grocery list can group itself), the pack you normally buy, what that pack
costs and the nutrition per 100 g or 100 ml. Everything else in the
workbook - recipe costs, the shopping list, the pantry value, the
nutrition tracker - looks its numbers up here, so a price only ever gets
typed in one place.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, id_col, min_number

NAME = "Ingredients"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "IngredientPrefix", digits=4, width=10),
        Col("name", "Ingredient", 24, "text",
            note="Use one spelling and stick to it - 'chicken breast', "
                 "not 'chicken breasts' as well. The grocery list adds "
                 "up by this name."),
        Col("category", "Category", 16, "text",
            validate=list_from("GroceryCategoryList"),
            note="Groups the shopping list."),
        Col("aisle", "Aisle", 17, "text", validate=list_from("AisleList"),
            note="Sorts the printed list into the order you walk round "
                 "the shop."),
        Col("store", "Usual store", 18, "text",
            validate=list_from("StoreNames")),
        Col("brand", "Preferred brand", 16, "text"),
        Col("base_unit", "Measured in", 12, "text",
            validate=list_from("UnitList"),
            note="The unit you cook with: g, ml or each. Recipe "
                 "quantities in any other unit are converted to this."),
        Col("pack_size", "Pack size", 11, "number1",
            validate=number_check(0, 1000000),
            note="How much is in the pack you normally buy, in the unit "
                 "above. A 400 g tin = 400."),
        Col("pack_price", "Pack price", 11, "money",
            validate=number_check(0, 100000)),
        Col("unit_cost", "Cost per unit", 12, "money",
            formula='=IF(OR({pack_size}{r}="",{pack_price}{r}=""),"",'
                    'IF({pack_size}{r}=0,"",{pack_price}{r}/'
                    '{pack_size}{r}))',
            note="Automatic. This is what every recipe cost is built "
                 "from."),
        Col("per_100", "Cost per 100", 11, "money",
            formula='=IF({unit_cost}{r}="","",{unit_cost}{r}*100)',
            note="Cost per 100 g or 100 ml - the number to compare "
                 "between brands and pack sizes."),
        Col("avg_paid", "Average price paid", 13, "money",
            formula='=IF({name}{r}="","",IFERROR(AVERAGEIFS('
                    '[[Grocery Spending.unit_price]],'
                    '[[Grocery Spending.item]],{name}{r},'
                    '[[Grocery Spending.unit_price]],">0"),""))',
            note="What you have actually been paying, from the purchase "
                 "log."),
        Col("best_price", "Cheapest store price", 14, "money",
            formula='=IF({name}{r}="","",IF(COUNTIFS('
                    '[[Store Prices.ingredient]],{name}{r})=0,"",'
                    'IF(' + min_number('[[Store Prices.unit_cost]]',
                                       '[[Store Prices.ingredient]]='
                                       '{name}{r}') + '>=999999,"",'
                    + min_number('[[Store Prices.unit_cost]]',
                                 '[[Store Prices.ingredient]]='
                                 '{name}{r}') + ')))',
            note="The lowest price per unit any store on the Store "
                 "Prices tab charges for it."),
        Col("best_store", "Cheapest at", 16, "text",
            formula='=IF({best_price}{r}="","",IFERROR(INDEX('
                    '[[Store Prices.store]],MATCH(1,INDEX(('
                    '[[Store Prices.ingredient]]={name}{r})*'
                    '([[Store Prices.unit_cost]]={best_price}{r}),0),0)),'
                    '""))'),
        Col("saving", "Saving per 100", 12, "money",
            formula='=IF(OR({best_price}{r}="",{unit_cost}{r}=""),"",'
                    'MAX(0,({unit_cost}{r}-{best_price}{r})*100))',
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "good"}],
            note="What you would save per 100 g / 100 ml by buying it "
                 "where it is cheapest."),
        Col("calories", "Calories per 100 / item", 12, "number",
            validate=number_check(0, 2000),
            note="Per 100 g or 100 ml from the label - or per item for "
                 "anything measured in 'each', like an egg."),
        Col("protein", "Protein per 100 (g)", 11, "number1",
            validate=number_check(0, 100)),
        Col("carbs", "Carbs per 100 (g)", 11, "number1",
            validate=number_check(0, 100)),
        Col("fat", "Fat per 100 (g)", 11, "number1",
            validate=number_check(0, 100)),
        Col("fibre", "Fibre per 100 (g)", 11, "number1",
            validate=number_check(0, 100)),
        Col("sugar", "Sugar per 100 (g)", 11, "number1",
            validate=number_check(0, 100)),
        Col("sodium", "Sodium per 100 (mg)", 12, "number",
            validate=number_check(0, 100000)),
        Col("allergen", "Allergen", 14, "text",
            validate=list_from("AllergenList"),
            note="Flags every recipe that uses this ingredient."),
        Col("diet", "Suits", 15, "text", validate=list_from("DietList"),
            note="The strictest diet this ingredient fits."),
        Col("season", "Best season", 12, "text",
            validate=list_from("SeasonList")),
        Col("shelf_days", "Keeps for (days)", 12, "int",
            validate=number_check(0, 3650),
            note="Typical shelf life once you get it home. Used to "
                 "suggest an expiry date when you add it to the pantry."),
        Col("storage", "Store it in", 13, "text",
            validate=list_from("StorageList")),
        Col("substitute", "Swap it for", 20, "text",
            note="What you use when the shop has run out."),
        Col("last_bought", "Last bought", 12, "date", validate=date_check(),
            note="Updated by hand, or left to the purchase log."),
        Col("in_pantry", "In the pantry now", 13, "number1",
            formula='=IF({name}{r}="","",SUMIFS([[Pantry.base_qty]],'
                    '[[Pantry.ingredient]],{name}{r}))',
            note="Automatic, in the unit above. Comes from the Pantry "
                 "tab."),
        Col("pantry_value", "Value in stock", 12, "money",
            formula='=IF(OR({in_pantry}{r}="",{unit_cost}{r}=""),"",'
                    '{in_pantry}{r}*{unit_cost}{r})'),
        Col("used_in", "Used in recipes", 12, "int",
            formula='=IF({name}{r}="","",COUNTIFS('
                    '[[Recipe Ingredients.ingredient]],{name}{r}))',
            note="How many of your recipes call for it."),
        Col("need_now", "Needed for the plan", 13, "number1",
            formula='=IF({name}{r}="","",SUMIFS('
                    '[[Recipe Ingredients.need_base]],'
                    '[[Recipe Ingredients.ingredient]],{name}{r}))',
            note="Everything your planned meals call for in the "
                 "shopping window, in base units."),
        Col("spend_key", "Spend key", 9, "number",
            formula='=IF({name}{r}="",0,IFERROR({per_100}{r}*1,0)*'
                    'IFERROR({used_in}{r}*1,0))',
            hidden=True,
            note="Cost per 100 times the number of recipes that use "
                 "it - a rough 'where the money goes' score."),
        Col("spend_rank", "Spend rank", 9, "int",
            formula='=IF({name}{r}="","",COUNTIFS({spend_key_first}:'
                    '{spend_key_last},">"&{spend_key}{r})+'
                    'COUNTIFS({spend_key_first}:{spend_key}{r},'
                    '{spend_key}{r}))',
            hidden=True, note="1 = the ingredient you spend most on."),
        Col("waste_count", "Times binned", 9, "int",
            formula='=IF({name}{r}="",0,COUNTIFS([[Food Waste.item]],'
                    '{name}{r}))',
            hidden=True),
        Col("waste_rank", "Waste rank", 9, "int",
            formula='=IF({name}{r}="","",COUNTIFS({waste_count_first}:'
                    '{waste_count_last},">"&{waste_count}{r})+'
                    'COUNTIFS({waste_count_first}:{waste_count}{r},'
                    '{waste_count}{r}))',
            hidden=True, note="1 = binned most often."),
        Col("need_rank", "Rank", 7, "int",
            formula='=IF({name}{r}="","",IF({need_now}{r}>0,'
                    'COUNTIF({need_now}${first}:{need_now}{r},">0"),""))',
            hidden=True,
            note="Position on the automatic grocery list."),
        Col("to_buy", "On the list now", 12, "number1",
            formula='=IF({name}{r}="","",SUMIFS([[Grocery List.to_buy]],'
                    '[[Grocery List.ingredient]],{name}{r}))'),
        Col("alert", "Check", 24, "status",
            formula=blank(
                "name",
                'IF({pack_price}{r}="","Add - no price yet",'
                'IF({pack_size}{r}="","Add - no pack size",'
                'IF(COUNTIFS({name_first}:{name_last},{name}{r})>1,'
                '"Duplicate - same name twice",'
                'IF({used_in}{r}=0,"Check - no recipe uses it","OK"))))'),
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Ingredients & Prices",
        subtitle="Your price book. Every recipe cost, every shopping "
                 "total and every nutrition figure in the workbook is "
                 "built from this one list.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).ingredients,
        help_text="Fill in the pack size and pack price for the twenty "
                  "or thirty things you buy every week and the rest of "
                  "the workbook starts doing the maths for you.",
        extra_notes=[
            "Pack size is in the unit you cook with: a 1 kg bag of "
            "flour measured in g is 1000, a 6 pack of eggs measured in "
            "'each' is 6.",
            "Cost per 100 is the honest comparison between a big bag "
            "and a small one - the supermarket shelf label does the "
            "same sum.",
            "Nutrition is per 100 g or 100 ml so it scales with any "
            "recipe quantity. Copy it off the label; leave it blank if "
            "you are not tracking.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Summary strip under the price book."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1        # 1 based Excel row
    last = spec_.last_col_index
    row = sheet.section(row, "Price book summary", last)
    sheet.table_head(row, 0, ["Ingredients listed", "With a price",
                              "Missing a price", "Value in the pantry",
                              "Average cost per 100",
                              "Could be cheaper elsewhere"],
                     [18, 14, 15, 18, 19, 24])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Ingredients.name]],"<>")',
        '=COUNTIFS([[Ingredients.name]],"<>",'
        '[[Ingredients.pack_price]],"<>")',
        '=COUNTIFS([[Ingredients.name]],"<>",'
        '[[Ingredients.pack_price]],"")',
        '=SUM([[Ingredients.pantry_value]])',
        '=IFERROR(AVERAGE([[Ingredients.per_100]]),"")',
        '=COUNTIFS([[Ingredients.saving]],">0")',
    ], ["int", "int", "int", "money", "money", "int"])
