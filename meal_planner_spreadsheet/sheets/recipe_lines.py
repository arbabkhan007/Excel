"""`Recipe Ingredients` - one line per ingredient in a recipe.

This is the engine room. Every line says *recipe R-0007 needs 300 g of
chicken breast*; the workbook converts that into the base unit, prices
it from the Ingredients tab, works out the nutrition, multiplies it by
how many times the recipe appears in your meal plan and hands the answer
to the grocery list. If the automatic shopping list ever looks wrong,
this is the tab to check.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import (alert_cf, base_unit_of, id_col, lookup_ingredient,
                     lookup_recipe, to_base)

NAME = "Recipe Ingredients"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("ingredient", '"RL-"', digits=4, width=9),
        Col("recipe", "Recipe ID", 11, "text",
            validate=list_from("RecipeIDs"),
            note="Pick the recipe this line belongs to. The drop-down "
                 "lists every ID on the Recipes tab."),
        Col("recipe_name", "Recipe", 26, "text",
            formula=lookup_recipe("name")),
        Col("ingredient", "Ingredient", 24, "text",
            validate=list_from("IngredientNames"),
            note="Must match the Ingredients tab exactly - choose from "
                 "the drop-down and it always will."),
        Col("qty", "Quantity", 10, "number1",
            validate=number_check(0, 100000),
            note="For the number of servings on the Recipes tab, not "
                 "per person."),
        Col("unit", "Unit", 11, "text", validate=list_from("UnitList"),
            note="Any unit from the table on the Lists tab - kg, g, ml, "
                 "tbsp, cup, each. It is converted for you."),
        Col("prep", "Prepared how", 18, "text",
            note="'finely chopped', 'drained', 'at room temperature'."),
        Col("section", "Part of the recipe", 16, "text",
            note="Optional: 'sauce', 'topping', 'marinade' - keeps long "
                 "recipes readable."),
        Col("optional", "Optional?", 10, "flag",
            validate=list_from("YesNoList"),
            note="Optional ingredients are still costed but are marked "
                 "on the shopping list so you can skip them."),
        # ------------------------------------------------- the conversion
        Col("base_qty", "In base units", 12, "number1",
            formula=to_base("qty", "unit"),
            note="Automatic: the quantity converted to g, ml or each "
                 "using the unit table on the Lists tab."),
        Col("base_unit", "Base unit", 10, "text",
            formula=base_unit_of("unit")),
        Col("price_unit", "Priced in", 10, "text",
            formula=lookup_ingredient("base_unit")),
        Col("unit_cost", "Cost per unit", 12, "money",
            formula=lookup_ingredient("unit_cost"), hidden=True),
        Col("line_cost", "Line cost", 11, "money",
            formula='=IF(OR({base_qty}{r}="",{unit_cost}{r}="",'
                    'NOT(ISNUMBER({unit_cost}{r}))),"",'
                    'IF({base_unit}{r}<>{price_unit}{r},"",'
                    '{base_qty}{r}*{unit_cost}{r}))',
            note="Quantity x the price per unit from the Ingredients "
                 "tab. Blank if the units do not match - fix the unit "
                 "and it fills itself in."),
        Col("cost_serving", "Cost a serving", 12, "money",
            formula='=IF({line_cost}{r}="","",IFERROR({line_cost}{r}/'
                    'INDEX([[Recipes.servings]],MATCH({recipe}{r},'
                    '[[Recipes.id]],0)),""))', hidden=True),
        # -------------------------------------------------- the nutrition
        Col("per100", "Per 100", 9, "number",
            formula='=IF({base_qty}{r}="","",'
                    'IF({base_unit}{r}="each",{base_qty}{r},'
                    '{base_qty}{r}/100))',
            hidden=True,
            note="How many 'per 100 g' blocks this line is worth."),
        Col("line_cals", "Calories", 10, "number",
            formula='=IFERROR(IF({per100}{r}="","",{per100}{r}*'
                    'INDEX([[Ingredients.calories]],MATCH('
                    '{ingredient}{r},[[Ingredients.name]],0))),"")'),
        Col("line_protein", "Protein (g)", 10, "number1",
            formula='=IFERROR(IF({per100}{r}="","",{per100}{r}*'
                    'INDEX([[Ingredients.protein]],MATCH('
                    '{ingredient}{r},[[Ingredients.name]],0))),"")',
            hidden=True),
        Col("line_carbs", "Carbs (g)", 10, "number1",
            formula='=IFERROR(IF({per100}{r}="","",{per100}{r}*'
                    'INDEX([[Ingredients.carbs]],MATCH('
                    '{ingredient}{r},[[Ingredients.name]],0))),"")',
            hidden=True),
        Col("line_fat", "Fat (g)", 9, "number1",
            formula='=IFERROR(IF({per100}{r}="","",{per100}{r}*'
                    'INDEX([[Ingredients.fat]],MATCH('
                    '{ingredient}{r},[[Ingredients.name]],0))),"")',
            hidden=True),
        Col("line_fibre", "Fibre (g)", 10, "number1",
            formula='=IFERROR(IF({per100}{r}="","",{per100}{r}*'
                    'INDEX([[Ingredients.fibre]],MATCH('
                    '{ingredient}{r},[[Ingredients.name]],0))),"")',
            hidden=True),
        Col("allergen", "Allergen", 13, "text",
            formula=lookup_ingredient("allergen"),
            cf=[{"type": "text", "criteria": "not containing",
                 "value": "None", "fmt": "warn"}]),
        Col("category", "Shopping category", 15, "text",
            formula=lookup_ingredient("category"), hidden=True),
        # -------------------------------------- what the meal plan needs
        Col("plan_factor", "Planned batches", 11, "number1",
            formula='=IFERROR(INDEX([[Recipes.plan_scale]],MATCH('
                    '{recipe}{r},[[Recipes.id]],0)),0)',
            hidden=True,
            note="How many recipes-worth of this dish the shopping "
                 "window asks for."),
        Col("need_base", "Needed for the plan", 13, "number1",
            formula='=IF(OR({base_qty}{r}="",{plan_factor}{r}=0),"",'
                    '{base_qty}{r}*{plan_factor}{r})',
            note="Quantity x planned batches. This is the number the "
                 "grocery list adds up."),
        Col("key", "Key", 10, "text",
            formula='=IF({ingredient}{r}="","",{recipe}{r}&"|"&'
                    '{ingredient}{r})', hidden=True),
        Col("alert", "Check", 26, "status",
            formula='=IF({ingredient}{r}="","",'
                    'IF({recipe}{r}="","Add - which recipe?",'
                    'IF(COUNTIFS([[Ingredients.name]],{ingredient}{r})=0,'
                    '"Missing - not in the Ingredients tab",'
                    'IF({qty}{r}="","Add - no quantity",'
                    'IF({unit}{r}="","Add - no unit",'
                    'IF({base_unit}{r}<>{price_unit}{r},'
                    '"Check - recipe in "&{base_unit}{r}&'
                    '", priced in "&{price_unit}{r},'
                    'IF(COUNTIFS({key_first}:{key_last},{key}{r})>1,'
                    '"Duplicate - same ingredient twice","OK")))))))',
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Recipe Ingredients",
        subtitle="One line per ingredient per recipe. This is what "
                 "turns a meal plan into a shopping list.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=4,
        samples=sample(cfg).recipe_lines,
        help_text="Sort or filter by Recipe ID to work on one recipe at "
                  "a time. Everything to the right of the unit column "
                  "fills itself in.",
        extra_notes=[
            "Quantities are for the recipe as written - the serves "
            "figure on the Recipes tab - and the planner scales them "
            "up or down for the number of people you are feeding.",
            "Units are converted through the table on the Lists tab. "
            "Weight and volume are deliberately kept apart: 100 ml of "
            "oil is not 100 g, so price oil in whichever one you cook "
            "with and the line cost appears.",
            "A line with a unit mismatch is flagged in the Check "
            "column and left out of the cost rather than quietly "
            "guessing at it.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Totals strip: how healthy is the engine room?"""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Ingredient line check", last)
    sheet.table_head(row, 0, ["Lines entered", "Recipes covered",
                              "Lines with a cost", "Unit mismatches",
                              "Not in the price book",
                              "Total value of all recipes"],
                     [15, 16, 16, 15, 20, 24])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Recipe Ingredients.ingredient]],"<>")',
        '=SUMPRODUCT((COUNTIF([[Recipe Ingredients.recipe]],'
        '[[Recipes.id]])>0)*([[Recipes.id]]<>""))',
        '=COUNTIFS([[Recipe Ingredients.line_cost]],">0")',
        '=COUNTIF([[Recipe Ingredients.alert]],"Check - *")',
        '=COUNTIF([[Recipe Ingredients.alert]],"Missing*")',
        '=SUM([[Recipe Ingredients.line_cost]])',
    ], ["int", "int", "int", "int", "int", "money"])
