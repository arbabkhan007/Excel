"""`Meal Plan` - every planned meal, one row at a time.

The Weekly Planner and Monthly Planner tabs are pretty views of this
list. The grocery list, the budget, the nutrition tracker and the
rotation report all read it. Add a row here - date, meal, recipe, how
many you are feeding - and the rest of the workbook reacts.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, lookup_recipe, id_col, month_col, status_cf,
                     week_col, weekday_col)

NAME = "Meal Plan"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "MealPrefix", digits=4, width=10),
        Col("date", "Date", 11, "date", validate=date_check(),
            note="One row per meal. Use the Weekly Planner tab to see "
                 "it laid out as a week."),
        weekday_col("date"),
        week_col("date"),
        month_col("date"),
        Col("slot", "Meal", 12, "text", validate=list_from("MealSlotList"),
            note="Breakfast, Lunch, Dinner, Snack... set your own on "
                 "the Settings tab."),
        Col("recipe", "Recipe ID", 11, "text",
            validate=list_from("RecipeIDs"),
            note="Leave blank for a meal you are not cooking from a "
                 "recipe - eating out, leftovers, cereal."),
        Col("recipe_name", "What you are eating", 28, "text",
            formula='=IF({recipe}{r}<>"",IFERROR(INDEX('
                    '[[Recipes.name]],MATCH({recipe}{r},'
                    '[[Recipes.id]],0)),"Not found"),{freetext}{r})',
            note="Filled in from the recipe, or from whatever you typed "
                 "in the next column."),
        Col("freetext", "Or type it in", 22, "text",
            note="'Leftover chilli', 'Fish and chips', 'Mum cooking'. "
                 "No recipe needed."),
        Col("servings", "People to feed", 11, "int",
            validate=number_check(0, 100, decimal=False),
            note="Defaults to your household size if you leave it "
                 "blank - set that on the Settings tab."),
        Col("eff_servings", "Servings used", 11, "number1",
            formula='=IF({date}{r}="","",IF({servings}{r}="",'
                    'DefaultServings,{servings}{r}))', hidden=True),
        Col("scale", "Recipe batches", 11, "number1",
            formula='=IF({recipe}{r}="",0,IFERROR({eff_servings}{r}/'
                    'INDEX([[Recipes.servings]],MATCH({recipe}{r},'
                    '[[Recipes.id]],0)),0))',
            note="People to feed divided by what the recipe serves. "
                 "1.5 means one and a half times the quantities."),
        Col("shop", "Shop for it?", 11, "flag",
            validate=list_from("YesNoList"),
            note="Yes puts its ingredients on the grocery list. Set it "
                 "to No for a meal you have already shopped for, or "
                 "one you are eating out."),
        Col("status", "Status", 13, "status",
            validate=list_from("MealStatusList"), cf=status_cf()),
        Col("cook", "Who is cooking", 15, "text",
            validate=list_from("MemberNames")),
        Col("eating", "Who is eating", 18, "text",
            note="Leave blank for everybody, or name the ones who are "
                 "in."),
        Col("leftovers", "Planned leftovers", 11, "number1",
            validate=number_check(0, 100),
            note="Extra servings you mean to make. They show up on the "
                 "Leftovers tab."),
        Col("prep_ahead", "Prep ahead?", 11, "flag",
            validate=list_from("YesNoList"),
            note="Flags it on the Meal Prep tab."),
        Col("event", "Occasion", 16, "text",
            validate=list_from("EventNames"),
            note="Links the meal to anything on the Special Events tab."),
        Col("cost", "Planned cost", 12, "money",
            formula='=IF({recipe}{r}="","",IFERROR(INDEX('
                    '[[Recipes.cost_serving]],MATCH({recipe}{r},'
                    '[[Recipes.id]],0))*{eff_servings}{r},""))',
            note="Cost per serving from the recipe x the number of "
                 "people. What it should cost, not what you spent."),
        Col("actual_cost", "Actual spend", 12, "money",
            validate=number_check(0, 100000),
            note="Only for meals out or takeaways - the weekly shop "
                 "goes on the Grocery Spending tab."),
        Col("calories", "Calories each", 11, "number",
            formula=lookup_recipe("calories")),
        Col("protein", "Protein each (g)", 11, "number1",
            formula=lookup_recipe("protein"), hidden=True),
        Col("carbs", "Carbs each (g)", 11, "number1",
            formula=lookup_recipe("carbs"), hidden=True),
        Col("fat", "Fat each (g)", 10, "number1",
            formula=lookup_recipe("fat"), hidden=True),
        Col("fibre", "Fibre each (g)", 10, "number1",
            formula=lookup_recipe("fibre"), hidden=True),
        Col("cals_total", "Calories on the table", 12, "number",
            formula='=IF(OR({calories}{r}="",NOT(ISNUMBER('
                    '{calories}{r}))),"",{calories}{r}*'
                    '{eff_servings}{r})', hidden=True),
        Col("minutes", "Minutes in the kitchen", 11, "int",
            formula=lookup_recipe("total_min"), hidden=True),
        Col("category", "Category", 13, "text",
            formula=lookup_recipe("category"), hidden=True),
        Col("home_cooked", "Cooked at home?", 11, "flag",
            formula='=IF({date}{r}="","",IF({recipe}{r}<>"","Yes",'
                    'IF(OR({freetext}{r}="",ISNUMBER(SEARCH("takeaway",'
                    '{freetext}{r})),ISNUMBER(SEARCH("out",'
                    '{freetext}{r}))),"No","Yes")))', hidden=True),
        Col("key", "Key", 12, "text",
            formula='=IF({date}{r}="","",TEXT({date}{r},"yyyy-mm-dd")&'
                    '"|"&{slot}{r})', hidden=True,
            note="Date and meal together - how the weekly and monthly "
                 "planners find this row."),
        Col("alert", "Check", 24, "status",
            formula='=IF({date}{r}="","",'
                    'IF(AND({recipe}{r}="",{freetext}{r}=""),'
                    '"Add - nothing planned yet",'
                    'IF(COUNTIFS({key_first}:{key_last},{key}{r})>1,'
                    '"Duplicate - two meals in one slot",'
                    'IF(AND({recipe}{r}<>"",{recipe_name}{r}='
                    '"Not found"),"Missing - recipe ID not on the '
                    'Recipes tab",IF(AND({date}{r}<TODAY(),'
                    'OR({status}{r}="Planned",{status}{r}="")),'
                    '"Check - did you cook it?","OK")))))',
            cf=alert_cf()),
        Col("notes", "Notes", 24, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Meal Plan",
        subtitle="Every planned meal in one list. The weekly and "
                 "monthly planners, the grocery list and the budget all "
                 "read from here.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).meal_plan,
        help_text="You can plan straight onto this tab, or use the "
                  "Weekly Planner and let it write the rows for you.",
        extra_notes=[
            "'Shop for it?' is the switch between a plan and a "
            "shopping list. Set it to No for anything you have already "
            "bought, and the ingredients drop off the list.",
            "'Recipe batches' is how the scaling works: feeding 6 from "
            "a recipe that serves 4 gives 1.5, and every ingredient is "
            "multiplied by 1.5 on the grocery list.",
            "A meal with no recipe is still worth a row - the budget "
            "and the eating-out counter both use it.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Totals strip under the plan."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Plan summary", last)
    sheet.table_head(row, 0, ["Meals planned", "From a recipe",
                              "Cooked at home", "Planned cost",
                              "Eaten out / takeaway", "Still to cook"],
                     [14, 14, 14, 16, 19, 14])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Meal Plan.date]],">0")',
        '=COUNTIF([[Meal Plan.recipe]],"<>")',
        '=COUNTIF([[Meal Plan.home_cooked]],"Yes")',
        '=SUM([[Meal Plan.cost]])',
        '=COUNTIF([[Meal Plan.home_cooked]],"No")',
        '=COUNTIFS([[Meal Plan.status]],"Planned")',
    ], ["int", "int", "int", "money", "int", "int"])
