"""`Meal Prep` - the Sunday afternoon plan.

Batch cooking only pays off if you know what to make, how many portions
it gives you and where they are going. One row per prep job: what you
are cooking, how long it takes, how many portions come out and which
days of the week they cover.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, id_col, lookup_recipe, status_cf, weekday_col

NAME = "Meal Prep"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("job", '"PR-"', digits=3, width=8),
        Col("job", "Prep job", 26, "text",
            note="'Chop all the veg', 'cook a big pot of bolognese', "
                 "'portion the chicken for the freezer'."),
        Col("date", "Prep day", 11, "date", validate=date_check(),
            note="Your usual prep day is set on the Settings tab."),
        weekday_col("date"),
        Col("recipe", "Recipe", 11, "text", validate=list_from("RecipeIDs")),
        Col("recipe_name", "What you are cooking", 24, "text",
            formula=lookup_recipe("name")),
        Col("batches", "Batches", 9, "number1",
            validate=number_check(0, 50),
            note="How many times the recipe. Two batches of a recipe "
                 "that serves four is eight portions."),
        Col("servings_each", "Serves (each)", 10, "int",
            formula=lookup_recipe("servings"), hidden=True),
        Col("portions", "Portions made", 11, "number1",
            formula='=IF(OR({batches}{r}="",{servings_each}{r}="",'
                    'NOT(ISNUMBER({servings_each}{r}))),"",'
                    '{batches}{r}*{servings_each}{r})'),
        Col("minutes_each", "Minutes (recipe)", 10, "int",
            formula=lookup_recipe("total_min"), hidden=True),
        Col("minutes", "Time needed", 11, "int",
            validate=number_check(0, 1440, decimal=False),
            note="Your estimate. Blank uses the recipe's own prep plus "
                 "cook time."),
        Col("eff_minutes", "Minutes used", 10, "int",
            formula='=IF({job}{r}="","",IF({minutes}{r}<>"",'
                    '{minutes}{r},IF(ISNUMBER({minutes_each}{r}),'
                    '{minutes_each}{r},0)))', hidden=True),
        Col("per_portion", "Minutes a portion", 11, "number1",
            formula='=IF(OR({portions}{r}="",{portions}{r}=0,'
                    '{eff_minutes}{r}=0),"",{eff_minutes}{r}/'
                    '{portions}{r})',
            note="The number that proves batch cooking works - watch "
                 "it fall as the batches get bigger.",
            cf=[{"type": "3_color_scale", "min_color": "#D9F2E4",
                 "mid_color": "#FFF1CF", "max_color": "#FBD9DD"}]),
        Col("storage", "Stored in", 12, "text",
            validate=list_from("StorageList")),
        Col("containers", "Containers", 11, "int",
            validate=number_check(0, 100, decimal=False)),
        Col("covers_from", "Covers meals from", 11, "date",
            validate=date_check()),
        Col("covers_to", "Covers meals to", 11, "date",
            validate=date_check()),
        Col("meals_covered", "Meals it covers", 11, "int",
            formula='=IF(OR({covers_from}{r}="",{covers_to}{r}=""),"",'
                    'COUNTIFS([[Meal Plan.date]],">="&{covers_from}{r},'
                    '[[Meal Plan.date]],"<="&{covers_to}{r},'
                    '[[Meal Plan.recipe]],{recipe}{r}))',
            note="Planned meals in that window that use this recipe."),
        Col("cook", "Who is doing it", 14, "text",
            validate=list_from("MemberNames")),
        Col("cost", "Cost of the batch", 11, "money",
            formula='=IF(OR({batches}{r}="",{recipe}{r}=""),"",'
                    'IFERROR(INDEX([[Recipes.cost]],MATCH({recipe}{r},'
                    '[[Recipes.id]],0))*{batches}{r},""))'),
        Col("cost_portion", "Cost a portion", 11, "money",
            formula='=IF(OR({cost}{r}="",{portions}{r}="",'
                    '{portions}{r}=0),"",{cost}{r}/{portions}{r})'),
        Col("todo_rank", "To-do rank", 9, "int",
            formula='=IF(OR({job}{r}="",{status}{r}="Done",'
                    'NOT(ISNUMBER({date}{r}))),"",COUNTIFS('
                    '{status_first}:{status_last},"<>Done",'
                    '{date_first}:{date_last},"<"&{date}{r})+'
                    'COUNTIFS({status_first}:{status}{r},"<>Done",'
                    '{date_first}:{date}{r},{date}{r}))',
            hidden=True,
            note="1 = the next unfinished prep job by date."),
        Col("status", "Status", 12, "status",
            validate=list_from("PrepStatusList"), cf=status_cf()),
        Col("alert", "Check", 24, "status",
            formula='=IF({job}{r}="","",'
                    'IF({status}{r}="Done","OK",'
                    'IF(AND({date}{r}<>"",{date}{r}<TODAY()),'
                    '"Overdue - still not made",'
                    'IF({date}{r}=TODAY(),"Today - prep day",'
                    'IF(AND({date}{r}<>"",{date}{r}-TODAY()<='
                    'PrepLeadDays),"Due - coming up","OK")))))',
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Meal Prep & Batch Cooking",
        subtitle="What you are making ahead, how many portions it "
                 "gives you and what each one costs in money and "
                 "minutes.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).prep,
        help_text="Plan two or three jobs for your prep day, not ten. "
                  "The minutes-a-portion column tells you which ones "
                  "are worth repeating.",
        extra_notes=[
            "Portions made feeds the leftovers tab: cook eight, eat "
            "four, and the other four are lunches with a value "
            "attached.",
            "Cost a portion is the number to compare with a shop "
            "bought lunch. It is usually the most convincing column "
            "in the workbook.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    spec_ = reg[NAME]
    s = sheet
    last = spec_.last_col_index
    row = spec_.last_data_row + 1
    row = s.section(row, "Batch cooking pays off", last)
    row = s.table_head(row, 0, ["Prep jobs", "Portions made",
                                "Hours in the kitchen",
                                "Average minutes a portion",
                                "Cost of the batches",
                                "Average cost a portion", "Still to make"],
                       [12, 14, 18, 22, 18, 20, 14])
    row = s.calc_row(row, 0, [
        '=COUNTIF([[Meal Prep.job]],"<>")',
        '=SUM([[Meal Prep.portions]])',
        '=SUM([[Meal Prep.eff_minutes]])/60',
        '=IFERROR(AVERAGE([[Meal Prep.per_portion]]),"")',
        '=SUM([[Meal Prep.cost]])',
        '=IFERROR(SUM([[Meal Prep.cost]])/SUM([[Meal Prep.portions]]),"")',
        '=COUNTIFS([[Meal Prep.status]],"<>Done",[[Meal Prep.job]],"<>")',
    ], ["int", "number1", "number1", "number1", "money", "money", "int"])
