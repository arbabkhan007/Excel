"""`Household` - who you are cooking for.

One row per person: what they will not eat, what they must not eat, how
hungry they are and which meals they are actually in for. The planner
uses the portion factor to work out how much food a meal really needs,
and the allergy column drives the warnings on the recipes.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, id_col, status_cf

NAME = "Household"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", '"HH-"', digits=2, width=8),
        Col("name", "Name", 18, "text",
            note="Everyone who eats here, including the ones who only "
                 "turn up at weekends."),
        Col("role", "Role", 14, "text", validate=list_from("MemberRoleList")),
        Col("age_group", "Age group", 13, "text",
            validate=list_from("AgeGroupList"),
            note="Only used to suggest a portion size."),
        Col("portion", "Portion factor", 11, "number1",
            validate=number_check(0.25, 3),
            note="1 is an adult portion. A small child might be 0.5, a "
                 "teenager 1.5. The planner multiplies by this."),
        Col("diet", "Diet", 15, "text", validate=list_from("DietList")),
        Col("allergies", "Allergies", 20, "text",
            validate=list_from("AllergenList"),
            note="The serious ones. Recipes containing them are "
                 "flagged - but always read the label yourself."),
        Col("intolerances", "Avoids", 20, "text",
            note="Things they would rather not eat: 'mushrooms', "
                 "'anything spicy'."),
        Col("dislikes", "Will not touch", 22, "text"),
        Col("loves", "Favourites", 22, "text",
            note="Worth filling in - it makes the 'what shall we have?' "
                 "question much easier."),
        Col("breakfast", "In for breakfast?", 11, "flag",
            validate=list_from("YesNoList")),
        Col("lunch", "In for lunch?", 11, "flag",
            validate=list_from("YesNoList")),
        Col("dinner", "In for dinner?", 11, "flag",
            validate=list_from("YesNoList")),
        Col("packed", "Takes a packed lunch?", 11, "flag",
            validate=list_from("YesNoList")),
        Col("cal_target", "Daily calories", 11, "number",
            validate=number_check(0, 6000),
            note="Leave blank to use the household target on the "
                 "Settings tab."),
        Col("eff_cals", "Calories used", 11, "number",
            formula='=IF({name}{r}="","",IF({cal_target}{r}="",'
                    'CalorieTarget*IF({portion}{r}="",1,{portion}{r}),'
                    '{cal_target}{r}))', hidden=True),
        Col("cooks", "Cooks?", 9, "flag", validate=list_from("YesNoList")),
        Col("cook_days", "Cooking nights", 16, "text",
            note="'Tuesday and Thursday'. The weekly planner shows it "
                 "so nobody can claim they forgot."),
        Col("shops", "Does the shopping?", 10, "flag",
            validate=list_from("YesNoList")),
        Col("jobs", "Kitchen jobs", 20, "text",
            note="Washing up, bins, packing lunches."),
        Col("meals_planned", "Meals planned for them", 11, "int",
            formula='=IF({name}{r}="","",COUNTIFS('
                    '[[Meal Plan.cook]],{name}{r}))',
            note="How many meals this person is down to cook."),
        Col("birthday", "Birthday", 11, "date", validate=date_check(),
            note="So the special events tab can remind you."),
        Col("status", "Status", 12, "status",
            validate=list_from("ActiveList"), cf=status_cf()),
        Col("alert", "Check", 24, "status",
            formula='=IF({name}{r}="","",IF({portion}{r}="",'
                    '"Add - no portion factor",IF(AND({allergies}{r}<>"",'
                    '{allergies}{r}<>"None"),"Check - allergy on file",'
                    '"OK")))',
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Household & Preferences",
        subtitle="Who eats here, what they can and cannot eat, and how "
                 "much of it - the numbers behind every portion in the "
                 "workbook.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).members,
        help_text="Fill this in once. The portion factors add up to the "
                  "household size the planner uses when you do not "
                  "type a number of servings.",
        extra_notes=[
            "Portion factor is the honest way to plan for a family "
            "with a toddler and a teenager: two adults, a 0.5 and a "
            "1.5 is 4 portions, not 4 identical plates.",
            "Allergy flags are an aid, not a medical check. Recipes "
            "are flagged when an ingredient carries the same allergen "
            "tag - always read the actual packaging.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    spec_ = reg[NAME]
    s = sheet
    last = spec_.last_col_index
    row = spec_.last_data_row + 1
    row = s.section(row, "The household in numbers", last)
    row = s.table_head(row, 0, ["People", "Total portions",
                                "Daily calories for the house",
                                "Allergies on file", "Packed lunches",
                                "Cooks in the house"],
                       [10, 14, 24, 16, 15, 17])
    row = s.calc_row(row, 0, [
        '=COUNTIFS([[Household.name]],"<>",[[Household.status]],'
        '"Active")',
        '=SUMIFS([[Household.portion]],[[Household.status]],"Active")',
        '=SUMIFS([[Household.eff_cals]],[[Household.status]],"Active")',
        '=COUNTIFS([[Household.allergies]],"<>",'
        '[[Household.allergies]],"<>None")',
        '=COUNTIF([[Household.packed]],"Yes")',
        '=COUNTIF([[Household.cooks]],"Yes")',
    ], ["int", "number1", "number", "int", "int", "int"])
