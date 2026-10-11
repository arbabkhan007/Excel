"""`Events` - birthdays, Christmas, the in-laws on Sunday.

Occasions need a head count, a menu, a budget and a cooking timetable.
This tab holds the occasion; the meals themselves go on the meal plan
with the event picked in the Occasion column, so the shopping list
picks them up like any other meal.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, id_col, status_cf

NAME = "Events"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "EventPrefix", digits=3, width=9),
        Col("name", "Occasion", 24, "text",
            note="'Dad's birthday dinner', 'Christmas Day', 'book "
                 "club'. The name appears in the meal plan drop-down."),
        Col("type", "Type", 14, "text", validate=list_from("EventTypeList")),
        Col("date", "Date", 11, "date", validate=date_check()),
        Col("days_to_go", "Days to go", 10, "int",
            formula='=IF({date}{r}="","",{date}{r}-TODAY())',
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "muted"},
                {"type": "cell", "criteria": "between", "minimum": 0,
                 "maximum": 7, "fmt": "warn"}]),
        Col("time", "Time", 10, "time"),
        Col("guests", "Guests", 9, "int",
            validate=number_check(0, 500, decimal=False),
            note="Including the household - it is the number you cook "
                 "for."),
        Col("adults", "Adults", 8, "int",
            validate=number_check(0, 500, decimal=False)),
        Col("children", "Children", 9, "int",
            validate=number_check(0, 500, decimal=False)),
        Col("portions", "Portions to cook", 11, "number1",
            formula='=IF({guests}{r}="","",IF(AND({adults}{r}="",'
                    '{children}{r}=""),{guests}{r},'
                    'IF({adults}{r}="",0,{adults}{r})+'
                    'IF({children}{r}="",0,{children}{r})*0.5))',
            note="Children count as half a portion - change the "
                 "formula if your children eat like mine."),
        Col("dietary", "Dietary needs", 24, "text",
            note="'One vegan, one coeliac, no nuts at all'. Check it "
                 "against the Allergens tab before you shop."),
        Col("budget", "Budget", 11, "money",
            validate=number_check(0, 100000)),
        Col("planned_cost", "Planned food cost", 12, "money",
            formula='=IF({id}{r}="","",SUMIFS([[Meal Plan.cost]],'
                    '[[Meal Plan.event]],{name}{r}))',
            note="From every meal on the plan tagged with this "
                 "occasion."),
        Col("actual_cost", "Actual cost", 11, "money",
            validate=number_check(0, 100000)),
        Col("vs_budget", "Against budget", 12, "money",
            formula='=IF({budget}{r}="","",{budget}{r}-'
                    'IF({actual_cost}{r}<>"",{actual_cost}{r},'
                    'IF({planned_cost}{r}="",0,{planned_cost}{r})))',
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">=", "value": 0,
                 "fmt": "good"}]),
        Col("meals", "Meals planned", 10, "int",
            formula='=IF({name}{r}="","",COUNTIFS('
                    '[[Meal Plan.event]],{name}{r}))'),
        Col("menu", "Menu", 30, "wrap",
            note="The running order: nibbles, starter, main, pudding."),
        Col("prep_start", "Start cooking", 11, "date",
            validate=date_check(),
            note="Work backwards from the meal - most of the stress "
                 "is in the timetable, not the food."),
        Col("shopping_done", "Shopping done?", 11, "flag",
            validate=list_from("YesNoList")),
        Col("host", "Who is hosting", 14, "text",
            validate=list_from("MemberNames")),
        Col("status", "Status", 12, "status",
            validate=list_from("EventStatusList"), cf=status_cf()),
        Col("alert", "Check", 26, "status",
            formula='=IF({name}{r}="","",'
                    'IF({status}{r}="Done","OK",'
                    'IF(AND({days_to_go}{r}<>"",{days_to_go}{r}<0),'
                    '"Check - date has passed",'
                    'IF(AND({days_to_go}{r}<=7,{meals}{r}=0),'
                    '"Action - nothing on the menu yet",'
                    'IF(AND({days_to_go}{r}<=3,{shopping_done}{r}<>'
                    '"Yes"),"Due - shopping not done","OK")))))',
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap",
            note="What worked, what to never attempt again."),
    ]

    return SheetSpec(
        name=NAME,
        title="Special Events",
        subtitle="Occasions with a head count, a budget and a "
                 "timetable - and a menu that flows into the normal "
                 "shopping list.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).events,
        help_text="Add the occasion here, then add its meals to the "
                  "Meal Plan with this occasion chosen - that is what "
                  "gets the ingredients onto the grocery list.",
    )
