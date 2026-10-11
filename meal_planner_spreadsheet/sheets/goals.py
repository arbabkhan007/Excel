"""`Goals` - what you are trying to change, and whether it is working.

Cook at home five nights a week. Spend less than £120. Try two new
recipes a month. Waste under a fiver. Each row is a goal with a target,
a live measurement pulled from the rest of the workbook and a simple
on-track flag.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, id_col, status_cf

NAME = "Goals"

#: the measures the 'Measured by' drop-down can read automatically
MEASURES = [
    ("Home cooked meals a week",
     'COUNTIFS([[Meal Plan.date]],">="&TODAY()-7,[[Meal Plan.date]],'
     '"<="&TODAY(),[[Meal Plan.home_cooked]],"Yes")'),
    ("Takeaways a month",
     'COUNTIFS([[Grocery Spending.month]],TEXT(TODAY(),"yyyy-mm"),'
     '[[Grocery Spending.type]],"Takeaway")'),
    ("Grocery spend this month",
     'SUMIFS([[Grocery Spending.amount]],[[Grocery Spending.month]],'
     'TEXT(TODAY(),"yyyy-mm"))'),
    ("New recipes this month",
     'COUNTIFS([[Recipes.added]],">="&DATE(YEAR(TODAY()),'
     'MONTH(TODAY()),1))'),
    ("Food waste this month",
     'SUMIFS([[Food Waste.value]],[[Food Waste.month]],'
     'TEXT(TODAY(),"yyyy-mm"))'),
    ("Packed lunches a week",
     'COUNTIFS([[Meal Plan.date]],">="&TODAY()-7,[[Meal Plan.slot]],'
     '"Lunch",[[Meal Plan.home_cooked]],"Yes")'),
    ("Leftovers eaten",
     'COUNTIFS([[Leftovers.status]],"Eaten")'),
    ("Meals prepped ahead",
     'COUNTIFS([[Meal Prep.status]],"Done")'),
]


def _measure_formula() -> str:
    """Nested IF that turns the chosen measure into a live number."""
    body = '""'
    for label, expr in reversed(MEASURES):
        body = f'IF({{measure}}{{r}}="{label}",{expr},{body})'
    return f'=IF({{measure}}{{r}}="","",IFERROR({body},""))'


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("goal", '"GL-"', digits=3, width=8),
        Col("goal", "Goal", 30, "text",
            note="Write it as something you would actually say: 'stop "
                 "ordering pizza on Fridays'."),
        Col("area", "Area", 14, "text", validate=list_from("GoalCategoryList")),
        Col("measure", "Measured by", 24, "text",
            validate=list_from("MeasureList"),
            note="Pick one of the measures the workbook can count for "
                 "itself, or choose 'Your own count' and fill in the "
                 "actual column by hand."),
        Col("target", "Target", 10, "number1",
            validate=number_check(0, 1000000)),
        Col("direction", "Aim to", 12, "text",
            validate=list_from("DirectionList"),
            note="'Hit or beat' for things you want more of, 'Stay "
                 "under' for spending and waste."),
        Col("auto", "Measured now", 11, "number1",
            formula=_measure_formula(), hidden=True),
        Col("manual", "Your own count", 11, "number1",
            validate=number_check(0, 1000000),
            note="Only needed for goals the workbook cannot count."),
        Col("actual", "Where you are", 11, "number1",
            formula='=IF({manual}{r}<>"",{manual}{r},{auto}{r})'),
        Col("progress", "Progress", 10, "percent0",
            formula='=IF(OR({target}{r}="",{target}{r}=0,'
                    '{actual}{r}=""),"",IF({direction}{r}='
                    '"Stay under",IF({actual}{r}=0,1,'
                    'MIN(2,{target}{r}/{actual}{r})),'
                    'MIN(2,{actual}{r}/{target}{r})))',
            cf=[{"type": "data_bar", "bar_color": "#4E8A63"}]),
        Col("on_track", "On track?", 11, "flag",
            formula='=IF({progress}{r}="","",IF({progress}{r}>=1,'
                    '"Yes","No"))',
            cf=[{"type": "text", "criteria": "containing", "value": "Yes",
                 "fmt": "good"},
                {"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "warn"}]),
        Col("why", "Why it matters", 26, "wrap",
            note="The reason you will remember in week three."),
        Col("start", "Started", 11, "date", validate=date_check()),
        Col("review", "Review on", 11, "date", validate=date_check()),
        Col("owner", "Whose goal", 14, "text",
            validate=list_from("MemberNames")),
        Col("status", "Status", 12, "status",
            validate=list_from("GoalStatusList"), cf=status_cf()),
        Col("alert", "Check", 24, "status",
            formula='=IF({goal}{r}="","",'
                    'IF({status}{r}="Done","OK",'
                    'IF({target}{r}="","Add - no target set",'
                    'IF(AND({review}{r}<>"",{review}{r}<TODAY()),'
                    '"Due - time to review this",'
                    'IF({on_track}{r}="No","Check - behind target",'
                    '"OK")))))',
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Food Goals & Habits",
        subtitle="Three or four goals, measured by the numbers already "
                 "in this workbook.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).goals,
        help_text="Three goals you will keep beats ten you will not. "
                  "The 'Measured now' column updates itself every time "
                  "the workbook recalculates.",
        extra_notes=[
            "Progress is capped at 200% so one brilliant week does "
            "not make the bar meaningless.",
            "'Stay under' goals invert the sum: hitting £90 against a "
            "£120 target reads as ahead, not behind.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    spec_ = reg[NAME]
    s = sheet
    last = spec_.last_col_index
    row = spec_.last_data_row + 1
    row = s.section(row, "How the habits are going", last)
    row = s.table_head(row, 0, ["Goals set", "On track", "Behind",
                                "Average progress", "Due a review",
                                "Finished"],
                       [12, 11, 11, 16, 14, 11])
    row = s.calc_row(row, 0, [
        '=COUNTIF([[Goals.goal]],"<>")',
        '=COUNTIF([[Goals.on_track]],"Yes")',
        '=COUNTIF([[Goals.on_track]],"No")',
        '=IFERROR(AVERAGE([[Goals.progress]]),"")',
        '=COUNTIFS([[Goals.review]],"<"&TODAY(),[[Goals.review]],">0",'
        '[[Goals.status]],"<>Done")',
        '=COUNTIF([[Goals.status]],"Done")',
    ], ["int", "int", "int", "percent0", "int", "int"])
