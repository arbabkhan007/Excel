"""`Meal Tracker` - what was served, who ate it, and who must not have it.

One row per child per sitting.  The reason this tab exists and a paper
menu does not is the two columns in the middle: every plate carries the
child's recorded diet and their allergy level, pulled live from the
medical tab, so the person serving lunch sees the warning on the same
line as the food.

It is a prompt, not a safety system.  The allergy plan on the wall is
the safety system.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, lookup_child,
                     lookup_staff, month_col, severity_cf, status_cf)

NAME = "Meal Tracker"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "MealPrefix", digits=5, width=11,
               header="Meal ID"),
        Col("date", "Date", 11, "date", validate=date_check()),
        month_col("date"),
        Col("sitting", "Sitting", 14, "text",
            validate=list_from("MealSittingList"), cf=status_cf(),
            note="Breakfast, lunch, snack, bottle. One row per child "
                 "per sitting."),
        Col("child", "Child", 11, "id", validate=list_from("ChildList")),
        Col("child_name", "Child's name", 18, "text",
            formula=lookup_child("name")),
        Col("room", "Room", 14, "text",
            formula=lookup_child("room", missing=""), hidden=True),
        Col("daykey", "Day key", 13, "text",
            formula=blank("child", '{child}{r}&"|"&{date}{r}'),
            hidden=True),
        # ------------------------------------------------ the warnings --
        Col("diet", "Diet on file", 15, "text",
            formula=lookup_child("diet", missing=""),
            cf=[{"type": "text", "criteria": "not containing",
                 "value": "No restrictions", "fmt": "info"}],
            note="From the child's enrollment record. Nothing to type "
                 "- change it there and every future plate follows."),
        Col("allergy_level", "Allergy risk", 13, "status",
            formula=lookup_child("allergy_level", missing=""),
            cf=severity_cf(),
            note="The worst allergy severity recorded for this child "
                 "on the Allergy & Medical tab. Red means read the "
                 "plan before you serve anything."),
        Col("allergy_main", "Allergic to", 18, "text",
            formula=lookup_child("first_allergy", missing=""),
            cf=[{"type": "text", "criteria": "not containing",
                 "value": "zzzz", "fmt": "warn"}],
            note="The first allergy on the child's medical record. "
                 "Where there is more than one, the medical tab has "
                 "the full list."),
        # --------------------------------------------------- the plate --
        Col("food", "Food served", 30, "text",
            note="What actually went on the plate, not what the menu "
                 "said. They differ more often than anyone admits."),
        Col("portion", "Amount eaten", 13, "text",
            validate=list_from("EatenAmountList"),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "All", "fmt": "good"},
                {"type": "text", "criteria": "containing",
                 "value": "Most", "fmt": "good"},
                {"type": "text", "criteria": "containing",
                 "value": "None", "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "Refused", "fmt": "bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Not offered", "fmt": "muted"}]),
        Col("refused", "Refused?", 9, "flag",
            formula=blank(
                "portion",
                'IF(({portion}{r}="Refused")+({portion}{r}="None")>0,'
                '"Yes","No")'),
            hidden=True),
        Col("alternative", "Alternative given", 22, "text",
            hidden=True,
            note="What you offered instead. Worth recording for a "
                 "fussy eater's parent."),
        Col("drink", "Drink", 16, "text"),
        Col("water_ml", "Fluid (ml)", 10, "int",
            validate=number_check(0, 3000, decimal=False),
            note="Only if you track it. Infant rooms usually do, "
                 "preschool rooms usually do not."),
        Col("own_food", "Brought own?", 10, "flag",
            validate=list_from("YesNoList"), hidden=True),
        Col("logged_by", "Logged by", 11, "id",
            validate=list_from("StaffList")),
        Col("logged_name", "Staff", 15, "text",
            formula=lookup_staff("name", "logged_by"), hidden=True),
        Col("notes", "Notes", 26, "wrap"),
        Col("alert", "Action", 32, "status",
            formula=blank(
                "date",
                'IF({child}{r}="","Add - no child picked",'
                'IF(({allergy_level}{r}="Anaphylaxis risk")+'
                '({allergy_level}{r}="Severe")>0,'
                '"Allergy - "&{allergy_main}{r}&", check the plan",'
                'IF({sitting}{r}="","Add - which sitting?",'
                'IF({food}{r}="","Add - what was served?",'
                'IF({portion}{r}="","Add - how much did they eat?",'
                'IF({refused}{r}="Yes","Check - refused, tell the '
                'parent",'
                'IF({allergy_level}{r}="Moderate","Check - moderate '
                'allergy on file","OK")))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Meal & Nutrition Tracker",
        subtitle="Every plate, with the child's diet and allergy "
                 "level on the same line as the food.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).meals,
        help_text="Pick the date, the sitting and the child. The diet "
                  "and the allergy warning appear on their own - you "
                  "type what was served and how much went in.",
        extra_notes=[
            "The allergy columns are pulled live from the Allergy & "
            "Medical tab. Record an allergy once, and it appears on "
            "every future meal for that child without anybody having "
            "to remember.",
            "This is a prompt, not a safeguard. The written allergy "
            "plan, the kitchen's own checks and the adult standing at "
            "the table are the safeguard.",
            "'Refused' twice in a day shows up on the child's daily "
            "report automatically, which saves the awkward "
            "conversation at the door.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Meal log summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "What got eaten", last)
    sheet.table_head(row, 0, [
        "Plates logged", "Ate it all", "Ate most", "Ate some",
        "Refused", "Refusal rate", "Allergy plates served",
        "Special diets served", "Fluids logged (ml)"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Meal Tracker.date]],"<>")',
        '=COUNTIFS([[Meal Tracker.portion]],"All")',
        '=COUNTIFS([[Meal Tracker.portion]],"Most")',
        '=COUNTIFS([[Meal Tracker.portion]],"About half")+'
        'COUNTIFS([[Meal Tracker.portion]],"A little")',
        '=COUNTIFS([[Meal Tracker.refused]],"Yes")',
        '=IF(COUNTIF([[Meal Tracker.portion]],"<>")<=0,"",'
        'COUNTIFS([[Meal Tracker.refused]],"Yes")/'
        'COUNTIF([[Meal Tracker.portion]],"<>"))',
        '=COUNTIFS([[Meal Tracker.allergy_level]],"Severe")+'
        'COUNTIFS([[Meal Tracker.allergy_level]],"Anaphylaxis risk")',
        '=COUNTIF([[Meal Tracker.diet]],"<>")-'
        'COUNTIFS([[Meal Tracker.diet]],"No restrictions")',
        '=SUM([[Meal Tracker.water_ml]])',
    ], ["int", "int", "int", "int", "int", "percent", "int", "int",
        "int"])
