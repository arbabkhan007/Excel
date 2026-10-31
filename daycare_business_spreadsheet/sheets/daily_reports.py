"""`Daily Reports` - the sheet that goes home with the child.

One row per child per day: how they were, what they did, what they ate,
how they slept, how many nappies, and the sentence the parent actually
reads at the kitchen table.

Arrival and departure times are pulled from the attendance register
rather than typed again, meal counts come from the meal tracker, and
the end-of-day summary writes itself out of the columns beside it.  The
staff member filling this in at 5:45pm types about eight things.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check, time_check
from .common import (alert_cf, blank, id_col, lookup_child,
                     lookup_staff, month_col, status_cf, weekday_col)

NAME = "Daily Reports"


def spec(cfg: Config) -> SheetSpec:
    def from_attendance(target: str) -> str:
        return blank(
            "daykey",
            f'IFERROR(INDEX([[Attendance.{target}]],'
            f'MATCH({{daykey}}{{r}},[[Attendance.daykey]],0))*1,"")')

    columns = [
        id_col("date", "ReportPrefix", digits=5, width=11,
               header="Report ID"),
        Col("date", "Date", 11, "date", validate=date_check()),
        weekday_col("date", hidden=True),
        month_col("date"),
        Col("child", "Child", 11, "id", validate=list_from("ChildList")),
        Col("child_name", "Child's name", 18, "text",
            formula=lookup_child("name")),
        Col("preferred", "Known as", 13, "text",
            formula=lookup_child("preferred", missing=""), hidden=True),
        Col("room", "Room", 14, "text",
            formula=lookup_child("room", missing="")),
        Col("daykey", "Day key", 13, "text",
            formula=blank("child", '{child}{r}&"|"&{date}{r}'),
            hidden=True),
        Col("arrived", "Arrived", 10, "time",
            formula=from_attendance("in_time"),
            note="From the attendance register, matched on child and "
                 "date. Nothing to type twice."),
        Col("departed", "Left", 10, "time",
            formula=from_attendance("out_time")),
        Col("hours", "Hours", 8, "number1",
            formula=from_attendance("hours"), hidden=True),
        # ------------------------------------------------- how they were
        Col("mood", "Mood", 12, "text", validate=list_from("MoodList"),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Tearful", "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "Unwell", "fmt": "bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Unsettled", "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "Happy", "fmt": "good"},
                {"type": "text", "criteria": "containing",
                 "value": "Settled", "fmt": "good"}]),
        Col("behaviour", "Behaviour note", 24, "wrap", hidden=True),
        Col("activities", "What we did", 34, "wrap",
            note="Two or three things. Parents want specifics - "
                 "'painted a rainbow and told everyone about it' "
                 "beats 'arts and crafts'."),
        Col("learning", "Learning focus", 26, "wrap",
            note="The skill behind the activity. This is the column "
                 "that makes a daily report feel like a nursery "
                 "rather than a babysitter."),
        Col("outdoor_min", "Outdoors (min)", 11, "int",
            validate=number_check(0, 600, decimal=False),
            cf=[{"type": "cell", "criteria": ">=", "value": 60,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 30,
                 "fmt": "warn"}]),
        # ------------------------------------------------------- naps ---
        Col("nap_from", "Nap from", 10, "time", validate=time_check()),
        Col("nap_to", "Nap to", 10, "time", validate=time_check()),
        Col("nap_hours", "Slept", 8, "number1",
            formula=blank(
                "nap_from",
                'IF({nap_to}{r}="","",'
                'ROUND((MOD({nap_to}{r}-{nap_from}{r},1))*24,2))')),
        Col("nap2_min", "Second nap (min)", 12, "int",
            validate=number_check(0, 400, decimal=False), hidden=True),
        Col("sleep_total", "Total sleep", 10, "number1",
            formula=blank(
                "date",
                'ROUND(IF({nap_hours}{r}="",0,{nap_hours}{r})+'
                'IF({nap2_min}{r}="",0,{nap2_min}{r}/60),2)'),
            hidden=True),
        # --------------------------------------------- nappies & potty --
        Col("nappies_wet", "Nappies wet", 10, "int",
            validate=number_check(0, 30, decimal=False)),
        Col("nappies_soiled", "Nappies soiled", 11, "int",
            validate=number_check(0, 30, decimal=False)),
        Col("nappies", "Nappy changes", 11, "int",
            formula=blank(
                "date",
                'IF({nappies_wet}{r}="",0,{nappies_wet}{r})+'
                'IF({nappies_soiled}{r}="",0,{nappies_soiled}{r})'),
            hidden=True),
        Col("toileting", "Toileting", 14, "text",
            validate=list_from("ToiletingList"),
            note="For children who are out of nappies or on their way."),
        Col("potty_notes", "Potty training", 22, "wrap", hidden=True,
            note="Progress, accidents, what worked. Parents want to "
                 "keep it consistent at home."),
        # -------------------------------------------------- food & drink
        Col("meals_logged", "Meals logged", 11, "int",
            formula=blank(
                "daykey",
                'COUNTIFS([[Meal Tracker.daykey]],{daykey}{r})'),
            note="From the meal tracker. Log the plates there and "
                 "they are counted here."),
        Col("ate_well", "Ate well", 9, "int",
            formula=blank(
                "daykey",
                'COUNTIFS([[Meal Tracker.daykey]],{daykey}{r},'
                '[[Meal Tracker.portion]],"All")+'
                'COUNTIFS([[Meal Tracker.daykey]],{daykey}{r},'
                '[[Meal Tracker.portion]],"Most")'),
            hidden=True),
        Col("refused", "Refused", 9, "int",
            formula=blank(
                "daykey",
                'COUNTIFS([[Meal Tracker.daykey]],{daykey}{r},'
                '[[Meal Tracker.portion]],"Refused")+'
                'COUNTIFS([[Meal Tracker.daykey]],{daykey}{r},'
                '[[Meal Tracker.portion]],"None")'),
            cf=[{"type": "cell", "criteria": ">=", "value": 2,
                 "fmt": "warn"}],
            hidden=True),
        Col("eating", "Eating", 13, "text",
            formula=blank(
                "date",
                'IF({meals_logged}{r}=0,"Not logged",'
                'IF({refused}{r}>={meals_logged}{r},"Refused most",'
                'IF({ate_well}{r}>={meals_logged}{r},"Ate everything",'
                '"Ate some")))'),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Refused", "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "everything", "fmt": "good"},
                {"type": "text", "criteria": "containing",
                 "value": "Not logged", "fmt": "muted"}]),
        Col("water", "Drinks", 8, "int",
            validate=number_check(0, 30, decimal=False),
            note="Cups or bottles offered and taken."),
        Col("bottles", "Bottles (ml)", 10, "int",
            validate=number_check(0, 3000, decimal=False), hidden=True),
        # --------------------------------------------------- the people -
        Col("teacher", "Written by", 11, "id",
            validate=list_from("StaffList")),
        Col("teacher_name", "Staff", 16, "text",
            formula=lookup_staff("name", "teacher"), hidden=True),
        Col("observations", "Teacher's observations", 34, "wrap",
            note="The one thing you noticed that the parent did not "
                 "see. This is the whole value of a daily report."),
        Col("parent_notes", "Message to parents", 30, "wrap",
            note="Bring wellies tomorrow. Nearly out of nappies. We "
                 "need the trip form back."),
        Col("incidents", "Incidents today", 10, "int",
            formula=blank(
                "date",
                'COUNTIFS([[Incident Reports.child]],{child}{r},'
                '[[Incident Reports.date]],{date}{r})'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "bad"}],
            note="Pulled from the incident tab so nothing is sent "
                 "home without you knowing it happened."),
        Col("summary", "End of day summary", 44, "wrap",
            formula=blank(
                "date",
                'IF({child_name}{r}="","",'
                'IF({preferred}{r}="",{child_name}{r},'
                '{preferred}{r})&" was "&'
                'LOWER(IF({mood}{r}="","settled",{mood}{r}))&" today. "&'
                'IF({sleep_total}{r}>0,"Slept "&'
                'TEXT({sleep_total}{r},"0.0")&"h. ","")&'
                'IF({meals_logged}{r}>0,{eating}{r}&" at "&'
                '{meals_logged}{r}&" meal(s). ","")&'
                'IF({nappies}{r}>0,{nappies}{r}&" nappy changes. ","")&'
                'IF({outdoor_min}{r}>0,{outdoor_min}{r}&'
                '" minutes outside. ","")&'
                'IF({incidents}{r}>0,"Please read the incident form '
                'we have sent home. ","")&'
                'IF({parent_notes}{r}="","",{parent_notes}{r}))'),
            note="Written for you out of the columns to the left. "
                 "Copy it into your app, print it, or hand the row "
                 "over as it is."),
        Col("shared", "Sent to parent?", 11, "flag",
            validate=list_from("YesNoList"), cf=status_cf()),
        Col("alert", "Action", 28, "status",
            formula=blank(
                "date",
                'IF({child}{r}="","Add - no child picked",'
                'IF({mood}{r}="","Add - how were they today?",'
                'IF(IF({activities}{r}="",{observations}{r}="",FALSE),'
                '"Add - nothing recorded for the parent",'
                'IF({incidents}{r}>0,"Action - incident form goes '
                'home too",'
                'IF(IF({shared}{r}="Yes",FALSE,{date}{r}<TodayDate),'
                '"Chase - not sent to the parent yet",'
                'IF({refused}{r}>=2,"Check - refused more than one '
                'meal","OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Daily Child Reports",
        subtitle="Mood, activities, naps, nappies, meals and the "
                 "end-of-day summary - written for you from the "
                 "columns beside it.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).daily_reports,
        help_text="Pick the date and the child. Arrival times, meals "
                  "and incidents come from the other tabs; you fill "
                  "in the mood, the activities and the observation.",
        extra_notes=[
            "The summary column writes a full sentence out of "
            "everything else on the row. Filter to one day, widen the "
            "column and print - that is your take-home report.",
            "Meals come from the Meal Tracker and arrival times from "
            "Attendance, both matched on the child and the date. If a "
            "time is blank here, it is blank on the attendance sheet.",
            "If a child had an incident today, the summary says so. "
            "No daily report should be the first a parent hears of it.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Daily report summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Reports written", last)
    sheet.table_head(row, 0, [
        "Reports logged", "Written today", "Sent to parents",
        "Not sent yet", "Average sleep (hours)",
        "Average outdoor minutes", "Average nappy changes",
        "Children unsettled", "Days with a refused meal"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Daily Reports.date]],"<>")',
        '=COUNTIFS([[Daily Reports.date]],TodayDate)',
        '=COUNTIFS([[Daily Reports.shared]],"Yes")',
        '=COUNTIF([[Daily Reports.date]],"<>")-'
        'COUNTIFS([[Daily Reports.shared]],"Yes")',
        '=IFERROR(ROUND(AVERAGEIF([[Daily Reports.sleep_total]],">0"),'
        '1),0)',
        '=IFERROR(ROUND(AVERAGEIF([[Daily Reports.outdoor_min]],">0"),'
        '0),0)',
        '=IFERROR(ROUND(AVERAGEIF([[Daily Reports.nappies]],">0"),1),0)',
        '=COUNTIFS([[Daily Reports.mood]],"Unsettled")+'
        'COUNTIFS([[Daily Reports.mood]],"Tearful")',
        '=COUNTIFS([[Daily Reports.refused]],">0")',
    ], ["int", "int", "int", "int", "number1", "int", "number1", "int",
        "int"])
