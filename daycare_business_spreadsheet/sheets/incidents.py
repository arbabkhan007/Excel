"""`Incident Reports` - accidents, injuries, near misses and who was told.

An incident form has one job beyond the obvious: to prove, months
later, that it was written up the same day, that a parent was told
within the hour and that somebody signed it.  Those three facts are
what gets checked, and all three are columns here.

The write-up column assembles a readable paragraph from the fields
beside it, so the form you print for the parent reads like a form and
not like a spreadsheet row.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, time_check
from .common import (alert_cf, blank, id_col, lookup_child,
                     lookup_staff, month_col, severity_cf, status_cf)

NAME = "Incident Reports"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "IncidentPrefix", digits=4, width=11,
               header="Incident ID",
               note="Automatic. Quote it to the parent and in your "
                    "own records."),
        Col("date", "Date", 11, "date", validate=date_check()),
        month_col("date"),
        Col("time", "Time", 10, "time", validate=time_check(),
            note="As close as you can. 'Around 2pm' helps nobody six "
                 "months later."),
        Col("child", "Child", 11, "id", validate=list_from("ChildList")),
        Col("child_name", "Child's name", 18, "text",
            formula=lookup_child("name")),
        Col("age_at", "Age then", 9, "text",
            formula=blank(
                "child",
                'IFERROR(DATEDIF(INDEX([[Child Enrollment.dob]],'
                'MATCH({child}{r},[[Child Enrollment.id]],0)),'
                '{date}{r},"y")&"y","")'),
            hidden=True),
        Col("room", "Room", 14, "text",
            formula=lookup_child("room", missing=""), hidden=True),
        Col("location", "Where", 16, "text",
            validate=list_from("IncidentLocationList")),
        Col("category", "What happened", 18, "text",
            validate=list_from("IncidentCategoryList"), cf=status_cf()),
        Col("severity", "Severity", 16, "status",
            validate=list_from("IncidentSeverityList"),
            cf=severity_cf()),
        Col("description", "Description", 40, "wrap",
            note="What happened, in order, in plain words. Facts "
                 "only - what you saw, not what you concluded."),
        Col("injury", "Injury & where", 28, "wrap",
            note="Body part and what it looked like. 'Graze, left "
                 "knee, 2cm' is a record; 'hurt leg' is not."),
        Col("first_aid", "First aid given", 30, "wrap"),
        Col("first_aider", "First aid by", 11, "id",
            validate=list_from("StaffList")),
        Col("first_aider_name", "First aider", 16, "text",
            formula=lookup_staff("name", "first_aider")),
        Col("staff", "Staff present", 11, "id",
            validate=list_from("StaffList"),
            note="Whoever was supervising at the time."),
        Col("staff_name", "Supervising", 16, "text",
            formula=lookup_staff("name", "staff"), hidden=True),
        Col("witness", "Witness", 18, "text"),
        # ------------------------------------------- telling the parent -
        Col("notified", "Parent told?", 11, "flag",
            validate=list_from("YesNoList"),
            cf=[{"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "bold_bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "good"}]),
        Col("notified_date", "Told on", 11, "date",
            validate=date_check()),
        Col("notified_time", "Told at", 10, "time",
            validate=time_check()),
        Col("notified_how", "Told by", 14, "text",
            validate=list_from("ContactMethodList")),
        Col("notified_by", "Told by whom", 11, "id",
            validate=list_from("StaffList"), hidden=True),
        Col("notify_mins", "Minutes to tell", 11, "int",
            formula=blank(
                "notified_date",
                'IF({time}{r}="","",IF({notified_time}{r}="","",'
                'ROUND((({notified_date}{r}+{notified_time}{r})-'
                '({date}{r}+{time}{r}))*1440,0)))'),
            cf=[{"type": "cell", "criteria": ">", "value": 240,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<=", "value": 60,
                 "fmt": "good"}],
            note="How long between the incident and telling the "
                 "parent. The number everybody asks about afterwards."),
        Col("signed", "Parent signed?", 11, "flag",
            validate=list_from("YesNoList"),
            cf=[{"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "good"}]),
        Col("signed_date", "Signed on", 11, "date",
            validate=date_check(), hidden=True),
        # ------------------------------------------------- afterwards ---
        Col("follow_up", "Follow-up action", 30, "wrap",
            note="What changed because of this. A gate catch, a rule, "
                 "a risk assessment. An incident with no follow-up "
                 "happens twice."),
        Col("follow_by", "Follow up by", 11, "date",
            validate=date_check()),
        Col("follow_done", "Done?", 8, "flag",
            validate=list_from("YesNoList"), cf=status_cf()),
        Col("reported_out", "Reported to", 18, "text",
            note="Only where your local rules require it - a "
                 "regulator, an insurer, a safeguarding lead. Blank "
                 "is fine and common."),
        Col("medical_sought", "Medical attention", 14, "text",
            hidden=True,
            note="None, first aid only, doctor, hospital. Record it "
                 "even when the answer is none."),
        Col("complete", "Form complete?", 11, "flag",
            formula=blank(
                "date",
                'IF({description}{r}="","No",'
                'IF({severity}{r}="","No",'
                'IF({notified}{r}<>"Yes","No",'
                'IF({first_aid}{r}="","No","Yes"))))'),
            cf=[{"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "good"}]),
        Col("writeup", "Printable write-up", 46, "wrap",
            formula=blank(
                "date",
                'IF({child_name}{r}="","",'
                '"On "&TEXT({date}{r},"dddd d mmmm yyyy")&'
                'IF({time}{r}="",""," at "&TEXT({time}{r},"h:mm AM/PM"))&'
                '", "&{child_name}{r}&" was involved in a "&'
                'LOWER(IF({category}{r}="","n incident",'
                '{category}{r}))&" in the "&'
                'LOWER(IF({location}{r}="","setting",{location}{r}))&'
                '". "&{description}{r}&'
                'IF({injury}{r}=""," "," Injury: "&{injury}{r}&".")&'
                'IF({first_aid}{r}=""," "," First aid given: "&'
                '{first_aid}{r}&".")&'
                'IF({first_aider_name}{r}=""," "," Treated by "&'
                '{first_aider_name}{r}&".")&'
                'IF({notified}{r}<>"Yes"," The parent has not yet '
                'been informed."," The parent was informed on "&'
                'TEXT({notified_date}{r},"d mmm yyyy")&'
                'IF({notified_time}{r}=""," "," at "&'
                'TEXT({notified_time}{r},"h:mm AM/PM"))&'
                'IF({notified_how}{r}=""," by "&LOWER('
                '{notified_how}{r})&".","."))&'
                'IF({follow_up}{r}=""," "," Follow-up: "&'
                '{follow_up}{r}&"."))'),
            note="A readable paragraph assembled from the columns on "
                 "this row. Widen the column, print the row, get it "
                 "signed."),
        Col("notes", "Internal notes", 26, "wrap",
            note="For you, not for the parent's copy."),
        Col("alert", "Action", 34, "status",
            formula=blank(
                "date",
                'IF({child}{r}="","Add - no child on this report",'
                'IF({description}{r}="","Missing - no description '
                'written",'
                'IF({severity}{r}="","Add - how serious was it?",'
                'IF({notified}{r}<>"Yes","Not notified - tell the '
                'parent today",'
                'IF({signed}{r}<>"Yes","Chase - parent has not signed '
                'the form",'
                'IF(IF({follow_by}{r}="",FALSE,'
                'IF({follow_done}{r}="Yes",FALSE,'
                '{follow_by}{r}<TodayDate)),'
                '"Overdue - follow-up action not done",'
                'IF(IF({follow_up}{r}="",'
                '({severity}{r}="Serious")+'
                '({severity}{r}="Emergency services called")>0,FALSE),'
                '"Action - serious incident with no follow-up",'
                'IF({complete}{r}="No","Action - form not complete",'
                '"OK"))))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Incident & Accident Reports",
        subtitle="What happened, who dealt with it, when the parent "
                 "was told and what changed afterwards.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.bad_text,
        freeze_cols=2,
        samples=sample(cfg).incidents,
        help_text="Fill the form in the same day. The write-up "
                  "column turns the row into a printable paragraph "
                  "for the parent to read and sign.",
        extra_notes=[
            "'Minutes to tell' is the column that matters if anybody "
            "ever questions how an incident was handled. Under an "
            "hour is green.",
            "A serious incident with nothing in the follow-up column "
            "is flagged until you write something there. That is "
            "deliberate - the follow-up is the part that stops it "
            "happening again.",
            "This workbook does not report anything to anybody. Where "
            "your local rules require an incident to be notified to a "
            "regulator or insurer, that is a separate job - record "
            "that you did it in 'Reported to'.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Incident summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Incident record", last)
    sheet.table_head(row, 0, [
        "Incidents logged", "This month", "No injury", "Minor",
        "Moderate or worse", "Parents told", "Not yet told",
        "Signed by parent", "Average minutes to tell",
        "Follow-ups outstanding", "Forms incomplete"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Incident Reports.date]],"<>")',
        '=COUNTIFS([[Incident Reports.month]],ThisMonthKey)',
        '=COUNTIFS([[Incident Reports.severity]],"No injury")',
        '=COUNTIFS([[Incident Reports.severity]],"Minor")',
        '=COUNTIFS([[Incident Reports.severity]],"Moderate")+'
        'COUNTIFS([[Incident Reports.severity]],"Serious")+'
        'COUNTIFS([[Incident Reports.severity]],'
        '"Emergency services called")',
        '=COUNTIFS([[Incident Reports.notified]],"Yes")',
        '=COUNTIF([[Incident Reports.date]],"<>")-'
        'COUNTIFS([[Incident Reports.notified]],"Yes")',
        '=COUNTIFS([[Incident Reports.signed]],"Yes")',
        '=IFERROR(ROUND(AVERAGEIF([[Incident Reports.notify_mins]],'
        '">=0"),0),0)',
        '=COUNTIFS([[Incident Reports.follow_by]],"<>",'
        '[[Incident Reports.follow_done]],"<>Yes")',
        '=COUNTIFS([[Incident Reports.complete]],"No")',
    ], ["int", "int", "int", "int", "int", "int", "int", "int", "int",
        "int", "int"])
