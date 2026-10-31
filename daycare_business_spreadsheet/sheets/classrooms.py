"""`Classroom Management` - how full each room is, and whether it is legal.

Two numbers run a childcare setting: how many children are in a room,
and how many adults are standing in it.  This tab holds both, works out
the seats you have left to sell and the staff the room needs right now,
and goes red before the inspector does.

Every ratio comes from the age band table on Settings, and every one of
them can be overridden per room, because the number of two-year-olds one
adult may supervise is a local licensing decision and not something a
spreadsheet should have an opinion about.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import (alert_cf, blank, bar_cf, id_col, lookup_staff,
                     status_cf)

NAME = "Classroom Management"


def spec(cfg: Config) -> SheetSpec:
    child_room = "[[Child Enrollment.room]]"
    child_status = "[[Child Enrollment.status]]"
    att_room = "[[Attendance.room]]"
    att_date = "[[Attendance.date]]"
    att_status = "[[Attendance.status]]"
    sch_room = "[[Staff Schedule.room]]"
    sch_date = "[[Staff Schedule.date]]"
    sch_counts = "[[Staff Schedule.counts]]"

    # Children holding a place in this room: Active, On hold or Notice
    # given.  Withdrawn and Graduated free the seat up.
    def occupying(extra: str = "") -> str:
        return ("+".join(
            f'COUNTIFS({child_room},{{name}}{{r}},{child_status},"{s}"'
            f'{extra})'
            for s in ("Active", "On hold", "Notice given")))

    columns = [
        id_col("name", "ClassroomPrefix", digits=2, width=10,
               header="Room ID",
               note="Automatic. Children and shifts point at the room "
                    "name, not this ID, so renaming a room is a one "
                    "cell job."),
        Col("name", "Room name", 20, "text",
            note="What you actually call it - Acorn Room, Baby Unit, "
                 "Pre-K 2. Every other tab picks the room from this "
                 "list, so type it once and type it carefully."),
        Col("age_band", "Age band", 15, "text",
            validate=list_from("AgeBandList"),
            note="Sets the default staff-to-child ratio for this room. "
                 "The bands and their ratios are yours to edit on "
                 "Settings."),
        Col("age_from", "From (months)", 11, "int",
            formula=blank(
                "age_band",
                'IFERROR(INDEX(AgeBandFrom,MATCH({age_band}{r},'
                'AgeBandList,0))*1,0)')),
        Col("age_to", "To (months)", 11, "int",
            formula=blank(
                "age_band",
                'IFERROR(INDEX(AgeBandTo,MATCH({age_band}{r},'
                'AgeBandList,0))*1,0)')),
        Col("capacity", "Licensed places", 13, "int",
            validate=number_check(0, 500, decimal=False),
            note="The most children allowed in this room at once, from "
                 "your licence - not the number of chairs."),
        Col("ratio_default", "Band ratio", 10, "int",
            formula=blank(
                "age_band",
                'IFERROR(INDEX(AgeBandRatio,MATCH({age_band}{r},'
                'AgeBandList,0))*1,0)'),
            hidden=True),
        Col("ratio_set", "Ratio override", 12, "int",
            validate=number_check(0, 40, decimal=False),
            note="Children per adult, if this room differs from the "
                 "band default. Leave blank to use the band."),
        Col("ratio", "Children per adult", 14, "int",
            formula=blank(
                "name",
                'IF({ratio_set}{r}<>"",{ratio_set}{r},'
                'IF({ratio_default}{r}>0,{ratio_default}{r},""))'),
            note="The ratio this room is actually judged against."),
        Col("max_group", "Max group size", 12, "int",
            formula=blank(
                "age_band",
                'IFERROR(INDEX(AgeBandMax,MATCH({age_band}{r},'
                'AgeBandList,0))*1,0)'),
            hidden=True),
        Col("teacher", "Lead teacher", 11, "id",
            validate=list_from("StaffList"),
            note="Staff ID of whoever leads the room."),
        Col("teacher_name", "Teacher", 18, "text",
            formula=lookup_staff("name", "teacher")),
        Col("assistant", "Assistant", 11, "id",
            validate=list_from("StaffList")),
        Col("assistant_name", "Assistant name", 18, "text",
            formula=lookup_staff("name", "assistant"), hidden=True),
        Col("room_status", "Room status", 12, "status",
            validate=list_from("RoomStatusList"), cf=status_cf(),
            note="Closed and Planned rooms are left out of the "
                 "occupancy figures on the dashboard."),
        # ------------------------------------------------- who is in it --
        Col("enrolled", "Children enrolled", 14, "int",
            formula=blank("name", occupying()),
            cf=bar_cf(),
            note="Active, On hold and Notice given - everyone still "
                 "holding a place. Withdrawn children free the seat."),
        Col("active", "Active only", 11, "int",
            formula=blank(
                "name",
                f'COUNTIFS({child_room},{{name}}{{r}},'
                f'{child_status},"Active")'),
            hidden=True),
        Col("starting", "Starting soon", 12, "int",
            formula=blank(
                "name",
                f'COUNTIFS({child_room},{{name}}{{r}},'
                f'{child_status},"Pending start")'),
            note="Offered a place and not in the room yet. They will "
                 "take a seat - plan for them."),
        Col("leaving", "Leaving soon", 11, "int",
            formula=blank(
                "name",
                f'COUNTIFS({child_room},{{name}}{{r}},'
                f'{child_status},"Notice given")'),
            hidden=True),
        Col("seats", "Seats free", 10, "int",
            formula=blank(
                "name",
                'IF({capacity}{r}="","",{capacity}{r}-{enrolled}{r})'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bold_bad"},
                {"type": "cell", "criteria": "=", "value": 0,
                 "fmt": "warn"},
                {"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "good"}],
            note="Places you could still sell today. A negative number "
                 "means you are over your licence."),
        Col("occupancy", "Occupancy", 10, "percent",
            formula=blank(
                "name",
                'IF({capacity}{r}="","",IF({capacity}{r}<=0,"",'
                '{enrolled}{r}/{capacity}{r}))'),
            cf=[{"type": "cell", "criteria": ">", "value": 1,
                 "fmt": "bold_bad"},
                {"type": "cell", "criteria": ">=",
                 "value": "=RatioWarning", "fmt": "warn"},
                {"type": "cell", "criteria": "<",
                 "value": "=RatioWarning", "fmt": "good"}],
            note="Enrolled divided by licensed places."),
        Col("waiting", "On the waitlist", 13, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Waitlist.room]],{name}{r},'
                '[[Waitlist.converted]],"No")'),
            note="Families still waiting for this room. A full room "
                 "with a queue behind it is a price rise, not a "
                 "problem."),
        # ------------------------------------------------------- today --
        Col("present_today", "In today", 10, "int",
            formula=blank(
                "name",
                "+".join(
                    f'COUNTIFS({att_room},{{name}}{{r}},'
                    f'{att_date},TodayDate,{att_status},"{s}")'
                    for s in ("Present", "Late", "Early pickup"))),
            note="From the attendance register, for today's date."),
        Col("staff_today", "Adults in today", 13, "number1",
            formula=blank(
                "name",
                f'SUMIFS([[Staff Schedule.hours]],{sch_room},'
                f'{{name}}{{r}},{sch_date},TodayDate,'
                f'{sch_counts},"Yes")/MAX(1,OpenHoursDay)'),
            note="Rota hours in this room today divided by the length "
                 "of the day - so a half-day cover counts as half an "
                 "adult, which is what the ratio really means."),
        Col("staff_needed", "Adults needed", 12, "number1",
            formula=blank(
                "name",
                'IF({ratio}{r}="","",IF({ratio}{r}<=0,"",'
                'ROUNDUP({present_today}{r}/{ratio}{r},0)))'),
            note="For the children actually in the building today."),
        Col("staff_needed_full", "Adults at full", 12, "number1",
            formula=blank(
                "name",
                'IF({ratio}{r}="","",IF({ratio}{r}<=0,"",'
                'ROUNDUP({enrolled}{r}/{ratio}{r},0)))'),
            hidden=True,
            note="What the room needs on a day when everybody turns "
                 "up. Staff to this, not to today."),
        Col("ratio_now", "Ratio today", 11, "text",
            formula=blank(
                "name",
                'IF({staff_today}{r}<=0,IF({present_today}{r}=0,'
                '"No session","No cover"),'
                '"1:"&ROUND({present_today}{r}/{staff_today}{r},1))'),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "No cover", "fmt": "bold_bad"}],
            note="One adult to this many children, right now."),
        Col("ratio_ok", "Ratio met?", 10, "flag",
            formula=blank(
                "name",
                'IF({present_today}{r}=0,"No session",'
                'IF({ratio}{r}="","Set ratio",'
                'IF({staff_today}{r}>={staff_needed}{r},"Yes","No")))'),
            cf=[{"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "bold_bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "good"}]),
        # -------------------------------------------------- the money ----
        Col("revenue", "Room revenue a week", 15, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Child Enrollment.weekly_net]],'
                '[[Child Enrollment.room]],{name}{r}),2)'),
            cf=bar_cf(Palette.accent),
            note="What the children in this room are contracted to pay "
                 "in a week, after discounts and subsidy."),
        Col("revenue_seat", "Per place a week", 13, "money",
            formula=blank(
                "name",
                'IF({capacity}{r}="","",IF({capacity}{r}<=0,"",'
                'ROUND({revenue}{r}/{capacity}{r},2)))'),
            hidden=True,
            note="Weekly revenue divided by licensed places - the "
                 "number to compare rooms on."),
        Col("empty_cost", "Empty places cost", 14, "money",
            formula=blank(
                "name",
                'IF({seats}{r}="","",IF({enrolled}{r}<=0,"",'
                'ROUND(MAX(0,{seats}{r})*({revenue}{r}/'
                'MAX(1,{enrolled}{r})),2)))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}],
            note="What those empty seats would bring in a week at the "
                 "rate the room already charges. This is the number to "
                 "put next to an advertising budget."),
        Col("notes", "Notes", 30, "wrap"),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "name",
                'IF({capacity}{r}="","Add - set the licensed places",'
                'IF({ratio}{r}="","Add - pick an age band or ratio",'
                'IF({enrolled}{r}>{capacity}{r},'
                '"Over capacity - stop enrolling",'
                'IF(IF({present_today}{r}=0,FALSE,'
                '{staff_today}{r}<{staff_needed}{r}),'
                '"Over ratio - get cover into the room",'
                'IF({teacher}{r}="","Add - no lead teacher",'
                'IF({enrolled}{r}>={capacity}{r},'
                '"Nearly full - open the waitlist",'
                'IF({occupancy}{r}="","OK",'
                'IF({occupancy}{r}>=RatioWarning,'
                '"Nearly full - last places",'
                'IF({seats}{r}>=3,"Space - '
                + str(cfg.currency_symbol) +
                '"&TEXT({empty_cost}{r},"#,##0")&" a week empty",'
                '"OK")))))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Classroom Management",
        subtitle="Capacity, occupancy, staff-to-child ratios and what "
                 "every empty place is costing you a week.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).classrooms,
        help_text="Name each room, give it an age band and the number "
                  "of places on your licence. Enrolment, attendance, "
                  "ratios and revenue all fill themselves in from the "
                  "other tabs.",
        extra_notes=[
            "Ratios come from the age band table on Settings & Lists. "
            "They are a starting point, not a legal opinion - set them "
            "to whatever your own licence says and the whole workbook "
            "follows.",
            "'Adults in today' is rota hours divided by the length of "
            "your day, so somebody covering a four hour shift counts "
            "as part of an adult rather than a whole one.",
            "'Empty places cost' is the fastest argument for a "
            "marketing budget you will find in this file.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Whole-setting capacity summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "The whole setting", last)
    sheet.table_head(row, 0, [
        "Rooms open", "Licensed places", "Children enrolled",
        "Seats free", "Occupancy", "In today", "Adults in today",
        "Rooms over ratio", "Weekly contracted revenue"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIFS([[Classroom Management.room_status]],"Open")',
        '=SUM([[Classroom Management.capacity]])',
        '=SUM([[Classroom Management.enrolled]])',
        '=SUM([[Classroom Management.capacity]])-'
        'SUM([[Classroom Management.enrolled]])',
        '=IF(SUM([[Classroom Management.capacity]])<=0,"",'
        'SUM([[Classroom Management.enrolled]])/'
        'SUM([[Classroom Management.capacity]]))',
        '=SUM([[Classroom Management.present_today]])',
        '=ROUND(SUM([[Classroom Management.staff_today]]),1)',
        '=COUNTIFS([[Classroom Management.ratio_ok]],"No")',
        '=ROUND(SUM([[Classroom Management.revenue]]),2)',
    ], ["int", "int", "int", "int", "percent", "int", "number1", "int",
        "money"])
