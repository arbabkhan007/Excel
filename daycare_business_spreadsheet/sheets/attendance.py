"""`Attendance` - the daily sign in and sign out sheet.

One row per child per day.  Everything else follows from it: the
attendance percentage on the enrollment tab, the ratio figures on the
classroom tab, the late pickup charges on the fee tab, the "who is in
today" tile on the dashboard and every monthly report in the file.

Three things this tab does that a paper sign-in sheet cannot.  It
measures late arrivals and late collections against each child's own
contracted times rather than against the centre's opening hours.  It
checks the name of whoever collected the child against the authorised
list and says so when it cannot find them.  And it refuses to let a
missing check-out time quietly disappear.
"""

from __future__ import annotations

from ..config import (ABSENT_STATUSES, NON_SESSION_STATUSES,
                      PRESENT_STATUSES, Config, Palette)
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, time_check
from .common import (alert_cf, blank, id_col, late_pickup_charge,
                     lookup_child, lookup_staff, month_col, status_cf,
                     week_start_col, weekday_col)

NAME = "Attendance"


def spec(cfg: Config) -> SheetSpec:
    present_test = "+".join(f'({{status}}{{r}}="{s}")'
                            for s in PRESENT_STATUSES)
    absent_test = "+".join(f'({{status}}{{r}}="{s}")'
                           for s in ABSENT_STATUSES)
    nonsession_test = "+".join(f'({{status}}{{r}}="{s}")'
                               for s in NON_SESSION_STATUSES)

    columns = [
        id_col("date", "AttendancePrefix", digits=5, width=11,
               header="Entry ID"),
        Col("date", "Date", 11, "date", validate=date_check(),
            note="One row per child per day. Fill a whole day down the "
                 "sheet in the morning and finish the times as people "
                 "arrive."),
        weekday_col("date"),
        month_col("date"),
        week_start_col("date"),
        Col("child", "Child", 11, "id", validate=list_from("ChildList"),
            note="Pick the child ID. Everything else about them fills "
                 "in automatically."),
        Col("child_name", "Child's name", 20, "text",
            formula=lookup_child("name")),
        Col("room", "Room", 15, "text",
            formula=lookup_child("room", missing=""),
            note="From the enrollment tab. The ratio figures on the "
                 "classroom tab count this column."),
        Col("program", "Programme", 13, "text",
            formula=lookup_child("program", missing=""), hidden=True),
        Col("daykey", "Day key", 14, "text",
            formula=blank("child", '{child}{r}&"|"&{date}{r}'),
            hidden=True,
            note="Child and date joined together. The daily report "
                 "and the meal tracker find this row with it."),
        Col("status", "Status", 16, "status",
            validate=list_from("AttendanceStatusList"), cf=status_cf(),
            note="Present, Late and Early pickup all mean the child "
                 "was here. Public holiday, Centre closed, Snow day "
                 "and Not scheduled mean nobody was expected, so they "
                 "do not count against the attendance percentage."),
        Col("session", "Session?", 8, "int",
            formula=blank(
                "status",
                f'IF(({nonsession_test})>0,0,1)'),
            hidden=True,
            note="1 when a session was expected. The denominator of "
                 "every attendance percentage in the workbook."),
        Col("present", "Present?", 8, "int",
            formula=blank("status", f'IF(({present_test})>0,1,0)'),
            hidden=True),
        Col("absent", "Absent?", 8, "int",
            formula=blank("status", f'IF(({absent_test})>0,1,0)'),
            hidden=True),
        # ----------------------------------------------------- times ----
        Col("in_time", "Checked in", 11, "time", validate=time_check(),
            note="The time they actually arrived. Type 8:15 AM or "
                 "08:15 - both work."),
        Col("out_time", "Checked out", 11, "time",
            validate=time_check(),
            note="Leave blank and the Action column will not let you "
                 "forget. A sign-in sheet with no sign-out is the one "
                 "thing an inspector always finds."),
        Col("hours", "Hours", 8, "number1",
            formula=blank(
                "in_time",
                'IF({out_time}{r}="","",'
                'ROUND((MOD({out_time}{r}-{in_time}{r},1))*24,2))'),
            note="Time in the building. Hourly fees and the "
                 "'contracted versus actual' comparison use it."),
        Col("due_in", "Due in", 10, "time",
            formula=blank(
                "child",
                'IFERROR(INDEX([[Child Enrollment.usual_in]],'
                'MATCH({child}{r},[[Child Enrollment.id]],0))*1,"")'),
            hidden=True),
        Col("due_out", "Due out", 10, "time",
            formula=blank(
                "child",
                'IFERROR(INDEX([[Child Enrollment.usual_out]],'
                'MATCH({child}{r},[[Child Enrollment.id]],0))*1,"")'),
            hidden=True),
        Col("late_in", "Late in (min)", 11, "int",
            formula=blank(
                "in_time",
                'IF({due_in}{r}="","",'
                'MAX(0,ROUND(({in_time}{r}-{due_in}{r})*1440,0)))'),
            cf=[{"type": "cell", "criteria": ">=", "value": 30,
                 "fmt": "warn"}],
            note="Minutes after this child's own contracted drop-off "
                 "time, not after opening time."),
        Col("early_out", "Left early (min)", 12, "int",
            formula=blank(
                "out_time",
                'IF({due_out}{r}="","",'
                'MAX(0,ROUND(({due_out}{r}-{out_time}{r})*1440,0)))'),
            hidden=True),
        Col("late_out", "Late pickup (min)", 13, "int",
            formula=blank(
                "out_time",
                'IF({due_out}{r}="","",'
                'MAX(0,ROUND(({out_time}{r}-{due_out}{r})*1440,0)))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}],
            note="Minutes after the contracted collection time. The "
                 "charge below is worked out from it."),
        Col("late_fee", "Late fee due", 11, "money",
            formula=blank(
                "late_out",
                'IF({late_out}{r}<=0,0,' +
                late_pickup_charge("{late_out}{r}") + ")"),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}],
            note="Free for the grace minutes, then the first block, "
                 "then per minute - all three set on Settings. It is "
                 "not charged until you raise it on the Tuition & "
                 "Fees tab."),
        Col("fee_raised", "Fee charged?", 11, "flag",
            validate=list_from("YesNoList"),
            note="Say Yes once you have added the charge to Tuition & "
                 "Fees, so the Action column stops reminding you."),
        # ------------------------------------------------ who and why ---
        Col("dropped_by", "Dropped off by", 18, "text"),
        Col("collected_by", "Collected by", 18, "text",
            note="Type the name of whoever actually took the child. "
                 "It is checked against the authorised list on the "
                 "child's and the family's record."),
        Col("auth_list", "Authorised list", 22, "text",
            formula=blank(
                "child",
                'IFERROR(INDEX([[Child Enrollment.pickup_list]],'
                'MATCH({child}{r},[[Child Enrollment.id]],0))&"","")&'
                '" "&IFERROR(INDEX([[Parent Database.pickup_list]],'
                'MATCH(IFERROR(INDEX([[Child Enrollment.family]],'
                'MATCH({child}{r},[[Child Enrollment.id]],0))&"",""),'
                '[[Parent Database.id]],0))&"","")&" "&'
                'IFERROR(INDEX([[Parent Database.parent1]],'
                'MATCH(IFERROR(INDEX([[Child Enrollment.family]],'
                'MATCH({child}{r},[[Child Enrollment.id]],0))&"",""),'
                '[[Parent Database.id]],0))&"","")&" "&'
                'IFERROR(INDEX([[Parent Database.parent2]],'
                'MATCH(IFERROR(INDEX([[Child Enrollment.family]],'
                'MATCH({child}{r},[[Child Enrollment.id]],0))&"",""),'
                '[[Parent Database.id]],0))&"","")'),
            hidden=True,
            note="The child's own list, the family's list and both "
                 "parents' names, joined together so the check below "
                 "can search it."),
        Col("pickup_ok", "Authorised?", 11, "flag",
            formula=blank(
                "collected_by",
                'IF({auth_list}{r}="","Not checked",'
                'IF(ISNUMBER(SEARCH({collected_by}{r},'
                '{auth_list}{r})),"Yes","No"))'),
            cf=[{"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "bold_bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "good"}],
            note="Searches the authorised names for whoever you typed. "
                 "'No' does not mean stop - it means check, and that "
                 "is exactly what it is for."),
        Col("reason", "Absence reason", 22, "text",
            note="Why they are not in. 'Unexplained' three days "
                 "running is a safeguarding prompt, not an admin "
                 "chore."),
        Col("recorded_by", "Recorded by", 11, "id",
            validate=list_from("StaffList")),
        Col("recorded_name", "Staff", 16, "text",
            formula=lookup_staff("name", "recorded_by"), hidden=True),
        Col("temp_check", "Temp / health check", 12, "text",
            hidden=True,
            note="Only if your setting does them. Left blank it costs "
                 "you nothing."),
        Col("notes", "Notes", 26, "wrap"),
        Col("run_absent", "Absent in a row", 11, "int",
            formula=blank(
                "child",
                'IF({absent}{r}=0,0,'
                'COUNTIFS({child_first}:{child}{r},{child}{r},'
                '{absent_first}:{absent}{r},1,'
                '{date_first}:{date}{r},">"&({date}{r}-'
                'AbsenceChaseDays)))'),
            hidden=True,
            note="Absences by this child inside the chase window set "
                 "on Settings."),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "date",
                'IF({child}{r}="","Add - no child picked",'
                'IF({child_name}{r}="Not in Child Enrollment",'
                '"Check - child ID not found",'
                'IF({status}{r}="","Add - no status yet",'
                'IF(IF({present}{r}=0,FALSE,{in_time}{r}=""),'
                '"Missing - no check-in time",'
                'IF(IF({present}{r}=0,FALSE,'
                'IF({out_time}{r}<>"",FALSE,{date}{r}<TodayDate)),'
                '"No check-out - find out who took them",'
                'IF({pickup_ok}{r}="No","Not authorised - check who '
                'collected",'
                'IF(IF({late_fee}{r}<=0,FALSE,{fee_raised}{r}<>"Yes"),'
                '"Invoice - late pickup fee not charged",'
                'IF(IF({absent}{r}=0,FALSE,{reason}{r}=""),'
                '"Unexplained - call the parent",'
                'IF({run_absent}{r}>=AbsenceChaseDays,'
                '"Chase - absent several days running",'
                'IF({status}{r}="Late","Late - "&{late_in}{r}&'
                '" minutes","OK"))))))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Daily Attendance, Check-In & Check-Out",
        subtitle="Who was in, when they arrived, when they left, who "
                 "collected them - and what that late pickup is worth.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.deep,
        freeze_cols=2,
        samples=sample(cfg).attendance,
        help_text="One row per child per day. Pick the date, the "
                  "child and the status, then fill the times in as "
                  "they happen. Everything to the right is worked out "
                  "for you.",
        extra_notes=[
            "Late arrivals and late collections are measured against "
            "each child's own contracted times on the enrollment tab, "
            "so a child contracted until 3pm is late at 3:10 even "
            "though you are open until 6.",
            "The authorised check searches the child's pickup list, "
            "the family's pickup list and both parents' names. It is "
            "a prompt to look, never a decision - the adult in the "
            "room makes that.",
            "Statuses that mean nobody was expected (Public holiday, "
            "Centre closed, Snow day, Not scheduled) are left out of "
            "the attendance percentage, which is the only way to get "
            "a number that is actually fair to the family.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Attendance summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Attendance summary", last)
    sheet.table_head(row, 0, [
        "Entries logged", "Sessions expected", "Attended", "Absences",
        "Attendance rate", "In today", "Absent today", "Late arrivals",
        "Late pickups", "Late fees earned", "Late fees not charged",
        "Missing check-outs"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Attendance.date]],"<>")',
        '=SUM([[Attendance.session]])',
        '=SUM([[Attendance.present]])',
        '=SUM([[Attendance.absent]])',
        '=IF(SUM([[Attendance.session]])<=0,"",'
        'SUM([[Attendance.present]])/SUM([[Attendance.session]]))',
        '=SUMIFS([[Attendance.present]],[[Attendance.date]],TodayDate)',
        '=SUMIFS([[Attendance.absent]],[[Attendance.date]],TodayDate)',
        '=COUNTIFS([[Attendance.status]],"Late")',
        '=COUNTIFS([[Attendance.late_out]],">0")',
        '=ROUND(SUM([[Attendance.late_fee]]),2)',
        '=ROUND(SUMIFS([[Attendance.late_fee]],'
        '[[Attendance.fee_raised]],"<>Yes"),2)',
        '=COUNTIFS([[Attendance.alert]],"No check-out*")',
    ], ["int", "int", "int", "int", "percent", "int", "int", "int",
        "int", "money", "money", "int"])
