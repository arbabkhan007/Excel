"""`Staff Schedule` - the rota, the clock and what the day actually cost.

One row per person per shift.  Scheduled hours and clocked hours are
kept in separate columns so the variance between the two is visible,
because in childcare the variance is where the money goes.

Overtime is allocated properly rather than fudged.  Each row knows how
many hours that person has already been rota'd in the same week up to
and including this shift, so the hours that tip them over the weekly
threshold are the ones charged at the overtime multiplier - not the
whole shift, and not nothing.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check, time_check
from .common import (alert_cf, bar_cf, blank, id_col, lookup_staff,
                     lookup_staff_num, month_col, status_cf,
                     week_start_col, weekday_col)

NAME = "Staff Schedule"

#: Shifts that do not put an adult on the floor.
NON_FLOOR = ["Off", "Leave", "Training", "Admin"]


def spec(cfg: Config) -> SheetSpec:
    floor_test = "+".join(f'({{shift}}{{r}}="{s}")' for s in NON_FLOOR)

    columns = [
        id_col("date", "ShiftPrefix", digits=5, width=10,
               header="Shift ID"),
        Col("date", "Date", 11, "date", validate=date_check()),
        weekday_col("date"),
        week_start_col("date"),
        month_col("date"),
        Col("staff", "Staff", 10, "id", validate=list_from("StaffList")),
        Col("staff_name", "Name", 18, "text",
            formula=lookup_staff("name")),
        Col("role", "Role", 16, "text",
            formula=lookup_staff("role"), hidden=True),
        Col("pay_type", "Paid as", 12, "text",
            formula=lookup_staff("pay_type"), hidden=True),
        Col("room", "Room", 15, "text",
            validate=list_from("ClassroomList"),
            note="Where they are working that shift. The ratio "
                 "figures on the classroom tab count this column, "
                 "not the person's usual room."),
        Col("shift", "Shift", 13, "text",
            validate=list_from("ShiftTypeList"), cf=status_cf(),
            note="Off, Leave, Training and Admin do not put an adult "
                 "on the floor, so they are excluded from the ratio "
                 "counts."),
        Col("counts", "Counts in ratio?", 11, "flag",
            formula=blank("shift",
                          f'IF(({floor_test})>0,"No","Yes")'),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "good"},
                {"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "muted"}]),
        Col("leave_type", "Leave type", 15, "text",
            validate=list_from("LeaveTypeList"), hidden=True),
        # ------------------------------------------------- rota'd hours
        Col("start", "Starts", 10, "time", validate=time_check()),
        Col("finish", "Ends", 10, "time", validate=time_check()),
        Col("break_min", "Break (min)", 10, "int",
            validate=number_check(0, 240, decimal=False),
            note="Unpaid break. Leave blank to use the default on "
                 "Settings."),
        Col("hours", "Rota'd hours", 11, "number1",
            # anchored on the date, not the start time: a leave day
            # has no shift but still has to read zero rather than
            # blank, or every hours column downstream turns to
            # #VALUE!
            formula=blank(
                "date",
                'IF({start}{r}="",0,IF({finish}{r}="",0,'
                'ROUND(MAX(0,(MOD({finish}{r}-{start}{r},1))*24-'
                '(IF({break_min}{r}="",BreakMinutes,{break_min}{r})/60))'
                ',2)))'),
            cf=bar_cf()),
        # ------------------------------------------------- actual hours
        Col("clock_in", "Clocked in", 10, "time",
            validate=time_check(),
            note="Leave blank and the rota'd hours are used. Fill it "
                 "in and the variance column shows the difference."),
        Col("clock_out", "Clocked out", 10, "time",
            validate=time_check()),
        Col("actual_hours", "Actual hours", 11, "number1",
            formula=blank(
                "date",
                'IF({clock_in}{r}="",{hours}{r},'
                'IF({clock_out}{r}="",{hours}{r},'
                'ROUND(MAX(0,(MOD({clock_out}{r}-{clock_in}{r},1))*24-'
                '(IF({break_min}{r}="",BreakMinutes,{break_min}{r})/60))'
                ',2)))'),
            note="What they actually worked. Falls back to the rota "
                 "when nobody clocked in or out, so the payroll is "
                 "never short."),
        Col("variance", "Over / under", 10, "number1",
            formula=blank(
                "date",
                'ROUND({actual_hours}{r}-{hours}{r},2)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0.25,
                 "fmt": "warn"},
                {"type": "cell", "criteria": "<", "value": -0.25,
                 "fmt": "info"}],
            hidden=True),
        # ----------------------------------------------------- overtime
        Col("week_to_date", "Week to date", 11, "number1",
            formula=blank(
                "date",
                'ROUND(SUMIFS({actual_hours_first}:{actual_hours}{r},'
                '{staff_first}:{staff}{r},{staff}{r},'
                '{week_start_first}:{week_start}{r},'
                '{week_start}{r}),2)'),
            hidden=True,
            note="Hours this person has worked in the same week up to "
                 "and including this shift. Overtime is worked out "
                 "from it."),
        Col("ot_hours", "Overtime", 9, "number1",
            formula=blank(
                "date",
                'ROUND(MAX(0,MIN({actual_hours}{r},'
                '{week_to_date}{r}-OvertimeAfter)),2)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}],
            note="Only the hours on this shift that tip the person "
                 "over the weekly threshold - not the whole shift."),
        Col("basic_hours", "Basic hours", 10, "number1",
            formula=blank(
                "date",
                'ROUND({actual_hours}{r}-{ot_hours}{r},2)'),
            hidden=True),
        Col("rate", "Rate", 10, "money",
            formula=lookup_staff_num("rate"), hidden=True),
        Col("cost", "Shift cost", 11, "money",
            formula=blank(
                "date",
                'ROUND(({basic_hours}{r}+{ot_hours}{r}*'
                'OvertimeMultiplier)*{rate}{r},2)'),
            cf=bar_cf(Palette.accent),
            note="Basic hours plus overtime at the multiplier, at "
                 "that person's effective rate. Before employer "
                 "on-costs."),
        Col("cost_full", "With on-costs", 11, "money",
            formula=blank(
                "date",
                'ROUND({cost}{r}*(1+IF(({pay_type}{r}="Contractor")+'
                '({pay_type}{r}="Volunteer - unpaid")>0,0,'
                'PayrollBurden)),2)'),
            hidden=True),
        # ----------------------------------------------------- cover ---
        Col("substitute", "Covered by", 10, "id",
            validate=list_from("StaffList"),
            note="When somebody else picks the shift up. The original "
                 "name stays so you can see what happened."),
        Col("sub_name", "Cover", 16, "text",
            formula=lookup_staff("name", "substitute"), hidden=True),
        Col("clash", "Double booked", 9, "int",
            formula=blank(
                "date",
                'IF({start}{r}="",0,'
                'COUNTIFS({staff_first}:{staff_last},{staff}{r},'
                '{date_first}:{date_last},{date}{r},'
                '{start_first}:{start_last},"<"&{finish}{r},'
                '{finish_first}:{finish_last},">"&{start}{r},'
                '{counts_first}:{counts_last},"Yes"))'),
            hidden=True,
            note="Overlapping shifts for the same person on the same "
                 "day. Anything above 1 is a clash."),
        Col("notes", "Notes", 24, "wrap"),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "date",
                'IF({staff}{r}="","Add - nobody on this shift",'
                'IF({staff_name}{r}="Not in Staff Database",'
                '"Check - staff ID not found",'
                'IF({shift}{r}="","Add - what kind of shift?",'
                'IF({clash}{r}>1,"Duplicate - double booked at the '
                'same time",'
                'IF(IF({counts}{r}<>"Yes",FALSE,{room}{r}=""),'
                '"Add - no room, so it does not count in ratio",'
                'IF(IF({counts}{r}<>"Yes",FALSE,{start}{r}=""),'
                '"Add - no start time",'
                'IF(IF({clock_in}{r}="",FALSE,'
                'IF({clock_out}{r}<>"",FALSE,{date}{r}<TodayDate)),'
                '"Missing - clocked in and never out",'
                'IF({ot_hours}{r}>0,"Check - "&'
                'TEXT({ot_hours}{r},"0.0")&" overtime hours",'
                '"OK"))))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Staff Schedule & Timesheet",
        subtitle="The rota, the clock, the overtime and what each "
                 "shift costs - with the ratio counts feeding the "
                 "classroom tab.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).schedule,
        help_text="One row per person per shift. Rota it in advance, "
                  "fill in the clock times afterwards, and the hours, "
                  "the overtime and the cost all follow.",
        extra_notes=[
            "Actual hours fall back to the rota'd hours when nobody "
            "clocked in or out, so a forgotten clock-out never "
            "silently removes somebody's pay.",
            "Overtime is allocated shift by shift: the workbook adds "
            "up what each person has already worked that week and "
            "only charges the hours above the threshold at the "
            "higher rate.",
            "'Counts in ratio' is what the classroom tab adds up. "
            "Training, admin, leave and days off are excluded, which "
            "is why a room can look staffed on the rota and still be "
            "short on the floor.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Rota summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Rota summary", last)
    sheet.table_head(row, 0, [
        "Shifts", "Rota'd hours", "Actual hours", "Variance",
        "Overtime hours", "Wage cost", "Cost with on-costs",
        "Average cost an hour", "Leave days", "Double bookings",
        "This month cost"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Staff Schedule.date]],"<>")',
        '=ROUND(SUM([[Staff Schedule.hours]]),1)',
        '=ROUND(SUM([[Staff Schedule.actual_hours]]),1)',
        '=ROUND(SUM([[Staff Schedule.actual_hours]])-'
        'SUM([[Staff Schedule.hours]]),1)',
        '=ROUND(SUM([[Staff Schedule.ot_hours]]),1)',
        '=ROUND(SUM([[Staff Schedule.cost]]),2)',
        '=ROUND(SUM([[Staff Schedule.cost_full]]),2)',
        '=IF(SUM([[Staff Schedule.actual_hours]])<=0,"",'
        'ROUND(SUM([[Staff Schedule.cost]])/'
        'SUM([[Staff Schedule.actual_hours]]),2))',
        '=COUNTIFS([[Staff Schedule.shift]],"Leave")',
        '=COUNTIFS([[Staff Schedule.clash]],">1")',
        '=ROUND(SUMIFS([[Staff Schedule.cost]],'
        '[[Staff Schedule.month]],ThisMonthKey),2)',
    ], ["int", "number1", "number1", "number1", "number1", "money",
        "money", "money", "int", "int", "money"])
