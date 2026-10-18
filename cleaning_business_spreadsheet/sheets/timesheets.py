"""`Timesheets` - hours worked, what they cost and whether they were billed.

One row per cleaner per job (or per shift). The hours come from the
start and finish times less the break, the cost comes from that
person's true hourly cost on the Employees tab, and the billable flag
is what makes the utilisation figure on the dashboard mean something.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import (date_check, list_from, number_check, time_check)
from .common import (alert_cf, blank, flag_yes_no_cf, id_col, month_col,
                     status_cf, week_col, weekday_col)

NAME = "Timesheets"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("employee", '"TS-"', digits=4, width=10),
        Col("date", "Date", 11, "date", validate=date_check()),
        weekday_col("date"),
        week_col("date"),
        month_col("date"),
        Col("employee", "Cleaner", 18, "text",
            validate=list_from("EmployeeNames")),
        Col("job", "Job ID", 11, "id", validate=list_from("JobIDs"),
            note="Leave blank for time that is not against a job - "
                 "training, admin, a supply run."),
        Col("client_name", "Client", 20, "text",
            formula='=IF({job}{r}="","",IFERROR(INDEX('
                    '[[Jobs.client_name]],MATCH({job}{r},[[Jobs.id]],0)),'
                    '""))'),
        Col("activity", "Activity", 16, "text",
            validate=list_from("MileagePurposeList"),
            note="Only needed for non-job time. Blank means cleaning."),
        Col("start", "Start", 9, "time", validate=time_check()),
        Col("finish", "Finish", 9, "time", validate=time_check()),
        Col("break_min", "Break (min)", 10, "int",
            validate=number_check(0, 480, decimal=False)),
        Col("hours_input", "Hours (override)", 11, "number1",
            validate=number_check(0, 24),
            note="Only if you are not recording start and finish "
                 "times. Otherwise leave it blank."),
        Col("overtime", "Of which overtime", 10, "number1",
            validate=number_check(0, 24),
            note="Hours paid at the overtime multiplier from the "
                 "Settings tab."),
        Col("attendance", "Attendance", 16, "status",
            validate=list_from("AttendanceList"), cf=status_cf()),
        Col("billable", "Billable?", 9, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="Yes for time a client pays for. Travel, training "
                 "and admin are No - and that gap is your "
                 "utilisation."),
        Col("approved", "Approved?", 9, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf()),
        Col("paid", "Paid?", 8, "flag", validate=list_from("YesNoList"),
            cf=flag_yes_no_cf()),
        Col("notes", "Notes", 30, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("clock_hours", "Clocked", 9, "number1",
            formula='=IF(OR({start}{r}="",{finish}{r}=""),"",'
                    'ROUND(({finish}{r}-{start}{r})*24-'
                    'IF({break_min}{r}="",0,{break_min}{r})/60,2))',
            hidden=True,
            note="Finish less start, less the break."),
        Col("hours", "Hours", 9, "number1",
            formula=blank("employee",
                          'IF({hours_input}{r}<>"",{hours_input}{r},'
                          'IF({clock_hours}{r}="",0,{clock_hours}{r}))'),
            note="The hours everything else counts: pay, job cost, "
                 "utilisation.",
            cf=[{"type": "data_bar", "bar_color": "#2F8F9D"}]),
        Col("rate", "Cost an hour", 10, "money",
            formula=blank(
                "employee",
                'IFERROR(INDEX([[Employees.cost_rate]],MATCH('
                '{employee}{r},[[Employees.name]],0)),'
                'PayRate*(1+EmployerOncost))'),
            hidden=True,
            note="That person's true cost from the Employees tab."),
        Col("basic_hours", "Basic hours", 9, "number1",
            formula=blank("employee",
                          'MAX(0,{hours}{r}-IF({overtime}{r}="",0,'
                          '{overtime}{r}))'),
            hidden=True),
        Col("cost", "Cost to you", 11, "money",
            formula=blank(
                "employee",
                'ROUND({basic_hours}{r}*{rate}{r}+IF({overtime}{r}="",0,'
                '{overtime}{r})*{rate}{r}*OvertimeRate,2)'),
            note="Basic hours plus overtime at the multiplier, at "
                 "that cleaner's true cost per hour."),
        Col("billed_value", "Billed value", 11, "money",
            formula=blank(
                "employee",
                'IF({billable}{r}<>"Yes",0,ROUND({hours}{r}*'
                'IF(IFERROR(INDEX([[Clients.rate]],MATCH(IFERROR(INDEX('
                '[[Jobs.client]],MATCH({job}{r},[[Jobs.id]],0)),""),'
                '[[Clients.id]],0)),0)>0,INDEX([[Clients.rate]],'
                'MATCH(INDEX([[Jobs.client]],MATCH({job}{r},[[Jobs.id]],'
                '0)),[[Clients.id]],0)),HourlyRate),2))'),
            hidden=True,
            note="What these hours are worth at the client's rate."),
        Col("alert", "Check", 26, "status",
            formula=blank(
                "employee",
                'IF({hours}{r}=0,"Add - no hours recorded",'
                'IF({attendance}{r}="Absent - no show","Action - no show",'
                'IF({hours}{r}>10,"Check - over ten hours in a day",'
                'IF({approved}{r}<>"Yes","Due - needs approving",'
                'IF({paid}{r}<>"Yes","Due - not paid yet","OK")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Timesheets & Attendance",
        subtitle="Hours worked, what they cost you and how much of "
                 "them a client actually paid for.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).timesheets,
        help_text="One row per cleaner per job. Record the start and "
                  "finish times and the hours, the cost and the job "
                  "costing all follow.",
        extra_notes=[
            "Hours entered here override the estimate on the job, so "
            "the profit figures are based on what really happened "
            "rather than what you hoped.",
            "Mark travel, training and admin as not billable. The "
            "difference between paid hours and billed hours is the "
            "single biggest hidden cost in a cleaning business.",
            "This is a record of hours, not a payroll system. It "
            "will not calculate tax, and it is not a substitute for "
            "whatever your country requires you to keep.",
        ],
    )
