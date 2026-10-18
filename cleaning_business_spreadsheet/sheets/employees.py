"""`Employees` - the team.

Who works for you, what they cost, what they are trained on and how
much of the work they are carrying. The hours, the pay and the revenue
each cleaner brings in are counted from the Timesheets and Jobs tabs,
so this tab is where you find out - without arguing about it - who is
overloaded and who is standing still.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, flag_yes_no_cf, id_col, rank_desc,
                     status_cf)

NAME = "Employees"

_T = "[[Timesheets.employee]]"
_J = "[[Jobs.employee]]"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "EmployeePrefix", digits=3, width=9),
        Col("name", "Name", 20, "text",
            note="The name that appears in every job and rota drop "
                 "down. Keep it exactly the same everywhere."),
        Col("role", "Role", 18, "text",
            validate=list_from("EmployeeRoleList")),
        Col("type", "Employment", 14, "text",
            validate=list_from("EmploymentTypeList"),
            note="Subcontractors are costed differently - they invoice "
                 "you rather than appearing on payroll."),
        Col("status", "Status", 12, "status",
            validate=list_from("EmployeeStatusList"), cf=status_cf()),
        Col("phone", "Phone", 15, "text"),
        Col("email", "Email", 26, "text"),
        Col("emergency", "Emergency contact", 22, "text",
            note="Worth having for anyone working alone in someone "
                 "else's house."),
        Col("start", "Started", 11, "date", validate=date_check()),
        Col("pay", "Pay an hour", 11, "money",
            validate=number_check(0, 1000),
            note="Blank uses the standard pay rate from Settings. The "
                 "true cost adds your on-costs on top."),
        Col("hours", "Contracted hours a week", 12, "number1",
            validate=number_check(0, 100),
            note="Used to work out how full their week is."),
        Col("skills", "Trained on", 24, "text",
            note="Deep cleans, carpet machine, out of hours, driving. "
                 "Check it before you assign the job."),
        Col("areas", "Areas covered", 18, "text"),
        Col("vehicle", "Vehicle", 12, "text",
            validate=list_from("VehicleList")),
        Col("licence", "Driver?", 8, "flag", validate=list_from("YesNoList"),
            cf=flag_yes_no_cf()),
        Col("checked", "Background check?", 11, "flag",
            validate=list_from("YesNoNAList"), cf=flag_yes_no_cf(),
            note="Essential for schools, nurseries and vulnerable "
                 "clients. The expiry goes on the Safety tab."),
        Col("insured", "Own insurance?", 11, "flag",
            validate=list_from("YesNoNAList"), cf=flag_yes_no_cf(),
            note="Subcontractors should have their own. Keep a copy."),
        Col("uniform", "Uniform issued", 11, "text"),
        Col("notes", "Notes", 34, "wrap"),
        # ------------------------------------------------ what they cost --
        Col("pay_rate", "Pay used", 11, "money",
            formula=blank("name", 'IF({pay}{r}="",PayRate,{pay}{r})'),
            hidden=True),
        Col("cost_rate", "True cost an hour", 12, "money",
            formula=blank(
                "name",
                'ROUND({pay_rate}{r}*IF({type}{r}="Subcontractor",1,'
                '1+EmployerOncost),2)'),
            note="Pay plus your on-costs - tax, insurance, holiday "
                 "pay. A subcontractor invoices you, so no on-costs."),
        # ------------------------------------------------- what they do --
        Col("jobs_assigned", "Jobs assigned", 11, "int",
            formula='=IF({name}{r}="","",COUNTIFS(' + _J + ',{name}{r}))'),
        Col("jobs_done", "Jobs done", 10, "int",
            formula='=IF({name}{r}="","",COUNTIFS(' + _J + ',{name}{r},'
                    '[[Jobs.status]],"Completed")+COUNTIFS(' + _J +
                    ',{name}{r},[[Jobs.status]],"Invoiced"))'),
        Col("upcoming", "Jobs ahead", 10, "int",
            formula='=IF({name}{r}="","",COUNTIFS(' + _J + ',{name}{r},'
                    '[[Jobs.date]],">="&TodayDate,[[Jobs.status]],'
                    '"<>Cancelled"))',
            cf=[{"type": "data_bar", "bar_color": "#2F8F9D"}]),
        Col("week_hours", "Hours booked this week", 12, "number1",
            formula='=IF({name}{r}="","",SUMIFS([[Jobs.hours_each]],' + _J +
                    ',{name}{r},[[Jobs.date]],">="&WeekStartDate,'
                    '[[Jobs.date]],"<="&WeekStartDate+6,[[Jobs.status]],'
                    '"<>Cancelled"))',
            note="Straight from the diary. Compare it with the "
                 "contracted hours in the next column.",
            cf=[{"type": "data_bar", "bar_color": "#E8A33D"}]),
        Col("load", "Week full", 9, "percent0",
            formula='=IF(OR({name}{r}="",{hours}{r}="",{hours}{r}=0),"",'
                    '{week_hours}{r}/{hours}{r})',
            cf=[{"type": "cell", "criteria": ">", "value": 1,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": 0.5,
                 "fmt": "warn"}],
            note="Over 100% means you have booked them for more than "
                 "they are contracted to work."),
        Col("paid_hours", "Hours paid (all time)", 12, "number1",
            formula='=IF({name}{r}="","",SUMIFS([[Timesheets.hours]],' +
                    _T + ',{name}{r}))'),
        Col("billed_hours", "Hours billed", 11, "number1",
            formula='=IF({name}{r}="","",SUMIFS([[Timesheets.hours]],' +
                    _T + ',{name}{r},[[Timesheets.billable]],"Yes"))'),
        Col("utilisation", "Utilisation", 10, "percent0",
            formula='=IF(OR({name}{r}="",{paid_hours}{r}=0),"",'
                    '{billed_hours}{r}/{paid_hours}{r})',
            cf=[{"type": "cell", "criteria": ">=",
                 "value": "=UtilisationTarget", "fmt": "good"},
                {"type": "cell", "criteria": "<",
                 "value": "=UtilisationTarget", "fmt": "warn"}],
            note="The share of the hours you pay for that a client "
                 "actually pays for. Travel, setup and waiting are the "
                 "usual culprits."),
        Col("wage_cost", "Wages paid", 12, "money",
            formula='=IF({name}{r}="","",SUMIFS([[Timesheets.cost]],' + _T +
                    ',{name}{r}))'),
        Col("revenue", "Revenue brought in", 13, "money",
            formula='=IF({name}{r}="","",SUMIFS([[Jobs.price]],' + _J +
                    ',{name}{r},[[Jobs.status]],"Completed")+SUMIFS('
                    '[[Jobs.price]],' + _J + ',{name}{r},[[Jobs.status]],'
                    '"Invoiced"))'),
        Col("contribution", "Contribution", 12, "money",
            formula='=IF({name}{r}="","",{revenue}{r}-{wage_cost}{r})',
            note="Revenue from their jobs less their wages. Not a "
                 "complete picture - overheads sit on top - but it is "
                 "the right shape.",
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}]),
        Col("rating", "Average rating", 11, "number1",
            formula='=IF({name}{r}="","",IFERROR(AVERAGEIFS('
                    '[[Jobs.rating]],' + _J + ',{name}{r},[[Jobs.rating]],'
                    '">0"),""))',
            cf=[{"type": "cell", "criteria": "<", "value": 3.5,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">=", "value": 4.5,
                 "fmt": "good"}]),
        Col("complaints", "Complaints", 10, "int",
            formula='=IF({name}{r}="","",COUNTIFS('
                    '[[Complaints.employee]],{name}{r}))',
            cf=[{"type": "cell", "criteria": ">=", "value": 2,
                 "fmt": "warn"}]),
        Col("late", "Late or absent", 11, "int",
            formula='=IF({name}{r}="","",COUNTIFS('
                    '[[Timesheets.attendance]],"Late",' + _T + ',{name}{r})'
                    '+COUNTIFS([[Timesheets.attendance]],'
                    '"Absent - no show",' + _T + ',{name}{r}))',
            cf=[{"type": "cell", "criteria": ">=", "value": 3,
                 "fmt": "warn"}]),
        Col("name_rank", "List order", 8, "int",
            formula='=IF({name}{r}="","",COUNTIF({name_first}:'
                    '{name}{r},"?*"))',
            hidden=True,
            note="1, 2, 3 down the filled rows. The rota reads it."),
        rank_desc("revenue", guard='{name}{r}=""',
                  note="1 = brings in the most revenue."),
        rank_desc("week_hours", guard='{name}{r}=""',
                  note="1 = the fullest week. Used by the rota."),
        Col("alert", "Check", 26, "status",
            formula=blank(
                "name",
                'IF({status}{r}="Left","Archived",'
                'IF(AND({load}{r}<>"",{load}{r}>1),'
                '"Overloaded - over contracted hours",'
                'IF({complaints}{r}>=2,"Check - complaints logged",'
                'IF(AND({rating}{r}<>"",{rating}{r}<3.5),'
                '"Check - low client ratings",'
                'IF({checked}{r}="No","Action - no background check",'
                'IF(AND({upcoming}{r}=0,{status}{r}="Active"),'
                '"Idle - nothing booked","OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Employee Database",
        subtitle="Your team, what they cost, what they bring in and "
                 "how full their week is.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).employees,
        help_text="Add everyone who cleans for you, including yourself "
                  "and any subcontractors. Their names fill the drop "
                  "downs on the Jobs, rota and timesheet tabs.",
        extra_notes=[
            "'True cost an hour' is the number to quote from, not the "
            "pay rate. An $18 cleaner on 18% on-costs really costs you "
            "$21.24 before they have put a cloth in a bucket.",
            "Utilisation is the quiet killer in a cleaning business. "
            "Paying for eight hours and billing five is a 37% hole in "
            "your margin, and it never shows up on a bank statement.",
            "Employee records are personal data. Only keep what you "
            "actually need, and keep the file somewhere only you can "
            "get at.",
        ],
    )
