"""`Staff Database` - the team, their pay and their certificates.

Childcare is the one small business where a lapsed certificate is not
an administrative annoyance but a reason you cannot open the room.
First aid, CPR, safeguarding training and background checks all have
dates, and all four are watched here: the soonest expiry across the
lot is calculated into one column, and the dashboard counts whatever
is about to run out.

Pay lives here too, as a rate rather than an amount, so the Staff
Schedule can turn hours into money and the Payroll Tracker can turn
money into a cost.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, id_col, status_cf)

NAME = "Staff Database"


def spec(cfg: Config) -> SheetSpec:
    doc_who = "[[Document Tracker.who]]"
    doc_req = "[[Document Tracker.required]]"
    doc_status = "[[Document Tracker.status]]"

    columns = [
        id_col("name", "StaffPrefix", digits=3, width=10,
               header="Staff ID",
               note="Automatic. The rota, the payroll, the attendance "
                    "register and every incident form point at it."),
        Col("name", "Name", 22, "text"),
        Col("role", "Role", 20, "text", validate=list_from("RoleList"),
            cf=status_cf()),
        Col("room", "Main room", 16, "text",
            validate=list_from("ClassroomList"),
            note="Where they usually are. The rota can send them "
                 "anywhere - this is just their base."),
        Col("phone", "Phone", 15, "text"),
        Col("email", "Email", 24, "text"),
        Col("address", "Address", 28, "text", hidden=True),
        Col("dob", "Date of birth", 11, "date", validate=date_check(),
            hidden=True),
        Col("hire_date", "Started", 11, "date", validate=date_check()),
        Col("status", "Status", 14, "status",
            validate=list_from("EmploymentStatusList"), cf=status_cf(),
            note="Only Active and Probation staff are counted as "
                 "available on the dashboard."),
        Col("end_date", "Left on", 11, "date", validate=date_check(),
            hidden=True),
        Col("years", "Years here", 9, "number1",
            formula=blank(
                "hire_date",
                'ROUND((IF({end_date}{r}="",TodayDate,{end_date}{r})-'
                '{hire_date}{r})/365,1)'),
            hidden=True),
        # ---------------------------------------------------- the pay ---
        Col("pay_type", "Paid as", 13, "text",
            validate=list_from("PayTypeList"),
            note="Hourly, salaried, daily or contractor. Payroll "
                 "handles all four differently, and volunteers cost "
                 "nothing at all."),
        Col("hourly", "Hourly rate", 11, "money",
            validate=number_check(0, 1000),
            note="For hourly staff. Leave blank for salaried ones - "
                 "the effective rate below works theirs out."),
        Col("salary", "Salary a year", 13, "money",
            validate=number_check(0, 1000000)),
        Col("contract_hours", "Hours a week", 11, "number1",
            validate=number_check(0, 80),
            note="Contracted hours. Used to turn a salary into an "
                 "hourly rate and to spot anyone being over-rota'd."),
        Col("day_rate", "Day rate", 11, "money",
            validate=number_check(0, 10000), hidden=True),
        Col("rate", "Effective hourly", 12, "money",
            formula=blank(
                "name",
                'ROUND(IF({pay_type}{r}="Volunteer - unpaid",0,'
                'IF({pay_type}{r}="Hourly",{hourly}{r},'
                'IF({pay_type}{r}="Salary",'
                'IF({contract_hours}{r}>0,'
                '{salary}{r}/(WeeksPaid*{contract_hours}{r}),0),'
                'IF({pay_type}{r}="Daily rate",'
                'IF({contract_hours}{r}>0,'
                '{day_rate}{r}/({contract_hours}{r}/5),0),'
                '{hourly}{r})))),2)'),
            cf=bar_cf(Palette.accent),
            note="One comparable hourly figure however somebody is "
                 "paid. The rota costs its hours at this rate."),
        Col("cost_hour", "Cost an hour", 11, "money",
            formula=blank(
                "name",
                'ROUND({rate}{r}*IF(({pay_type}{r}="Contractor")+'
                '({pay_type}{r}="Volunteer - unpaid")>0,1,'
                '1+PayrollBurden),2)'),
            hidden=True,
            note="The rate plus employer costs. Contractors and "
                 "volunteers carry none."),
        # ------------------------------------------------ certificates --
        Col("first_aid_exp", "First aid expires", 12, "date",
            validate=date_check()),
        Col("cpr_exp", "CPR expires", 12, "date",
            validate=date_check()),
        Col("safeguard_exp", "Safeguarding expires", 13, "date",
            validate=date_check()),
        Col("check_exp", "Background check expires", 14, "date",
            validate=date_check(),
            note="Whatever the check is called where you are - DBS, "
                 "Garda vetting, state background check. Record when "
                 "it needs renewing."),
        Col("check_done", "Check completed", 12, "date",
            validate=date_check(), hidden=True),
        Col("qualification", "Qualification", 22, "text",
            note="Level and subject. The thing an inspector asks to "
                 "see the certificate for."),
        Col("training", "Other training", 26, "wrap", hidden=True),
        Col("next_expiry", "Next expiry", 12, "date",
            formula=blank(
                "name",
                'IF(({first_aid_exp}{r}="")*({cpr_exp}{r}="")*'
                '({safeguard_exp}{r}="")*({check_exp}{r}="")=1,"",'
                'MIN(IF({first_aid_exp}{r}="",99999,'
                '{first_aid_exp}{r}),'
                'IF({cpr_exp}{r}="",99999,{cpr_exp}{r}),'
                'IF({safeguard_exp}{r}="",99999,{safeguard_exp}{r}),'
                'IF({check_exp}{r}="",99999,{check_exp}{r})))'),
            note="The soonest of the four. One column to filter on "
                 "instead of four to remember."),
        Col("days_to_expiry", "Days left", 9, "int",
            formula=blank(
                "next_expiry", '{next_expiry}{r}-TodayDate'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bold_bad"},
                {"type": "cell", "criteria": "<=",
                 "value": "=DocUrgentDays", "fmt": "bad"},
                {"type": "cell", "criteria": "<=",
                 "value": "=DocWarningDays", "fmt": "warn"}]),
        Col("certs_ok", "Certificates", 13, "status",
            formula=blank(
                "name",
                'IF({next_expiry}{r}="","Nothing recorded",'
                'IF({days_to_expiry}{r}<0,"Expired",'
                'IF({days_to_expiry}{r}<=DocUrgentDays,"Expiring",'
                'IF({days_to_expiry}{r}<=DocWarningDays,"Due",'
                '"On file"))))'),
            cf=status_cf()),
        Col("docs_missing", "Docs missing", 10, "int",
            formula=blank(
                "name",
                f'COUNTIFS({doc_who},{{id}}{{r}},{doc_req},"Yes")-'
                f'COUNTIFS({doc_who},{{id}}{{r}},{doc_req},"Yes",'
                f'{doc_status},"On file")-'
                f'COUNTIFS({doc_who},{{id}}{{r}},{doc_req},"Yes",'
                f'{doc_status},"Signed")-'
                f'COUNTIFS({doc_who},{{id}}{{r}},{doc_req},"Yes",'
                f'{doc_status},"Received")'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "bad"}]),
        # ------------------------------------------------- their week ---
        Col("hours_week", "Rota'd this week", 12, "number1",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Staff Schedule.hours]],'
                '[[Staff Schedule.staff]],{id}{r},'
                '[[Staff Schedule.week_start]],'
                'TodayDate-WEEKDAY(TodayDate,3)),2)'),
            note="Hours on the rota for the week that contains today."),
        Col("hours_month", "Hours this month", 12, "number1",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Staff Schedule.actual_hours]],'
                '[[Staff Schedule.staff]],{id}{r},'
                '[[Staff Schedule.month]],ThisMonthKey),2)'),
            hidden=True),
        Col("over_contract", "Over contract", 10, "number1",
            formula=blank(
                "name",
                'IF({contract_hours}{r}="","",'
                'ROUND(MAX(0,{hours_week}{r}-{contract_hours}{r}),2))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}],
            hidden=True),
        Col("pay_month", "Paid this month", 12, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Payroll Tracker.gross]],'
                '[[Payroll Tracker.staff]],{id}{r},'
                '[[Payroll Tracker.month]],ThisMonthKey),2)'),
            hidden=True),
        Col("children_room", "Children in room", 11, "int",
            formula=blank(
                "room",
                'COUNTIFS([[Child Enrollment.room]],{room}{r},'
                '[[Child Enrollment.holds_place]],"Yes")'),
            hidden=True),
        Col("leave_days", "Leave taken", 10, "number1",
            formula=blank(
                "name",
                'COUNTIFS([[Staff Schedule.staff]],{id}{r},'
                '[[Staff Schedule.shift]],"Leave")'),
            hidden=True),
        Col("em_name", "Emergency contact", 20, "text"),
        Col("em_phone", "Emergency phone", 15, "text"),
        Col("notes", "Notes", 26, "wrap"),
        Col("alert", "Action", 34, "status",
            formula=blank(
                "name",
                'IF(({status}{r}="Left")+({status}{r}="Inactive")>0,"",'
                'IF({next_expiry}{r}="","Missing - no certificate '
                'dates recorded",'
                'IF({days_to_expiry}{r}<0,"Expired - "&'
                'TEXT({next_expiry}{r},"d mmm")&", they cannot be '
                'counted in ratio",'
                'IF({days_to_expiry}{r}<=DocUrgentDays,'
                '"Expiring - renew within "&{days_to_expiry}{r}&" days",'
                'IF({docs_missing}{r}>0,"Missing - "&'
                '{docs_missing}{r}&" document(s) outstanding",'
                'IF({pay_type}{r}="","Add - how are they paid?",'
                'IF(IF({rate}{r}=0,{pay_type}{r}<>"Volunteer - '
                'unpaid",FALSE),"Add - no pay rate set",'
                'IF({days_to_expiry}{r}<=DocWarningDays,'
                '"Due - book the renewal",'
                'IF({em_phone}{r}="","Missing - no emergency contact",'
                '"OK")))))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Staff & Employee Database",
        subtitle="The team, what they are paid, and exactly how many "
                 "days until a certificate runs out.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.deep,
        freeze_cols=2,
        samples=sample(cfg).staff,
        help_text="One row per person. Fill in the pay rate and the "
                  "four certificate dates - the rota, the payroll and "
                  "the renewal warnings all run off them.",
        extra_notes=[
            "'Effective hourly' turns a salary, a day rate or an "
            "hourly wage into one comparable number, so the rota can "
            "cost a shift however that person happens to be paid.",
            "'Next expiry' is the soonest of first aid, CPR, "
            "safeguarding and the background check. Sort by it once "
            "a month and you will never be caught out.",
            "Employer on-costs are added on top of the rate for "
            "everyone except contractors and volunteers. Set the "
            "percentage on the Settings tab.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Team summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "The team", last)
    sheet.table_head(row, 0, [
        "On the books", "Active", "On leave", "Average hourly rate",
        "Average cost an hour", "Contracted hours a week",
        "Certificates expired", "Expiring soon", "Documents missing",
        "Average years here"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Staff Database.name]],"<>")',
        '=COUNTIFS([[Staff Database.status]],"Active")+'
        'COUNTIFS([[Staff Database.status]],"Probation")',
        '=COUNTIFS([[Staff Database.status]],"On leave")',
        '=IFERROR(ROUND(AVERAGEIF([[Staff Database.rate]],">0"),2),0)',
        '=IFERROR(ROUND(AVERAGEIF([[Staff Database.cost_hour]],">0"),'
        '2),0)',
        '=ROUND(SUMIFS([[Staff Database.contract_hours]],'
        '[[Staff Database.status]],"Active"),1)',
        '=COUNTIFS([[Staff Database.certs_ok]],"Expired")',
        '=COUNTIFS([[Staff Database.certs_ok]],"Expiring")+'
        'COUNTIFS([[Staff Database.certs_ok]],"Due")',
        '=SUMIFS([[Staff Database.docs_missing]],'
        '[[Staff Database.status]],"Active")',
        '=IFERROR(ROUND(AVERAGEIF([[Staff Database.years]],">0"),1),0)',
    ], ["int", "int", "int", "money", "money", "number1", "int", "int",
        "int", "number1"])
