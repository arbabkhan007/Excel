"""`Payroll Tracker` - turning rota hours into a wage bill.

One row per person per pay run.  Hours come straight off the Staff
Schedule for the dates the run covers, split into basic and overtime,
and costed at that person's effective rate.  Salaried staff are paid
their slice of the year regardless of hours; contractors and
volunteers carry no employer on-costs at all.

`Total cost` is the only wage figure the profit and loss reads, which
is why the expense tracker deliberately skips its own wage categories
unless you tell it otherwise on Settings.  Paying twice for the same
staff is the easiest mistake to make in a small business spreadsheet
and the hardest to spot afterwards.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, id_col, lookup_staff,
                     lookup_staff_num, month_col, status_cf)

NAME = "Payroll Tracker"


def spec(cfg: Config) -> SheetSpec:
    sch_staff = "[[Staff Schedule.staff]]"
    sch_date = "[[Staff Schedule.date]]"

    def hours_in_period(column: str) -> str:
        return (f'ROUND(SUMIFS([[Staff Schedule.{column}]],'
                f'{sch_staff},{{staff}}{{r}},'
                f'{sch_date},">="&{{period_start}}{{r}},'
                f'{sch_date},"<="&{{period_end}}{{r}}),2)')

    columns = [
        id_col("period_end", "PayrollPrefix", digits=4, width=11,
               header="Pay run ID"),
        Col("period_end", "Period ends", 12, "date",
            validate=date_check(),
            note="The last day the run covers. Monthly totals group "
                 "on this date."),
        month_col("period_end"),
        Col("period_start", "Period starts", 12, "date",
            validate=date_check(),
            note="The first day. The hours below are added up between "
                 "these two dates, inclusive."),
        Col("pay_period", "Run type", 13, "text",
            validate=list_from("PayPeriodList"), cf=status_cf()),
        Col("staff", "Staff", 10, "id", validate=list_from("StaffList")),
        Col("staff_name", "Name", 18, "text",
            formula=lookup_staff("name")),
        Col("pay_type", "Paid as", 13, "text",
            formula=lookup_staff("pay_type")),
        Col("rate", "Rate", 10, "money",
            formula=lookup_staff_num("rate"),
            note="The effective hourly rate from the staff tab."),
        # --------------------------------------------------- the hours -
        Col("basic_hours", "Basic hours", 11, "number1",
            formula=blank("staff", hours_in_period("basic_hours")),
            cf=bar_cf(),
            note="Straight off the rota for the dates this run "
                 "covers. Nothing to add up by hand."),
        Col("ot_hours", "Overtime hours", 11, "number1",
            formula=blank("staff", hours_in_period("ot_hours")),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}]),
        Col("total_hours", "Total hours", 11, "number1",
            formula=blank(
                "staff",
                'ROUND({basic_hours}{r}+{ot_hours}{r},2)'),
            hidden=True),
        Col("leave_days", "Leave days", 9, "int",
            formula=blank(
                "staff",
                f'COUNTIFS({sch_staff},{{staff}}{{r}},'
                f'{sch_date},">="&{{period_start}}{{r}},'
                f'{sch_date},"<="&{{period_end}}{{r}},'
                f'[[Staff Schedule.shift]],"Leave")'),
            hidden=True),
        # ---------------------------------------------------- the pay --
        Col("periods_year", "Runs a year", 9, "int",
            formula=blank(
                "pay_period",
                'IF({pay_period}{r}="Weekly",52,'
                'IF({pay_period}{r}="Fortnightly",26,'
                'IF({pay_period}{r}="Four weekly",13,'
                'IF({pay_period}{r}="Twice monthly",24,'
                'IF({pay_period}{r}="Monthly",12,1)))))'),
            hidden=True),
        Col("basic_pay", "Basic pay", 11, "money",
            formula=blank(
                "staff",
                'ROUND(IF({pay_type}{r}="Salary",'
                'IFERROR(INDEX([[Staff Database.salary]],'
                'MATCH({staff}{r},[[Staff Database.id]],0))*1,0)/'
                'MAX(1,{periods_year}{r}),'
                '{basic_hours}{r}*{rate}{r}),2)'),
            note="Hourly staff are paid for the hours they worked. "
                 "Salaried staff get their slice of the year however "
                 "many hours the rota shows."),
        Col("ot_pay", "Overtime pay", 11, "money",
            formula=blank(
                "staff",
                'ROUND(IF({pay_type}{r}="Salary",0,'
                '{ot_hours}{r}*{rate}{r}*OvertimeMultiplier),2)')),
        Col("bonus", "Bonus", 10, "money",
            validate=number_check(0, 100000)),
        Col("adjust", "Adjustment", 11, "money",
            validate=number_check(-100000, 100000),
            note="Anything one-off: a backdated rise, a correction "
                 "from last run, an agreed deduction (type it "
                 "negative)."),
        Col("gross", "Gross pay", 12, "money",
            formula=blank(
                "staff",
                'ROUND({basic_pay}{r}+{ot_pay}{r}+'
                'IF({bonus}{r}="",0,{bonus}{r})+'
                'IF({adjust}{r}="",0,{adjust}{r}),2)'),
            cf=bar_cf(Palette.accent)),
        Col("employer_cost", "Employer costs", 12, "money",
            formula=blank(
                "staff",
                'ROUND(IF(({pay_type}{r}="Contractor")+'
                '({pay_type}{r}="Volunteer - unpaid")>0,0,'
                '{gross}{r}*PayrollBurden),2)'),
            note="Employer tax, insurance and holiday cover as a "
                 "percentage of gross - set it on Settings. "
                 "Contractors and volunteers carry none."),
        Col("total_cost", "Total cost", 12, "money",
            formula=blank(
                "staff",
                'ROUND({gross}{r}+{employer_cost}{r},2)'),
            cf=bar_cf(Palette.deep),
            note="What this person cost the business this run. The "
                 "only wage figure the profit and loss reads."),
        Col("deductions", "Deductions", 11, "money",
            validate=number_check(0, 100000), hidden=True,
            note="Tax and anything else withheld. For your own "
                 "records - this workbook does not calculate anybody's "
                 "tax."),
        Col("net_pay", "Net to pay", 12, "money",
            formula=blank(
                "staff",
                'ROUND({gross}{r}-IF({deductions}{r}="",0,'
                '{deductions}{r}),2)'),
            note="Gross less whatever you entered as deductions. It "
                 "is arithmetic, not a tax calculation."),
        Col("cost_hour", "Cost an hour", 10, "money",
            formula=blank(
                "staff",
                'IF({total_hours}{r}<=0,"",'
                'ROUND({total_cost}{r}/{total_hours}{r},2))'),
            hidden=True),
        # ------------------------------------------------------- paid --
        Col("paid", "Paid?", 8, "flag",
            validate=list_from("YesNoList"), cf=status_cf()),
        Col("paid_date", "Paid on", 11, "date", validate=date_check()),
        Col("method", "Method", 14, "text",
            validate=list_from("PaymentMethodList"), hidden=True),
        Col("notes", "Notes", 24, "wrap"),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "period_end",
                'IF({staff}{r}="","Add - nobody on this pay run",'
                'IF({staff_name}{r}="Not in Staff Database",'
                '"Check - staff ID not found",'
                'IF({period_start}{r}="","Add - no period start date",'
                'IF({period_start}{r}>{period_end}{r},'
                '"Check - the period runs backwards",'
                'IF(IF({pay_type}{r}="Salary",FALSE,'
                '{total_hours}{r}<=0),'
                '"Check - no rota hours in this period",'
                'IF(IF({pay_type}{r}="Volunteer - unpaid",FALSE,'
                '{gross}{r}<=0),"Check - gross pay is zero",'
                'IF(IF({paid}{r}="Yes",FALSE,'
                '{period_end}{r}<TodayDate-7),'
                '"Overdue - pay run not marked as paid",'
                'IF({paid}{r}<>"Yes","Due - not paid yet","OK")))))))'
                ')'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Payroll Tracker",
        subtitle="Hours from the rota, pay from the staff tab, "
                 "employer costs on top - one row per person per run.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.deep,
        freeze_cols=2,
        samples=sample(cfg).payroll,
        help_text="Set the two period dates and pick the person. The "
                  "hours, the overtime, the gross and the employer "
                  "costs all come from the rota and the staff record.",
        extra_notes=[
            "This is not a payroll bureau. It works out what staff "
            "cost you; it does not calculate anybody's income tax, "
            "social security or statutory deductions, and it files "
            "nothing with anybody.",
            "'Total cost' is gross plus employer on-costs, and it is "
            "the figure the profit and loss uses. Leave the wage "
            "categories out of the Expense Tracker, or switch "
            "'Payroll sits in expenses' to Yes on Settings - one or "
            "the other, never both.",
            "Salaried staff are paid their slice of the year whatever "
            "the rota says, so a quiet week does not quietly cut "
            "their pay.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Payroll summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Payroll summary", last)
    sheet.table_head(row, 0, [
        "Pay runs", "People paid", "Basic hours", "Overtime hours",
        "Gross pay", "Employer costs", "Total cost",
        "Average cost an hour", "This month cost", "Not yet paid"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Payroll Tracker.period_end]],"<>")',
        '=COUNTIFS([[Payroll Tracker.paid]],"Yes")',
        '=ROUND(SUM([[Payroll Tracker.basic_hours]]),1)',
        '=ROUND(SUM([[Payroll Tracker.ot_hours]]),1)',
        '=ROUND(SUM([[Payroll Tracker.gross]]),2)',
        '=ROUND(SUM([[Payroll Tracker.employer_cost]]),2)',
        '=ROUND(SUM([[Payroll Tracker.total_cost]]),2)',
        '=IF(SUM([[Payroll Tracker.total_hours]])<=0,"",'
        'ROUND(SUM([[Payroll Tracker.total_cost]])/'
        'SUM([[Payroll Tracker.total_hours]]),2))',
        '=ROUND(SUMIFS([[Payroll Tracker.total_cost]],'
        '[[Payroll Tracker.month]],ThisMonthKey),2)',
        '=ROUND(SUMIFS([[Payroll Tracker.net_pay]],'
        '[[Payroll Tracker.paid]],"<>Yes"),2)',
    ], ["int", "int", "number1", "number1", "money", "money", "money",
        "money", "money", "money"])
