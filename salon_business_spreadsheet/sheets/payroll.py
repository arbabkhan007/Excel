"""`Payroll` - what each pay run actually costs the salon.

Hours at a rate, plus the commission you are releasing this run, plus
the tips you are passing on, plus the employer costs on top. The last
column is the one that belongs in the profit and loss: the total cost
of employing somebody for that period, not the number on their payslip.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, lookup_staff, month_col,
                     week_col, yesno)

NAME = "Payroll"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("period_end", "PayrollPrefix", digits=4, width=11),
        Col("period_end", "Period ending", 13, "date",
            validate=date_check(),
            note="The last day of the pay period. Everything on this "
                 "tab is grouped by the month this date falls in."),
        Col("period", "Run", 12, "text", validate=list_from("PayPeriodList"),
            note="Weekly, fortnightly, monthly, four weekly or one "
                 "off - label only, so your runs can be whatever "
                 "shape they really are."),
        Col("staff", "Stylist ID", 11, "id",
            validate=list_from("StaffList")),
        Col("staff_name", "Name", 18, "text",
            formula=lookup_staff("name")),
        Col("hours", "Hours", 9, "number1",
            validate=number_check(0, 400),
            note="Hours actually worked in this period."),
        Col("rate", "Hourly rate", 11, "money",
            validate=number_check(0, 1000),
            note="Leave blank to use their rate from the Staff tab, "
                 "or the salon default if they have none."),
        Col("commission", "Commission paid", 13, "money",
            validate=number_check(-100000, 100000),
            note="What you are releasing this run. The column beside "
                 "it shows what the commissions tab calculated for "
                 "the whole month, so you can see what is left."),
        Col("tips", "Tips passed on", 12, "money",
            validate=number_check(0, 100000),
            note="Tips handed over with this run. They are the "
                 "stylist's money passing through, so they are kept "
                 "out of the salon's wage cost."),
        Col("bonus", "Bonus", 10, "money",
            validate=number_check(0, 100000)),
        Col("holiday", "Holiday pay", 11, "money",
            validate=number_check(0, 100000)),
        Col("deductions", "Deductions", 11, "money",
            validate=number_check(0, 100000),
            note="Anything taken off before it is paid - a product "
                 "account, a uniform, an advance."),
        Col("paid", "Paid?", 9, "flag", validate=yesno()),
        Col("paid_on", "Paid on", 11, "date", validate=date_check()),
        Col("method", "Method", 14, "text",
            validate=list_from("PaymentMethodList")),
        Col("notes", "Notes", 22, "wrap"),
        # ---------------------------------------------------- worked out --
        month_col("period_end"),
        week_col("period_end"),
        Col("rate_used", "Rate used", 10, "money",
            formula=blank(
                "period_end",
                'IF({rate}{r}<>"",{rate}{r},IFERROR(IF(INDEX('
                '[[Staff.hourly]],MATCH({staff}{r},[[Staff.id]],0))="",'
                'DefaultHourly,INDEX([[Staff.hourly]],MATCH({staff}{r},'
                '[[Staff.id]],0))),DefaultHourly))'),
            hidden=True),
        Col("base_pay", "Hourly pay", 11, "money",
            formula=blank(
                "period_end",
                'ROUND(IF({hours}{r}="",0,{hours}{r})*'
                '{rate_used}{r},2)')),
        Col("gross_pay", "Gross pay", 11, "money",
            formula=blank(
                "period_end",
                'ROUND({base_pay}{r}+IF({commission}{r}="",0,'
                '{commission}{r})+IF({bonus}{r}="",0,{bonus}{r})+'
                'IF({holiday}{r}="",0,{holiday}{r}),2)'),
            cf=[{"type": "data_bar", "bar_color": Palette.primary_mid}],
            note="Hourly pay plus commission, bonus and holiday pay. "
                 "Tips are not in here - they were never yours."),
        Col("net_pay", "After deductions", 13, "money",
            formula=blank(
                "period_end",
                'ROUND({gross_pay}{r}-IF({deductions}{r}="",0,'
                '{deductions}{r}),2)'),
            note="Before income tax and any statutory deductions - "
                 "this workbook does not calculate those."),
        Col("to_hand_over", "Including tips", 12, "money",
            formula=blank(
                "period_end",
                'ROUND({net_pay}{r}+IF({tips}{r}="",0,{tips}{r}),2)'),
            hidden=True,
            note="What physically leaves you on pay day, tips "
                 "included."),
        Col("employer_cost", "Employer costs", 12, "money",
            formula=blank(
                "period_end",
                'IF(IFERROR(INDEX([[Staff.type]],MATCH({staff}{r},'
                '[[Staff.id]],0))&"","")="Self-employed",0,'
                'IF(IFERROR(INDEX([[Staff.type]],MATCH({staff}{r},'
                '[[Staff.id]],0))&"","")="Chair renter",0,'
                'ROUND({gross_pay}{r}*PayrollBurden,2)))'),
            note="Employer tax, insurance and holiday accrual at the "
                 "percentage on Settings. Self-employed stylists and "
                 "chair renters carry none."),
        Col("total_cost", "Cost to salon", 12, "money",
            formula=blank(
                "period_end",
                'ROUND({gross_pay}{r}+{employer_cost}{r},2)'),
            cf=[{"type": "data_bar", "bar_color": Palette.accent}],
            note="Gross pay plus employer costs. This is the figure "
                 "the profit and loss uses for wages."),
        Col("month_commission", "Month calculated", 13, "money",
            formula=blank(
                "period_end",
                'ROUND(SUMIFS([[Commissions.payout]],'
                '[[Commissions.staff]],{staff}{r},'
                '[[Commissions.month]],{month}{r}),2)'),
            hidden=True),
        Col("month_paid", "Month released", 12, "money",
            formula=blank(
                "period_end",
                'ROUND(SUMIFS({commission_first}:{commission_last},'
                '{staff_first}:{staff_last},{staff}{r},'
                '{month_first}:{month_last},{month}{r}),2)'),
            hidden=True),
        Col("commission_left", "Commission left", 13, "money",
            formula=blank(
                "period_end",
                'ROUND({month_commission}{r}-{month_paid}{r},2)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0.004,
                 "fmt": "warn"},
                {"type": "cell", "criteria": "<", "value": -0.004,
                 "fmt": "bad"}],
            note="This month's calculated commission less everything "
                 "already released in the month's pay runs. Zero means "
                 "you are square; negative means you have overpaid."),
        Col("cost_hour", "Cost an hour", 11, "money",
            formula=blank(
                "period_end",
                'IF(IF({hours}{r}="",0,{hours}{r})=0,"",'
                'ROUND({total_cost}{r}/{hours}{r},2))'),
            hidden=True),
        Col("alert", "Check", 28, "status",
            formula=blank(
                "period_end",
                'IF({staff}{r}="","Add - who is this run for?",'
                'IF(IF({hours}{r}="",{commission}{r}="",FALSE),'
                '"Add - no hours and no commission",'
                'IF({commission_left}{r}<-0.004,'
                '"Check - more commission paid than earned",'
                'IF({paid}{r}<>"Yes","Due - not paid yet",'
                'IF({paid_on}{r}="","Check - marked paid, no date",'
                '"OK")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Payroll & Tips",
        subtitle="Hours, commission released, tips passed on and the "
                 "employer costs on top - the real cost of a pay run.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.deep,
        freeze_cols=2,
        samples=sample(cfg).payroll,
        help_text="One row per person per pay run. Hours and the "
                  "commission you are releasing are the only numbers "
                  "to type.",
        extra_notes=[
            "Commission is typed here rather than pulled in "
            "automatically, because a monthly commission is often "
            "paid across two or three weekly runs. The 'commission "
            "left' column keeps the score so nobody is paid twice.",
            "Tips pass straight through. They are added to what you "
            "hand over and left out of the wage cost, because they "
            "were never the salon's money.",
            "This is a cost tracker, not a payroll system. It does "
            "not calculate income tax, national insurance or social "
            "security - your payroll provider or accountant does "
            "that, and these totals are what you give them.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Payroll summary under the grid."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Pay runs, all time", last)
    sheet.table_head(row, 0, [
        "Runs logged", "Hours", "Hourly pay", "Commission released",
        "Bonuses", "Gross pay", "Employer costs", "Cost to salon",
        "Tips passed on", "Runs not paid"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Payroll.period_end]],"<>")',
        '=ROUND(SUM([[Payroll.hours]]),1)',
        '=ROUND(SUM([[Payroll.base_pay]]),2)',
        '=ROUND(SUM([[Payroll.commission]]),2)',
        '=ROUND(SUM([[Payroll.bonus]]),2)',
        '=ROUND(SUM([[Payroll.gross_pay]]),2)',
        '=ROUND(SUM([[Payroll.employer_cost]]),2)',
        '=ROUND(SUM([[Payroll.total_cost]]),2)',
        '=ROUND(SUM([[Payroll.tips]]),2)',
        '=COUNTIFS([[Payroll.paid]],"<>Yes",'
        '[[Payroll.period_end]],"<>")',
    ], ["int", "number1", "money", "money", "money", "money", "money",
        "money", "money", "int"])
