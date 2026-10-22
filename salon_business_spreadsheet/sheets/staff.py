"""`Staff` - the team, and what each chair is actually bringing in.

Hours available against hours booked gives utilisation; revenue
against hours gives the rate per chair hour. Those two numbers,
side by side, are the whole argument for or against taking somebody
on, moving a shift, or putting the prices up.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, rank_desc, status_cf)

NAME = "Staff"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "StaffPrefix", digits=3, width=10,
               note="Automatic. Appointments, sales, commissions and "
                    "payroll all point at this ID."),
        Col("name", "Name", 20, "text"),
        Col("role", "Role", 18, "text",
            note="Stylist, colourist, nail tech, lash artist, front "
                 "desk, apprentice - whatever you call it."),
        Col("level", "Level", 14, "text",
            validate=list_from("StylistLevelList")),
        Col("type", "Employment", 15, "text",
            validate=list_from("StaffTypeList"),
            note="Employed, self-employed, chair renter, apprentice, "
                 "owner or contractor. Chair renters pay you rent "
                 "instead of earning commission."),
        Col("start_date", "Started", 11, "date", validate=date_check()),
        Col("status", "Status", 11, "text",
            validate=list_from("StaffStatusList"), cf=status_cf()),
        Col("phone", "Phone", 15, "text"),
        Col("email", "Email", 24, "text"),
        Col("hours_week", "Hours a week", 12, "number1",
            validate=number_check(0, 100),
            note="Contracted or usual hours on the floor. The "
                 "utilisation figure divides booked hours by these."),
        Col("days", "Days worked", 18, "text",
            note="Mon, Tue, Thu, Fri, Sat - free text, for the rota at "
                 "a glance."),
        Col("method", "Commission method", 24, "text",
            validate=list_from("CommissionMethodList"),
            note="Yours to choose, per person: none, percent of "
                 "service, service and retail, percent after product "
                 "cost, tiered, flat per service, hourly only, or "
                 "rent a chair. The commissions tab follows whatever "
                 "you pick here."),
        Col("rate_service", "Service rate", 11, "percent",
            validate=number_check(0, 1),
            note="Leave blank to use the salon default on Settings."),
        Col("rate_retail", "Retail rate", 11, "percent",
            validate=number_check(0, 1)),
        Col("flat_rate", "Flat per service", 12, "money",
            validate=number_check(0, 1000),
            note="Only used by the 'flat amount per service' method."),
        Col("hourly", "Hourly rate", 11, "money",
            validate=number_check(0, 1000),
            note="Blank uses the default hourly rate on Settings."),
        Col("rent_week", "Chair rent", 11, "money",
            validate=number_check(0, 10000),
            note="Weekly rent for a chair renter. It is income to the "
                 "salon, not a cost - log the money itself on the "
                 "Sales tab as kind Other."),
        Col("target_month", "Monthly target", 13, "money",
            validate=number_check(0, 1000000),
            note="What you expect this person to bill in a month."),
        Col("notes", "Notes", 24, "wrap"),
        # -------------------------------------------------- performance ---
        Col("appts", "Appointments", 12, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Appointments.staff]],{id}{r})'),
            hidden=True),
        Col("done", "Completed", 10, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Appointments.staff]],{id}{r},'
                '[[Appointments.status]],"Completed")')),
        Col("no_shows", "No shows", 9, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Appointments.staff]],{id}{r},'
                '[[Appointments.status]],"No show")'),
            hidden=True),
        Col("clients", "Clients seen", 11, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Clients.staff]],{id}{r})'),
            hidden=True,
            note="Clients who have this stylist down as their usual."),
        Col("hours_booked", "Hours booked", 12, "number1",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Appointments.hours]],'
                '[[Appointments.staff]],{id}{r},'
                '[[Appointments.status]],"Completed",'
                '[[Appointments.date]],">="&TodayDate-28),1)'),
            note="Chair hours actually worked in the last four weeks."),
        Col("hours_avail", "Hours available", 12, "number1",
            formula=blank(
                "name",
                'IF({hours_week}{r}="",0,ROUND({hours_week}{r}*4,1))'),
            hidden=True),
        Col("utilisation", "Utilisation", 11, "percent",
            formula=blank(
                "name",
                'IF({hours_avail}{r}=0,"",'
                '{hours_booked}{r}/{hours_avail}{r})'),
            cf=[{"type": "cell", "criteria": ">=",
                 "value": "=TargetUtilisation", "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0.5,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<",
                 "value": "=TargetUtilisation", "fmt": "warn"}],
            note="Hours booked against hours available, last four "
                 "weeks. Below half means an empty chair you are "
                 "paying for."),
        Col("rev_service", "Service sales", 12, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.net]],[[Sales.staff]],{id}{r},'
                '[[Sales.kind]],"Service"),2)'),
            cf=[{"type": "data_bar", "bar_color": Palette.primary_mid}]),
        Col("rev_retail", "Retail sales", 11, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.net]],[[Sales.staff]],{id}{r},'
                '[[Sales.kind]],"Retail"),2)')),
        Col("rev_total", "Total sales", 12, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.net]],[[Sales.staff]],'
                '{id}{r}),2)')),
        Col("rev_month", "Sales this month", 13, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.net]],[[Sales.staff]],{id}{r},'
                '[[Sales.month]],ThisMonthKey),2)'),
            cf=[{"type": "data_bar", "bar_color": Palette.accent}]),
        Col("target_gap", "Against target", 12, "money",
            formula=blank(
                "name",
                'IF({target_month}{r}="","",'
                'ROUND({rev_month}{r}-{target_month}{r},2))'),
            cf=[{"type": "cell", "criteria": ">=", "value": 0,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "warn"}]),
        Col("avg_ticket", "Average ticket", 12, "money",
            formula=blank(
                "name",
                'IF({done}{r}=0,"",ROUND({rev_total}{r}/{done}{r},2))'),
            cf=[{"type": "cell", "criteria": ">=", "value": "=TargetTicket",
                 "fmt": "good"}]),
        Col("retail_share", "Retail share", 11, "percent",
            formula=blank(
                "name",
                'IF({rev_total}{r}=0,"",'
                '{rev_retail}{r}/{rev_total}{r})'),
            cf=[{"type": "cell", "criteria": ">=",
                 "value": "=TargetRetailShare", "fmt": "good"},
                {"type": "cell", "criteria": "<",
                 "value": "=TargetRetailShare", "fmt": "warn"}],
            note="How much of this stylist's money comes from product. "
                 "The easiest number in the salon to improve."),
        Col("rate_hour", "Earned an hour", 12, "money",
            formula=blank(
                "name",
                'IF({hours_booked}{r}=0,"",ROUND(SUMIFS('
                '[[Sales.net]],[[Sales.staff]],{id}{r},'
                '[[Sales.date]],">="&TodayDate-28)/'
                '{hours_booked}{r},2))'),
            cf=[{"type": "cell", "criteria": ">=",
                 "value": "=TargetChairRate", "fmt": "good"},
                {"type": "cell", "criteria": "<",
                 "value": "=TargetChairRate", "fmt": "warn"}]),
        Col("rebook_rate", "Rebook rate", 11, "percent",
            formula=blank(
                "name",
                'IF(COUNTIFS([[Appointments.staff]],{id}{r},'
                '[[Appointments.status]],"Completed")=0,"",'
                'COUNTIFS([[Appointments.staff]],{id}{r},'
                '[[Appointments.rebooked]],"Yes")/'
                'COUNTIFS([[Appointments.staff]],{id}{r},'
                '[[Appointments.status]],"Completed"))'),
            cf=[{"type": "cell", "criteria": ">=",
                 "value": "=TargetRebook", "fmt": "good"},
                {"type": "cell", "criteria": "<",
                 "value": "=TargetRebook", "fmt": "warn"}],
            note="How often their clients leave with the next "
                 "appointment already in the book."),
        Col("tips", "Tips", 10, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Payments.amount]],[[Payments.staff]],'
                '{id}{r},[[Payments.kind]],"Tip"),2)'),
            hidden=True),
        Col("commission_paid", "Commission", 12, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Commissions.commission]],'
                '[[Commissions.staff]],{id}{r}),2)'),
            hidden=True,
            note="Everything calculated for this person on the "
                 "commissions tab."),
        Col("cost_month", "Pay this month", 12, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Payroll.total_cost]],[[Payroll.staff]],'
                '{id}{r},[[Payroll.month]],ThisMonthKey),2)'),
            note="From the payroll tab - wages, commission and the "
                 "employer costs on top."),
        Col("margin_month", "Leaves behind", 12, "money",
            formula=blank(
                "name",
                'ROUND({rev_month}{r}-{cost_month}{r},2)'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "good"}],
            note="This month's sales less this month's pay. What is "
                 "left to cover the rent and everything else."),
        rank_desc("rev_total", header="Sales rank"),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "name",
                'IF({status}{r}<>"Active","",'
                'IF({method}{r}="","Add - pick a commission method",'
                'IF(IF({hours_week}{r}="",TRUE,{hours_week}{r}=0),'
                '"Add - no hours, no utilisation",'
                'IF({utilisation}{r}<0.5,"Check - chair is half empty",'
                'IF({margin_month}{r}<0,"Loss - pay is ahead of sales",'
                'IF(IF({target_month}{r}="",FALSE,'
                '{rev_month}{r}<{target_month}{r}*0.8),'
                '"Below - behind the monthly target",'
                'IF({retail_share}{r}<TargetRetailShare,'
                '"Action - barely selling any product","OK")))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Staff Management",
        subtitle="The team, their pay basis, and what each chair "
                 "brings in an hour.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.deep,
        freeze_cols=2,
        samples=sample(cfg).staff,
        help_text="Set the commission method per person. Everything "
                  "from 'hours booked' rightwards is worked out from "
                  "the diary, the till and the payroll tab.",
        extra_notes=[
            "There is no house commission rate baked into this "
            "workbook. Each person carries their own method and their "
            "own rate, and the commission tab does what that method "
            "says - including paying nothing at all to a chair renter.",
            "Utilisation uses the last four weeks, so it reacts to a "
            "quiet month rather than being flattered by a busy year.",
            "Leaves behind is the test that matters: this month's "
            "sales less this month's full cost of employing them. If "
            "it is negative, nothing else on the row makes up for it.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Team summary under the grid."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "The team at a glance", last)
    sheet.table_head(row, 0, [
        "On the team", "Active", "Hours a week", "Completed visits",
        "Service sales", "Retail sales", "Average utilisation",
        "Average rebook rate", "Pay this month"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Staff.name]],"<>")',
        '=COUNTIFS([[Staff.status]],"Active")',
        '=ROUND(SUM([[Staff.hours_week]]),1)',
        '=SUM([[Staff.done]])',
        '=ROUND(SUM([[Staff.rev_service]]),2)',
        '=ROUND(SUM([[Staff.rev_retail]]),2)',
        '=IFERROR(SUMPRODUCT(IFERROR([[Staff.utilisation]]*1,0))/'
        'MAX(1,COUNTIFS([[Staff.status]],"Active")),0)',
        '=IFERROR(SUMPRODUCT(IFERROR([[Staff.rebook_rate]]*1,0))/'
        'MAX(1,COUNTIFS([[Staff.status]],"Active")),0)',
        '=ROUND(SUM([[Staff.cost_month]]),2)',
    ], ["int", "int", "number1", "int", "money", "money", "percent",
        "percent", "money"])
