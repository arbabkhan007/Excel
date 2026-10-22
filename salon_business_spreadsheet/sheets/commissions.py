"""`Commissions` - one row per stylist per month, eight ways to pay.

The method is whatever the owner chose on the Staff tab. This tab
reads it and applies it: a straight percentage, a percentage of
service and retail, a percentage after the colour has been paid for,
a tiered rate that climbs with the month, a flat amount per service,
hourly only, rent a chair, or nothing at all. No house rate is
assumed anywhere.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, lookup_staff, rank_desc,
                     status_cf, yesno)

NAME = "Commissions"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("month", "CommissionPrefix", digits=4, width=11),
        Col("month", "Month", 11, "text", validate=list_from("MonthKeyList"),
            note="The month this covers, as yyyy-mm. Pick it from the "
                 "list so it matches the sales tab exactly."),
        Col("staff", "Stylist ID", 11, "id",
            validate=list_from("StaffList")),
        Col("staff_name", "Stylist", 18, "text",
            formula=lookup_staff("name")),
        Col("method", "Method used", 24, "text",
            formula=lookup_staff("method", missing="None"),
            cf=status_cf(),
            note="Taken from the Staff tab. Change it there and this "
                 "month recalculates."),
        Col("adjust", "Adjustment", 11, "money",
            validate=number_check(-100000, 100000),
            note="A correction, a clawback on a refund, a guarantee "
                 "top-up. Plus or minus."),
        Col("bonus", "Bonus", 11, "money",
            validate=number_check(0, 100000),
            note="Anything on top - a retail competition, a full "
                 "column bonus."),
        Col("approved", "Approved?", 11, "flag", validate=yesno()),
        Col("paid_on", "Paid on", 11, "date", validate=date_check(),
            note="Leave blank until it has actually gone out."),
        Col("notes", "Notes", 24, "wrap"),
        # ------------------------------------------------- what they sold --
        Col("service_net", "Service sales", 12, "money",
            formula=blank(
                "month",
                'ROUND(SUMIFS([[Sales.net]],[[Sales.staff]],{staff}{r},'
                '[[Sales.month]],{month}{r},[[Sales.kind]],'
                '"Service"),2)'),
            cf=[{"type": "data_bar", "bar_color": Palette.primary_mid}]),
        Col("service_gross", "Before discount", 12, "money",
            formula=blank(
                "month",
                'ROUND(SUMIFS([[Sales.gross]],[[Sales.staff]],'
                '{staff}{r},[[Sales.month]],{month}{r},'
                '[[Sales.kind]],"Service"),2)'),
            hidden=True),
        Col("retail_net", "Retail sales", 11, "money",
            formula=blank(
                "month",
                'ROUND(SUMIFS([[Sales.net]],[[Sales.staff]],{staff}{r},'
                '[[Sales.month]],{month}{r},[[Sales.kind]],'
                '"Retail"),2)')),
        Col("retail_gross", "Retail before", 11, "money",
            formula=blank(
                "month",
                'ROUND(SUMIFS([[Sales.gross]],[[Sales.staff]],'
                '{staff}{r},[[Sales.month]],{month}{r},'
                '[[Sales.kind]],"Retail"),2)'),
            hidden=True),
        Col("product_cost", "Product used", 11, "money",
            formula=blank(
                "month",
                'ROUND(SUMIFS([[Sales.cost]],[[Sales.staff]],{staff}{r},'
                '[[Sales.month]],{month}{r},[[Sales.kind]],'
                '"Service"),2)'),
            note="Backbar cost of the services they did - the figure "
                 "the 'percent after product cost' method takes off "
                 "first."),
        Col("services", "Services done", 11, "int",
            formula=blank(
                "month",
                'COUNTIFS([[Sales.staff]],{staff}{r},[[Sales.month]],'
                '{month}{r},[[Sales.kind]],"Service")'),
            hidden=True),
        # ------------------------------------------------- the base and rate
        Col("base_service", "Service base", 12, "money",
            formula=blank(
                "month",
                'IF(CommissionOn="Gross before discounts",'
                '{service_gross}{r},{service_net}{r})'),
            note="Which figure commission is paid on is your choice - "
                 "set it once on Settings, net of discounts or gross "
                 "before them."),
        Col("base_retail", "Retail base", 11, "money",
            formula=blank(
                "month",
                'IF(CommissionOn="Gross before discounts",'
                '{retail_gross}{r},{retail_net}{r})'),
            hidden=True),
        Col("rate_service", "Service rate", 11, "percent",
            formula=blank(
                "month",
                'IF({method}{r}="Tiered on monthly service revenue",'
                'IFERROR(INDEX(TierRate,MATCH({base_service}{r},'
                'TierFrom,1)),0),'
                'IFERROR(IF(INDEX([[Staff.rate_service]],'
                'MATCH({staff}{r},[[Staff.id]],0))="",ServiceCommission,'
                'INDEX([[Staff.rate_service]],MATCH({staff}{r},'
                '[[Staff.id]],0))),ServiceCommission))'),
            note="Their own rate if they have one, otherwise the salon "
                 "default. On the tiered method it is read from the "
                 "bands on Settings instead."),
        Col("rate_retail", "Retail rate", 11, "percent",
            formula=blank(
                "month",
                'IFERROR(IF(INDEX([[Staff.rate_retail]],MATCH({staff}{r},'
                '[[Staff.id]],0))="",RetailCommission,'
                'INDEX([[Staff.rate_retail]],MATCH({staff}{r},'
                '[[Staff.id]],0))),RetailCommission)'),
            hidden=True),
        Col("flat_rate", "Flat per service", 12, "money",
            formula=blank(
                "month",
                'IFERROR(INDEX([[Staff.flat_rate]],MATCH({staff}{r},'
                '[[Staff.id]],0))*1,0)'),
            hidden=True),
        # ----------------------------------------------------- the payment --
        Col("commission", "Commission", 12, "money",
            formula=blank(
                "month",
                'ROUND(IF({method}{r}="Percent of service",'
                '{base_service}{r}*{rate_service}{r},'
                'IF({method}{r}="Percent of service and retail",'
                '{base_service}{r}*{rate_service}{r}+'
                '{base_retail}{r}*{rate_retail}{r},'
                'IF({method}{r}="Percent after product cost",'
                'MAX(0,{base_service}{r}-{product_cost}{r})*'
                '{rate_service}{r},'
                'IF({method}{r}="Tiered on monthly service revenue",'
                '{base_service}{r}*{rate_service}{r},'
                'IF({method}{r}="Flat amount per service",'
                '{services}{r}*{flat_rate}{r},0))))),2)'),
            cf=[{"type": "data_bar", "bar_color": Palette.accent}],
            note="Eligible revenue x the rate, exactly as the method "
                 "says. None, hourly only and rent a chair all come "
                 "out at zero here on purpose - those people are paid "
                 "on the payroll tab or pay you rent instead."),
        Col("tips", "Tips", 10, "money",
            formula=blank(
                "month",
                'ROUND(SUMIFS([[Payments.amount]],[[Payments.staff]],'
                '{staff}{r},[[Payments.month]],{month}{r},'
                '[[Payments.kind]],"Tip"),2)'),
            note="Tips logged to this stylist. They are theirs - shown "
                 "here so the payout is complete, kept out of salon "
                 "revenue everywhere else."),
        Col("payout", "To pay", 12, "money",
            formula=blank(
                "month",
                'ROUND({commission}{r}+IF({adjust}{r}="",0,'
                '{adjust}{r})+IF({bonus}{r}="",0,{bonus}{r}),2)'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}],
            note="Commission plus adjustments and bonus. Tips are not "
                 "added in - they are not yours to pay."),
        Col("share", "Share of their sales", 13, "percent",
            formula=blank(
                "month",
                'IF({base_service}{r}+{base_retail}{r}=0,"",'
                '{payout}{r}/({base_service}{r}+{base_retail}{r}))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0.55,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 0.45,
                 "fmt": "warn"}],
            note="What this month costs you as a percentage of what "
                 "they billed. Over about 45% is worth a conversation."),
        rank_desc("commission", guard='{month}{r}=""',
                  header="Commission rank"),
        Col("alert", "Check", 28, "status",
            formula=blank(
                "month",
                'IF({staff}{r}="","Add - whose commission is this?",'
                'IF(COUNTIFS({month_first}:{month_last},{month}{r},'
                '{staff_first}:{staff_last},{staff}{r})>1,'
                '"Duplicate - two rows for the same month",'
                'IF({method}{r}="None","",'
                'IF({method}{r}="Rent a chair","",'
                'IF({method}{r}="Hourly only","",'
                'IF({payout}{r}=0,"Check - nothing calculated",'
                'IF({approved}{r}<>"Yes","Due - waiting for approval",'
                'IF({paid_on}{r}="","Due - approved but not paid",'
                '"OK"))))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Staff Commissions",
        subtitle="One row per stylist per month. The method is yours; "
                 "the arithmetic is the workbook's.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.deep,
        freeze_cols=2,
        samples=sample(cfg).commissions,
        help_text="Add a row for each stylist each month: month, "
                  "stylist, and any adjustment. The sales, the base, "
                  "the rate and the payout are all read from elsewhere.",
        extra_notes=[
            "Eight methods are supported and none of them is assumed. "
            "Set the method on the Staff tab, per person, and this "
            "tab follows it: percent of service, service and retail, "
            "percent after product cost, tiered, flat per service, "
            "hourly only, rent a chair, or none.",
            "Commission is paid on whichever base you chose on "
            "Settings - net of discounts, or gross before them. One "
            "switch, applied to everybody, so there is never an "
            "argument about which figure was used.",
            "The tiered method reads the bands on the Settings tab. "
            "Edit the bands there to match your own ladder; the "
            "lookup finds the right rung on its own.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Commission summary under the grid."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Commission totals", last)
    sheet.table_head(row, 0, [
        "Rows", "Service sales", "Retail sales", "Commission",
        "Adjustments", "Bonuses", "To pay", "Tips passed on",
        "Waiting for approval", "Approved, not paid"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Commissions.month]],"<>")',
        '=ROUND(SUM([[Commissions.service_net]]),2)',
        '=ROUND(SUM([[Commissions.retail_net]]),2)',
        '=ROUND(SUM([[Commissions.commission]]),2)',
        '=ROUND(SUM([[Commissions.adjust]]),2)',
        '=ROUND(SUM([[Commissions.bonus]]),2)',
        '=ROUND(SUM([[Commissions.payout]]),2)',
        '=ROUND(SUM([[Commissions.tips]]),2)',
        '=COUNTIFS([[Commissions.approved]],"<>Yes",'
        '[[Commissions.month]],"<>")',
        '=COUNTIFS([[Commissions.approved]],"Yes",'
        '[[Commissions.paid_on]],"")',
    ], ["int", "money", "money", "money", "money", "money", "money",
        "money", "int", "int"])
