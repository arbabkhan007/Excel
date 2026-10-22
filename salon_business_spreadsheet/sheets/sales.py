"""`Sales` - the till. One row per thing sold.

Gross, discount, refund, net, tax, total, cost and profit, on every
line. Net sales is gross less discounts less refunds, and that is the
number every other tab in the workbook calls revenue - so a discount
can never quietly inflate your takings and a refund can never be
forgotten.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, lookup_client, lookup_staff,
                     month_col, owing_cf, rank_desc, status_cf, week_col)

NAME = "Sales"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "SalePrefix", digits=4, width=11),
        Col("date", "Date", 11, "date", validate=date_check(),
            note="The day the money was earned, not the day it landed "
                 "in the bank."),
        Col("appointment", "Appointment", 12, "id",
            note="Optional. The APT ID this sale belongs to - it links "
                 "the till back to the diary and brings any deposit "
                 "with it."),
        Col("client", "Client ID", 11, "id",
            validate=list_from("ClientList"),
            note="Blank is fine for a walk-in who is not on the "
                 "client book."),
        Col("client_name", "Client", 20, "text",
            formula=lookup_client("name", missing="Walk-in"),
            hidden=True),
        Col("staff", "Stylist ID", 11, "id",
            validate=list_from("StaffList"),
            note="Who did the work, or who sold the product. This is "
                 "the column the commission is calculated from."),
        Col("staff_name", "Stylist", 16, "text",
            formula=lookup_staff("name", missing="House"), hidden=True),
        Col("kind", "Kind", 13, "text", validate=list_from("SaleKindList"),
            note="Service, Retail, Package, Gift certificate, Course "
                 "or Other. Keeping these apart is what makes the "
                 "retail share and the commission split honest."),
        Col("item", "Service / product", 12, "id",
            note="The service ID for a treatment, the product ID for "
                 "retail. Leave blank for anything that is neither."),
        Col("description", "Description", 26, "text",
            note="Only needed when there is no ID - 'gift voucher', "
                 "'course of six', 'fringe trim'."),
        Col("qty", "Qty", 7, "number1", validate=number_check(0, 10000),
            note="Blank counts as 1."),
        Col("unit_price", "Unit price", 11, "money",
            validate=number_check(0, 100000),
            note="Leave blank to take the menu or retail price. Fill "
                 "it in to override for this sale only."),
        Col("discount_pct", "Discount %", 10, "percent",
            validate=number_check(0, 1),
            note="As a percentage of the line - 10% is 0.1 or 10%."),
        Col("discount_amt", "Discount $", 11, "money",
            validate=number_check(0, 100000),
            note="A straight amount off, on top of any percentage."),
        Col("promo", "Promo code", 13, "text",
            note="Matches the code on the Marketing tab so a campaign "
                 "can be measured."),
        Col("refund", "Refunded", 11, "money",
            validate=number_check(0, 100000),
            note="Money handed back on this line. It comes straight "
                 "off net sales - no separate correction needed."),
        Col("notes", "Notes", 22, "wrap"),
        # --------------------------------------------------- the maths ----
        month_col("date"),
        week_col("date"),
        Col("qty_used", "Qty used", 9, "number1",
            formula=blank("date", 'IF({qty}{r}="",1,{qty}{r})'),
            hidden=True),
        Col("item_name", "Item", 24, "text",
            formula=blank(
                "date",
                'IF({description}{r}<>"",{description}{r},'
                'IF({kind}{r}="Retail",'
                'IFERROR(INDEX([[Inventory.name]],MATCH({item}{r},'
                '[[Inventory.id]],0))&"","Not in Inventory"),'
                'IFERROR(INDEX([[Services.name]],MATCH({item}{r},'
                '[[Services.id]],0))&"","Not in Services")))'),
            note="Your description if you typed one, otherwise the "
                 "name behind the ID."),
        Col("price_used", "Price used", 11, "money",
            formula=blank(
                "date",
                'IF({unit_price}{r}<>"",{unit_price}{r},'
                'IF({kind}{r}="Retail",'
                'IFERROR(INDEX([[Inventory.retail_price]],'
                'MATCH({item}{r},[[Inventory.id]],0))*1,0),'
                'IFERROR(INDEX([[Services.price]],MATCH({item}{r},'
                '[[Services.id]],0))*1,0)))'),
            hidden=True),
        Col("gross", "Gross", 11, "money",
            formula=blank(
                "date",
                'ROUND({qty_used}{r}*{price_used}{r},2)'),
            note="Quantity x price, before anything is taken off."),
        Col("discount", "Discount", 11, "money",
            formula=blank(
                "date",
                'ROUND(IF({discount_pct}{r}="",0,'
                '{gross}{r}*{discount_pct}{r})+'
                'IF({discount_amt}{r}="",0,{discount_amt}{r}),2)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}]),
        Col("net", "Net sale", 11, "money",
            formula=blank(
                "date",
                'ROUND({gross}{r}-{discount}{r}-'
                'IF({refund}{r}="",0,{refund}{r}),2)'),
            cf=[{"type": "data_bar", "bar_color": Palette.primary_mid}],
            note="Gross less discounts less refunds. This is the "
                 "revenue figure the whole workbook uses."),
        Col("tax", "Tax", 10, "money",
            formula=blank(
                "date",
                'IF({kind}{r}="Retail",'
                'IF(TaxOnRetail="Yes",ROUND({net}{r}*TaxRate,2),0),'
                'IF({kind}{r}="Service",'
                'IF(TaxOnServices="Yes",ROUND({net}{r}*TaxRate,2),0),'
                '0))'),
            note="Worked out from the two switches and the rate on "
                 "Settings. Set both to No if sales tax is not your "
                 "problem, or change the rate to your own."),
        Col("total", "Total due", 11, "money",
            formula=blank("date", 'ROUND({net}{r}+{tax}{r},2)'),
            note="What the client owes for this line, tax included."),
        Col("cost", "Cost", 10, "money",
            formula=blank(
                "date",
                'ROUND({qty_used}{r}*IF({kind}{r}="Retail",'
                'IFERROR(INDEX([[Inventory.unit_cost]],MATCH({item}{r},'
                '[[Inventory.id]],0))*1,0),'
                'IFERROR(INDEX([[Services.cost_product]],'
                'MATCH({item}{r},[[Services.id]],0))*1,0)),2)'),
            note="Trade cost for retail, backbar cost for a service. "
                 "Wages and commission are handled on the staff tabs, "
                 "not here."),
        Col("profit", "Gross profit", 11, "money",
            formula=blank("date", 'ROUND({net}{r}-{cost}{r},2)'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}]),
        # ---------------------------------------------------- the money ----
        Col("first_of_appt", "First line?", 10, "flag",
            formula=blank(
                "date",
                'IF({appointment}{r}="","No",'
                'IF(COUNTIFS({appointment_first}:{appointment}{r},'
                '{appointment}{r})=1,"Yes","No"))'),
            hidden=True,
            note="Only the first sale line of an appointment picks up "
                 "that appointment's deposit, so a deposit can never "
                 "be applied twice."),
        Col("paid_direct", "Paid to this sale", 12, "money",
            formula=blank(
                "date",
                'ROUND(SUMIFS([[Payments.amount]],[[Payments.sale]],'
                '{id}{r},[[Payments.counts]],"Yes"),2)'),
            hidden=True),
        Col("paid_deposit", "Deposit applied", 12, "money",
            formula=blank(
                "date",
                'IF({first_of_appt}{r}="No",0,'
                'ROUND(SUMIFS([[Payments.amount]],'
                '[[Payments.appointment]],{appointment}{r},'
                '[[Payments.sale]],"",[[Payments.counts]],"Yes"),2))'),
            hidden=True),
        Col("paid", "Paid", 11, "money",
            formula=blank(
                "date",
                'ROUND({paid_direct}{r}+{paid_deposit}{r},2)'),
            note="Payments booked against this sale, plus any deposit "
                 "taken on the appointment that has not already been "
                 "applied somewhere else."),
        Col("balance", "Balance", 11, "money",
            formula=blank("date", 'ROUND({total}{r}-{paid}{r},2)'),
            cf=owing_cf(),
            note="Total due less everything paid. Positive is money "
                 "still owed to you."),
        Col("status", "Paid?", 12, "status",
            formula=blank(
                "date",
                'IF({total}{r}<=0,"No charge",'
                'IF({balance}{r}<=0.004,"Paid",'
                'IF({paid}{r}<=0.004,"Unpaid","Part paid")))'),
            cf=status_cf()),
        Col("retail_with", "Retail too?", 11, "flag",
            formula=blank(
                "date",
                'IF({kind}{r}<>"Service","",IF({client}{r}="","No",'
                'IF(COUNTIFS({client_first}:{client_last},{client}{r},'
                '{date_first}:{date_last},{date}{r},'
                '{kind_first}:{kind_last},"Retail")>0,"Yes","No")))'),
            hidden=True,
            note="Did this client also take something home that day? "
                 "The retail attachment rate is built from it."),
        rank_desc("net", guard='{date}{r}=""', header="Sale rank"),
        Col("alert", "Check", 28, "status",
            formula=blank(
                "date",
                'IF({kind}{r}="","Add - what kind of sale is this?",'
                'IF(IF({price_used}{r}<=0,{refund}{r}="",FALSE),'
                '"Check - no price found for this item",'
                'IF({staff}{r}="","Assign - nobody credited",'
                'IF({balance}{r}>0.004,"Unpaid - balance outstanding",'
                'IF({profit}{r}<0,"Loss - sold below cost",'
                'IF({discount}{r}>{gross}{r}*0.5,'
                '"Check - more than half off","OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Sales & Transactions",
        subtitle="Every service, product and voucher sold - gross, "
                 "discount, refund, net, tax and the profit left over.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).sales,
        help_text="One row per thing sold. Two lines for a client who "
                  "had a colour and bought shampoo - that is how the "
                  "retail figures stay separate from the service "
                  "figures.",
        extra_notes=[
            "Net sale = gross - discounts - refunds. Nothing else in "
            "the workbook calls anything revenue, so the dashboard, "
            "the profit and loss and the commission calculator can "
            "never disagree with each other.",
            "Refunds go on the line they relate to, as a positive "
            "number in the refunded column. No negative sales, no "
            "correcting entries, no double subtraction.",
            "Tax is driven by the two switches on Settings. Tax on "
            "retail only, on services only, on both or on neither - "
            "set it once to match where you trade.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Till summary under the grid."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "The till, all time", last)
    sheet.table_head(row, 0, [
        "Lines", "Gross", "Discounts", "Refunds", "Net sales", "Tax",
        "Cost of sales", "Gross profit", "Paid", "Outstanding"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Sales.date]],"<>")',
        '=ROUND(SUM([[Sales.gross]]),2)',
        '=ROUND(SUM([[Sales.discount]]),2)',
        '=ROUND(SUM([[Sales.refund]]),2)',
        '=ROUND(SUM([[Sales.net]]),2)',
        '=ROUND(SUM([[Sales.tax]]),2)',
        '=ROUND(SUM([[Sales.cost]]),2)',
        '=ROUND(SUM([[Sales.profit]]),2)',
        '=ROUND(SUM([[Sales.paid]]),2)',
        '=ROUND(SUM([[Sales.balance]]),2)',
    ], ["int", "money", "money", "money", "money", "money", "money",
        "money", "money", "money"])
