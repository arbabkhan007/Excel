"""`Wholesale Customers` - the shops that buy from you, and what they are worth.

Wholesale is a different business from retail: bigger orders, half
the price, no marketplace fees and a shop that either reorders or
does not. This tab keeps the accounts, the terms and the honest
answer to "is this stockist worth the discount?".
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..layout import max_if
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, id_col, loss_cf, status_cf

NAME = "Wholesale Customers"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "CustomerPrefix", digits=4, width=10),
        Col("name", "Shop", 26, "text"),
        Col("kind", "Type", 18, "text",
            validate=list_from("CustomerTypeList")),
        Col("contact", "Contact", 18, "text"),
        Col("email", "Email", 24, "text"),
        Col("phone", "Phone", 15, "text"),
        Col("city", "Town / city", 16, "text"),
        Col("terms", "Payment terms", 22, "text",
            validate=list_from("TermsList"),
            note="Net 30 means you are lending a shop money for a "
                 "month. Price that in or ask for a deposit."),
        Col("discount", "Their discount", 11, "percent",
            validate=number_check(0, 0.9),
            note="Off your retail price. Blank uses the standard "
                 "wholesale discount on Settings."),
        Col("min_order", "Their minimum", 11, "money",
            validate=number_check(0, 1000000),
            hidden=True),
        Col("first_order", "First order", 11, "date",
            validate=date_check()),
        Col("stockist_since", "Stockist since", 12, "date",
            validate=date_check(), hidden=True),
        Col("exclusive", "Exclusive area?", 12, "flag",
            validate=list_from("YesNoList"),
            note="Granting exclusivity is worth something. Make "
                 "sure you got something for it."),
        Col("active", "Still buying?", 11, "flag",
            validate=list_from("YesNoList")),
        Col("notes", "Notes", 28, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("discount_used", "Discount used", 11, "percent",
            formula=blank(
                "name",
                'IF({discount}{r}<>"",{discount}{r},WholesaleDiscount)'),
            hidden=True),
        Col("orders", "Orders", 9, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Wholesale Orders.customer]],{id}{r})')),
        Col("units", "Items bought", 11, "int",
            formula=blank(
                "name",
                'SUMIFS([[Wholesale Orders.qty]],'
                '[[Wholesale Orders.customer]],{id}{r})')),
        Col("value", "Total spent", 12, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Wholesale Orders.order_value]],'
                '[[Wholesale Orders.customer]],{id}{r}),2)')),
        Col("profit", "Profit from them", 13, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Wholesale Orders.profit]],'
                '[[Wholesale Orders.customer]],{id}{r}),2)'),
            cf=loss_cf()),
        Col("margin", "Margin", 9, "percent",
            formula=blank(
                "name",
                'IF(OR({value}{r}="",{value}{r}=0),"",{profit}{r}/'
                '{value}{r})'),
            cf=[{"type": "cell", "criteria": "<", "value": 0, "fmt": "bad"},
                {"type": "cell", "criteria": "<",
                 "value": "=TargetWholesaleMargin", "fmt": "warn"},
                {"type": "cell", "criteria": ">=",
                 "value": "=TargetWholesaleMargin", "fmt": "good"}]),
        Col("avg_order", "Average order", 11, "money",
            formula=blank(
                "name",
                'IF({orders}{r}=0,"",ROUND({value}{r}/{orders}{r},2))')),
        Col("last_order", "Last order", 11, "date",
            formula=blank(
                "name",
                'IF(' + max_if("[[Wholesale Orders.date]]",
                               "[[Wholesale Orders.customer]]={id}{r}") +
                '=0,"",' + max_if("[[Wholesale Orders.date]]",
                                  "[[Wholesale Orders.customer]]={id}{r}") +
                ')')),
        Col("days_quiet", "Days since", 10, "int",
            formula=blank(
                "name",
                'IF({last_order}{r}="","",TodayDate-{last_order}{r})'),
            cf=[{"type": "cell", "criteria": ">", "value": 180,
                 "fmt": "warn"}],
            hidden=True),
        Col("outstanding", "Unpaid", 11, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Wholesale Orders.order_value]],'
                '[[Wholesale Orders.customer]],{id}{r},'
                '[[Wholesale Orders.paid]],"No"),2)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}]),
        Col("status", "Status", 13, "status",
            formula=blank(
                "name",
                'IF({orders}{r}=0,"Prospect",'
                'IF({active}{r}="No","Inactive",'
                'IF(AND({days_quiet}{r}<>"",{days_quiet}{r}>180),'
                '"Going quiet","Active")))'),
            cf=status_cf()),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "name",
                'IF({orders}{r}=0,"New - no orders yet",'
                'IF({outstanding}{r}>0,"Chase - money outstanding",'
                'IF({margin}{r}<0,"Loss - this stockist costs you money",'
                'IF(AND({margin}{r}<>"",{margin}{r}<'
                'TargetWholesaleMargin),'
                '"Low margin - their discount is too deep",'
                'IF({status}{r}="Going quiet",'
                '"Chase - six months without an order","OK")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Wholesale Customers",
        subtitle="The shops that stock you, their terms, and what "
                 "each one is really worth.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).wholesale_customers,
        help_text="Set each shop's discount and terms here; the "
                  "order tab prices against them automatically.",
        extra_notes=[
            "A stockist who reorders every quarter at 50% off is "
            "worth more than three who buy once at 40%. Look at "
            "the order count before the discount.",
            "Net 30 is a loan. If a shop wants terms, a deposit "
            "or a slightly smaller discount is a fair trade.",
            "Exclusivity has a price. If a shop wants to be the "
            "only one in town, that is worth a bigger minimum "
            "order, not a bigger discount.",
        ],
    )
