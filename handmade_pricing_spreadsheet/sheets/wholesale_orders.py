"""`Wholesale Orders` - line by line, priced at the shop's own discount.

One row per product per order. The price comes from the catalog at
whatever discount that customer gets, the cost comes from the
catalog too, and the margin is what is left - with no marketplace
fees, because a wholesale order does not pay any.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, loss_cf, lookup_product,
                     month_col, status_cf)

NAME = "Wholesale Orders"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "OrderPrefix", digits=4, width=10),
        Col("date", "Date", 11, "date", validate=date_check()),
        Col("customer", "Customer", 12, "id",
            validate=list_from("CustomerIDs")),
        Col("po_ref", "Their order ref", 14, "text"),
        Col("product", "Product", 12, "id",
            validate=list_from("ProductIDs")),
        Col("qty", "Quantity", 9, "int",
            validate=number_check(1, 1000000, decimal=False)),
        Col("price_input", "Agreed price each", 13, "money",
            validate=number_check(0, 1000000),
            note="Leave blank to use the catalog price at this "
                 "customer's discount."),
        Col("due", "Due", 11, "date", validate=date_check()),
        Col("despatched", "Despatched", 11, "date", validate=date_check()),
        Col("carriage", "Carriage charged", 13, "money",
            validate=number_check(0, 100000),
            hidden=True),
        Col("carriage_cost", "Carriage cost", 12, "money",
            validate=number_check(0, 100000),
            hidden=True),
        Col("status", "Status", 14, "status",
            validate=list_from("OrderStatusList"), cf=status_cf()),
        Col("paid", "Paid?", 8, "flag", validate=list_from("YesNoList")),
        Col("notes", "Notes", 26, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("customer_name", "Shop", 22, "text",
            formula=blank(
                "customer",
                'IFERROR(INDEX([[Wholesale Customers.name]],'
                'MATCH({customer}{r},[[Wholesale Customers.id]],0)),'
                '"Not found")')),
        Col("product_name", "Product name", 24, "text",
            formula=lookup_product("name")),
        month_col("date"),
        Col("discount", "Discount", 9, "percent",
            formula=blank(
                "customer",
                'IFERROR(INDEX([[Wholesale Customers.discount_used]],'
                'MATCH({customer}{r},[[Wholesale Customers.id]],0)),'
                'WholesaleDiscount)'),
            hidden=True),
        Col("retail", "Retail each", 11, "money",
            formula=lookup_product("price", missing=""),
            hidden=True),
        Col("unit_price", "Price each", 11, "money",
            formula=blank(
                "product",
                'IF({price_input}{r}<>"",{price_input}{r},'
                'IF({retail}{r}="","",ROUND({retail}{r}*'
                '(1-{discount}{r}),2)))')),
        Col("unit_cost", "Cost each", 11, "money",
            formula=lookup_product("cost_base", missing=""),
            hidden=True,
            note="The making cost without the retail postage - a "
                 "wholesale order ships as one box."),
        Col("order_value", "Order value", 12, "money",
            formula=blank(
                "product",
                'IF(OR({qty}{r}="",{unit_price}{r}=""),"",'
                'ROUND({qty}{r}*{unit_price}{r},2))')),
        Col("order_cost", "Order cost", 11, "money",
            formula=blank(
                "product",
                'IF(OR({qty}{r}="",{unit_cost}{r}=""),"",'
                'ROUND({qty}{r}*{unit_cost}{r}+'
                'MAX(0,IF({carriage_cost}{r}="",0,{carriage_cost}{r})-'
                'IF({carriage}{r}="",0,{carriage}{r})),2))')),
        Col("profit", "Profit", 11, "money",
            formula=blank(
                "product",
                'IF({order_value}{r}="","",ROUND({order_value}{r}-'
                '{order_cost}{r},2))'),
            cf=loss_cf()),
        Col("margin", "Margin", 9, "percent",
            formula=blank(
                "product",
                'IF(OR({order_value}{r}="",{order_value}{r}=0),"",'
                '{profit}{r}/{order_value}{r})'),
            cf=[{"type": "cell", "criteria": "<", "value": 0, "fmt": "bad"},
                {"type": "cell", "criteria": "<",
                 "value": "=TargetWholesaleMargin", "fmt": "warn"},
                {"type": "cell", "criteria": ">=",
                 "value": "=TargetWholesaleMargin", "fmt": "good"}]),
        Col("hours", "Hours of making", 12, "number",
            formula=blank(
                "product",
                'IFERROR(ROUND({qty}{r}*INDEX([[Products.minutes]],'
                'MATCH({product}{r},[[Products.id]],0))/60,1),"")'),
            note="A 60 piece order at 25 minutes each is 25 hours "
                 "of your life. Check the diary before you say "
                 "yes."),
        Col("per_hour", "Profit an hour", 11, "money",
            formula=blank(
                "product",
                'IF(OR({hours}{r}="",{hours}{r}=0),"",'
                'ROUND({profit}{r}/{hours}{r},2))'),
            hidden=True),
        Col("meets_min", "Meets minimum?", 12, "flag",
            formula=blank(
                "product",
                'IF({order_value}{r}="","",IF(AND({order_value}{r}>='
                'MinOrderValue,{qty}{r}>=MinOrderQty),"Yes","No"))'),
            hidden=True),
        Col("lead_days", "Days to despatch", 12, "int",
            formula=blank(
                "date",
                'IF({despatched}{r}="","",{despatched}{r}-{date}{r})'),
            hidden=True),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "date",
                'IF({product}{r}="","Add - which product?",'
                'IF({profit}{r}<0,"Loss - this order costs you money",'
                'IF(AND({margin}{r}<>"",{margin}{r}<'
                'TargetWholesaleMargin),'
                '"Low margin - the discount is too deep",'
                'IF(AND({due}{r}<>"",{despatched}{r}="",'
                '{due}{r}<TodayDate),"Overdue - not despatched",'
                'IF(AND({paid}{r}="No",{despatched}{r}<>"",'
                'TodayDate-{despatched}{r}>30),'
                '"Chase - unpaid over 30 days",'
                'IF({meets_min}{r}="No",'
                '"Check - below your minimum order","OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Wholesale Orders",
        subtitle="Line by line, at the shop's own discount, with "
                 "the hours it will take you.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=3,
        samples=sample(cfg).wholesale_orders,
        help_text="One row per product per order. The price comes "
                  "from the catalog at that customer's discount "
                  "unless you agreed something else.",
        extra_notes=[
            "Wholesale margin is thinner on purpose - there are "
            "no marketplace fees and the shop does the selling. "
            "It still has to be positive after your own wage.",
            "Check the hours before you accept. A big order at a "
            "thin margin that fills six weeks is how makers end "
            "up unable to take the profitable work.",
            "If a line will not make the target wholesale margin, "
            "the retail price is too low. Fix the retail price "
            "rather than refusing the shop.",
        ],
    )
