"""`Sales` - what really happened, order by order.

Estimated profit is a plan. This is the tab where the plan meets
discounts, refunds, a postage label that cost more than you charged
and a fee that was not quite what the fee page said.

Costs come from the catalog, with the actual postage swapped in for
the estimate, so nothing is counted twice. Type the fee the
marketplace really took whenever you have the statement in front of
you - the Channels tab compares it with your settings.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, loss_cf, lookup_product,
                     month_col, status_cf)

NAME = "Sales"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "SalePrefix", digits=4, width=10),
        Col("date", "Date", 11, "date", validate=date_check()),
        Col("product", "Product", 12, "id",
            validate=list_from("ProductIDs"),
            note="A product ID, or a bundle ID - both work, the "
                 "costs come from whichever tab the ID lives on."),
        Col("qty", "Qty", 7, "int",
            validate=number_check(1, 10000, decimal=False)),
        Col("channel", "Channel", 16, "text",
            validate=list_from("ChannelList")),
        Col("unit_price", "Price each", 11, "money",
            validate=number_check(0, 1000000),
            note="What you actually charged, before any discount. "
                 "Blank uses the catalog price."),
        Col("discount", "Discount given", 12, "money",
            validate=number_check(0, 1000000),
            note="In money, for the whole order. A 15% code on a "
                 "$28 order is 4.20."),
        Col("promo", "Promo code", 12, "text",
            validate=list_from("PromoCodes")),
        Col("ship_charged", "Postage charged", 12, "money",
            validate=number_check(0, 10000)),
        Col("ship_paid", "Postage paid", 11, "money",
            validate=number_check(0, 10000),
            note="What the label actually cost. The difference is "
                 "yours either way."),
        Col("fees_input", "Fees actually taken", 14, "money",
            validate=number_check(0, 1000000),
            note="From the marketplace statement. Leave blank and "
                 "the workbook estimates it from the Channels tab."),
        Col("refund", "Refunded", 10, "money",
            validate=number_check(0, 1000000)),
        Col("status", "Status", 14, "status",
            validate=list_from("SaleStatusList"), cf=status_cf()),
        Col("customer", "Customer", 18, "text"),
        Col("event", "Event", 16, "text", validate=list_from("EventNames"),
            note="For craft fair sales - ties the order to the "
                 "event so you can see if the stall paid."),
        Col("notes", "Notes", 24, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("product_name", "Product name", 24, "text",
            formula=lookup_product("name", missing="Bundle or deleted")),
        month_col("date"),
        Col("catalog_price", "Catalog price", 11, "money",
            formula=lookup_product("price", missing=""),
            hidden=True),
        Col("price_used", "Price used", 11, "money",
            formula=blank(
                "date",
                'IF({unit_price}{r}<>"",{unit_price}{r},'
                'IF({catalog_price}{r}="",0,{catalog_price}{r}))'),
            hidden=True),
        Col("gross", "Gross", 11, "money",
            formula=blank(
                "date",
                'ROUND(IF({qty}{r}="",0,{qty}{r})*{price_used}{r},2)'),
            hidden=True),
        Col("revenue", "Revenue", 11, "money",
            formula=blank(
                "date",
                'ROUND({gross}{r}-IF({discount}{r}="",0,{discount}{r})-'
                'IF({refund}{r}="",0,{refund}{r}),2)'),
            note="What you actually took for the items, after "
                 "discounts and refunds. Postage is kept separate."),
        Col("unit_cost_total", "Catalog cost", 11, "money",
            formula=lookup_product("total_cost", missing=""),
            hidden=True),
        Col("unit_ship", "Catalog postage", 11, "money",
            formula=lookup_product("ship_cost", missing=""),
            hidden=True),
        Col("cogs", "Cost of the goods", 13, "money",
            formula=blank(
                "date",
                'ROUND(IF({qty}{r}="",0,{qty}{r})*'
                'MAX(0,IF({unit_cost_total}{r}="",0,{unit_cost_total}{r})-'
                'IF({unit_ship}{r}="",0,{unit_ship}{r})),2)'),
            note="Materials, labour, packaging and overhead from "
                 "the catalog. The estimated postage is taken out "
                 "because the real postage is on this row."),
        Col("ship_absorbed", "Postage you absorbed", 14, "money",
            formula=blank(
                "date",
                'ROUND(MAX(0,IF({ship_paid}{r}="",0,{ship_paid}{r})-'
                'IF({ship_charged}{r}="",0,{ship_charged}{r})),2)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}]),
        Col("fee_pct", "Fee rate", 9, "percent",
            formula=blank(
                "date",
                'IFERROR(INDEX([[Channels.pct_total]],'
                'MATCH(IF({channel}{r}="",DefaultChannel,{channel}{r}),'
                '[[Channels.name]],0)),0)'),
            hidden=True),
        Col("fee_fixed", "Fixed fees", 9, "money",
            formula=blank(
                "date",
                'IFERROR(INDEX([[Channels.fixed_total]],'
                'MATCH(IF({channel}{r}="",DefaultChannel,{channel}{r}),'
                '[[Channels.name]],0)),0)'),
            hidden=True),
        Col("fee_ship", "Fees on postage?", 10, "flag",
            formula=blank(
                "date",
                'IFERROR(INDEX([[Channels.fees_on_shipping]],'
                'MATCH(IF({channel}{r}="",DefaultChannel,{channel}{r}),'
                '[[Channels.name]],0)),"No")'),
            hidden=True),
        Col("fees_est", "Fees estimated", 12, "money",
            formula=blank(
                "date",
                'ROUND(({revenue}{r}+IF({fee_ship}{r}="Yes",'
                'IF({ship_charged}{r}="",0,{ship_charged}{r}),0))*'
                '{fee_pct}{r}+{fee_fixed}{r},2)'),
            hidden=True),
        Col("fees", "Fees", 11, "money",
            formula=blank(
                "date",
                'IF({fees_input}{r}<>"",{fees_input}{r},'
                '{fees_est}{r})')),
        Col("fee_gap", "Fees versus estimate", 13, "money",
            formula=blank(
                "date",
                'IF({fees_input}{r}="","",ROUND({fees_input}{r}-'
                '{fees_est}{r},2))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0.25,
                 "fmt": "bad"}],
            hidden=True,
            note="Real fees consistently above the estimate means "
                 "the fee settings need updating."),
        Col("profit", "Profit", 11, "money",
            formula=blank(
                "date",
                'ROUND({revenue}{r}-{cogs}{r}-{fees}{r}-'
                '{ship_absorbed}{r},2)'),
            cf=loss_cf(),
            note="Revenue less the goods, the fees and whatever "
                 "postage you swallowed."),
        Col("margin", "Margin", 9, "percent",
            formula=blank(
                "date",
                'IF(OR({revenue}{r}="",{revenue}{r}=0),"",{profit}{r}/'
                '{revenue}{r})'),
            cf=[{"type": "cell", "criteria": "<", "value": 0, "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": "=MarginWarning",
                 "fmt": "warn"},
                {"type": "cell", "criteria": ">=", "value": "=MarginWarning",
                 "fmt": "good"}]),
        Col("planned_margin", "Planned margin", 12, "percent",
            formula=lookup_product("margin", missing=""),
            hidden=True),
        Col("margin_gap", "Versus plan", 10, "percent",
            formula=blank(
                "date",
                'IF(OR({margin}{r}="",{planned_margin}{r}=""),"",'
                '{margin}{r}-{planned_margin}{r})'),
            hidden=True),
        Col("hours", "Hours of making", 12, "number",
            formula=blank(
                "date",
                'IFERROR(ROUND(IF({qty}{r}="",0,{qty}{r})*'
                'INDEX([[Products.minutes]],MATCH({product}{r},'
                '[[Products.id]],0))/60,2),"")'),
            hidden=True),
        Col("profit_hour", "Profit an hour", 11, "money",
            formula=blank(
                "date",
                'IF(OR({hours}{r}="",{hours}{r}=0),"",'
                'ROUND({profit}{r}/{hours}{r},2))'),
            hidden=True),
        Col("alert", "Action", 28, "status",
            formula=blank(
                "date",
                'IF({product}{r}="","Add - which product?",'
                'IF({profit}{r}<0,"Loss - this order cost you money",'
                'IF({status}{r}="Refunded","Refunded - check why",'
                'IF(AND({fee_gap}{r}<>"",{fee_gap}{r}>0.25),'
                '"Check - fees higher than your settings",'
                'IF({ship_absorbed}{r}>1,'
                '"Check - postage is eating the margin",'
                'IF(AND({margin}{r}<>"",{margin}{r}<MarginWarning),'
                '"Low margin on this order","OK")))))) '),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Sales & Actual Profit",
        subtitle="Every order, with the real postage, the real fees "
                 "and the real margin.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=3,
        samples=sample(cfg).sales,
        help_text="Enter the order. Leave the price blank to use "
                  "the catalog price, and type the real fees "
                  "whenever you have the statement.",
        extra_notes=[
            "Estimated and actual will not match, and that is the "
            "point. The gap is where a price quietly stopped "
            "working.",
            "Postage is a separate line because it behaves "
            "differently: the buyer usually pays it, the fee is "
            "usually charged on it, and when you offer free "
            "shipping it comes straight out of your margin.",
            "Record craft fair sales against the event. A stall "
            "that took $380 and cost $95 plus two days is a "
            "decision you can only make with the numbers in front "
            "of you.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Summary strip under the order book."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Sales summary", last)
    sheet.table_head(row, 0, ["Orders", "Items", "Revenue", "Cost of goods",
                              "Fees", "Postage absorbed", "Profit",
                              "Margin", "Loss making orders"],
                     [10, 10, 14, 15, 12, 17, 13, 10, 18])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Sales.date]],"<>")',
        '=SUM([[Sales.qty]])',
        '=ROUND(SUM([[Sales.revenue]]),2)',
        '=ROUND(SUM([[Sales.cogs]]),2)',
        '=ROUND(SUM([[Sales.fees]]),2)',
        '=ROUND(SUM([[Sales.ship_absorbed]]),2)',
        '=ROUND(SUM([[Sales.profit]]),2)',
        '=IFERROR(ROUND(SUM([[Sales.profit]])/SUM([[Sales.revenue]]),4),"")',
        '=COUNTIFS([[Sales.profit]],"<0")',
    ], ["int", "int", "money", "money", "money", "money", "money",
        "percent", "int"])
