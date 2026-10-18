"""`Shipping` - your postage rate card, and who is really paying it.

Two numbers per line that most sellers merge into one and regret:
what the postage costs *you*, and what the buyer is charged. When the
buyer pays it, postage is not a cost of the product - it passes
through. When you offer free shipping, it is entirely your cost, and
it has to come out of the price.

This tab keeps them apart. Every product picks a shipping option by
name and the catalog takes only the part you actually absorb.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import (alert_cf, blank, flag_yes_no_cf, loss_cf,
                     min_number)

NAME = "Shipping"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("service", "Service", 22, "text",
            validate=list_from("ShipServiceList"),
            note="The postage product you actually buy."),
        Col("zone", "Zone", 16, "text", validate=list_from("ShipZoneList")),
        Col("max_weight", "Up to (g)", 11, "number1",
            validate=number_check(0, 1000000),
            note="The weight limit of this band. Product weight plus "
                 "packaging weight has to fit under it."),
        Col("cost", "Postage costs you", 13, "money",
            validate=number_check(0, 10000),
            note="What the label costs, including any fuel or "
                 "handling surcharge."),
        Col("packaging_extra", "Extra packaging", 12, "money",
            validate=number_check(0, 10000),
            note="Anything this service needs that the profile does "
                 "not include - a box for fragile items, a tracked "
                 "label."),
        Col("charge", "Buyer is charged", 13, "money",
            validate=number_check(0, 10000),
            note="0 for free shipping. Whatever you put here, the "
                 "workbook never pretends it covers more than it "
                 "does."),
        Col("payer", "Who pays", 20, "text",
            validate=list_from("ShipPayerList"),
            note="'Free shipping (I absorb it)' moves the whole cost "
                 "into the product price - which is exactly why free "
                 "shipping needs a higher price, not a thinner "
                 "margin."),
        Col("tracked", "Tracked?", 10, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf()),
        Col("days", "Days", 8, "int", validate=number_check(0, 365,
                                                            decimal=False)),
        Col("carrier", "Carrier", 16, "text"),
        Col("notes", "Notes", 24, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("label", "Option", 34, "text",
            formula=blank(
                "service",
                '{service}{r}&" - "&{zone}{r}&IF({max_weight}{r}="",""," to "'
                '&TEXT({max_weight}{r},"#,##0")&"g")'),
            note="The name products pick from. Built automatically so "
                 "two bands of the same service never collide."),
        Col("total_cost", "Costs you", 11, "money",
            formula=blank(
                "service",
                'ROUND(IF({cost}{r}="",0,{cost}{r})+'
                'IF({packaging_extra}{r}="",0,{packaging_extra}{r}),2)')),
        Col("absorbed", "You absorb", 11, "money",
            formula=blank(
                "service",
                'ROUND(MAX(0,{total_cost}{r}-IF({charge}{r}="",0,'
                '{charge}{r})),2)'),
            note="The part that has to come out of the product price. "
                 "This is what the catalog adds to a product's cost."),
        Col("surplus", "You keep", 11, "money",
            formula=blank(
                "service",
                'ROUND(IF({charge}{r}="",0,{charge}{r})-'
                '{total_cost}{r},2)'),
            cf=loss_cf(),
            note="Positive means the postage you charge more than "
                 "covers the label. Negative is the per-order hole "
                 "the price has to fill."),
        Col("per_gram", "Cost per 100 g", 12, "money",
            formula=blank(
                "service",
                'IF(OR({max_weight}{r}="",{max_weight}{r}=0),"",'
                'ROUND({total_cost}{r}/{max_weight}{r}*100,3))'),
            hidden=True),
        Col("used", "Products using it", 12, "int",
            formula=blank(
                "service",
                'COUNTIFS([[Products.ship_option]],{label}{r})')),
        Col("check", "Check", 26, "status",
            formula=blank(
                "service",
                'IF({cost}{r}="","Add - what does the label cost?",'
                'IF({max_weight}{r}="","Add - what is the weight limit?",'
                'IF(AND({payer}{r}="Buyer pays",{charge}{r}=""),'
                '"Add - what are you charging?",'
                'IF({surplus}{r}<-0.5,"Loss - postage is costing you",'
                'IF(AND({charge}{r}>0,{surplus}{r}>1),'
                '"Check - you are over charging postage","OK")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Shipping Rates",
        subtitle="What postage costs you, what the buyer pays, and the "
                 "difference you have to price for.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).shipping,
        help_text="One row per service and weight band. Products pick "
                  "an option by name; only the part you absorb goes "
                  "into the product cost.",
        extra_notes=[
            "Free shipping is a price decision, not a kindness. If "
            "you absorb $4.20 of postage on a $18 item, your margin "
            "just dropped 23 points - the workbook shows you "
            "exactly that.",
            "Weigh a finished, packed order on kitchen scales. "
            "Guessing the band is the most expensive guess in "
            "handmade selling.",
            "Most marketplaces charge their transaction fee on the "
            "postage the buyer paid as well. The fee setting on the "
            "Settings tab controls whether the workbook does too.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Summary strip under the rate card."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Postage summary", last)
    sheet.table_head(row, 0, ["Options listed", "Free shipping options",
                              "Average label cost", "Average absorbed",
                              "Options losing money", "Cheapest domestic"],
                     [15, 20, 18, 17, 20, 18])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Shipping.service]],"<>")',
        '=COUNTIFS([[Shipping.payer]],"Free shipping (I absorb it)")',
        '=IFERROR(ROUND(AVERAGE([[Shipping.total_cost]]),2),"")',
        '=IFERROR(ROUND(AVERAGE([[Shipping.absorbed]]),2),"")',
        '=COUNTIFS([[Shipping.surplus]],"<-0.5")',
        '=IF(COUNTIFS([[Shipping.zone]],"Domestic")=0,"",ROUND(' +
        min_number("[[Shipping.total_cost]]",
                   '[[Shipping.zone]]="Domestic"') + ',2))',
    ], ["int", "int", "money", "money", "int", "money"])
