"""`Bundles` - gift sets and multi-buys, priced so the discount is a choice.

A bundle is three products in a box with a ribbon and a discount, and
it is the easiest place in handmade to lose money without noticing:
the discount comes off the retail price while every cost stays
exactly where it was, plus a bigger box and the time to assemble it.

This tab adds the components up from the Bundle Lines tab, adds the
assembly time and the packaging, and shows you what the discount is
really costing.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import (alert_cf, blank, id_col, loss_cf, round_price,
                     status_cf)

NAME = "Bundles"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "BundlePrefix", digits=4, width=10),
        Col("name", "Bundle", 28, "text",
            note="'Cosy night in gift set', 'Starter crochet kit'."),
        Col("kind", "Type", 18, "text",
            validate=list_from("BundleTypeList")),
        Col("channel", "Channel", 16, "text",
            validate=list_from("ChannelList")),
        Col("pack_profile", "Packaging", 12, "text",
            validate=list_from("PackProfileList"),
            note="Bundles usually need the bigger box. Pick the "
                 "profile that actually goes out."),
        Col("assembly_mins", "Assembly minutes", 13, "int",
            validate=number_check(0, 10000, decimal=False),
            note="Boxing, wrapping, tying, the card. Ten minutes a "
                 "set at $25 an hour is $4.17 - more than most "
                 "bundle discounts."),
        Col("discount", "Discount off retail", 13, "percent",
            validate=number_check(0, 0.9),
            note="What the buyer saves against buying the pieces "
                 "separately. Leave blank for no discount."),
        Col("price_input", "Your price", 11, "money",
            validate=number_check(0, 1000000)),
        Col("status", "Status", 13, "status",
            validate=list_from("StatusList"), cf=status_cf()),
        Col("season", "Season", 14, "text", validate=list_from("SeasonList")),
        Col("notes", "Notes", 26, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("items", "Items", 8, "int",
            formula=blank(
                "name",
                'SUMIFS([[Bundle Lines.qty]],[[Bundle Lines.bundle]],'
                '{id}{r})')),
        Col("lines", "Products", 9, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Bundle Lines.bundle]],{id}{r})'),
            hidden=True),
        Col("parts_cost", "Components cost", 13, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Bundle Lines.line_cost]],'
                '[[Bundle Lines.bundle]],{id}{r}),2)'),
            note="What the pieces cost to make, before the box and "
                 "the ribbon."),
        Col("parts_price", "Bought separately", 13, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Bundle Lines.line_price]],'
                '[[Bundle Lines.bundle]],{id}{r}),2)'),
            note="What the same items would cost the buyer one at "
                 "a time. The discount comes off this."),
        Col("pack_cost", "Packaging", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({pack_profile}{r}="Budget",PackBudget,'
                'IF({pack_profile}{r}="Standard",PackStandard,'
                'IF({pack_profile}{r}="Gift",PackGift,'
                'IF({pack_profile}{r}="Premium",PackPremium,'
                'DefaultPackaging)))),2)')),
        Col("assembly_cost", "Assembly", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({assembly_mins}{r}="",0,{assembly_mins}{r})/60*'
                'LaborCostHour,2)')),
        Col("total_cost", "Total cost", 11, "money",
            formula=blank(
                "name",
                'ROUND({parts_cost}{r}+{pack_cost}{r}+'
                '{assembly_cost}{r},2)'),
            note="Components already carry their own packaging "
                 "only if you listed them that way - check you are "
                 "not boxing twice."),
        Col("fee_pct", "Fee rate", 9, "percent",
            formula=blank(
                "name",
                'IFERROR(INDEX([[Channels.pct_total]],'
                'MATCH(IF({channel}{r}="",DefaultChannel,{channel}{r}),'
                '[[Channels.name]],0)),0)'),
            hidden=True),
        Col("fee_fixed", "Fixed fees", 9, "money",
            formula=blank(
                "name",
                'IFERROR(INDEX([[Channels.fixed_total]],'
                'MATCH(IF({channel}{r}="",DefaultChannel,{channel}{r}),'
                '[[Channels.name]],0)),0)'),
            hidden=True),
        Col("denom", "Denominator", 9, "number",
            formula=blank("name", '1-TargetRetailMargin-{fee_pct}{r}'),
            hidden=True),
        Col("suggested_raw", "Unrounded", 10, "money",
            formula=blank(
                "name",
                'IF({denom}{r}<=0,"",({total_cost}{r}+{fee_fixed}{r})/'
                '{denom}{r})'),
            hidden=True),
        Col("suggested", "Suggested price", 13, "money",
            formula=blank(
                "name",
                'IF({suggested_raw}{r}="","Check the settings",ROUND(' +
                round_price("{suggested_raw}{r}") + ',2))'),
            note="Priced from the bundle's own costs - ignore the "
                 "discount for a moment and see what it needs."),
        Col("discounted", "Discounted price", 13, "money",
            formula=blank(
                "name",
                'IF({parts_price}{r}=0,"",ROUND({parts_price}{r}*'
                '(1-IF({discount}{r}="",0,{discount}{r})),2))'),
            hidden=True),
        Col("price", "Price used", 11, "money",
            formula=blank(
                "name",
                'IF({price_input}{r}<>"",{price_input}{r},'
                'IF({discounted}{r}<>"",{discounted}{r},'
                'IF(ISNUMBER({suggested}{r}),{suggested}{r},"")))'),
            note="Your price, otherwise the discounted total of the "
                 "pieces, otherwise what the maths suggests."),
        Col("fees", "Channel fees", 11, "money",
            formula=blank(
                "name",
                'IF({price}{r}="","",ROUND({price}{r}*{fee_pct}{r}+'
                '{fee_fixed}{r},2))')),
        Col("profit", "Profit", 11, "money",
            formula=blank(
                "name",
                'IF({price}{r}="","",ROUND({price}{r}-{total_cost}{r}-'
                '{fees}{r},2))'),
            cf=loss_cf()),
        Col("margin", "Margin", 9, "percent",
            formula=blank(
                "name",
                'IF(OR({price}{r}="",{price}{r}=0),"",{profit}{r}/'
                '{price}{r})'),
            cf=[{"type": "cell", "criteria": "<", "value": 0, "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": "=MarginWarning",
                 "fmt": "warn"},
                {"type": "cell", "criteria": ">=", "value": "=MarginWarning",
                 "fmt": "good"}]),
        Col("saving", "Buyer saves", 11, "money",
            formula=blank(
                "name",
                'IF(OR({price}{r}="",{parts_price}{r}=0),"",'
                'ROUND({parts_price}{r}-{price}{r},2))'),
            note="The headline for the listing. Make sure you can "
                 "afford the headline."),
        Col("saving_pct", "Saving", 9, "percent",
            formula=blank(
                "name",
                'IF(OR({saving}{r}="",{parts_price}{r}=0),"",'
                '{saving}{r}/{parts_price}{r})'),
            hidden=True),
        Col("parts_margin", "Separate margin", 12, "percent",
            formula=blank(
                "name",
                'IF({parts_price}{r}=0,"",({parts_price}{r}-'
                '{parts_cost}{r})/{parts_price}{r})'),
            hidden=True,
            note="The margin those items make when sold one at a "
                 "time. Compare it with the bundle margin - that "
                 "gap is what the discount cost you."),
        Col("margin_lost", "Margin given away", 12, "percent",
            formula=blank(
                "name",
                'IF(OR({margin}{r}="",{parts_margin}{r}=""),"",'
                '{parts_margin}{r}-{margin}{r})'),
            cf=[{"type": "cell", "criteria": ">", "value": 0.15,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<=", "value": 0.05,
                 "fmt": "good"}]),
        Col("sold", "Bundles sold", 11, "int",
            formula=blank(
                "name",
                'SUMIFS([[Sales.qty]],[[Sales.product]],{id}{r})')),
        Col("alert", "Action", 28, "status",
            formula=blank(
                "name",
                'IF({items}{r}=0,"Add - no products in it yet",'
                'IF({price}{r}="","Check - no price yet",'
                'IF({profit}{r}<0,"Loss - the discount is too deep",'
                'IF({margin}{r}<MarginWarning,'
                '"Low margin - trim the discount",'
                'IF(AND({margin_lost}{r}<>"",{margin_lost}{r}>0.15),'
                '"Check - giving away a lot of margin",'
                'IF({assembly_mins}{r}="",'
                '"Add - how long does it take to box?","OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Bundles & Gift Sets",
        subtitle="Multi-item sets, with the discount, the box and "
                 "the assembly time all counted.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).bundles,
        help_text="Build the contents on the Bundle Lines tab, then "
                  "set a discount here and watch what it does to "
                  "the margin.",
        extra_notes=[
            "A 15% bundle discount on products that make 45% "
            "margin leaves 35% - before the bigger box and ten "
            "minutes of ribbon. Bundles are a volume play, not a "
            "margin one.",
            "Sell the saving, not the discount. 'Save $9' reads "
            "better than '15% off' and costs you exactly the same.",
            "Check you are not paying for packaging twice. If each "
            "piece is individually wrapped inside the gift box, "
            "both costs are real - put the extra in the assembly "
            "line or pick a richer profile.",
        ],
    )
