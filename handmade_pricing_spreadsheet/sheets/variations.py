"""`Variations` - sizes, colours, personalisation and rush orders.

The same product in three sizes is three different costs, and
"can you add a name on it?" is twenty minutes nobody charged for.
This tab starts from a product in the catalog, applies the size
factors, adds whatever extra material and time the variation takes,
and prices it - including a surcharge for a turnaround you did not
offer.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import (alert_cf, blank, id_col, loss_cf, lookup_product,
                     round_price)

NAME = "Variations"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "VariationPrefix", digits=4, width=10),
        Col("product", "Base product", 12, "id",
            validate=list_from("ProductIDs"),
            note="The variation inherits this product's costs and "
                 "then adjusts them."),
        Col("name", "Variation", 26, "text",
            note="'Large, personalised', 'Rush order, navy'."),
        Col("kind", "What changes", 18, "text",
            validate=list_from("VariationTypeList")),
        Col("size", "Size", 14, "text", validate=list_from("SizeNames"),
            note="Applies the material and time factors from the "
                 "Lists tab to the base product."),
        Col("extra_material", "Extra materials", 13, "money",
            validate=number_check(0, 100000),
            note="Anything this version uses that the base does "
                 "not - a bigger clasp, a second colour, a gift "
                 "box insert."),
        Col("extra_mins", "Extra minutes", 12, "int",
            validate=number_check(0, 100000, decimal=False),
            note="Personalisation, colour changes, re-threading "
                 "the machine, an extra fitting."),
        Col("rush", "Turnaround", 14, "text",
            validate=list_from("RushNames"),
            note="Adds the surcharge from the Lists tab. Saying "
                 "yes to Friday should be worth something."),
        Col("personalisation", "Personalisation fee", 14, "money",
            validate=number_check(0, 100000),
            note="A flat charge you add on top, if you prefer that "
                 "to pricing the minutes."),
        Col("channel", "Channel", 15, "text",
            validate=list_from("ChannelList")),
        Col("price_input", "Your price", 11, "money",
            validate=number_check(0, 1000000)),
        Col("offered", "Offered?", 10, "flag",
            validate=list_from("YesNoList"),
            note="No for a variation you costed and decided "
                 "against - the workbook keeps the reasoning."),
        Col("notes", "Notes", 24, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("base_name", "Base product name", 24, "text",
            formula=lookup_product("name"),
            hidden=True),
        Col("base_mat", "Base materials", 11, "money",
            formula=lookup_product("mat_cost", missing=""),
            hidden=True),
        Col("base_labour", "Base labour", 11, "money",
            formula=lookup_product("labour_adj", missing=""),
            hidden=True),
        Col("base_other", "Base other", 11, "money",
            formula=blank(
                "product",
                'IFERROR(INDEX([[Products.cost_base]],MATCH({product}{r},'
                '[[Products.id]],0))-INDEX([[Products.mat_cost]],'
                'MATCH({product}{r},[[Products.id]],0))-'
                'INDEX([[Products.labour_adj]],MATCH({product}{r},'
                '[[Products.id]],0)),0)'),
            hidden=True,
            note="Consumables, packaging and postage from the base "
                 "product - unchanged by the variation."),
        Col("base_minutes", "Base minutes", 11, "number1",
            formula=lookup_product("minutes", missing=""),
            hidden=True),
        Col("base_price", "Base price", 11, "money",
            formula=lookup_product("price", missing="")),
        Col("size_mat", "Size factor", 9, "number",
            formula=blank(
                "product",
                'IFERROR(INDEX(SizeMaterialFactors,MATCH({size}{r},'
                'SizeNames,0)),1)'),
            hidden=True),
        Col("size_time", "Time factor", 9, "number",
            formula=blank(
                "product",
                'IFERROR(INDEX(SizeTimeFactors,MATCH({size}{r},'
                'SizeNames,0)),1)'),
            hidden=True),
        Col("rush_rate", "Rush surcharge", 10, "percent",
            formula=blank(
                "product",
                'IFERROR(INDEX(RushFactors,MATCH({rush}{r},'
                'RushNames,0)),0)'),
            hidden=True),
        Col("mat_cost", "Materials", 11, "money",
            formula=blank(
                "product",
                'ROUND(IF({base_mat}{r}="",0,{base_mat}{r})*'
                '{size_mat}{r}+IF({extra_material}{r}="",0,'
                '{extra_material}{r}),2)')),
        Col("minutes", "Minutes", 10, "number1",
            formula=blank(
                "product",
                'ROUND(IF({base_minutes}{r}="",0,{base_minutes}{r})*'
                '{size_time}{r}+IF({extra_mins}{r}="",0,'
                '{extra_mins}{r}),1)')),
        Col("labor_cost", "Labour", 11, "money",
            formula=blank(
                "product",
                'ROUND(IF({base_labour}{r}="",0,{base_labour}{r})*'
                '{size_time}{r}+IF({extra_mins}{r}="",0,'
                '{extra_mins}{r})/60*LaborCostHour,2)')),
        Col("total_cost", "Total cost", 11, "money",
            formula=blank(
                "product",
                'ROUND({mat_cost}{r}+{labor_cost}{r}+'
                'IF({base_other}{r}="",0,{base_other}{r}),2)')),
        Col("fee_pct", "Fee rate", 9, "percent",
            formula=blank(
                "product",
                'IFERROR(INDEX([[Channels.pct_total]],'
                'MATCH(IF({channel}{r}="",DefaultChannel,{channel}{r}),'
                '[[Channels.name]],0)),0)'),
            hidden=True),
        Col("fee_fixed", "Fixed fees", 9, "money",
            formula=blank(
                "product",
                'IFERROR(INDEX([[Channels.fixed_total]],'
                'MATCH(IF({channel}{r}="",DefaultChannel,{channel}{r}),'
                '[[Channels.name]],0)),0)'),
            hidden=True),
        Col("denom", "Denominator", 9, "number",
            formula=blank("product", '1-TargetRetailMargin-{fee_pct}{r}'),
            hidden=True),
        Col("suggested_raw", "Unrounded", 10, "money",
            formula=blank(
                "product",
                'IF({denom}{r}<=0,"",(({total_cost}{r}+{fee_fixed}{r})/'
                '{denom}{r})*(1+{rush_rate}{r})+'
                'IF({personalisation}{r}="",0,{personalisation}{r}))'),
            hidden=True),
        Col("suggested", "Suggested price", 13, "money",
            formula=blank(
                "product",
                'IF({suggested_raw}{r}="","Check the settings",ROUND(' +
                round_price("{suggested_raw}{r}") + ',2))'),
            note="The base maths, stretched by the size, plus the "
                 "rush surcharge and any flat personalisation fee."),
        Col("price", "Price used", 11, "money",
            formula=blank(
                "product",
                'IF({price_input}{r}<>"",{price_input}{r},'
                'IF(ISNUMBER({suggested}{r}),{suggested}{r},""))')),
        Col("uplift", "Versus the base", 11, "money",
            formula=blank(
                "product",
                'IF(OR({price}{r}="",{base_price}{r}=""),"",'
                'ROUND({price}{r}-{base_price}{r},2))'),
            note="What you are charging extra. Put this in the "
                 "listing as the upgrade price."),
        Col("fees", "Channel fees", 10, "money",
            formula=blank(
                "product",
                'IF({price}{r}="","",ROUND({price}{r}*{fee_pct}{r}+'
                '{fee_fixed}{r},2))'),
            hidden=True),
        Col("profit", "Profit", 11, "money",
            formula=blank(
                "product",
                'IF({price}{r}="","",ROUND({price}{r}-{total_cost}{r}-'
                '{fees}{r},2))'),
            cf=loss_cf()),
        Col("margin", "Margin", 9, "percent",
            formula=blank(
                "product",
                'IF(OR({price}{r}="",{price}{r}=0),"",{profit}{r}/'
                '{price}{r})'),
            cf=[{"type": "cell", "criteria": "<", "value": 0, "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": "=MarginWarning",
                 "fmt": "warn"},
                {"type": "cell", "criteria": ">=", "value": "=MarginWarning",
                 "fmt": "good"}]),
        Col("wage_hour", "You earn an hour", 12, "money",
            formula=blank(
                "product",
                'IF(OR({minutes}{r}=0,{minutes}{r}="",{profit}{r}=""),"",'
                'ROUND(({profit}{r}+{labor_cost}{r})/'
                '({minutes}{r}/60),2))'),
            cf=[{"type": "cell", "criteria": "<", "value": "=MinimumWage",
                 "fmt": "bad"}]),
        Col("alert", "Action", 28, "status",
            formula=blank(
                "product",
                'IF({base_name}{r}="Not found","Check - no such product ID",'
                'IF({price}{r}="","Check - no price yet",'
                'IF({profit}{r}<0,"Loss - this version costs you money",'
                'IF(AND({wage_hour}{r}<>"",{wage_hour}{r}<MinimumWage),'
                '"Below - pays under minimum wage",'
                'IF(AND({extra_mins}{r}>0,{uplift}{r}<=0),'
                '"Check - extra work, no extra charge",'
                'IF({margin}{r}<MarginWarning,'
                '"Low margin - re-price it","OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Variations & Custom Orders",
        subtitle="Sizes, personalisation and rush work, priced from "
                 "the base product.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).variations,
        help_text="Pick a base product, say what changes, and the "
                  "variation prices itself from the size factors "
                  "and the extra minutes.",
        extra_notes=[
            "Personalisation is not free. Twenty minutes of "
            "embroidery at $25 an hour is $8.33, and a $3 "
            "'personalisation fee' is you paying the customer to "
            "give you more work.",
            "Rush orders should cost more. The surcharges on the "
            "Lists tab are a starting point - the real question is "
            "what dropping everything else is worth to you.",
            "Costing a variation and deciding against it is a win. "
            "Set 'Offered' to No and the reasoning stays here for "
            "the next time somebody asks.",
        ],
    )
