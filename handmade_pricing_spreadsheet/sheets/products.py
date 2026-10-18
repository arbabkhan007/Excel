"""`Products` - the catalog, and the price every other tab argues about.

This is where it all lands. Materials from the bill of materials (or
from one of the six specialist calculators), labour from the time you
logged, packaging from the profile you picked, the postage you absorb,
your share of the overhead, and the fees the channel takes.

    Cost  = materials + labour + packaging + shipping + overhead
    Price = (Cost + fixed fees) / (1 - target margin - fee rate)

Type a price of your own in "Your price" whenever you want to - the
workbook keeps showing you what the maths said, what you actually
charge, and what the difference is doing to your margin.
"""

from __future__ import annotations

from ..config import CRAFT_SHEETS, Config, Palette
from ..layout import Col, SheetSpec, max_if
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, loss_cf, rank_desc,
                     round_price, status_cf)

NAME = "Products"


def _craft(field: str) -> str:
    """Sum one column of all six specialist calculators for this product."""
    return "+".join(
        f'SUMIFS([[{sheet}.{field}]],[[{sheet}.product]],{{id}}{{r}})'
        for sheet in CRAFT_SHEETS)


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "ProductPrefix", digits=4, width=10,
               note="Generated automatically. Every other tab refers "
                    "to a product by this ID, so filter the catalog "
                    "rather than re-sorting it."),
        Col("name", "Product", 30, "text",
            note="The name a buyer would recognise. 'Cream cable "
                 "beanie' beats 'Hat 3'."),
        Col("sku_input", "SKU override", 14, "text",
            note="Leave blank and the workbook builds one from the "
                 "category and the row number."),
        Col("category", "Category", 22, "text",
            validate=list_from("CategoryList")),
        Col("status", "Status", 14, "status", validate=list_from("StatusList"),
            cf=status_cf()),
        Col("season", "Season", 14, "text", validate=list_from("SeasonList")),
        Col("difficulty", "Difficulty", 14, "text",
            validate=list_from("DifficultyList"),
            note="Feeds the time factor on the Lists tab, which "
                 "stretches your estimated minutes."),
        Col("channel", "Main channel", 16, "text",
            validate=list_from("ChannelList"),
            note="Which fee schedule this product is priced for. "
                 "Change it and watch the suggested price move."),
        Col("pack_profile", "Packaging", 13, "text",
            validate=list_from("PackProfileList"),
            note="Picks up the cost and the weight of that profile "
                 "from the Packaging tab."),
        Col("ship_option", "Shipping option", 30, "text",
            validate=list_from("ShipOptions"),
            note="From the Shipping rate card. Only the part you "
                 "absorb becomes a cost of this product."),
        Col("weight", "Item weight (g)", 12, "number1",
            validate=number_check(0, 1000000),
            note="The finished piece, before packaging. The workbook "
                 "warns you when item plus packaging goes over the "
                 "weight band you chose."),
        Col("stock", "In stock", 9, "int",
            validate=number_check(0, 100000, decimal=False)),
        Col("price_input", "Your price", 11, "money",
            validate=number_check(0, 1000000),
            note="What you actually charge. Leave it blank to use "
                 "the suggested price."),
        Col("target_input", "Target margin", 11, "percent",
            validate=number_check(0, 0.95),
            note="Per product override. Blank uses the target on the "
                 "Settings tab."),
        Col("made", "Made so far", 10, "int",
            validate=number_check(0, 1000000, decimal=False),
            note="Lifetime units made, for your own reference - the "
                 "sales figures come from the Sales tab."),
        Col("launched", "Listed on", 11, "date", validate=date_check()),
        Col("url", "Listing link", 22, "text"),
        Col("notes", "Notes", 26, "wrap"),
        # ----------------------------------------------- what it costs --
        Col("sku", "SKU", 13, "id",
            formula=blank(
                "name",
                'IF({sku_input}{r}<>"",{sku_input}{r},'
                'UPPER(LEFT(IF({category}{r}="","GEN",{category}{r}),3))'
                '&"-"&TEXT(ROW()-4,"000"))'),
            note="Automatic unless you override it. Printed on the "
                 "costing sheet and the order forms."),
        Col("mat_cost", "Materials", 11, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Material Usage.line_cost]],'
                '[[Material Usage.product]],{id}{r})+' + _craft("mat_cost") +
                ',2)'),
            note="The bill of materials plus anything costed on a "
                 "specialist calculator tab. Use one or the other "
                 "for a given product, never both."),
        Col("labor_cost", "Labour", 11, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Labor.cost]],[[Labor.product]],{id}{r})+' +
                _craft("labor_cost") + ',2)')),
        Col("minutes_raw", "Minutes logged", 11, "number1",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Labor.mins_item]],[[Labor.product]],'
                '{id}{r})+' + _craft("minutes") + ',1)'),
            hidden=True),
        Col("time_factor", "Difficulty factor", 10, "number",
            formula=blank(
                "name",
                'IFERROR(INDEX(DiffTimeFactors,MATCH({difficulty}{r},'
                'DiffNames,0)),1)'),
            hidden=True),
        Col("minutes", "Minutes", 10, "number1",
            formula=blank(
                "name",
                'ROUND({minutes_raw}{r}*{time_factor}{r},1)'),
            note="Your logged minutes, stretched by the difficulty "
                 "factor on the Lists tab."),
        Col("hours", "Hours", 9, "number",
            formula=blank("name", 'ROUND({minutes}{r}/60,3)'),
            hidden=True),
        Col("labour_adj", "Labour at difficulty", 11, "money",
            formula=blank(
                "name",
                'ROUND({labor_cost}{r}*{time_factor}{r},2)'),
            note="The labour cost after the difficulty factor. This "
                 "is the figure the price is built on."),
        Col("consumables", "Consumables", 11, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Labor.consumables]],[[Labor.product]],'
                '{id}{r})+' + _craft("other_cost") + ',2)'),
            note="Thread, glue, sandpaper - charged per hour worked "
                 "- plus anything a specialist tab calls an extra."),
        Col("pack_cost", "Packaging", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({pack_profile}{r}="Budget",PackBudget,'
                'IF({pack_profile}{r}="Standard",PackStandard,'
                'IF({pack_profile}{r}="Gift",PackGift,'
                'IF({pack_profile}{r}="Premium",PackPremium,'
                'DefaultPackaging)))),2)')),
        Col("pack_weight", "Packaging weight", 11, "number1",
            formula=blank(
                "name",
                'IF({pack_profile}{r}="Budget",PackWeightBudget,'
                'IF({pack_profile}{r}="Standard",PackWeightStandard,'
                'IF({pack_profile}{r}="Gift",PackWeightGift,'
                'IF({pack_profile}{r}="Premium",PackWeightPremium,0))))'),
            hidden=True),
        Col("ship_pay", "Postage costs", 11, "money",
            formula=blank(
                "name",
                'IFERROR(INDEX([[Shipping.total_cost]],'
                'MATCH({ship_option}{r},[[Shipping.label]],0)),'
                'DefaultShipping)'),
            hidden=True),
        Col("ship_charge", "Postage charged", 11, "money",
            formula=blank(
                "name",
                'IFERROR(INDEX([[Shipping.charge]],'
                'MATCH({ship_option}{r},[[Shipping.label]],0)),0)'),
            hidden=True),
        Col("ship_limit", "Weight band", 11, "number1",
            formula=blank(
                "name",
                'IFERROR(INDEX([[Shipping.max_weight]],'
                'MATCH({ship_option}{r},[[Shipping.label]],0)),"")'),
            hidden=True),
        Col("ship_cost", "Postage you absorb", 12, "money",
            formula=blank(
                "name",
                'ROUND(MAX(0,{ship_pay}{r}-{ship_charge}{r}),2)'),
            note="What the label costs less what the buyer pays. "
                 "Free shipping lands here in full."),
        Col("cost_base", "Cost before overhead", 13, "money",
            formula=blank(
                "name",
                'ROUND({mat_cost}{r}+{labour_adj}{r}+{consumables}{r}+'
                '{pack_cost}{r}+{ship_cost}{r},2)'),
            note="Materials + labour + consumables + packaging + the "
                 "postage you absorb."),
        # ------------------------------------------------- what it sells
        Col("fee_pct", "Fee rate", 10, "percent",
            formula=blank(
                "name",
                'IFERROR(INDEX([[Channels.pct_total]],'
                'MATCH(IF({channel}{r}="",DefaultChannel,{channel}{r}),'
                '[[Channels.name]],0)),0)'),
            note="Every percentage fee that channel charges, added "
                 "up. The 'f' in the pricing formula."),
        Col("fee_ship", "Fees on postage?", 10, "flag",
            formula=blank(
                "name",
                'IFERROR(INDEX([[Channels.fees_on_shipping]],'
                'MATCH(IF({channel}{r}="",DefaultChannel,{channel}{r}),'
                '[[Channels.name]],0)),"No")'),
            hidden=True),
        Col("fee_fixed", "Fixed fees", 10, "money",
            formula=blank(
                "name",
                'ROUND(IFERROR(INDEX([[Channels.fixed_total]],'
                'MATCH(IF({channel}{r}="",DefaultChannel,{channel}{r}),'
                '[[Channels.name]],0)),0)+IF({fee_ship}{r}="Yes",'
                '{ship_charge}{r}*{fee_pct}{r},0),4)'),
            note="Listing and payment fees per order, plus the fee "
                 "the channel charges on the postage the buyer paid."),
        Col("oh_rate", "Overhead rate", 10, "percent",
            formula=blank(
                "name",
                'IF(OverheadMethod="Share of revenue",'
                'IF(RevenueTargetMonth<=0,0,OverheadMonthly/'
                'RevenueTargetMonth),0)'),
            hidden=True),
        Col("oh_fixed", "Overhead in cost", 10, "money",
            formula=blank(
                "name",
                'IF(OverheadMethod="Per hour worked",'
                'ROUND(OverheadPerHour*{hours}{r},2),'
                'IF(OR(OverheadMethod="Per unit made",'
                'OverheadMethod="Per order"),OverheadPerUnit,0))'),
            hidden=True),
        Col("target", "Target margin used", 12, "percent",
            formula=blank(
                "name",
                'IF({target_input}{r}<>"",{target_input}{r},'
                'TargetRetailMargin)')),
        Col("denom", "Pricing denominator", 10, "number",
            formula=blank(
                "name",
                '1-{target}{r}-{fee_pct}{r}-{oh_rate}{r}'),
            hidden=True),
        Col("suggested_raw", "Suggested, unrounded", 12, "money",
            formula=blank(
                "name",
                'IF({denom}{r}<=0,"",({cost_base}{r}+{oh_fixed}{r}+'
                '{fee_fixed}{r})/{denom}{r})'),
            hidden=True),
        Col("suggested", "Suggested price", 12, "money",
            formula=blank(
                "name",
                'IF({suggested_raw}{r}="","Check the settings",ROUND(' +
                round_price("{suggested_raw}{r}") + ',2))'),
            note="(Cost + fixed fees) / (1 - target margin - fee "
                 "rate), rounded by the rule on Settings. The one "
                 "number this whole workbook exists to produce."),
        Col("price", "Price used", 11, "money",
            formula=blank(
                "name",
                'IF({price_input}{r}<>"",{price_input}{r},'
                'IF(ISNUMBER({suggested}{r}),{suggested}{r},""))'),
            note="Your price if you typed one, otherwise the "
                 "suggested price. Everything below uses this."),
        Col("overhead", "Overhead", 10, "money",
            formula=blank(
                "name",
                'IF({price}{r}="","",ROUND({oh_fixed}{r}+{price}{r}*'
                '{oh_rate}{r},2))')),
        Col("total_cost", "Total cost", 11, "money",
            formula=blank(
                "name",
                'IF({price}{r}="","",ROUND({cost_base}{r}+'
                '{overhead}{r},2))'),
            note="C = materials + labour + packaging + shipping + "
                 "overhead. Fees are not a cost of making it, so "
                 "they are shown separately."),
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
            cf=loss_cf(),
            note="What is left after everything - including paying "
                 "yourself for the hours. A small number here is "
                 "still a wage; a negative one is a hobby."),
        Col("margin", "Margin", 9, "percent",
            formula=blank(
                "name",
                'IF(OR({price}{r}="",{price}{r}=0),"",{profit}{r}/'
                '{price}{r})'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": "=MarginWarning",
                 "fmt": "warn"},
                {"type": "cell", "criteria": ">=", "value": "=MarginWarning",
                 "fmt": "good"}],
            note="Profit / price. The share of every dollar you keep."),
        Col("markup", "Markup", 9, "percent",
            formula=blank(
                "name",
                'IF(OR({total_cost}{r}="",{total_cost}{r}=0),"",'
                '{profit}{r}/{total_cost}{r})'),
            note="Profit / cost. The same product, a bigger looking "
                 "number - 100% markup is a 50% margin. Never "
                 "confuse the two again."),
        Col("profit_hour", "Profit an hour", 11, "money",
            formula=blank(
                "name",
                'IF(OR({hours}{r}="",{hours}{r}=0,{profit}{r}=""),"",'
                'ROUND({profit}{r}/{hours}{r},2))'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}],
            note="On top of the wage already inside the labour cost. "
                 "This is what the business earns over and above "
                 "paying you."),
        Col("wage_hour", "You earn an hour", 11, "money",
            formula=blank(
                "name",
                'IF(OR({hours}{r}="",{hours}{r}=0,{profit}{r}=""),"",'
                'ROUND(({profit}{r}+{labour_adj}{r})/{hours}{r},2))'),
            cf=[{"type": "cell", "criteria": "<", "value": "=MinimumWage",
                 "fmt": "bad"}],
            note="Wage plus profit, divided by the hours. If this is "
                 "under your local minimum wage, the price is wrong "
                 "- not your speed."),
        Col("min_price", "Never sell below", 12, "money",
            formula=blank(
                "name",
                'IF(1-{fee_pct}{r}-{oh_rate}{r}<=0,"",ROUND(({cost_base}{r}'
                '+{oh_fixed}{r}+{fee_fixed}{r})/(1-{fee_pct}{r}-'
                '{oh_rate}{r}),2))'),
            note="The price at which this product makes exactly "
                 "nothing. Sale prices below this are costing you "
                 "money to give away."),
        Col("keystone", "Keystone (x2)", 11, "money",
            formula=blank(
                "name",
                'IF({total_cost}{r}="","",ROUND({total_cost}{r}*2,2))'),
            hidden=True,
            note="The old craft fair rule of thumb, for comparison "
                 "only. It ignores fees entirely, which is why it "
                 "leaves makers short."),
        Col("wholesale", "Wholesale price", 12, "money",
            formula=blank(
                "name",
                'IF({price}{r}="","",ROUND({price}{r}*'
                '(1-WholesaleDiscount),2))'),
            note="Your retail price less the wholesale discount on "
                 "Settings. The Wholesale Calculator shows whether "
                 "it survives."),
        Col("ws_profit", "Wholesale profit", 12, "money",
            formula=blank(
                "name",
                'IF({wholesale}{r}="","",ROUND({wholesale}{r}-'
                '{cost_base}{r}-{oh_fixed}{r},2))'),
            cf=loss_cf(),
            note="No marketplace fees on a wholesale order, but no "
                 "retail price either."),
        Col("ws_margin", "Wholesale margin", 11, "percent",
            formula=blank(
                "name",
                'IF(OR({wholesale}{r}="",{wholesale}{r}=0),"",'
                '{ws_profit}{r}/{wholesale}{r})'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<",
                 "value": "=TargetWholesaleMargin", "fmt": "warn"},
                {"type": "cell", "criteria": ">=",
                 "value": "=TargetWholesaleMargin", "fmt": "good"}]),
        Col("comp_price", "Competitors average", 12, "money",
            formula=blank(
                "name",
                'IFERROR(ROUND(AVERAGEIFS([[Competitors.price]],'
                '[[Competitors.product]],{id}{r}),2),"")')),
        Col("comp_gap", "Versus the market", 11, "percent",
            formula=blank(
                "name",
                'IF(OR({comp_price}{r}="",{comp_price}{r}=0,'
                '{price}{r}=""),"",{price}{r}/{comp_price}{r}-1)'),
            note="Above the market is not a problem if the work "
                 "justifies it - but you should know you are there."),
        Col("sold", "Units sold", 10, "int",
            formula=blank(
                "name",
                'SUMIFS([[Sales.qty]],[[Sales.product]],{id}{r})')),
        Col("revenue", "Revenue", 11, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.revenue]],[[Sales.product]],'
                '{id}{r}),2)')),
        Col("actual_profit", "Actual profit", 11, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.profit]],[[Sales.product]],'
                '{id}{r}),2)'),
            cf=loss_cf()),
        Col("actual_margin", "Actual margin", 11, "percent",
            formula=blank(
                "name",
                'IF(OR({revenue}{r}="",{revenue}{r}=0),"",'
                '{actual_profit}{r}/{revenue}{r})'),
            note="What really happened, discounts and refunds "
                 "included. Compare it with the margin you planned."),
        Col("last_sold", "Last sold", 11, "date",
            formula=blank(
                "name",
                'IF(' + max_if("[[Sales.date]]",
                               "[[Sales.product]]={id}{r}") +
                '=0,"",' + max_if("[[Sales.date]]",
                                  "[[Sales.product]]={id}{r}") + ')')),
        Col("last_priced", "Price changed", 11, "date",
            formula=blank(
                "name",
                'IF(' + max_if("[[Price History.date]]",
                               "[[Price History.product]]={id}{r}") +
                '=0,IF({launched}{r}="","",{launched}{r}),' +
                max_if("[[Price History.date]]",
                       "[[Price History.product]]={id}{r}") + ')'),
            hidden=True),
        Col("price_age", "Days since review", 11, "int",
            formula=blank(
                "name",
                'IF({last_priced}{r}="","",TodayDate-{last_priced}{r})'),
            cf=[{"type": "cell", "criteria": ">", "value": "=PriceReviewDays",
                 "fmt": "warn"}],
            hidden=True),
        rank_desc("profit", guard='{name}{r}=""', header="Profit rank"),
        rank_desc("margin", guard='{name}{r}=""', header="Margin rank"),
        rank_desc("revenue", guard='{name}{r}=""', header="Revenue rank"),
        rank_desc("profit_hour", guard='{name}{r}=""',
                  header="Hourly rank"),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "name",
                'IF({cost_base}{r}=0,"Add - no costs entered yet",'
                'IF({price}{r}="","Check - no price yet",'
                'IF({profit}{r}<0,"Loss - every sale costs you money",'
                'IF(AND({wage_hour}{r}<>"",{wage_hour}{r}<MinimumWage),'
                '"Below - pays under minimum wage",'
                'IF({margin}{r}<MarginWarning,"Low margin - re-price it",'
                'IF(AND({ship_limit}{r}<>"",{weight}{r}<>"",'
                '{weight}{r}+{pack_weight}{r}>{ship_limit}{r}),'
                '"Check - too heavy for that postage band",'
                'IF(AND({price_age}{r}<>"",'
                '{price_age}{r}>PriceReviewDays),'
                '"Due - price has not been reviewed",'
                'IF(AND({comp_gap}{r}<>"",{comp_gap}{r}>0.5),'
                '"Check - well above the market","OK")))))))) '),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Product Catalog",
        subtitle="Every product, every cost, the price the maths "
                 "asks for and the margin you actually make.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).products,
        help_text="Give a product a name, a category, a packaging "
                  "profile and a channel. Cost it on the Materials "
                  "and Labor tabs (or a specialist tab) and the "
                  "price appears here.",
        extra_notes=[
            "The suggested price is (cost + fixed fees) / (1 - "
            "target margin - fee rate). Dividing by what is left "
            "is the step most pricing advice skips, and it is why "
            "'cost x 2' leaves you short on a marketplace.",
            "Margin and markup are not the same thing. 50% margin "
            "is 100% markup. Both are shown so nobody has to "
            "remember which one the Facebook group meant.",
            "'You earn an hour' is the number to look at when a "
            "price feels too high. If it is under minimum wage, "
            "the price is not too high - it is too low.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Summary strip under the catalog."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Catalog summary", last)
    sheet.table_head(row, 0, ["Products", "Active", "Average price",
                              "Average cost", "Average margin",
                              "Losing money", "Below target margin",
                              "Under minimum wage", "Average profit an hour"],
                     [12, 10, 14, 14, 15, 14, 19, 19, 21])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Products.name]],"<>")',
        '=COUNTIFS([[Products.status]],"Active")+'
        'COUNTIFS([[Products.status]],"Made to order")',
        '=IFERROR(ROUND(AVERAGE([[Products.price]]),2),"")',
        '=IFERROR(ROUND(AVERAGE([[Products.total_cost]]),2),"")',
        '=IFERROR(AVERAGE([[Products.margin]]),"")',
        '=COUNTIFS([[Products.profit]],"<0")',
        '=COUNTIFS([[Products.margin]],"<"&MarginWarning)',
        '=COUNTIFS([[Products.wage_hour]],"<"&MinimumWage,'
        '[[Products.name]],"<>")',
        '=IFERROR(ROUND(AVERAGE([[Products.profit_hour]]),2),"")',
    ], ["int", "int", "money", "money", "percent", "int", "int", "int",
        "money"])
