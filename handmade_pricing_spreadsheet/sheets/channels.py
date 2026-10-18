"""`Channels` - what each place you sell costs you to sell there.

Etsy, your own site, a craft fair, wholesale, Instagram. Each one
takes a different cut, and a product that makes 48% on your own
website can make 22% on a marketplace once listing, transaction,
processing, ads and regulatory fees have all had a bite.

Every fee on this tab is a number you can edit. Nothing is hard coded
anywhere in this workbook - and nothing here is advice about what the
fees are. Marketplace fees change and they differ by country: check
the current official fee schedule for your shop and type what you
find.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, flag_yes_no_cf

NAME = "Channels"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("name", "Channel", 22, "text",
            note="The name products pick from. Add your own - a "
                 "second marketplace, a gallery, a subscription box."),
        Col("kind", "Fee shape", 18, "text",
            validate=list_from("FeeTypeList"),
            note="Documentation only - the individual fees below are "
                 "what the maths uses."),
        Col("listing_fee", "Listing fee", 11, "money",
            validate=number_check(0, 1000),
            note="Charged per listing, and again each time it "
                 "renews. Spread over the items a listing sells."),
        Col("listing_sales", "Sales per listing", 13, "number1",
            validate=number_check(0, 100000),
            note="How many items one listing sells before it renews. "
                 "2 means each sale carries half a listing fee."),
        Col("transaction", "Transaction fee", 12, "percent",
            validate=number_check(0, 1)),
        Col("processing_pct", "Processing %", 12, "percent",
            validate=number_check(0, 1)),
        Col("processing_fixed", "Processing fixed", 13, "money",
            validate=number_check(0, 100),
            note="The flat part of the payment fee. On a $6 order "
                 "this hurts far more than the percentage."),
        Col("ads_fee", "Offsite ads fee", 12, "percent",
            validate=number_check(0, 1)),
        Col("ads_share", "Orders from ads", 12, "percent",
            validate=number_check(0, 1),
            note="What share of your orders arrive through the "
                 "marketplace's own advertising. Your statements "
                 "will tell you."),
        Col("regulatory", "Regulatory fee", 11, "percent",
            validate=number_check(0, 1)),
        Col("currency_fee", "Currency fee", 11, "percent",
            validate=number_check(0, 1)),
        Col("other_pct", "Other %", 10, "percent",
            validate=number_check(0, 1),
            note="Anything else that scales with the order - a "
                 "gallery commission, an affiliate cut."),
        Col("other_fixed", "Other fixed", 11, "money",
            validate=number_check(0, 10000),
            note="Stall fee per item, packing surcharge, a per-order "
                 "charge."),
        Col("subscription", "Monthly subscription", 14, "money",
            validate=number_check(0, 100000),
            note="A shop plan or a pitch rent. Not charged per "
                 "order - it belongs in your overhead, and the "
                 "reminder column says so."),
        Col("fees_on_shipping", "Fees on postage?", 13, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="Most marketplaces charge their percentage on the "
                 "postage the buyer paid as well as the item."),
        Col("reviewed", "Fees checked", 11, "date", validate=date_check(),
            note="The date you last read the official fee page. Put "
                 "today's date in when you do."),
        Col("url", "Fee schedule", 26, "text",
            note="Where you read it. Paste the link to the official "
                 "page so the next check takes ten seconds."),
        Col("notes", "Notes", 24, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("pct_total", "Percentage fees", 12, "percent",
            formula=blank(
                "name",
                'IF({transaction}{r}="",0,{transaction}{r})+'
                'IF({processing_pct}{r}="",0,{processing_pct}{r})+'
                'IF({regulatory}{r}="",0,{regulatory}{r})+'
                'IF({currency_fee}{r}="",0,{currency_fee}{r})+'
                'IF({other_pct}{r}="",0,{other_pct}{r})+'
                'IF({ads_fee}{r}="",0,{ads_fee}{r})*'
                'IF({ads_share}{r}="",0,{ads_share}{r})'),
            note="Everything that scales with the price, added up, "
                 "with offsite ads weighted by how many orders "
                 "actually come from them. This is the 'r' in the "
                 "pricing formula."),
        Col("fixed_total", "Fixed fees", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({processing_fixed}{r}="",0,'
                '{processing_fixed}{r})+'
                'IF({other_fixed}{r}="",0,{other_fixed}{r})+'
                'IF(OR({listing_fee}{r}="",{listing_sales}{r}="",'
                '{listing_sales}{r}=0),0,{listing_fee}{r}/'
                '{listing_sales}{r}),4)'),
            note="Everything charged per order regardless of price, "
                 "with the listing fee spread over the sales one "
                 "listing makes. This is the 'F'."),
        Col("fee_20", "Fee on a 20 sale", 12, "money",
            formula=blank(
                "name",
                'ROUND(20*{pct_total}{r}+{fixed_total}{r},2)'),
            note="What this channel would take from a sale of 20 in "
                 "your currency. The quickest way to compare two "
                 "places to sell."),
        Col("keep_20", "You keep", 11, "money",
            formula=blank("name", 'ROUND(20-{fee_20}{r},2)')),
        Col("take_rate", "Effective take", 11, "percent",
            formula=blank("name", '{fee_20}{r}/20'),
            cf=[{"type": "cell", "criteria": ">", "value": 0.15,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": 0.06,
                 "fmt": "good"}],
            note="The honest headline number: what this channel "
                 "costs you as a share of a typical sale."),
        Col("products", "Products on it", 11, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Products.channel]],{name}{r})')),
        Col("sales_count", "Sales recorded", 11, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Sales.channel]],{name}{r})')),
        Col("sales_value", "Sold through it", 12, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.revenue]],[[Sales.channel]],'
                '{name}{r}),2)')),
        Col("fees_paid", "Fees actually paid", 13, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.fees]],[[Sales.channel]],'
                '{name}{r}),2)'),
            note="From the Sales tab - what this channel really took "
                 "once the orders were in. Compare it with the "
                 "estimate above."),
        Col("actual_rate", "Actual take", 11, "percent",
            formula=blank(
                "name",
                'IF(OR({sales_value}{r}="",{sales_value}{r}=0),"",'
                '{fees_paid}{r}/{sales_value}{r})'),
            note="If this is well above the estimate, your fee "
                 "settings are out of date."),
        Col("check", "Check", 26, "status",
            formula=blank(
                "name",
                'IF({reviewed}{r}="","Check - fees never verified",'
                'IF(TodayDate-{reviewed}{r}>FeeReviewDays,'
                '"Check - verify the fee schedule",'
                'IF({pct_total}{r}=0,"Check - no fees entered",'
                'IF(AND({actual_rate}{r}<>"",{actual_rate}{r}>'
                '{take_rate}{r}+0.02),'
                '"Check - real fees are higher than these settings",'
                'IF({subscription}{r}>0,'
                '"Note - put the subscription in Overhead too",'
                '"OK")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Sales Channels & Fees",
        subtitle="Every fee, editable, per place you sell - and what "
                 "each one really takes from a sale.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).channels,
        help_text="Check the official fee schedule for your country, "
                  "type what you find, and date the row. Every price "
                  "in the workbook uses these numbers.",
        extra_notes=[
            "VERIFY THESE FIGURES. The defaults this workbook ships "
            "with are a plausible starting point from public fee "
            "pages, not a quotation - fees change, and they are "
            "different in different countries.",
            "The fixed part of a payment fee is what kills small "
            "items. A 30 cent fee on a $4 sticker is 7.5% before "
            "anybody has taken a percentage.",
            "Comparing channels is the point of this tab. The same "
            "product can be worth making for one channel and not "
            "another, and that is a decision you can only make with "
            "the take rate in front of you.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Summary strip under the fee table."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Channel comparison", last)
    sheet.table_head(row, 0, ["Channels set up", "Cheapest to sell on",
                              "Dearest to sell on", "Average take",
                              "Fees paid this year", "Needing a fee check"],
                     [16, 22, 22, 15, 19, 21])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Channels.name]],"<>")',
        '=IFERROR(INDEX([[Channels.name]],MATCH(MIN([[Channels.take_rate]]),'
        '[[Channels.take_rate]],0)),"")',
        '=IFERROR(INDEX([[Channels.name]],MATCH(MAX([[Channels.take_rate]]),'
        '[[Channels.take_rate]],0)),"")',
        '=IFERROR(AVERAGE([[Channels.take_rate]]),"")',
        '=ROUND(SUM([[Channels.fees_paid]]),2)',
        '=COUNTIFS([[Channels.check]],"Check*")',
    ], ["int", "text", "text", "percent", "money", "int"])
