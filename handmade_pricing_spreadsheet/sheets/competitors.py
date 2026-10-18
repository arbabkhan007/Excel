"""`Competitors` - what the market charges, and where you sit in it.

Competitor research is not about copying a price. It is about
knowing whether you are the cheap option, the middle one or the
expensive one - and making that a decision instead of an accident.

Log what comparable pieces sell for, including their postage,
because a $24 item with free shipping beats a $21 one with $5
postage in every buyer's head.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, lookup_product

NAME = "Competitors"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("product", "Your product", 12, "id",
            validate=list_from("ProductIDs"),
            note="Which of your products this one competes with."),
        Col("product_name", "Your product name", 24, "text",
            formula=lookup_product("name")),
        Col("competitor", "Their item", 28, "text"),
        Col("shop", "Shop", 20, "text"),
        Col("channel", "Where", 15, "text",
            validate=list_from("ChannelList")),
        Col("price", "Their price", 11, "money",
            validate=number_check(0, 1000000)),
        Col("shipping", "Their postage", 11, "money",
            validate=number_check(0, 10000),
            note="0 for free shipping - which usually means it is "
                 "in the price."),
        Col("quality", "Quality", 16, "text",
            validate=list_from("QualityList"),
            note="Honestly. 'Above mine' is useful information, "
                 "not an insult."),
        Col("handmade", "Really handmade?", 13, "flag",
            validate=list_from("YesNoNAList"),
            note="A resold import is not your competition, "
                 "whatever the listing says."),
        Col("reviews", "Reviews", 9, "int",
            validate=number_check(0, 10000000, decimal=False),
            note="A rough proxy for how much they sell."),
        Col("checked", "Checked on", 11, "date", validate=date_check()),
        Col("url", "Link", 24, "text"),
        Col("notes", "Notes", 26, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("their_total", "Their total", 11, "money",
            formula=blank(
                "competitor",
                'ROUND(IF({price}{r}="",0,{price}{r})+'
                'IF({shipping}{r}="",0,{shipping}{r}),2)'),
            note="Price plus postage: what the buyer actually "
                 "pays, which is the only number they compare."),
        Col("your_price", "Your price", 11, "money",
            formula=lookup_product("price", missing=""),
            hidden=True),
        Col("your_ship", "Your postage", 11, "money",
            formula=lookup_product("ship_charge", missing=""),
            hidden=True),
        Col("your_total", "Your total", 11, "money",
            formula=blank(
                "competitor",
                'IF({your_price}{r}="","",ROUND({your_price}{r}+'
                'IF({your_ship}{r}="",0,{your_ship}{r}),2))')),
        Col("gap", "Difference", 11, "money",
            formula=blank(
                "competitor",
                'IF(OR({your_total}{r}="",{their_total}{r}=0),"",'
                'ROUND({your_total}{r}-{their_total}{r},2))')),
        Col("gap_pct", "You are", 10, "percent",
            formula=blank(
                "competitor",
                'IF(OR({your_total}{r}="",{their_total}{r}=0),"",'
                '{your_total}{r}/{their_total}{r}-1)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0.35,
                 "fmt": "warn"},
                {"type": "cell", "criteria": "<", "value": -0.25,
                 "fmt": "bad"}],
            note="Well below the market is not a bargain, it is "
                 "an unpaid wage. Well above it needs better "
                 "photographs and a reason."),
        Col("position", "Position", 16, "status",
            formula=blank(
                "competitor",
                'IF({gap_pct}{r}="","",'
                'IF({gap_pct}{r}>0.15,"Dearer",'
                'IF({gap_pct}{r}<-0.15,"Cheaper","Similar")))'),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Cheaper", "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "Similar", "fmt": "good"},
                {"type": "text", "criteria": "containing",
                 "value": "Dearer", "fmt": "info"}]),
        Col("stale", "Days old", 9, "int",
            formula=blank(
                "competitor",
                'IF({checked}{r}="","",TodayDate-{checked}{r})'),
            hidden=True),
        Col("check", "Note", 30, "status",
            formula=blank(
                "competitor",
                'IF({price}{r}="","Add - what do they charge?",'
                'IF({your_price}{r}="","Check - your product has no price",'
                'IF(AND({stale}{r}<>"",{stale}{r}>180),'
                '"Check - research is over six months old",'
                'IF(AND({position}{r}="Cheaper",{quality}{r}="Below mine"),'
                '"Raise - better work, lower price",'
                'IF({position}{r}="Dearer",'
                '"Fine - make sure the listing shows why","OK")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Competitor Research",
        subtitle="What comparable work sells for, postage included, "
                 "and where that puts you.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).competitors,
        help_text="Three or four honest comparisons per product "
                  "are worth more than twenty screenshots. The "
                  "catalog shows the average against your price.",
        extra_notes=[
            "Compare the total the buyer pays. Free shipping is a "
            "price, not a kindness, and buyers compare totals.",
            "Being cheaper than everybody is the most common "
            "mistake in handmade. If your work is better and your "
            "price is lower, the market is not telling you to "
            "stay there.",
            "Check again every few months. Everybody else's "
            "materials went up too.",
        ],
    )
