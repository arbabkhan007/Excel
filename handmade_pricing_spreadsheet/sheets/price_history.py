"""`Price History` - every price change, and whether it worked.

Raising a price feels like a risk until you can see what happened
last time you did it. Log the change, the reason and what the
product cost that week, and this tab keeps the story: what the
margin was before, what it became, and how long the price has stood
since.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, loss_cf, lookup_product, month_col

NAME = "Price History"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("date", "Date", 11, "date", validate=date_check()),
        Col("product", "Product", 12, "id",
            validate=list_from("ProductIDs")),
        Col("product_name", "Product name", 24, "text",
            formula=lookup_product("name")),
        Col("old_price", "Old price", 11, "money",
            validate=number_check(0, 1000000)),
        Col("new_price", "New price", 11, "money",
            validate=number_check(0, 1000000)),
        Col("cost_then", "Cost at the time", 13, "money",
            validate=number_check(0, 1000000),
            note="What the product cost to make that week. Blank "
                 "uses today's cost from the catalog, which is "
                 "fine for a recent change and misleading for an "
                 "old one."),
        Col("reason", "Reason", 22, "text",
            validate=list_from("PriceReasonList")),
        Col("channel", "Channel", 15, "text",
            validate=list_from("ChannelList")),
        Col("sold_before", "Sold in the 30 days before", 16, "int",
            validate=number_check(0, 1000000, decimal=False),
            note="Optional, but this is what turns a price change "
                 "into evidence."),
        Col("sold_after", "Sold in the 30 days after", 16, "int",
            validate=number_check(0, 1000000, decimal=False)),
        Col("notes", "Notes", 28, "wrap"),
        # ------------------------------------------------- the automatics --
        month_col("date"),
        Col("change", "Change", 10, "money",
            formula=blank(
                "date",
                'IF(OR({old_price}{r}="",{new_price}{r}=""),"",'
                'ROUND({new_price}{r}-{old_price}{r},2))'),
            cf=loss_cf()),
        Col("change_pct", "Change", 9, "percent",
            formula=blank(
                "date",
                'IF(OR({old_price}{r}="",{old_price}{r}=0,'
                '{new_price}{r}=""),"",{new_price}{r}/'
                '{old_price}{r}-1)')),
        Col("cost_used", "Cost used", 11, "money",
            formula=blank(
                "date",
                'IF({cost_then}{r}<>"",{cost_then}{r},'
                'IFERROR(INDEX([[Products.total_cost]],MATCH({product}{r},'
                '[[Products.id]],0)),""))'),
            hidden=True),
        Col("margin_before", "Margin before", 12, "percent",
            formula=blank(
                "date",
                'IF(OR({old_price}{r}="",{old_price}{r}=0,'
                '{cost_used}{r}=""),"",({old_price}{r}-'
                '{cost_used}{r})/{old_price}{r})')),
        Col("margin_after", "Margin after", 12, "percent",
            formula=blank(
                "date",
                'IF(OR({new_price}{r}="",{new_price}{r}=0,'
                '{cost_used}{r}=""),"",({new_price}{r}-'
                '{cost_used}{r})/{new_price}{r})'),
            cf=[{"type": "cell", "criteria": "<", "value": 0, "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": "=MarginWarning",
                 "fmt": "warn"},
                {"type": "cell", "criteria": ">=", "value": "=MarginWarning",
                 "fmt": "good"}]),
        Col("margin_move", "Margin moved", 11, "percent",
            formula=blank(
                "date",
                'IF(OR({margin_before}{r}="",{margin_after}{r}=""),"",'
                '{margin_after}{r}-{margin_before}{r})'),
            hidden=True),
        Col("volume_move", "Sales moved", 10, "percent",
            formula=blank(
                "date",
                'IF(OR({sold_before}{r}="",{sold_before}{r}=0,'
                '{sold_after}{r}=""),"",{sold_after}{r}/'
                '{sold_before}{r}-1)'),
            note="Did the units fall? A 15% rise that cost you 5% "
                 "of the volume made you money."),
        Col("revenue_move", "Revenue effect", 12, "percent",
            formula=blank(
                "date",
                'IF(OR({volume_move}{r}="",{change_pct}{r}=""),"",'
                '(1+{change_pct}{r})*(1+{volume_move}{r})-1)'),
            cf=[{"type": "cell", "criteria": "<", "value": 0, "fmt": "bad"},
                {"type": "cell", "criteria": ">=", "value": 0,
                 "fmt": "good"}],
            note="Price change and volume change together. "
                 "Positive means the change was worth making."),
        Col("seq", "Change number", 9, "int",
            formula=blank(
                "product",
                'COUNTIFS({product_first}:{product}{r},{product}{r})'),
            hidden=True),
        Col("days_since", "Days since", 10, "int",
            formula=blank(
                "date", 'TodayDate-{date}{r}'),
            hidden=True),
        Col("verdict", "Verdict", 26, "status",
            formula=blank(
                "date",
                'IF({change}{r}="","Add - old and new price",'
                'IF({margin_after}{r}<0,"Loss - priced below cost",'
                'IF({revenue_move}{r}="","Watch - no volume data yet",'
                'IF({revenue_move}{r}>0.02,"Worked - revenue up",'
                'IF({revenue_move}{r}<-0.02,'
                '"Hurt - revenue down","Neutral - no real change")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Price History",
        subtitle="Every change, the reason for it, and what it did "
                 "to the margin and the volume.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).price_history,
        help_text="Log the change on the day you make it. The "
                  "catalog uses the latest date here to tell you "
                  "when a price is going stale.",
        extra_notes=[
            "Record the units sold before and after and you stop "
            "guessing about price sensitivity. Most makers "
            "discover their buyers never noticed.",
            "A price rise that loses 10% of the volume on a 45% "
            "margin product still leaves you better off. The "
            "revenue effect column does that sum.",
            "Put the reason in. 'Yarn went up 18%' is the "
            "difference between a price rise you can explain and "
            "one that feels like greed six months later.",
        ],
    )
