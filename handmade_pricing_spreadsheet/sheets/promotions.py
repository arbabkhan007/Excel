"""`Promotions` - what a sale costs, and how much more you have to sell.

A 20% discount on a 45% margin product does not cost you 20%. It
costs you 44% of your profit, and you need to sell 80% more units to
stand still. That arithmetic is on this tab, next to what each code
actually did.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, loss_cf, status_cf

NAME = "Promotions"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("code", "Code", 14, "id",
            note="The code buyers type, or a short name for a sale "
                 "with no code. Sales rows pick from these."),
        Col("name", "Promotion", 26, "text"),
        Col("kind", "Type", 18, "text", validate=list_from("PromoTypeList")),
        Col("value", "Value", 10, "number",
            validate=number_check(0, 100000),
            note="A percentage as a decimal (0.15 for 15% off), or "
                 "an amount for 'Amount off'."),
        Col("scope", "Applies to", 18, "text",
            note="'Everything', a category, or a product ID. Free "
                 "text - it is for your own records."),
        Col("start", "Starts", 11, "date", validate=date_check()),
        Col("ends", "Ends", 11, "date", validate=date_check()),
        Col("min_spend", "Minimum spend", 12, "money",
            validate=number_check(0, 100000)),
        Col("limit", "Use limit", 10, "int",
            validate=number_check(0, 1000000, decimal=False)),
        Col("base_margin", "Normal margin", 12, "percent",
            validate=number_check(0, 0.95),
            note="The margin these products normally make. Blank "
                 "uses the target margin on Settings."),
        Col("notes", "Notes", 26, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("margin_used", "Margin used", 11, "percent",
            formula=blank(
                "code",
                'IF({base_margin}{r}<>"",{base_margin}{r},'
                'TargetRetailMargin)'),
            hidden=True),
        Col("discount_pct", "Discount", 10, "percent",
            formula=blank(
                "code",
                'IF({kind}{r}="Percent off",{value}{r},'
                'IF({kind}{r}="Free shipping","",'
                'IF(OR({min_spend}{r}="",{min_spend}{r}=0),"",'
                '{value}{r}/{min_spend}{r})))'),
            hidden=True,
            note="An amount off is turned into a percentage using "
                 "the minimum spend, which is the only honest way "
                 "to compare the two."),
        Col("margin_after", "Margin after it", 12, "percent",
            formula=blank(
                "code",
                'IF({discount_pct}{r}="","",'
                '({margin_used}{r}-{discount_pct}{r})/'
                '(1-{discount_pct}{r}))'),
            cf=[{"type": "cell", "criteria": "<", "value": 0, "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": "=MarginWarning",
                 "fmt": "warn"},
                {"type": "cell", "criteria": ">=", "value": "=MarginWarning",
                 "fmt": "good"}],
            note="What is left of the margin once the discount is "
                 "off both the price and the profit."),
        Col("profit_lost", "Profit given away", 12, "percent",
            formula=blank(
                "code",
                'IF(OR({discount_pct}{r}="",{margin_used}{r}=0),"",'
                '{discount_pct}{r}/{margin_used}{r})'),
            cf=[{"type": "cell", "criteria": ">", "value": 0.5,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 0.3,
                 "fmt": "warn"}],
            note="The share of your profit the discount hands "
                 "over. 20% off a 45% margin gives away 44% of the "
                 "profit on every order."),
        Col("uplift_needed", "Extra sales needed", 14, "percent",
            formula=blank(
                "code",
                'IF({discount_pct}{r}="","Never - it loses money",'
                'IF({margin_used}{r}-{discount_pct}{r}<=0,'
                '"Never - it loses money",'
                '{discount_pct}{r}/({margin_used}{r}-'
                '{discount_pct}{r})))'),
            note="How much more you have to sell to make the same "
                 "profit as before the sale. This is the number "
                 "that stops a 30% off weekend."),
        Col("orders", "Orders", 9, "int",
            formula=blank(
                "code",
                'COUNTIFS([[Sales.promo]],{code}{r})')),
        Col("revenue", "Revenue", 11, "money",
            formula=blank(
                "code",
                'ROUND(SUMIFS([[Sales.revenue]],[[Sales.promo]],'
                '{code}{r}),2)')),
        Col("given", "Discount given", 12, "money",
            formula=blank(
                "code",
                'ROUND(SUMIFS([[Sales.discount]],[[Sales.promo]],'
                '{code}{r}),2)')),
        Col("profit", "Profit made", 11, "money",
            formula=blank(
                "code",
                'ROUND(SUMIFS([[Sales.profit]],[[Sales.promo]],'
                '{code}{r}),2)'),
            cf=loss_cf()),
        Col("actual_margin", "Actual margin", 11, "percent",
            formula=blank(
                "code",
                'IF(OR({revenue}{r}="",{revenue}{r}=0),"",'
                '{profit}{r}/{revenue}{r})')),
        Col("avg_order", "Average order", 11, "money",
            formula=blank(
                "code",
                'IF({orders}{r}=0,"",ROUND({revenue}{r}/'
                '{orders}{r},2))'),
            hidden=True),
        Col("status", "Status", 13, "status",
            formula=blank(
                "code",
                'IF(AND({start}{r}<>"",{start}{r}>TodayDate),"Scheduled",'
                'IF(AND({ends}{r}<>"",{ends}{r}<TodayDate),"Expired",'
                'IF(AND({limit}{r}<>"",{orders}{r}>={limit}{r}),'
                '"Used up","Active")))'),
            cf=status_cf()),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "code",
                'IF({value}{r}="","Add - how much off?",'
                'IF(AND({margin_after}{r}<>"",{margin_after}{r}<0),'
                '"Loss - this discount sells below cost",'
                'IF(AND({profit_lost}{r}<>"",{profit_lost}{r}>0.5),'
                '"Check - over half the profit given away",'
                'IF({status}{r}="Expired","Ended - review what it did",'
                'IF(AND({orders}{r}=0,{status}{r}="Active"),'
                '"Waiting - no orders yet","OK")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Discounts & Promotions",
        subtitle="What each sale costs in margin, and how many "
                 "extra orders it has to bring.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).promotions,
        help_text="Set the discount and your normal margin, and "
                  "the workbook shows what is left and how much "
                  "more you would have to sell.",
        extra_notes=[
            "Discounting is the fastest way to turn a good margin "
            "into a busy fortnight with nothing to show for it. "
            "The 'extra sales needed' column is the reality check.",
            "Free shipping is a discount too. Price it as one: "
            "work out what you absorb and treat it as money off.",
            "A code that brings new buyers who come back is worth "
            "more than the margin it cost. Keep notes here - the "
            "numbers alone will not tell you that.",
        ],
    )
