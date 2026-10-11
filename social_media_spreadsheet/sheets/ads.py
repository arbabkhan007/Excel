"""`Ads` - paid promotion, what it cost and what it brought back."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, id_col, money_cf, status_cf)

NAME = "Ads"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", '"AD-"', digits=3, width=9),
        Col("name", "Ad / ad set name", 24, "text"),
        Col("platform", "Platform", 13, "text",
            validate=list_from("PlatformList")),
        Col("objective", "Objective", 15, "text",
            validate=list_from("AdObjectiveList")),
        Col("campaign", "Linked campaign", 20, "text",
            validate=list_from("CampaignNames")),
        Col("audience", "Audience", 28, "wrap"),
        Col("segment", "Targeting", 24, "wrap"),
        Col("creative", "Creative used", 13, "text",
            validate=list_from("AssetIDs")),
        Col("copy", "Ad copy", 32, "wrap"),
        Col("headline", "Headline", 26, "wrap"),
        Col("landing", "Landing page", 26, "link"),
        Col("start", "Start", 12, "date", validate=date_check()),
        Col("end", "End", 12, "date", validate=date_check()),
        Col("budget", "Budget", 12, "money", validate=number_check(0, 1e9),
            cf=bar_cf()),
        Col("spend", "Spent", 12, "money", validate=number_check(0, 1e9)),
        Col("impressions", "Impressions", 12, "int",
            validate=number_check(0, 1e10, decimal=False)),
        Col("reach", "Reach", 11, "int",
            validate=number_check(0, 1e10, decimal=False)),
        Col("clicks", "Clicks", 10, "int",
            validate=number_check(0, 1e9, decimal=False)),
        Col("conversions", "Sales", 10, "int",
            validate=number_check(0, 1e7, decimal=False)),
        Col("revenue", "Revenue", 13, "money",
            validate=number_check(0, 1e9), cf=money_cf()),
        Col("status", "Status", 12, "status",
            formula='=IF({name}{r}="","",IF({start}{r}="","Planned",'
                    'IF(AND({start}{r}<=TODAY(),OR({end}{r}="",'
                    '{end}{r}>=TODAY())),"Running",IF({start}{r}>TODAY(),'
                    '"Planned","Complete"))))',
            cf=status_cf()),
        Col("days", "Days live", 9, "int",
            formula='=IF({start}{r}="","",IF({end}{r}="",'
                    'TODAY()-{start}{r}+1,MIN(TODAY(),{end}{r})'
                    '-{start}{r}+1))'),
        Col("daily_spend", "Spend per day", 11, "money",
            formula='=IF(OR({days}{r}="",{days}{r}<=0),"",'
                    'N({spend}{r})/{days}{r})'),
        Col("ctr", "Click rate", 10, "percent",
            formula='=IF(OR({impressions}{r}="",{impressions}{r}=0),"",'
                    'N({clicks}{r})/{impressions}{r})',
            note="1% and up is healthy on most platforms.",
            cf=[{"type": "cell", "criteria": ">=", "value": 0.01,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0.005,
                 "fmt": "warn"}]),
        Col("cpc", "Cost per click", 11, "money",
            formula='=IF(OR({clicks}{r}="",{clicks}{r}=0),"",'
                    'N({spend}{r})/{clicks}{r})'),
        Col("cpm", "Cost per 1,000 shown", 12, "money",
            formula='=IF(OR({impressions}{r}="",{impressions}{r}=0),"",'
                    'N({spend}{r})/({impressions}{r}/1000))'),
        Col("conv_rate", "Clicks to sales", 11, "percent",
            formula='=IF(OR({clicks}{r}="",{clicks}{r}=0),"",'
                    'N({conversions}{r})/{clicks}{r})'),
        Col("cpa", "Cost per sale", 11, "money",
            formula='=IF(OR({conversions}{r}="",{conversions}{r}=0),"",'
                    'N({spend}{r})/{conversions}{r})'),
        Col("roas", "Return on spend", 12, "number1",
            formula='=IF(OR({spend}{r}="",{spend}{r}=0),"",'
                    'N({revenue}{r})/{spend}{r})',
            cf=[{"type": "cell", "criteria": ">=", "value": "=RoasTarget",
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 1,
                 "fmt": "bad"}]),
        Col("profit", "Revenue less spend", 13, "money",
            formula='=IF(AND({revenue}{r}="",{spend}{r}=""),"",'
                    'N({revenue}{r})-N({spend}{r}))',
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}]),
        Col("budget_used", "Budget used", 11, "percent",
            formula='=IF(OR({budget}{r}="",{budget}{r}=0),"",'
                    'N({spend}{r})/{budget}{r})',
            cf=[{"type": "cell", "criteria": ">", "value": 1,
                 "fmt": "bad"},
                {"type": "data_bar", "bar_color": "#F2665E"}]),
        Col("alert", "Action needed", 28, "status",
            formula=blank(
                "name",
                'IF(N({spend}{r})>N({budget}{r}),"Over budget - pause or '
                'top up",IF(AND({roas}{r}<>"",{roas}{r}<1),'
                '"Action - losing money, change it",'
                'IF(AND({ctr}{r}<>"",{ctr}{r}<0.005),'
                '"Check - the creative is not landing",'
                'IF(AND({status}{r}="Complete",{revenue}{r}=""),'
                '"Add - record the revenue","OK"))))'),
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Advertising Planner & Tracker",
        subtitle="Boosted posts and proper ad sets. Keep the spend, the "
                 "creative and the result in one row so you can see "
                 "which ads actually pay.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).ads,
        help_text="Boost what already worked organically - check the "
                  "Performance tab for your top posts first.",
    )
