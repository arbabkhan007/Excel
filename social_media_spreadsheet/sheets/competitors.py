"""`Competitors` - what the people in your niche are doing."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, bar_cf, blank, id_col, rate_cf

NAME = "Competitors"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", '"COM-"', digits=3, width=10),
        Col("name", "Competitor", 20, "text"),
        Col("platform", "Platform", 13, "text",
            validate=list_from("PlatformList")),
        Col("url", "Profile link", 26, "link"),
        Col("niche", "Their niche", 22, "text"),
        Col("audience", "Followers", 12, "int",
            validate=number_check(0, 1e9, decimal=False), cf=bar_cf()),
        Col("frequency", "How often they post", 18, "text"),
        Col("pillars", "Their content themes", 30, "wrap"),
        Col("formats", "Formats they use", 24, "wrap"),
        Col("best_post", "Their best post", 30, "wrap"),
        Col("engagement", "Their engagement rate", 13, "percent",
            validate=number_check(0, 1),
            note="Engagements divided by followers on a handful of recent "
                 "posts - near enough is good enough.",
            cf=rate_cf()),
        Col("campaign", "Campaign they are running", 26, "wrap"),
        Col("questions", "Questions their audience asks", 28, "wrap"),
        Col("trending", "What is working for them", 26, "wrap"),
        Col("news", "Platform / industry news", 26, "wrap"),
        Col("gap", "Gap you could fill", 28, "wrap",
            note="The most valuable column on the sheet."),
        Col("opportunity", "Your move", 28, "wrap"),
        Col("reviewed", "Last reviewed", 12, "date", validate=date_check()),
        Col("source", "Where you looked", 22, "text"),
        Col("vs_you", "Versus your rate", 13, "percent",
            formula='=IF({engagement}{r}="","",'
                    '{engagement}{r}-TargetEngagement)',
            note="Their engagement rate minus your target rate.",
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<=", "value": 0,
                 "fmt": "good"}]),
        Col("ideas", "Ideas taken from them", 11, "int",
            formula=blank("name",
                          'COUNTIF([[Idea Bank.competitor]],{name}{r})')),
        Col("days_since", "Days since review", 11, "int",
            formula=blank("reviewed", "TODAY()-{reviewed}{r}"),
            cf=[{"type": "cell", "criteria": ">", "value": 60,
                 "fmt": "warn"}]),
        Col("alert", "Action needed", 28, "status",
            formula=blank(
                "name",
                'IF(OR({reviewed}{r}="",{days_since}{r}>60),'
                '"Action - review this account again",'
                'IF({gap}{r}="","Add - no gap identified yet",'
                'IF({ideas}{r}=0,"Add - no ideas captured from them",'
                '"OK")))'),
            cf=alert_cf()),
        Col("notes", "Notes", 24, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Competitor Research",
        subtitle="Watch four or five accounts properly rather than forty "
                 "badly. Look for the gap, then fill it.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).competitors,
        help_text="Review every month or two. Note their best post and "
                  "the questions people ask in their comments - that is "
                  "free content research.",
    )
