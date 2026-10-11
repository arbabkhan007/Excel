"""`Trends` - sounds, formats and topics worth jumping on quickly."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, scale_cf, score_validation,
                     status_cf)

NAME = "Trends"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", '"TR-"', digits=3, width=9),
        Col("name", "Trend / sound / format", 24, "text"),
        Col("type", "Type", 14, "text", validate=list_from("TrendTypeList")),
        Col("platform", "Platform", 13, "text",
            validate=list_from("PlatformList")),
        Col("found", "Spotted on", 12, "date", validate=date_check()),
        Col("url", "Example link", 26, "link"),
        Col("audio", "Audio / sound name", 22, "text"),
        Col("topic", "Topic", 20, "text"),
        Col("audience", "Who it suits", 18, "text",
            validate=list_from("PersonaNames")),
        Col("score", "Fit for your brand (1-10)", 13, "rating",
            validate=score_validation(),
            note="Only ride a trend if it fits. A forced trend reads as "
                 "a forced trend.",
            cf=scale_cf()),
        Col("idea", "How you would use it", 32, "wrap"),
        Col("owner", "Owner", 16, "text", validate=list_from("TeamNames")),
        Col("deadline", "Use it before", 12, "date", validate=date_check(),
            note="Trends go stale fast - a fortnight is usually the "
                 "whole window."),
        Col("status", "Status", 15, "status",
            validate=list_from("ContentStatusList"), cf=status_cf()),
        Col("posted", "Posted on", 12, "date", validate=date_check()),
        Col("reach", "Reach it got", 12, "int",
            validate=number_check(0, 1e9, decimal=False)),
        Col("engagement", "Engagements", 12, "int",
            validate=number_check(0, 1e8, decimal=False)),
        Col("er", "Engagement rate", 12, "percent",
            formula='=IF(OR({reach}{r}="",{reach}{r}=0),"",'
                    'N({engagement}{r})/{reach}{r})'),
        Col("age", "Days old", 9, "int",
            formula=blank("found", "TODAY()-{found}{r}")),
        Col("window", "Days left", 10, "int",
            formula='=IF(OR({deadline}{r}="",{posted}{r}<>""),"",'
                    '{deadline}{r}-TODAY())',
            cf=[{"type": "cell", "criteria": "<", "value": 3,
                 "fmt": "warn"},
                {"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}]),
        Col("outcome", "What happened", 28, "wrap"),
        Col("alert", "Action needed", 28, "status",
            formula=blank(
                "name",
                'IF(AND({posted}{r}="",{deadline}{r}<>"",'
                '{deadline}{r}<TODAY()),"Expired - the moment has passed",'
                'IF(AND({posted}{r}="",{score}{r}>=8),'
                '"Action - strong fit, make it now",'
                'IF(AND({posted}{r}="",{age}{r}>TrendWindowDays),'
                '"Check - this has been sitting here a while","OK")))'),
            cf=alert_cf()),
        Col("notes", "Notes", 24, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Trend Tracker",
        subtitle="Catch a trend in the first week or leave it alone. "
                 "Score the fit, set a use-by date, and record what "
                 "actually happened.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).trends,
        help_text="Check your platform's trending page once a week and "
                  "add anything that fits. Anything older than your trend "
                  "window gets flagged.",
    )
