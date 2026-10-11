"""`Repurposing` - one good post, used five more times."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from
from .common import (alert_cf, blank, id_col, lookup_content, rate_cf,
                     status_cf)

NAME = "Repurposing"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("source", '"RP-"', digits=3, width=9),
        Col("source", "Original content ID", 13, "text",
            validate=list_from("ContentIDs"),
            note="The post you are reusing. Pick its C-0000 ID."),
        Col("source_title", "Original post", 30, "text",
            formula=lookup_content("title", "source")),
        Col("source_platform", "Was on", 13, "text",
            formula=lookup_content("platform", "source")),
        Col("source_date", "Original date", 12, "date",
            formula=lookup_content("date", "source")),
        Col("source_er", "How it performed", 13, "percent",
            formula='=IF({source}{r}="","",IFERROR(SUMIFS('
                    '[[Performance.engagements]],[[Performance.content]],'
                    '{source}{r})/SUMIFS([[Performance.reach]],'
                    '[[Performance.content]],{source}{r}),""))',
            note="Repurpose what already worked - this is the engagement "
                 "rate of the original.",
            cf=rate_cf()),
        Col("destination_platform", "Reuse on", 14, "text",
            validate=list_from("PlatformList")),
        Col("new_format", "New format", 14, "text",
            validate=list_from("FormatList")),
        Col("idea", "What you will change", 32, "wrap"),
        Col("angle", "New angle", 28, "wrap"),
        Col("caption", "New caption / title", 30, "wrap"),
        Col("cta", "Call to action", 18, "text",
            validate=list_from("CTAList")),
        Col("asset", "Asset to use", 13, "text",
            validate=list_from("AssetIDs")),
        Col("owner", "Owner", 16, "text", validate=list_from("TeamNames")),
        Col("deadline", "Do it by", 12, "date", validate=date_check()),
        Col("status", "Status", 14, "status",
            validate=list_from("RepurposeStatusList"), cf=status_cf()),
        Col("new_date", "Reposted on", 12, "date", validate=date_check()),
        Col("new_url", "New link", 26, "link"),
        Col("times_used", "Times reused", 11, "int",
            formula=blank("source",
                          'COUNTIF([[Repurposing.source]],{source}{r})')),
        Col("days_late", "Days late", 10, "int",
            formula='=IF(OR({deadline}{r}="",{status}{r}="Published",'
                    '{status}{r}="Not suitable"),"",'
                    'MAX(0,TODAY()-{deadline}{r}))',
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}]),
        Col("alert", "Action needed", 28, "status",
            formula=blank(
                "source",
                'IF({source_title}{r}="Not found","Check - content ID not '
                'on the calendar",IF(AND({status}{r}="Published",'
                '{new_url}{r}=""),"Add - link missing",'
                'IF(AND({days_late}{r}<>"",{days_late}{r}>0),'
                '"Overdue - past the deadline",'
                'IF(AND({status}{r}="Planned",{deadline}{r}=""),'
                '"Add - no deadline set","OK"))))'),
            cf=alert_cf()),
        Col("notes", "Notes", 24, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Content Repurposing Tracker",
        subtitle="Your best reel is a carousel, a pin, a short and a "
                 "newsletter. Plan the second life of every post that "
                 "worked.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=3,
        samples=sample(cfg).repurposing,
        help_text="Sort by 'How it performed' and repurpose from the top "
                  "down. Add the new post to the Content Calendar too, so "
                  "it gets its own ID and results.",
    )
