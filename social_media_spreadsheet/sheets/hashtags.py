"""`Hashtag Bank` - hashtag sets, keywords and search intent."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, dup_cf, id_col, status_cf,
                     yesno)

NAME = "Hashtag Bank"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("tag", "HashtagPrefix", digits=3),
        Col("tag", "Hashtag", 22, "text",
            note="Include the # so you can copy a whole group straight "
                 "into the app."),
        Col("type", "Type", 14, "text",
            validate=list_from("HashtagTypeList"),
            note="Mix broad, niche, community and branded tags in every "
                 "group."),
        Col("group", "Group", 11, "text",
            note="Groups let you reuse a tested set. Type the same group "
                 "name on the Content Calendar."),
        Col("platform", "Platform", 13, "text",
            validate=list_from("PlatformList")),
        Col("pillar", "Content pillar", 16, "text",
            validate=list_from("PillarNames")),
        Col("topic", "Topic", 16, "text"),
        Col("keyword", "Keyword / search term", 26, "text",
            note="The words people actually type into the search box - "
                 "these matter on Pinterest, YouTube and TikTok."),
        Col("related", "Related keywords", 30, "wrap"),
        Col("intent", "Search intent", 15, "text",
            validate=list_from("SearchIntentList")),
        Col("difficulty", "Competition", 12, "text",
            validate=list_from("DifficultyList"), cf=status_cf()),
        Col("volume", "Posts / search volume", 14, "int",
            validate=number_check(0, 10000000000, decimal=False),
            cf=bar_cf()),
        Col("size", "Size band", 13, "text",
            formula='=IF({volume}{r}="","",IF({volume}{r}>=1000000,'
                    '"Huge - very competitive",IF({volume}{r}>=250000,'
                    '"Large",IF({volume}{r}>=25000,"Medium - sweet spot",'
                    '"Small - easy to rank"))))'),
        Col("uses", "Times used (manual)", 12, "int",
            validate=number_check(0, 100000, decimal=False)),
        Col("posts", "Posts using this group", 13, "int",
            formula=blank("group",
                          'COUNTIF([[Content Calendar.hashtag_group]],'
                          '{group}{r})')),
        Col("last_used", "Last used", 12, "date", validate=date_check()),
        Col("rest_days", "Days rested", 10, "int",
            formula=blank("last_used", "TODAY()-{last_used}{r}")),
        Col("banned", "Banned / flagged?", 12, "flag", validate=yesno(),
            note="Some tags get hidden by the platform. Check before a "
                 "campaign.",
            cf=[{"type": "text", "criteria": "containing", "value": "Yes",
                 "fmt": "bad"},
                {"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "good"}]),
        Col("dup_check", "Duplicate check", 13, "status",
            formula=blank("tag",
                          'IF(COUNTIF([[Hashtag Bank.tag]],{tag}{r})>1,'
                          '"Duplicate","OK")'),
            cf=dup_cf()),
        Col("alert", "Action needed", 28, "status",
            formula=blank(
                "tag",
                'IF({banned}{r}="Yes","Remove - flagged by the platform",'
                'IF(AND({volume}{r}<>"",{volume}{r}>=1000000),'
                '"Check - huge tag, pair with niche ones",'
                'IF({posts}{r}=0,"Add - group not on the calendar yet",'
                '"OK")))'),
            cf=alert_cf()),
        Col("notes", "Notes", 24, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Hashtag & Keyword Bank",
        subtitle="Build tested groups of 20-30 tags, keep the keywords "
                 "next to them for search, and rotate the groups so your "
                 "posts do not look identical.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).hashtags,
        help_text="Filter on Group to copy a whole set at once. The size "
                  "band tells you whether a tag is realistic to rank for.",
    )
