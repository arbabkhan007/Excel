"""`Caption Bank` - reusable hooks, captions and calls to action."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from
from .common import (alert_cf, blank, dup_cf, id_col, rate_cf, status_cf,
                     yesno)

NAME = "Caption Bank"

#: character limits used by the length check (generous, platform published)
LIMIT = ('IF({platform}{r}="Instagram",2200,IF({platform}{r}="TikTok",2200,'
         'IF({platform}{r}="Pinterest",500,IF({platform}{r}="YouTube",5000,'
         'IF({platform}{r}="Facebook",63206,IF({platform}{r}="X (Twitter)",'
         '280,IF({platform}{r}="LinkedIn",3000,IF({platform}{r}="Threads",'
         '500,2200))))))))')


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("title", "CaptionPrefix", digits=3),
        Col("title", "Caption name", 24, "text",
            note="This name appears in the caption drop down on the "
                 "Content Calendar, so keep it unique."),
        Col("platform", "Platform", 13, "text",
            validate=list_from("PlatformList")),
        Col("pillar", "Content pillar", 16, "text",
            validate=list_from("PillarNames")),
        Col("persona", "Written for", 18, "text",
            validate=list_from("PersonaNames")),
        Col("tone", "Tone of voice", 13, "text",
            validate=list_from("ToneList")),
        Col("hook", "Hook (first line)", 34, "wrap",
            note="The only line most people read. Make it earn the tap."),
        Col("body", "Body", 44, "wrap"),
        Col("cta", "Closing call to action", 26, "wrap"),
        Col("full", "Full caption", 48, "wrap",
            note="Paste the finished caption here, ready to copy into the "
                 "app."),
        Col("short", "Short version", 34, "wrap"),
        Col("long", "Long version", 44, "wrap"),
        Col("cta_type", "CTA type", 18, "text",
            validate=list_from("CTAList")),
        Col("emojis", "Emoji guidance", 20, "text"),
        Col("offer", "Offer mentioned", 18, "text"),
        Col("group", "Hashtag group", 13, "text",
            note="Which hashtag group pairs with this caption."),
        Col("reusable", "Reusable?", 10, "flag", validate=yesno(),
            cf=status_cf()),
        Col("approval", "Approval", 15, "status",
            validate=list_from("ApprovalStatusList"), cf=status_cf()),
        Col("created", "Written on", 12, "date", validate=date_check()),
        Col("edited", "Last edited", 12, "date", validate=date_check()),
        Col("chars", "Characters", 10, "int",
            formula=blank("full", "LEN({full}{r})")),
        Col("limit", "Platform limit", 11, "int",
            formula=blank("platform", LIMIT)),
        Col("fits", "Length check", 12, "status",
            formula='=IF(OR({full}{r}="",{limit}{r}=""),"",'
                    'IF({chars}{r}>{limit}{r},"Too long - trim it",'
                    'IF({chars}{r}<40,"Check - very short","OK")))',
            cf=alert_cf() + [{"type": "text", "criteria": "containing",
                              "value": "Too long", "fmt": "bad"}]),
        Col("uses", "Times used", 10, "int",
            formula=blank("title",
                          'COUNTIF([[Content Calendar.caption]],'
                          '{title}{r})')),
        Col("last_used", "Last used", 12, "date",
            formula='=IF({uses}{r}=0,"",IF({title}{r}="","",'
                    'SUMPRODUCT(MAX(([[Content Calendar.caption]]='
                    '{title}{r})*[[Content Calendar.date]]))))'),
        Col("avg_er", "Average engagement", 13, "percent",
            formula='=IF({uses}{r}=0,"",IFERROR(SUMIFS('
                    '[[Performance.engagements]],'
                    '[[Performance.caption]],{title}{r})/SUMIFS('
                    '[[Performance.reach]],[[Performance.caption]],'
                    '{title}{r}),""))',
            cf=rate_cf()),
        Col("dup_check", "Duplicate check", 13, "status",
            formula=blank("title",
                          'IF(COUNTIF([[Caption Bank.title]],{title}{r})>1,'
                          '"Duplicate","OK")'),
            cf=dup_cf()),
        Col("alert", "Action needed", 26, "status",
            formula=blank(
                "title",
                'IF({approval}{r}="Rejected","Rewrite - not approved",'
                'IF({full}{r}="","Add - caption not written yet",'
                'IF({approval}{r}<>"Approved","Check - waiting for '
                'approval",IF({uses}{r}=0,"Use - never used yet","OK"))))'),
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Caption & Copy Bank",
        subtitle="Write it once, use it again. The length check compares "
                 "your caption against the platform limit, and the grey "
                 "columns show how often each caption has been used and "
                 "how it performed.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).captions,
        help_text="Keep the caption name short - it is what you pick from "
                  "the drop down on the Content Calendar.",
    )
