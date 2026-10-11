"""`Performance` - the numbers for one individual post.

Type the content ID and the metrics; the title, platform, pillar, campaign
and date are pulled from the Content Calendar, and the rates are worked
out for you.  This is the table the top-content lists, the pillar
comparison and the platform planners all read.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import (alert_cf, bar_cf, blank, dup_cf, id_col, lookup_content,
                     money_cf, rate_cf, status_cf)

NAME = "Performance"


def _int(hi: float = 1e9):
    return number_check(0, hi, decimal=False)


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("content", '"PF-"', digits=4, width=9),
        Col("content", "Content ID", 12, "text",
            validate=list_from("ContentIDs"),
            note="Pick the post from the Content Calendar - everything "
                 "grey fills itself in."),
        Col("title", "Content title", 30, "text",
            formula=lookup_content("title")),
        Col("date", "Published", 12, "date",
            formula=lookup_content("date")),
        Col("month", "Month", 10, "text",
            formula=lookup_content("month")),
        Col("platform", "Platform", 13, "text",
            formula=lookup_content("platform")),
        Col("format", "Format", 13, "text",
            formula=lookup_content("format")),
        Col("pillar", "Pillar", 15, "text",
            formula=lookup_content("pillar")),
        Col("campaign", "Campaign", 18, "text",
            formula=lookup_content("campaign")),
        Col("caption", "Caption used", 20, "text",
            formula=lookup_content("caption")),
        Col("persona", "Written for", 16, "text",
            formula=lookup_content("persona")),
        Col("funnel", "Funnel stage", 13, "text",
            formula=lookup_content("funnel")),
        Col("slot", "Time slot", 12, "text",
            formula=lookup_content("slot")),
        Col("weekday", "Day", 9, "text",
            formula=lookup_content("weekday")),
        # ------------------------------------------------- typed numbers
        Col("reach", "Reach", 11, "int", validate=_int(), cf=bar_cf(),
            note="How many people saw it. If the platform only gives you "
                 "impressions, use that and be consistent."),
        Col("impressions", "Impressions", 11, "int", validate=_int()),
        Col("likes", "Likes", 10, "int", validate=_int(1e8)),
        Col("comments", "Comments", 10, "int", validate=_int(1e8)),
        Col("shares", "Shares", 10, "int", validate=_int(1e8)),
        Col("saves", "Saves", 10, "int", validate=_int(1e8)),
        Col("clicks", "Link clicks", 11, "int", validate=_int(1e8)),
        Col("watch", "Avg watch time (sec)", 13, "number1",
            validate=number_check(0, 100000)),
        Col("conversions", "Sales from it", 11, "int",
            validate=_int(1e7)),
        Col("revenue", "Revenue", 13, "money",
            validate=number_check(0, 1e9), cf=money_cf()),
        Col("target_rate", "Target rate", 11, "percent",
            validate=number_check(0, 1),
            note="Leave blank to use the workbook target from the "
                 "Settings tab."),
        # ---------------------------------------------------- calculated
        Col("engagements", "Engagements", 12, "int",
            formula='=IF({reach}{r}="","",N({likes}{r})+N({comments}{r})'
                    '+N({shares}{r})+N({saves}{r}))'),
        Col("er", "Engagement rate", 12, "percent",
            formula='=IF(OR({reach}{r}="",{reach}{r}=0),"",'
                    '{engagements}{r}/{reach}{r})',
            cf=rate_cf()),
        Col("save_rate", "Save rate", 10, "percent",
            formula='=IF(OR({reach}{r}="",{reach}{r}=0),"",'
                    'N({saves}{r})/{reach}{r})',
            note="Saves are the strongest signal on most platforms."),
        Col("share_rate", "Share rate", 10, "percent",
            formula='=IF(OR({reach}{r}="",{reach}{r}=0),"",'
                    'N({shares}{r})/{reach}{r})'),
        Col("comment_rate", "Comment rate", 11, "percent",
            formula='=IF(OR({reach}{r}="",{reach}{r}=0),"",'
                    'N({comments}{r})/{reach}{r})'),
        Col("ctr", "Click rate", 10, "percent",
            formula='=IF(OR({reach}{r}="",{reach}{r}=0),"",'
                    'N({clicks}{r})/{reach}{r})'),
        Col("conv_rate", "Clicks to sales", 11, "percent",
            formula='=IF(OR({clicks}{r}="",{clicks}{r}=0),"",'
                    'N({conversions}{r})/{clicks}{r})'),
        Col("rev_per_1k", "Revenue per 1,000 reached", 14, "money",
            formula='=IF(OR({reach}{r}="",{reach}{r}=0),"",'
                    'N({revenue}{r})/({reach}{r}/1000))'),
        Col("target_used", "Target used", 10, "percent",
            formula='=IF({reach}{r}="","",IF({target_rate}{r}="",'
                    'TargetEngagement,{target_rate}{r}))',
            hidden=True),
        Col("vs_target", "Versus target", 11, "percent",
            formula='=IF(OR({er}{r}="",{target_used}{r}=""),"",'
                    '{er}{r}-{target_used}{r})',
            cf=[{"type": "cell", "criteria": ">=", "value": 0,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "warn"}]),
        Col("rank_key", "Rank key", 10, "number1",
            formula='=IF({er}{r}="","",{er}{r}+ROW()/1000000)',
            hidden=True,
            note="Hidden helper: makes the engagement rate unique so the "
                 "top-post lookups never tie."),
        Col("rank", "Rank by engagement", 12, "int",
            formula='=IF({er}{r}="","",RANK({rank_key}{r},'
                    '[[Performance.rank_key]]))',
            cf=[{"type": "cell", "criteria": "<=", "value": 3,
                 "fmt": "good"}]),
        Col("verdict", "Verdict", 20, "status",
            formula='=IF({er}{r}="","",IF({er}{r}>={target_used}{r}*1.5,'
                    '"Top performer",IF({er}{r}>={target_used}{r},'
                    '"Above target",IF({er}{r}>={target_used}{r}*0.6,'
                    '"Below target","Weak - review it"))))',
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Top performer", "fmt": "bold_good"},
                {"type": "text", "criteria": "containing",
                 "value": "Above target", "fmt": "good"},
                {"type": "text", "criteria": "containing",
                 "value": "Below target", "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "Weak", "fmt": "bad"}]),
        Col("dup_check", "Duplicate check", 12, "status",
            formula=blank("content",
                          'IF(COUNTIF([[Performance.content]],'
                          '{content}{r})>1,"Duplicate","OK")'),
            cf=dup_cf()),
        Col("alert", "Action needed", 26, "status",
            formula=blank(
                "content",
                'IF({title}{r}="Not found","Check - content ID not on the '
                'calendar",IF({reach}{r}="","Add - no numbers entered '
                'yet",IF({dup_check}{r}="Duplicate","Duplicate - this '
                'post is entered twice",IF({verdict}{r}="Top performer",'
                '"Repurpose this one","OK"))))'),
            cf=alert_cf() + status_cf() + [
                {"type": "text", "criteria": "containing",
                 "value": "Repurpose", "fmt": "info"}]),
        Col("notes", "What you learned", 28, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Content Performance Tracker",
        subtitle="Post by post results. Fill in the cream columns a few "
                 "days after publishing and the workbook ranks every "
                 "post, finds your best pillar and tells you what to "
                 "make more of.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=3,
        samples=sample(cfg).performance,
        help_text="Wait 48-72 hours before recording numbers. Be "
                  "consistent about whether you use reach or impressions "
                  "- the rates only compare if you are.",
    )
