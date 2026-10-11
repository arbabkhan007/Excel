"""`Content Calendar` - the master register. Every post lives on this tab.

Everything else in the workbook either feeds this sheet (captions, hashtag
groups, assets, campaigns) or reads it (the dashboards, the workflow view,
the reports and the printable briefs).
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from
from .common import (alert_cf, blank, dup_cf, id_col, month_col, rate_cf,
                     status_cf, week_col, weekday_col)

NAME = "Content Calendar"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("title", "ContentPrefix",
               note="C-0001, C-0002 ... Quote this ID on the Performance, "
                    "Repurposing and platform tabs."),
        Col("date", "Publish date", 13, "date", validate=date_check(),
            note="The day it goes live. Draft and approval deadlines are "
                 "worked out from this."),
        Col("slot", "Time slot", 12, "text",
            validate=list_from("TimeSlotList")),
        month_col("date"),
        week_col("date"),
        weekday_col("date"),
        Col("platform", "Platform", 13, "text",
            validate=list_from("PlatformList")),
        Col("title", "Content title", 34, "text",
            note="Short and recognisable - this is what shows up in the "
                 "reports and the top content tables."),
        Col("description", "What it is about", 38, "wrap"),
        Col("pillar", "Content pillar", 16, "text",
            validate=list_from("PillarNames")),
        Col("format", "Format", 14, "text",
            validate=list_from("FormatList")),
        Col("campaign", "Campaign", 20, "text",
            validate=list_from("CampaignNames")),
        Col("objective", "Objective", 15, "text",
            validate=list_from("ObjectiveList")),
        Col("funnel", "Funnel stage", 14, "text",
            validate=list_from("FunnelStageList")),
        Col("persona", "Written for", 18, "text",
            validate=list_from("PersonaNames")),
        Col("caption", "Caption (from the bank)", 24, "text",
            validate=list_from("CaptionTitles")),
        Col("caption_check", "Caption ready", 12, "status",
            formula='=IF({title}{r}="","",IF({caption}{r}="",'
                    '"Add - no caption",IFERROR(INDEX('
                    '[[Caption Bank.approval]],MATCH({caption}{r},'
                    '[[Caption Bank.title]],0)),"Check - not in bank")))',
            cf=alert_cf() + status_cf()),
        Col("hashtag_group", "Hashtag group", 14, "text",
            note="Type a group name you have used on the Hashtag Bank "
                 "tab, e.g. 'Group 1'."),
        Col("hashtag_count", "Tags in group", 11, "int",
            formula='=IF({hashtag_group}{r}="","",COUNTIF('
                    '[[Hashtag Bank.group]],{hashtag_group}{r}))',
            cf=[{"type": "cell", "criteria": "<", "value": 3,
                 "fmt": "warn"}]),
        Col("asset", "Creative asset", 13, "text",
            validate=list_from("AssetIDs")),
        Col("asset_check", "Asset approved", 12, "status",
            formula='=IF({title}{r}="","",IF({asset}{r}="",'
                    '"Add - no asset",IFERROR(INDEX([[Assets.approval]],'
                    'MATCH({asset}{r},[[Assets.id]],0)),'
                    '"Check - not in library")))',
            cf=alert_cf() + status_cf()),
        Col("cta", "Call to action", 18, "text",
            validate=list_from("CTAList")),
        Col("owner", "Owner", 16, "text",
            validate=list_from("TeamNames")),
        Col("status", "Status", 15, "status",
            validate=list_from("ContentStatusList"), cf=status_cf(),
            note="Move it along: Idea > Brief > In production > Draft > "
                 "Awaiting review > Approved > Scheduled > Published."),
        Col("draft_due", "Draft due", 12, "date",
            formula=blank("date", "{date}{r}-ProductionDays")),
        Col("approve_due", "Approval due", 12, "date",
            formula=blank("date", "{date}{r}-ApprovalDays")),
        Col("days_to_go", "Days to go", 10, "int",
            formula=blank("date", "{date}{r}-TODAY()"),
            cf=[{"type": "cell", "criteria": "between", "minimum": 0,
                 "maximum": 2, "fmt": "warn"},
                {"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "muted"}]),
        Col("url", "Published link", 28, "link"),
        Col("published", "Live?", 10, "status",
            formula='=IF({title}{r}="","",IF({status}{r}="Published",'
                    'IF({url}{r}="","Add - link missing","Yes"),"No"))',
            cf=alert_cf() + [{"type": "text", "criteria": "containing",
                              "value": "Yes", "fmt": "good"}]),
        Col("reach", "Reach", 11, "int",
            formula='=IF({id}{r}="","",IF(COUNTIF([[Performance.content]],'
                    '{id}{r})=0,"",SUMIFS([[Performance.reach]],'
                    '[[Performance.content]],{id}{r})))',
            note="Pulled from the Performance tab - enter the numbers "
                 "there once and they appear here."),
        Col("engagements", "Engagements", 12, "int",
            formula='=IF({reach}{r}="","",SUMIFS('
                    '[[Performance.engagements]],'
                    '[[Performance.content]],{id}{r}))'),
        Col("er", "Engagement rate", 12, "percent",
            formula='=IF(OR({reach}{r}="",{reach}{r}=0),"",'
                    '{engagements}{r}/{reach}{r})',
            cf=rate_cf()),
        Col("clicks", "Clicks", 9, "int",
            formula='=IF({reach}{r}="","",SUMIFS([[Performance.clicks]],'
                    '[[Performance.content]],{id}{r}))'),
        Col("repurposed", "Repurposed", 10, "int",
            formula='=IF({id}{r}="","",COUNTIF([[Repurposing.source]],'
                    '{id}{r}))'),
        Col("day_key", "Day key", 11, "number",
            formula='=IF(OR({title}{r}="",{date}{r}=""),"",'
                    '{date}{r}+ROW()/100000)',
            hidden=True,
            note="Hidden helper: lets the Month View grid list the posts "
                 "that fall on each day without sorting the sheet."),
        Col("queue_key", "Queue key", 11, "number",
            formula='=IF(OR({title}{r}="",{date}{r}=""),"",'
                    'IF(OR({status}{r}="Published",'
                    '{status}{r}="Cancelled",{status}{r}="Archived"),"",'
                    'IF({date}{r}<TODAY(),"",{date}{r}+ROW()/100000)))',
            hidden=True,
            note="Hidden helper: lets the dashboard list the next posts "
                 "due without sorting the sheet."),
        Col("late_key", "Late key", 11, "number",
            formula='=IF(OR({title}{r}="",{date}{r}=""),"",'
                    'IF(OR({status}{r}="Published",'
                    '{status}{r}="Cancelled",{status}{r}="Archived"),"",'
                    'IF({date}{r}>=TODAY(),"",{date}{r}+ROW()/100000)))',
            hidden=True,
            note="Hidden helper: feeds the overdue list on the "
                 "Dashboard and the Workflow tab."),
        Col("dup_check", "Duplicate check", 13, "status",
            formula=blank("title",
                          'IF(COUNTIFS([[Content Calendar.title]],'
                          '{title}{r},[[Content Calendar.platform]],'
                          '{platform}{r},[[Content Calendar.date]],'
                          '{date}{r})>1,"Duplicate","OK")'),
            cf=dup_cf()),
        Col("alert", "Action needed", 30, "status",
            formula=blank(
                "title",
                'IF(AND({date}{r}<>"",{date}{r}<TODAY(),'
                '{status}{r}<>"Published",{status}{r}<>"Cancelled",'
                '{status}{r}<>"Archived"),"Overdue - date has passed",'
                'IF(AND({status}{r}="Published",{url}{r}=""),'
                '"Add - publish link missing",'
                'IF(AND({date}{r}=TODAY(),{status}{r}<>"Published"),'
                '"Today - due to go live",'
                'IF(AND({draft_due}{r}<>"",{draft_due}{r}<TODAY(),'
                'OR({status}{r}="Idea",{status}{r}="Brief",'
                '{status}{r}="In production")),"Late - draft not ready",'
                'IF(AND({days_to_go}{r}>=0,{days_to_go}{r}<=3,'
                'OR({status}{r}="Draft",{status}{r}="Awaiting review")),'
                '"Due - needs approval",'
                'IF(AND({title}{r}<>"",{caption}{r}=""),'
                '"Add - caption missing","OK"))))))'),
            cf=alert_cf()),
        Col("notes", "Notes", 28, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="12-Month Content Calendar",
        subtitle="One row per post, on any platform, for the whole year. "
                 "Filter by month, platform, pillar or status - the grey "
                 "columns keep the deadlines, links and results up to "
                 "date for you.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).calendar,
        help_text="Type the publish date and a title - the ID, month, "
                  "week, deadlines and alerts fill themselves in. Use the "
                  "filter arrows rather than re-sorting so the content IDs "
                  "stay matched to the other tabs.",
    )
