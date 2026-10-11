"""`Campaigns` - themed pushes with a budget, a goal and a result."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, id_col, money_cf, rate_cf,
                     status_cf)

NAME = "Campaigns"

_POSTS = 'COUNTIF([[Content Calendar.campaign]],{name}{r})'
_PERF = ('IF(COUNTIF([[Performance.campaign]],{name}{r})=0,"",'
         'SUMIFS([[Performance.%s]],[[Performance.campaign]],{name}{r}))')


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "CampaignPrefix", digits=3),
        Col("name", "Campaign name", 24, "text",
            note="This name appears in the campaign drop down on the "
                 "Content Calendar, so keep it short and unique."),
        Col("objective", "Objective", 15, "text",
            validate=list_from("ObjectiveList")),
        Col("description", "What it is", 34, "wrap"),
        Col("product", "Product / service", 22, "text"),
        Col("audience", "Audience", 18, "text",
            validate=list_from("PersonaNames")),
        Col("message", "Key message", 30, "wrap"),
        Col("platforms", "Platforms", 22, "text"),
        Col("start", "Start date", 12, "date", validate=date_check()),
        Col("end", "End date", 12, "date", validate=date_check()),
        Col("owner", "Owner", 16, "text", validate=list_from("TeamNames")),
        Col("offer", "Offer / hook", 20, "text"),
        Col("landing", "Landing page", 26, "link"),
        Col("cta", "Call to action", 18, "text",
            validate=list_from("CTAList")),
        Col("utm", "Tracking link / UTM", 28, "text",
            note="Add a UTM tag so your website analytics can tell you "
                 "which campaign sent the visitor."),
        Col("budget", "Budget", 12, "money", validate=number_check(0, 1e9),
            cf=bar_cf()),
        Col("spend", "Spent so far", 12, "money",
            validate=number_check(0, 1e9)),
        Col("leads", "Leads", 9, "int",
            validate=number_check(0, 1e7, decimal=False)),
        Col("conversions", "Sales / conversions", 12, "int",
            validate=number_check(0, 1e7, decimal=False)),
        Col("revenue", "Revenue", 13, "money",
            validate=number_check(0, 1e9), cf=money_cf()),
        Col("status", "Status", 13, "status",
            validate=list_from("CampaignStatusList"), cf=status_cf()),
        Col("days", "Length (days)", 11, "int",
            formula='=IF(OR({start}{r}="",{end}{r}=""),"",'
                    '{end}{r}-{start}{r}+1)'),
        Col("posts", "Posts planned", 11, "int",
            formula=blank("name", _POSTS), cf=bar_cf()),
        Col("published", "Published", 10, "int",
            formula=blank("name",
                          'COUNTIFS([[Content Calendar.campaign]],'
                          '{name}{r},[[Content Calendar.status]],'
                          '"Published")')),
        Col("reach", "Reach", 12, "int",
            formula=blank("name", _PERF % "reach"), cf=bar_cf()),
        Col("engagements", "Engagements", 12, "int",
            formula=blank("name", _PERF % "engagements")),
        Col("er", "Engagement rate", 12, "percent",
            formula='=IF(OR({reach}{r}="",{reach}{r}=0),"",'
                    '{engagements}{r}/{reach}{r})',
            cf=rate_cf()),
        Col("clicks", "Clicks", 9, "int",
            formula=blank("name", _PERF % "clicks")),
        Col("remaining", "Budget left", 12, "money",
            formula='=IF({budget}{r}="","",{budget}{r}-N({spend}{r}))',
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}]),
        Col("budget_used", "Budget used", 11, "percent",
            formula='=IF(OR({budget}{r}="",{budget}{r}=0),"",'
                    'N({spend}{r})/{budget}{r})',
            cf=[{"type": "cell", "criteria": ">", "value": 1,
                 "fmt": "bad"},
                {"type": "data_bar", "bar_color": "#F2665E"}]),
        Col("cpl", "Cost per lead", 12, "money",
            formula='=IF(OR({leads}{r}="",{leads}{r}=0),"",'
                    'N({spend}{r})/{leads}{r})',
            cf=[{"type": "cell", "criteria": ">", "value": "=CplTarget",
                 "fmt": "warn"}]),
        Col("cpa", "Cost per sale", 12, "money",
            formula='=IF(OR({conversions}{r}="",{conversions}{r}=0),"",'
                    'N({spend}{r})/{conversions}{r})'),
        Col("conv_rate", "Click to sale", 11, "percent",
            formula='=IF(OR({clicks}{r}="",{clicks}{r}=0),"",'
                    'N({conversions}{r})/{clicks}{r})'),
        Col("roas", "Return on spend", 12, "number1",
            formula='=IF(OR({spend}{r}="",{spend}{r}=0),"",'
                    'N({revenue}{r})/{spend}{r})',
            note="Revenue divided by spend. 3.0 means three back for "
                 "every one spent.",
            cf=[{"type": "cell", "criteria": ">=", "value": "=RoasTarget",
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": "=RoasTarget",
                 "fmt": "warn"}]),
        Col("profit", "Revenue less spend", 13, "money",
            formula='=IF(AND({revenue}{r}="",{spend}{r}=""),"",'
                    'N({revenue}{r})-N({spend}{r}))',
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}]),
        Col("alert", "Action needed", 28, "status",
            formula=blank(
                "name",
                'IF(N({spend}{r})>N({budget}{r}),"Over budget",'
                'IF(AND({status}{r}="Running",{posts}{r}=0),'
                '"Action - running with no posts planned",'
                'IF(AND({end}{r}<>"",{end}{r}<TODAY(),'
                '{status}{r}="Running"),"Overdue - mark it complete",'
                'IF(AND({roas}{r}<>"",{roas}{r}<RoasTarget),'
                '"Review - return below target","OK"))))'),
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]
    return SheetSpec(
        name=NAME,
        title="Campaign Planner",
        subtitle="Give every push a name, a window and a budget, then let "
                 "the grey columns count the posts, pull the results and "
                 "work out the return.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).campaigns,
        help_text="Tag posts with the campaign name on the Content "
                  "Calendar - reach, engagement and clicks then roll up "
                  "here automatically.",
    )
