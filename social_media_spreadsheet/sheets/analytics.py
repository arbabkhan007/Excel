"""`Analytics` - the monthly numbers you copy from each platform.

One row per account per month.  This is the only place account level
numbers are typed; every dashboard, report and growth chart reads them
from here.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, dup_cf, id_col, money_cf,
                     rate_cf)

NAME = "Analytics"

_PREV = ('SUMIFS([[Analytics.%s]],[[Analytics.account]],{account}{r},'
         '[[Analytics.month_key]],TEXT(EDATE({month}{r},-1),"yyyy-mm"))')


def _int(hi: float = 1e9):
    return number_check(0, hi, decimal=False)


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("account", '"AN-"', digits=4, width=9),
        Col("month", "Month", 12, "date", validate=date_check(),
            note="Use the 1st of the month - 01/03/2026 means March 2026."),
        Col("month_key", "Month key", 10, "text",
            formula='=IF({month}{r}="","",TEXT({month}{r},"yyyy-mm"))',
            note="Automatic. Every monthly report groups on this."),
        Col("platform", "Platform", 13, "text",
            validate=list_from("PlatformList")),
        Col("account", "Account", 20, "text",
            validate=list_from("AccountNames")),
        Col("followers_start", "Followers at start", 13, "int",
            validate=_int()),
        Col("followers_end", "Followers at end", 13, "int",
            validate=_int(), cf=bar_cf()),
        Col("lost", "Unfollows", 10, "int", validate=_int(1e8)),
        Col("posts", "Posts published", 11, "int", validate=_int(10000),
            note="What the platform reports. The grey column next to it "
                 "counts what your calendar says."),
        Col("reach", "Reach", 12, "int", validate=_int(), cf=bar_cf()),
        Col("impressions", "Impressions", 12, "int", validate=_int()),
        Col("engagements", "Engagements", 12, "int", validate=_int(),
            note="Likes + comments + shares + saves, as the platform "
                 "reports it."),
        Col("likes", "Likes", 10, "int", validate=_int()),
        Col("comments", "Comments", 10, "int", validate=_int()),
        Col("shares", "Shares", 10, "int", validate=_int()),
        Col("saves", "Saves", 10, "int", validate=_int()),
        Col("profile_visits", "Profile visits", 12, "int", validate=_int()),
        Col("clicks", "Link clicks", 11, "int", validate=_int()),
        Col("leads", "Leads / sign-ups", 12, "int", validate=_int(1e8)),
        Col("conversions", "Sales", 10, "int", validate=_int(1e8)),
        Col("revenue", "Revenue", 13, "money",
            validate=number_check(0, 1e9), cf=money_cf()),
        Col("spend", "Spend", 12, "money", validate=number_check(0, 1e9)),
        # ---------------------------------------------------- calculated
        Col("gained", "New followers", 12, "int",
            formula='=IF({followers_end}{r}="","",{followers_end}{r}'
                    '-N({followers_start}{r})+N({lost}{r}))'),
        Col("net_growth", "Net growth", 11, "int",
            formula='=IF({followers_end}{r}="","",{followers_end}{r}'
                    '-N({followers_start}{r}))',
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}]),
        Col("growth_rate", "Growth rate", 11, "percent",
            formula='=IF(OR({followers_start}{r}="",'
                    '{followers_start}{r}=0),"",{net_growth}{r}'
                    '/{followers_start}{r})',
            cf=[{"type": "cell", "criteria": ">=", "value": 0.02,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}]),
        Col("planned", "On your calendar", 12, "int",
            formula='=IF(OR({month}{r}="",{platform}{r}=""),"",'
                    'COUNTIFS([[Content Calendar.platform]],{platform}{r},'
                    '[[Content Calendar.month]],{month_key}{r},'
                    '[[Content Calendar.status]],"Published"))'),
        Col("er", "Engagement rate", 12, "percent",
            formula='=IF(OR({reach}{r}="",{reach}{r}=0),"",'
                    'N({engagements}{r})/{reach}{r})',
            note="Engagements divided by reach - the number to watch.",
            cf=rate_cf()),
        Col("reach_per_post", "Reach per post", 12, "int",
            formula='=IF(OR({posts}{r}="",{posts}{r}=0),"",'
                    'ROUND(N({reach}{r})/{posts}{r},0))'),
        Col("eng_per_post", "Engagements per post", 13, "int",
            formula='=IF(OR({posts}{r}="",{posts}{r}=0),"",'
                    'ROUND(N({engagements}{r})/{posts}{r},0))'),
        Col("reach_rate", "Reach vs followers", 12, "percent",
            formula='=IF(OR({followers_end}{r}="",{followers_end}{r}=0),'
                    '"",N({reach}{r})/{followers_end}{r})',
            note="Over 100% means you reached beyond your own audience."),
        Col("ctr", "Click rate", 11, "percent",
            formula='=IF(OR({reach}{r}="",{reach}{r}=0),"",'
                    'N({clicks}{r})/{reach}{r})'),
        Col("lead_rate", "Clicks to leads", 11, "percent",
            formula='=IF(OR({clicks}{r}="",{clicks}{r}=0),"",'
                    'N({leads}{r})/{clicks}{r})'),
        Col("conv_rate", "Leads to sales", 11, "percent",
            formula='=IF(OR({leads}{r}="",{leads}{r}=0),"",'
                    'N({conversions}{r})/{leads}{r})'),
        Col("cpl", "Cost per lead", 11, "money",
            formula='=IF(OR({leads}{r}="",{leads}{r}=0,{spend}{r}=""),"",'
                    '{spend}{r}/{leads}{r})',
            cf=[{"type": "cell", "criteria": ">", "value": "=CplTarget",
                 "fmt": "warn"}]),
        Col("roas", "Return on spend", 12, "number1",
            formula='=IF(OR({spend}{r}="",{spend}{r}=0),"",'
                    'N({revenue}{r})/{spend}{r})',
            cf=[{"type": "cell", "criteria": ">=", "value": "=RoasTarget",
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": "=RoasTarget",
                 "fmt": "warn"}]),
        Col("prev_reach", "Reach last month", 12, "int",
            formula='=IF({month}{r}="","",' + (_PREV % "reach") + ")",
            hidden=True),
        Col("reach_change", "Reach vs last month", 13, "percent",
            formula='=IF(OR({prev_reach}{r}="",{prev_reach}{r}=0),"",'
                    'N({reach}{r})/{prev_reach}{r}-1)',
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": -0.1,
                 "fmt": "bad"}]),
        Col("prev_followers", "Followers last month", 12, "int",
            formula='=IF({month}{r}="","",'
                    + (_PREV % "followers_end") + ")",
            hidden=True),
        Col("dup_check", "Duplicate check", 13, "status",
            formula='=IF({account}{r}="","",IF(COUNTIFS('
                    '[[Analytics.account]],{account}{r},'
                    '[[Analytics.month_key]],{month_key}{r})>1,'
                    '"Duplicate","OK"))',
            cf=dup_cf()),
        Col("alert", "Action needed", 28, "status",
            formula=blank(
                "account",
                'IF({dup_check}{r}="Duplicate","Duplicate - this month is '
                'entered twice",IF(AND({reach}{r}<>"",{er}{r}<'
                'TargetEngagement),"Check - engagement below target",'
                'IF({net_growth}{r}<0,"Check - followers went down",'
                'IF(AND({posts}{r}<>"",{planned}{r}<>"",'
                '{posts}{r}<>{planned}{r}),"Check - post count does not '
                'match the calendar","OK"))))'),
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Monthly Analytics Data Entry",
        subtitle="Once a month, copy the headline numbers from each "
                 "platform into the cream cells. Everything else in the "
                 "workbook - dashboards, growth, reports - is calculated "
                 "from this one table.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=5,
        samples=sample(cfg).analytics,
        help_text="One row per account per month, always dated the 1st. "
                  "Set a reminder for the first working day of each "
                  "month - it takes about ten minutes.",
    )
