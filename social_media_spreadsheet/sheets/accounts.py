"""`Accounts` - one row per social media account, with a profile checklist."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, flag_yes_no_cf, id_col,
                     rate_cf, status_cf)

NAME = "Accounts"

_YEAR_START = 'DATE(PlanYear,1,1)'
_YEAR_END = 'DATE(PlanYear,12,31)'

#: the six profile checks that make up the optimisation score
CHECKS = [
    ("bio_ok", "Bio / description checked"),
    ("photo_ok", "Profile picture on brand"),
    ("cover_ok", "Cover image current"),
    ("link_ok", "Website link works"),
    ("contact_ok", "Contact details correct"),
    ("branding_ok", "Platform branding done"),
]

_LAST_POST = ('SUMPRODUCT(MAX(([[Content Calendar.platform]]={platform}{r})'
              '*([[Content Calendar.status]]="Published")'
              '*[[Content Calendar.date]]))')


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("account", '"ACC-"', digits=3),
        Col("platform", "Platform", 14, "text",
            validate=list_from("PlatformList"),
            note="Pick from the Lists tab. Everything else in the workbook "
                 "groups on this."),
        Col("account", "Account / profile name", 22, "text"),
        Col("handle", "Username / handle", 20, "text"),
        Col("url", "Profile URL", 30, "link"),
        Col("type", "Account type", 16, "text",
            validate=list_from("AccountTypeList")),
        Col("owner", "Account owner", 16, "text",
            validate=list_from("TeamNames")),
        Col("purpose", "What this account is for", 34, "wrap"),
        Col("audience", "Primary audience", 22, "text"),
        Col("followers", "Followers today", 13, "int",
            validate=number_check(0, 1000000000, decimal=False),
            note="Type today's number. The monthly history lives on the "
                 "Analytics tab.",
            cf=bar_cf()),
        Col("launched", "Account started", 13, "date",
            validate=date_check()),
        Col("target_per_week", "Posts per week target", 11, "int",
            validate=number_check(0, 100, decimal=False)),
    ]

    for key, header in CHECKS:
        columns.append(Col(key, header, 12, "flag",
                           validate=list_from("YesNoList"),
                           cf=flag_yes_no_cf()))

    checks = "+".join(f'(({{{key}}}{{r}}="Yes")*1)' for key, _h in CHECKS)
    columns += [
        Col("profile_score", "Profile score", 11, "percent",
            formula=blank("account", f"({checks})/{len(CHECKS)}"),
            cf=[{"type": "cell", "criteria": ">=", "value": 1,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 1,
                 "fmt": "warn"}],
            note="Share of the six profile checks marked Yes."),
        Col("access", "Access & responsibility notes", 30, "wrap"),
        Col("posts_year", "Posts this year", 11, "int",
            formula=blank("platform",
                          'COUNTIFS([[Content Calendar.platform]],'
                          '{platform}{r},[[Content Calendar.date]],">="&'
                          f'{_YEAR_START},[[Content Calendar.date]],"<="&'
                          f'{_YEAR_END})')),
        Col("published_year", "Published", 10, "int",
            formula=blank("platform",
                          'COUNTIFS([[Content Calendar.platform]],'
                          '{platform}{r},[[Content Calendar.status]],'
                          '"Published",[[Content Calendar.date]],">="&'
                          f'{_YEAR_START},[[Content Calendar.date]],"<="&'
                          f'{_YEAR_END})')),
        Col("reach_year", "Reach this year", 13, "int",
            formula=blank("account",
                          'SUMIFS([[Analytics.reach]],'
                          '[[Analytics.account]],{account}{r},'
                          '[[Analytics.platform]],{platform}{r})'),
            cf=bar_cf()),
        Col("engagements_year", "Engagements", 12, "int",
            formula=blank("account",
                          'SUMIFS([[Analytics.engagements]],'
                          '[[Analytics.account]],{account}{r},'
                          '[[Analytics.platform]],{platform}{r})')),
        Col("er_year", "Engagement rate", 12, "percent",
            formula='=IF(OR({reach_year}{r}="",{reach_year}{r}=0),"",'
                    '{engagements_year}{r}/{reach_year}{r})',
            cf=rate_cf()),
        Col("growth_year", "Follower growth", 12, "int",
            formula=blank("account",
                          'SUMIFS([[Analytics.followers_end]],'
                          '[[Analytics.account]],{account}{r})'
                          '-SUMIFS([[Analytics.followers_start]],'
                          '[[Analytics.account]],{account}{r})')),
        Col("last_post", "Last published", 13, "date",
            formula=('=IF({platform}{r}="","",'
                     f'IF({_LAST_POST}=0,"",{_LAST_POST}))')),
        Col("days_quiet", "Days since", 9, "int",
            formula='=IF({last_post}{r}="","",TODAY()-{last_post}{r})',
            cf=[{"type": "cell", "criteria": ">", "value": "=QuietAccountDays",
                 "fmt": "warn"}]),
        Col("alert", "Action needed", 26, "status",
            formula=blank(
                "account",
                'IF({profile_score}{r}<1,"Action - finish the profile '
                'checklist",IF(AND(ISNUMBER({days_quiet}{r}),'
                '{days_quiet}{r}>QuietAccountDays),"Check - account has '
                'gone quiet",IF(AND({er_year}{r}<>"",'
                '{er_year}{r}<TargetEngagement),"Check - engagement below '
                'target","OK")))'),
            cf=alert_cf()),
        Col("notes", "Notes", 28, "wrap"),
    ]

    # keep the status colouring available for the Yes/No checks
    for col in columns:
        if col.style == "flag" and not col.cf:
            col.cf = status_cf()

    return SheetSpec(
        name=NAME,
        title="Social Media Accounts",
        subtitle="One row per account. The grey columns count your posts, "
                 "pull the reach and engagement from the Analytics tab and "
                 "tell you when a profile needs attention.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=3,
        samples=sample(cfg).accounts,
        help_text="Add every account you post to - the dashboards compare "
                  "them side by side. Mark the six profile checks Yes once "
                  "you have optimised each profile.",
    )
