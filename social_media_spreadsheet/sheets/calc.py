"""
`Calc` - hidden helper sheet.

The 2007-2016 formula set has no SORT or FILTER, so the twelve month trend
series, the platform and pillar breakdowns and the "top ten" tables are
pre-computed here with plain SUMIFS / COUNTIFS and then pulled into the
dashboards and reports with LARGE + INDEX/MATCH.

Nothing here needs touching: the tab is hidden and every cell is a locked
formula.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec
from . import lists

NAME = "Calc"

# Dashboard period window (DashStart / DashEnd are defined on the Dashboard
# tab - one month, or the whole year).
_DASH = ('[[Content Calendar.date]],">="&DashStart,'
         '[[Content Calendar.date]],"<="&DashEnd')
_DASH_AN = ('[[Analytics.month]],">="&DashStart,'
            '[[Analytics.month]],"<="&DashEnd')
#: year window on the Content Calendar
_YEAR = ('[[Content Calendar.date]],">="&DATE(PlanYear,1,1),'
         '[[Content Calendar.date]],"<="&DATE(PlanYear,12,31)')


def _mirror(sheet: str, key: str) -> str:
    return f'=IF([[{sheet}.{key}#]]="","",[[{sheet}.{key}#]])'


def _from_list(name: str) -> str:
    ref = lists.ref(name)
    return f'=IF({ref}="","",{ref})'


def spec(cfg: Config) -> SheetSpec:
    rows = max(cfg.row_count("Performance"),
               cfg.row_count("Content Calendar"),
               cfg.row_count("Analytics"),
               cfg.row_count("Campaigns"), 60)

    columns = [
        # ------------------------------------------- 12 month series ------
        Col("m_start", "Month start", 12, "date",
            formula='=IF(ROW()-4>12,"",DATE(PlanYear,ROW()-4,1))'),
        Col("m_end", "Month end", 12, "date",
            formula='=IF({m_start}{r}="","",EOMONTH({m_start}{r},0))'),
        Col("m_key", "Month key", 10, "text",
            formula='=IF({m_start}{r}="","",TEXT({m_start}{r},"yyyy-mm"))'),
        Col("m_label", "Month", 9, "text",
            formula='=IF({m_start}{r}="","",TEXT({m_start}{r},"mmm"))'),
        Col("m_planned", "Posts planned", 11, "int",
            formula='=IF({m_key}{r}="","",COUNTIFS('
                    '[[Content Calendar.month]],{m_key}{r}))'),
        Col("m_published", "Published", 10, "int",
            formula='=IF({m_key}{r}="","",COUNTIFS('
                    '[[Content Calendar.month]],{m_key}{r},'
                    '[[Content Calendar.status]],"Published"))'),
        Col("m_reach", "Reach", 12, "int",
            formula='=IF({m_key}{r}="","",SUMIFS([[Analytics.reach]],'
                    '[[Analytics.month_key]],{m_key}{r}))'),
        Col("m_impressions", "Impressions", 12, "int",
            formula='=IF({m_key}{r}="","",SUMIFS('
                    '[[Analytics.impressions]],[[Analytics.month_key]],'
                    '{m_key}{r}))'),
        Col("m_engagements", "Engagements", 12, "int",
            formula='=IF({m_key}{r}="","",SUMIFS('
                    '[[Analytics.engagements]],[[Analytics.month_key]],'
                    '{m_key}{r}))'),
        Col("m_er", "Engagement rate", 12, "percent",
            formula='=IF(OR({m_reach}{r}="",{m_reach}{r}=0),"",'
                    '{m_engagements}{r}/{m_reach}{r})'),
        Col("m_followers", "Followers", 11, "int",
            formula='=IF({m_key}{r}="","",SUMIFS('
                    '[[Analytics.followers_end]],[[Analytics.month_key]],'
                    '{m_key}{r}))'),
        Col("m_growth", "Follower growth", 11, "int",
            formula='=IF({m_key}{r}="","",SUMIFS('
                    '[[Analytics.net_growth]],[[Analytics.month_key]],'
                    '{m_key}{r}))'),
        Col("m_clicks", "Clicks", 10, "int",
            formula='=IF({m_key}{r}="","",SUMIFS([[Analytics.clicks]],'
                    '[[Analytics.month_key]],{m_key}{r}))'),
        Col("m_leads", "Leads", 10, "int",
            formula='=IF({m_key}{r}="","",SUMIFS([[Analytics.leads]],'
                    '[[Analytics.month_key]],{m_key}{r}))'),
        Col("m_conversions", "Sales", 10, "int",
            formula='=IF({m_key}{r}="","",SUMIFS('
                    '[[Analytics.conversions]],[[Analytics.month_key]],'
                    '{m_key}{r}))'),
        Col("m_revenue", "Revenue", 13, "money",
            formula='=IF({m_key}{r}="","",SUMIFS([[Analytics.revenue]],'
                    '[[Analytics.month_key]],{m_key}{r}))'),
        Col("m_spend", "Spend", 12, "money",
            formula='=IF({m_key}{r}="","",SUMIFS([[Budget.actual]],'
                    '[[Budget.month_key]],{m_key}{r}))'),
        Col("m_profit", "Revenue less spend", 13, "money",
            formula='=IF({m_key}{r}="","",{m_revenue}{r}-{m_spend}{r})'),
        Col("m_target", "Target posts", 10, "int",
            formula='=IF({m_key}{r}="","",MonthlyPostTarget)'),
        # --------------------------------------------- platform block -----
        Col("p_name", "Platform", 14, "text",
            formula=_from_list("PlatformList")),
        Col("p_posts_m", "Posts (month)", 11, "int",
            formula='=IF({p_name}{r}="","",COUNTIFS('
                    '[[Content Calendar.platform]],{p_name}{r},'
                    + _DASH + "))"),
        Col("p_published_m", "Published (month)", 11, "int",
            formula='=IF({p_name}{r}="","",COUNTIFS('
                    '[[Content Calendar.platform]],{p_name}{r},'
                    '[[Content Calendar.status]],"Published",'
                    + _DASH + "))"),
        Col("p_reach_m", "Reach (month)", 12, "int",
            formula='=IF({p_name}{r}="","",SUMIFS([[Analytics.reach]],'
                    '[[Analytics.platform]],{p_name}{r},'
                    + _DASH_AN + "))"),
        Col("p_eng_m", "Engagements (month)", 12, "int",
            formula='=IF({p_name}{r}="","",SUMIFS('
                    '[[Analytics.engagements]],[[Analytics.platform]],'
                    '{p_name}{r},' + _DASH_AN + "))"),
        Col("p_er_m", "Engagement rate (month)", 12, "percent",
            formula='=IF(OR({p_reach_m}{r}="",{p_reach_m}{r}=0),"",'
                    '{p_eng_m}{r}/{p_reach_m}{r})'),
        Col("p_followers", "Followers (month)", 11, "int",
            formula='=IF({p_name}{r}="","",SUMIFS('
                    '[[Analytics.followers_end]],[[Analytics.platform]],'
                    '{p_name}{r},' + _DASH_AN + "))"),
        Col("p_growth_m", "Growth (month)", 11, "int",
            formula='=IF({p_name}{r}="","",SUMIFS('
                    '[[Analytics.net_growth]],[[Analytics.platform]],'
                    '{p_name}{r},' + _DASH_AN + "))"),
        Col("p_clicks_m", "Clicks (month)", 10, "int",
            formula='=IF({p_name}{r}="","",SUMIFS([[Analytics.clicks]],'
                    '[[Analytics.platform]],{p_name}{r},'
                    + _DASH_AN + "))"),
        Col("p_revenue_m", "Revenue (month)", 13, "money",
            formula='=IF({p_name}{r}="","",SUMIFS([[Analytics.revenue]],'
                    '[[Analytics.platform]],{p_name}{r},'
                    + _DASH_AN + "))"),
        Col("p_posts_y", "Posts (year)", 11, "int",
            formula='=IF({p_name}{r}="","",COUNTIFS('
                    '[[Content Calendar.platform]],{p_name}{r},'
                    + _YEAR + "))"),
        Col("p_reach_y", "Reach (year)", 12, "int",
            formula='=IF({p_name}{r}="","",SUMIFS([[Analytics.reach]],'
                    '[[Analytics.platform]],{p_name}{r}))'),
        Col("p_eng_y", "Engagements (year)", 12, "int",
            formula='=IF({p_name}{r}="","",SUMIFS('
                    '[[Analytics.engagements]],[[Analytics.platform]],'
                    '{p_name}{r}))'),
        Col("p_er_y", "Engagement rate (year)", 12, "percent",
            formula='=IF(OR({p_reach_y}{r}="",{p_reach_y}{r}=0),"",'
                    '{p_eng_y}{r}/{p_reach_y}{r})'),
        Col("p_growth_y", "Growth (year)", 11, "int",
            formula='=IF({p_name}{r}="","",SUMIFS('
                    '[[Analytics.net_growth]],[[Analytics.platform]],'
                    '{p_name}{r}))'),
        Col("p_key", "Platform rank key", 12, "number",
            formula='=IF({p_name}{r}="","",IF({p_reach_y}{r}=0,"",'
                    '{p_reach_y}{r}+ROW()/1000000))'),
        # ------------------------------------------------ pillar block ----
        Col("pl_name", "Pillar", 16, "text",
            formula=_mirror("Content Pillars", "name")),
        Col("pl_posts", "Pillar posts", 11, "int",
            formula='=IF({pl_name}{r}="","",COUNTIFS('
                    '[[Content Calendar.pillar]],{pl_name}{r},'
                    + _YEAR + "))"),
        Col("pl_reach", "Pillar reach", 12, "int",
            formula='=IF({pl_name}{r}="","",SUMIFS([[Performance.reach]],'
                    '[[Performance.pillar]],{pl_name}{r}))'),
        Col("pl_eng", "Pillar engagements", 12, "int",
            formula='=IF({pl_name}{r}="","",SUMIFS('
                    '[[Performance.engagements]],[[Performance.pillar]],'
                    '{pl_name}{r}))'),
        Col("pl_er", "Pillar engagement rate", 12, "percent",
            formula='=IF(OR({pl_reach}{r}="",{pl_reach}{r}=0),"",'
                    '{pl_eng}{r}/{pl_reach}{r})'),
        Col("pl_key", "Pillar rank key", 12, "number",
            formula='=IF({pl_er}{r}="","",{pl_er}{r}+ROW()/1000000)'),
        # ------------------------------------------------ format block ----
        Col("fm_name", "Format", 14, "text",
            formula=_from_list("FormatList")),
        Col("fm_posts", "Format posts", 11, "int",
            formula='=IF({fm_name}{r}="","",COUNTIF('
                    '[[Performance.format]],{fm_name}{r}))'),
        Col("fm_reach", "Format reach", 12, "int",
            formula='=IF({fm_name}{r}="","",SUMIFS([[Performance.reach]],'
                    '[[Performance.format]],{fm_name}{r}))'),
        Col("fm_eng", "Format engagements", 12, "int",
            formula='=IF({fm_name}{r}="","",SUMIFS('
                    '[[Performance.engagements]],[[Performance.format]],'
                    '{fm_name}{r}))'),
        Col("fm_er", "Format engagement rate", 12, "percent",
            formula='=IF(OR({fm_reach}{r}="",{fm_reach}{r}=0),"",'
                    '{fm_eng}{r}/{fm_reach}{r})'),
        Col("fm_key", "Format rank key", 12, "number",
            formula='=IF({fm_posts}{r}=0,"",IF({fm_er}{r}="","",'
                    '{fm_er}{r}+ROW()/1000000))'),
        # ------------------------------------------- day and slot block ---
        Col("dw_day", "Weekday", 9, "text",
            formula='=IF(ROW()-4>7,"",TEXT(DATE(2024,1,ROW()-4),"ddd"))'),
        Col("dw_posts", "Posts that day", 11, "int",
            formula='=IF({dw_day}{r}="","",COUNTIF('
                    '[[Performance.weekday]],{dw_day}{r}))'),
        Col("dw_reach", "Reach that day", 12, "int",
            formula='=IF({dw_day}{r}="","",SUMIFS([[Performance.reach]],'
                    '[[Performance.weekday]],{dw_day}{r}))'),
        Col("dw_eng", "Engagements that day", 12, "int",
            formula='=IF({dw_day}{r}="","",SUMIFS('
                    '[[Performance.engagements]],'
                    '[[Performance.weekday]],{dw_day}{r}))'),
        Col("dw_er", "Engagement rate that day", 12, "percent",
            formula='=IF(OR({dw_reach}{r}="",{dw_reach}{r}=0),"",'
                    '{dw_eng}{r}/{dw_reach}{r})'),
        Col("dw_key", "Weekday rank key", 12, "number",
            formula='=IF({dw_posts}{r}=0,"",IF({dw_er}{r}="","",'
                    '{dw_er}{r}+ROW()/1000000))'),
        Col("sl_slot", "Time slot", 13, "text",
            formula=_from_list("TimeSlotList")),
        Col("sl_posts", "Posts in slot", 11, "int",
            formula='=IF({sl_slot}{r}="","",COUNTIF('
                    '[[Performance.slot]],{sl_slot}{r}))'),
        Col("sl_reach", "Reach in slot", 12, "int",
            formula='=IF({sl_slot}{r}="","",SUMIFS([[Performance.reach]],'
                    '[[Performance.slot]],{sl_slot}{r}))'),
        Col("sl_eng", "Engagements in slot", 12, "int",
            formula='=IF({sl_slot}{r}="","",SUMIFS('
                    '[[Performance.engagements]],[[Performance.slot]],'
                    '{sl_slot}{r}))'),
        Col("sl_er", "Engagement rate in slot", 12, "percent",
            formula='=IF(OR({sl_reach}{r}="",{sl_reach}{r}=0),"",'
                    '{sl_eng}{r}/{sl_reach}{r})'),
        Col("sl_key", "Slot rank key", 12, "number",
            formula='=IF({sl_posts}{r}=0,"",IF({sl_er}{r}="","",'
                    '{sl_er}{r}+ROW()/1000000))'),
        # ------------------------------------------- performance mirror ---
        Col("pf_title", "Post", 28, "text",
            formula=_mirror("Performance", "title")),
        Col("pf_platform", "Post platform", 13, "text",
            formula=_mirror("Performance", "platform")),
        Col("pf_pillar", "Post pillar", 14, "text",
            formula=_mirror("Performance", "pillar")),
        Col("pf_date", "Post date", 12, "date",
            formula=_mirror("Performance", "date")),
        Col("pf_month", "Post month", 10, "text",
            formula=_mirror("Performance", "month")),
        Col("pf_reach", "Post reach", 11, "int",
            formula=_mirror("Performance", "reach")),
        Col("pf_er", "Post engagement rate", 12, "percent",
            formula=_mirror("Performance", "er")),
        Col("pf_key_m", "Top post key (period)", 13, "number",
            formula='=IF({pf_title}{r}="","",IF({pf_date}{r}="","",'
                    'IF(OR({pf_date}{r}<DashStart,{pf_date}{r}>DashEnd),'
                    '"",IF({pf_er}{r}="","",{pf_er}{r}+ROW()/1000000))))'),
        Col("pf_reach_key_m", "Top reach key (period)", 13, "number",
            formula='=IF({pf_reach}{r}="","",IF({pf_date}{r}="","",'
                    'IF(OR({pf_date}{r}<DashStart,{pf_date}{r}>DashEnd),'
                    '"",{pf_reach}{r}+ROW()/1000000)))'),
        Col("pf_key_y", "Top post key (year)", 13, "number",
            formula='=IF({pf_title}{r}="","",IF({pf_date}{r}="","",'
                    'IF(YEAR({pf_date}{r})<>PlanYear,"",'
                    'IF({pf_er}{r}="","",{pf_er}{r}+ROW()/1000000))))'),
        Col("pf_key_r", "Top post key (report range)", 13, "number",
            formula='=IF({pf_title}{r}="","",IF({pf_date}{r}="","",'
                    'IF(OR({pf_date}{r}<RepFrom,{pf_date}{r}>RepTo),"",'
                    'IF({pf_er}{r}="","",{pf_er}{r}+ROW()/1000000))))'),
        Col("pf_reach_key", "Top reach key (year)", 13, "number",
            formula='=IF({pf_reach}{r}="","",IF({pf_date}{r}="","",'
                    'IF(YEAR({pf_date}{r})<>PlanYear,"",'
                    '{pf_reach}{r}+ROW()/1000000)))'),
        # --------------------------------------------- workflow block -----
        Col("ws_status", "Workflow status", 15, "text",
            formula=_from_list("ContentStatusList")),
        Col("ws_count", "In this status", 11, "int",
            formula='=IF({ws_status}{r}="","",COUNTIF('
                    '[[Content Calendar.status]],{ws_status}{r}))'),
        Col("ws_month", "In status (month)", 11, "int",
            formula='=IF({ws_status}{r}="","",COUNTIFS('
                    '[[Content Calendar.status]],{ws_status}{r},'
                    + _DASH + "))"),
        Col("ow_name", "Owner", 16, "text", formula=_mirror("Team", "name")),
        Col("ow_posts", "Owner posts", 11, "int",
            formula='=IF({ow_name}{r}="","",COUNTIF('
                    '[[Content Calendar.owner]],{ow_name}{r}))'),
        Col("ow_open", "Owner in progress", 11, "int",
            formula='=IF({ow_name}{r}="","",COUNTIFS('
                    '[[Content Calendar.owner]],{ow_name}{r},'
                    '[[Content Calendar.status]],"<>Published"))'),
        Col("ow_tasks", "Owner open tasks", 11, "int",
            formula='=IF({ow_name}{r}="","",COUNTIFS([[Tasks.owner]],'
                    '{ow_name}{r},[[Tasks.status]],"<>Complete"))'),
        # --------------------------------------------- campaign block -----
        Col("cm_name", "Campaign", 20, "text",
            formula=_mirror("Campaigns", "name")),
        Col("cm_revenue", "Campaign revenue", 13, "money",
            formula=_mirror("Campaigns", "revenue")),
        Col("cm_spend", "Campaign spend", 12, "money",
            formula=_mirror("Campaigns", "spend")),
        Col("cm_reach", "Campaign reach", 12, "int",
            formula=_mirror("Campaigns", "reach")),
        Col("cm_key", "Campaign rank key", 12, "number",
            formula='=IF({cm_name}{r}="","",IF(N({cm_revenue}{r})<=0,"",'
                    'N({cm_revenue}{r})+ROW()/1000000))'),
        # ------------------------------------------------ account block ---
        Col("ac_name", "Account", 20, "text",
            formula=_mirror("Accounts", "account")),
        Col("ac_platform", "Account platform", 13, "text",
            formula=_mirror("Accounts", "platform")),
        Col("ac_followers", "Account followers", 12, "int",
            formula=_mirror("Accounts", "followers")),
        Col("ac_growth", "Account growth (month)", 12, "int",
            formula='=IF({ac_name}{r}="","",SUMIFS('
                    '[[Analytics.net_growth]],[[Analytics.account]],'
                    '{ac_name}{r},' + _DASH_AN + "))"),
        Col("ac_key", "Account rank key", 12, "number",
            formula='=IF({ac_followers}{r}="","",IF({ac_followers}{r}=0,'
                    '"",{ac_followers}{r}+ROW()/1000000))'),
    ]

    return SheetSpec(
        name=NAME,
        title="Calculation engine (do not edit)",
        subtitle="Helper tables for the dashboards and reports. This tab "
                 "is hidden by default - unhide it only if you want to "
                 "check the maths.",
        columns=columns,
        rows=rows,
        tab_color="#9AA7AD",
        freeze_cols=1,
        help_text="Every cell on this tab is a locked formula.",
    )
