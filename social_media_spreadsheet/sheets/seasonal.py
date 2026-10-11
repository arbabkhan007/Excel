"""`Seasonal Planner` - holidays, awareness days and seasonal moments."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, id_col, month_col, status_cf)

NAME = "Seasonal Planner"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", '"SEA-"', digits=3, width=10),
        Col("date", "Date", 12, "date", validate=date_check()),
        month_col("date"),
        Col("name", "Event / moment", 24, "text"),
        Col("type", "Type", 16, "text",
            validate=list_from("SeasonalTypeList")),
        Col("angle", "Your angle on it", 34, "wrap",
            note="Only join in where you have something real to say."),
        Col("pillar", "Content pillar", 16, "text",
            validate=list_from("PillarNames")),
        Col("campaign", "Linked campaign", 20, "text",
            validate=list_from("CampaignNames")),
        Col("platform", "Platforms", 20, "text"),
        Col("lead_days", "Plan this many days ahead", 13, "int",
            validate=number_check(0, 365, decimal=False),
            note="Pinterest wants 30-45 days, Instagram about 14."),
        Col("hashtags", "Hashtags", 22, "text"),
        Col("budget", "Budget", 12, "money", validate=number_check(0, 1e9)),
        Col("status", "Status", 13, "status",
            validate=list_from("CampaignStatusList"), cf=status_cf()),
        Col("start_by", "Start work by", 12, "date",
            formula='=IF({date}{r}="","",{date}{r}-N({lead_days}{r}))'),
        Col("days_to_go", "Days to go", 10, "int",
            formula=blank("date", "{date}{r}-TODAY()"),
            cf=[{"type": "cell", "criteria": "between", "minimum": 0,
                 "maximum": 14, "fmt": "warn"}]),
        Col("posts", "Posts planned", 11, "int",
            formula='=IF({date}{r}="","",COUNTIFS('
                    '[[Content Calendar.date]],">="&{start_by}{r},'
                    '[[Content Calendar.date]],"<="&{date}{r}))',
            cf=bar_cf()),
        Col("alert", "Action needed", 30, "status",
            formula=blank(
                "name",
                'IF(AND({date}{r}<TODAY(),{status}{r}<>"Complete",'
                '{status}{r}<>"Cancelled"),"Expired - the date has '
                'passed",IF(AND({start_by}{r}<=TODAY(),'
                '{date}{r}>=TODAY(),{posts}{r}=0),'
                '"Action - start planning this now",'
                'IF(AND({days_to_go}{r}>=0,{days_to_go}{r}<=14),'
                '"Upcoming - within a fortnight","OK")))'),
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Seasonal & Holiday Content Planner",
        subtitle="The dates that come round every year. Set how far ahead "
                 "each one needs planning and the sheet tells you when to "
                 "start.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=3,
        samples=sample(cfg).seasonal,
        help_text="Fill this in once and it works for every year - just "
                  "update the dates each January.",
    )
