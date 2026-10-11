"""`Content Pillars` - the themes you post about, and the mix you planned."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import alert_cf, bar_cf, blank, id_col, rate_cf

NAME = "Content Pillars"

_YEAR = ('[[Content Calendar.date]],">="&DATE(PlanYear,1,1),'
         '[[Content Calendar.date]],"<="&DATE(PlanYear,12,31)')


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", '"PIL-"', digits=2, width=9),
        Col("name", "Pillar name", 20, "text",
            note="Five or six is the sweet spot. These feed the pillar "
                 "drop down on the Content Calendar."),
        Col("target", "Target share of posts", 13, "percent",
            validate=number_check(0, 1),
            note="Aim for the shares to add up to 100%. The total is "
                 "checked at the bottom of the grey columns.",
            cf=bar_cf()),
        Col("purpose", "What this pillar is for", 34, "wrap"),
        Col("objective", "Main objective", 16, "text",
            validate=list_from("ObjectiveList")),
        Col("format", "Usual format", 15, "text",
            validate=list_from("FormatList")),
        Col("persona", "Mainly for", 18, "text",
            validate=list_from("PersonaNames")),
        Col("examples", "Example post ideas", 36, "wrap"),
        Col("planned", "Posts planned", 11, "int",
            formula=blank("name",
                          'COUNTIFS([[Content Calendar.pillar]],{name}{r},'
                          f'{_YEAR})'),
            cf=bar_cf()),
        Col("published", "Published", 10, "int",
            formula=blank("name",
                          'COUNTIFS([[Content Calendar.pillar]],{name}{r},'
                          '[[Content Calendar.status]],"Published",'
                          f'{_YEAR})')),
        Col("actual", "Actual share", 11, "percent",
            formula='=IF(OR({name}{r}="",SUM([[Content Pillars.planned]])=0)'
                    ',"",{planned}{r}/SUM([[Content Pillars.planned]]))'),
        Col("variance", "Off target by", 11, "percent",
            formula='=IF(OR({actual}{r}="",{target}{r}=""),"",'
                    '{actual}{r}-{target}{r})',
            cf=[{"type": "cell", "criteria": "between", "minimum": -0.05,
                 "maximum": 0.05, "fmt": "good"},
                {"type": "cell", "criteria": "not between",
                 "minimum": -0.1, "maximum": 0.1, "fmt": "bad"}]),
        Col("reach", "Reach", 12, "int",
            formula=blank("name",
                          'SUMIFS([[Performance.reach]],'
                          '[[Performance.pillar]],{name}{r})'),
            cf=bar_cf()),
        Col("engagements", "Engagements", 12, "int",
            formula=blank("name",
                          'SUMIFS([[Performance.engagements]],'
                          '[[Performance.pillar]],{name}{r})')),
        Col("er", "Engagement rate", 12, "percent",
            formula='=IF(OR({reach}{r}="",{reach}{r}=0),"",'
                    '{engagements}{r}/{reach}{r})',
            cf=rate_cf()),
        Col("clicks", "Clicks", 10, "int",
            formula=blank("name",
                          'SUMIFS([[Performance.clicks]],'
                          '[[Performance.pillar]],{name}{r})')),
        Col("best", "Best post", 30, "text",
            formula='=IF(OR({name}{r}="",{published}{r}=0),"",'
                    'IFERROR(INDEX([[Performance.title]],MATCH('
                    'SUMPRODUCT(MAX(([[Performance.pillar]]={name}{r})'
                    '*[[Performance.rank_key]])),'
                    '[[Performance.rank_key]],0)),""))',
            note="The post with the highest engagement rate for this "
                 "pillar."),
        Col("alert", "Action needed", 26, "status",
            formula=blank(
                "name",
                'IF({planned}{r}=0,"Action - no posts planned",'
                'IF({variance}{r}>0.1,"Check - over-used pillar",'
                'IF({variance}{r}<-0.1,"Check - under-used pillar","OK")))'),
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Content Pillars & Strategy",
        subtitle="Set the mix you want, then watch the grey columns tell "
                 "you the mix you actually published.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).pillars,
        help_text="Target shares should add up to 100%. 'Off target by' "
                  "turns red when a pillar drifts more than 10 points "
                  "away from its target.",
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Totals strip under the grid: do the target shares add up?"""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1          # 1 based Excel row
    last = spec_.last_col_index
    s = sheet
    s.merge(row, 0, 1, "Totals", fmts.chrome("total_label"))
    total_fmt = fmts.chrome("total_value", num_format="0%")
    s.ws.write_formula(
        row, spec_.index("target"),
        f"=SUM({spec_.letter('target')}{spec_.first_data_row}:"
        f"{spec_.letter('target')}{spec_.last_data_row})", total_fmt)
    s.ws.write_formula(
        row, spec_.index("planned"),
        f"=SUM({spec_.letter('planned')}{spec_.first_data_row}:"
        f"{spec_.letter('planned')}{spec_.last_data_row})",
        fmts.chrome("total_value", num_format="#,##0"))
    s.merge(row, spec_.index("purpose"), min(last, spec_.index("examples")),
            "The target column should add up to 100%.",
            fmts.chrome("total_value", align="left"))
