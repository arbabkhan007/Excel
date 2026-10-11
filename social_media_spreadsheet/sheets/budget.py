"""`Budget` - what social media costs you, month by month."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, id_col, money_cf, month_col)

NAME = "Budget"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("category", '"BU-"', digits=3, width=9),
        Col("month", "Month", 12, "date", validate=date_check(),
            note="Use the 1st of the month."),
        month_col("month", key="month_key", header="Month key"),
        Col("category", "Cost category", 20, "text",
            validate=list_from("CostCategoryList")),
        Col("platform", "Platform", 15, "text"),
        Col("campaign", "Linked campaign", 20, "text",
            validate=list_from("CampaignNames")),
        Col("planned", "Planned", 12, "money",
            validate=number_check(0, 1e9), cf=bar_cf()),
        Col("actual", "Actual", 12, "money",
            validate=number_check(0, 1e9)),
        Col("variance", "Difference", 12, "money",
            formula='=IF(AND({planned}{r}="",{actual}{r}=""),"",'
                    'N({planned}{r})-N({actual}{r}))',
            note="Positive means you came in under budget.",
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">=", "value": 0,
                 "fmt": "good"}]),
        Col("variance_pct", "Difference %", 11, "percent",
            formula='=IF(OR({planned}{r}="",{planned}{r}=0),"",'
                    '{variance}{r}/{planned}{r})'),
        Col("share", "Share of the month", 12, "percent",
            formula='=IF(OR({actual}{r}="",{month_key}{r}=""),"",'
                    'IFERROR({actual}{r}/SUMIFS([[Budget.actual]],'
                    '[[Budget.month_key]],{month_key}{r}),""))'),
        Col("revenue", "Revenue credited", 13, "money",
            validate=number_check(0, 1e9), cf=money_cf(),
            note="Only if you can tie revenue to this line - leave it "
                 "blank otherwise."),
        Col("roi", "Return", 11, "number1",
            formula='=IF(OR({actual}{r}="",{actual}{r}=0,'
                    '{revenue}{r}=""),"",{revenue}{r}/{actual}{r})',
            cf=[{"type": "cell", "criteria": ">=", "value": "=RoasTarget",
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 1,
                 "fmt": "bad"}]),
        Col("month_total", "Spent that month", 13, "money",
            formula=blank("month_key",
                          'SUMIFS([[Budget.actual]],'
                          '[[Budget.month_key]],{month_key}{r})')),
        Col("vs_budget", "Versus monthly budget", 13, "percent",
            formula='=IF(OR({month_total}{r}="",MonthlyBudget=0),"",'
                    '{month_total}{r}/MonthlyBudget)',
            cf=[{"type": "cell", "criteria": ">", "value": 1,
                 "fmt": "bad"},
                {"type": "data_bar", "bar_color": "#F2665E"}]),
        Col("alert", "Action needed", 26, "status",
            formula=blank(
                "category",
                'IF({vs_budget}{r}>1,"Over budget this month",'
                'IF(AND({actual}{r}<>"",{planned}{r}<>"",'
                '{actual}{r}>{planned}{r}*1.2),"Check - 20% over plan",'
                'IF(AND({roi}{r}<>"",{roi}{r}<1),"Review - costs more '
                'than it returns","OK")))'),
            cf=alert_cf()),
        Col("notes", "Notes", 28, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Marketing Budget & Return",
        subtitle="Ads, tools, props, editing and freelancers. One row "
                 "per cost per month - the dashboard totals it and "
                 "compares it with the revenue you recorded.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=3,
        samples=sample(cfg).budget,
        help_text="Set your monthly budget on the Settings tab. Anything "
                  "over it is flagged here and on the dashboard.",
    )
