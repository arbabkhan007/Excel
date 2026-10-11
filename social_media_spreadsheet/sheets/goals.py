"""`Goals` - the handful of numbers you are actually trying to move."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, id_col, status_cf

NAME = "Goals"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", '"GOAL-"', digits=2, width=10),
        Col("name", "Goal", 26, "text",
            note="Write it as a number you can check: 'Instagram "
                 "followers to 5,000'."),
        Col("category", "Category", 16, "text",
            validate=list_from("GoalCategoryList")),
        Col("platform", "Platform", 13, "text",
            validate=list_from("PlatformList")),
        Col("description", "What success looks like", 34, "wrap"),
        Col("owner", "Owner", 16, "text", validate=list_from("TeamNames")),
        Col("campaign", "Linked campaign", 18, "text",
            validate=list_from("CampaignNames")),
        Col("plan", "How you will get there", 26, "wrap"),
        Col("baseline", "Starting point", 12, "number",
            validate=number_check(-1e12, 1e12)),
        Col("current", "Where you are now", 13, "number",
            validate=number_check(-1e12, 1e12),
            note="Update this once a month when you fill in the "
                 "Analytics tab."),
        Col("target", "Target", 12, "number",
            validate=number_check(-1e12, 1e12)),
        Col("target_date", "By when", 12, "date", validate=date_check()),
        Col("milestone_month", "Monthly step", 12, "number",
            validate=number_check(-1e12, 1e12)),
        Col("milestone_quarter", "Quarterly step", 12, "number",
            validate=number_check(-1e12, 1e12)),
        Col("budget", "Budget for it", 12, "money",
            validate=number_check(0, 1e9)),
        Col("status", "Status", 13, "status",
            validate=list_from("GoalStatusList"), cf=status_cf()),
        Col("to_go", "Still to go", 12, "number",
            formula='=IF(OR({target}{r}="",{current}{r}=""),"",'
                    '{target}{r}-{current}{r})'),
        Col("progress", "Progress", 11, "percent",
            formula='=IF(OR({target}{r}="",{baseline}{r}="",'
                    '{target}{r}={baseline}{r}),"",MAX(0,'
                    '({current}{r}-{baseline}{r})'
                    '/({target}{r}-{baseline}{r})))',
            cf=[{"type": "data_bar", "bar_color": "#6350A8"},
                {"type": "cell", "criteria": ">=", "value": 1,
                 "fmt": "good"}]),
        Col("days_left", "Days left", 10, "int",
            formula=blank("target_date", "{target_date}{r}-TODAY()"),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "muted"},
                {"type": "cell", "criteria": "between", "minimum": 0,
                 "maximum": 30, "fmt": "warn"}]),
        Col("expected", "Where you should be", 13, "percent",
            formula='=IF(OR({target_date}{r}="",{start_ref}{r}="",'
                    '{target_date}{r}<={start_ref}{r}),"",'
                    'MIN(1,MAX(0,(TODAY()-{start_ref}{r})'
                    '/({target_date}{r}-{start_ref}{r}))))',
            note="How far through the time you are - compare it with "
                 "the progress column."),
        Col("start_ref", "Year start", 10, "date",
            formula='=IF({name}{r}="","",DATE(PlanYear,1,1))',
            hidden=True),
        Col("pace", "On track?", 20, "status",
            formula='=IF(OR({progress}{r}="",{expected}{r}=""),"",'
                    'IF({progress}{r}>=1,"Achieved",'
                    'IF({progress}{r}>={expected}{r},"On track",'
                    'IF({progress}{r}>={expected}{r}*0.8,"Slightly '
                    'behind","Behind - needs a push"))))',
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Achieved", "fmt": "bold_good"},
                {"type": "text", "criteria": "containing",
                 "value": "On track", "fmt": "good"},
                {"type": "text", "criteria": "containing",
                 "value": "Slightly", "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "Behind", "fmt": "bad"}]),
        Col("needed_per_month", "Needed each month", 13, "number1",
            formula='=IF(OR({to_go}{r}="",{days_left}{r}="",'
                    '{days_left}{r}<=0),"",{to_go}{r}'
                    '/MAX(1,{days_left}{r}/30.4))'),
        Col("alert", "Action needed", 28, "status",
            formula=blank(
                "name",
                'IF(AND({days_left}{r}<>"",{days_left}{r}<0,'
                '{progress}{r}<1),"Overdue - the date has passed",'
                'IF({pace}{r}="Behind - needs a push","Action - behind '
                'the pace you need",IF({current}{r}="","Add - no current '
                'number yet","OK")))'),
            cf=alert_cf()),
        Col("notes", "Notes", 26, "wrap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Goals & KPI Tracker",
        subtitle="Four or five goals, each with a number, a date and an "
                 "owner. The sheet compares how far you have come with "
                 "how far through the year you are.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).goals,
        help_text="Update the 'where you are now' column once a month. "
                  "Everything else works itself out.",
    )
