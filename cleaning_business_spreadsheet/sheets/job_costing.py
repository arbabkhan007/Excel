"""`Job Costing` - what every job really made, line by line.

There is nothing to type on this tab. Each row mirrors the job on the
same row of the Jobs tab and works out the wages, the supplies, the
travel and the profit. It is the tab that answers the question most
cleaning businesses never ask: which of these jobs is actually worth
doing?
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..writer import number_check
from .common import alert_cf, blank, month_col, rank_desc, status_cf

NAME = "Job Costing"

#: the job on the same row of the Jobs tab
_MIRROR = 'IFERROR(INDEX([[Jobs.id]],ROW()-4),"")'


def _from_job(key: str, fallback: str = '""') -> str:
    return ('=IF({job}{r}="","",IFERROR(INDEX([[Jobs.' + key + ']],'
            'MATCH({job}{r},[[Jobs.id]],0)),' + fallback + '))')


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("job", "Job ID", 11, "id",
            formula='=IF(' + _MIRROR + '="","",' + _MIRROR + ')',
            note="Mirrors the Jobs tab row for row. Nothing to type."),
        Col("date", "Date", 11, "date", formula=_from_job("date")),
        month_col("date"),
        Col("client", "Client ID", 11, "id", formula=_from_job("client")),
        Col("client_name", "Client", 20, "text",
            formula=_from_job("client_name")),
        Col("property_name", "Property", 20, "text",
            formula=_from_job("property_name")),
        Col("service", "Service", 20, "text", formula=_from_job("service")),
        Col("employee", "Cleaner", 16, "text",
            formula=_from_job("employee")),
        Col("status", "Status", 13, "status", formula=_from_job("status"),
            cf=status_cf()),
        # ----------------------------------------------------- the money --
        Col("revenue", "Revenue", 12, "money",
            formula='=IF({job}{r}="","",IF(OR({status}{r}="Completed",'
                    '{status}{r}="Invoiced"),IFERROR(INDEX('
                    '[[Jobs.price]],MATCH({job}{r},[[Jobs.id]],0)),0),0))',
            note="Only counted once the job is Completed or Invoiced "
                 "- a booking is not money.",
            cf=[{"type": "data_bar", "bar_color": "#E8A33D"}]),
        Col("est_hours", "Hours quoted", 10, "number1",
            formula=_from_job("est_hours", "0"), hidden=True),
        Col("actual_hours", "Hours worked", 10, "number1",
            formula=_from_job("actual_hours", "0"),
            note="From the timesheets if you keep them, otherwise "
                 "the figure typed on the job."),
        Col("hours_used", "Hours costed", 10, "number1",
            formula='=IF({job}{r}="","",IF(IFERROR({actual_hours}{r}*1,0)'
                    '>0,{actual_hours}{r},IFERROR({est_hours}{r}*1,0)))',
            hidden=True),
        Col("timesheet_cost", "Wages from timesheets", 11, "money",
            formula='=IF({job}{r}="","",SUMIFS([[Timesheets.cost]],'
                    '[[Timesheets.job]],{job}{r}))',
            hidden=True),
        Col("labour", "Wages", 11, "money",
            formula='=IF({job}{r}="","",IF({timesheet_cost}{r}>0,'
                    '{timesheet_cost}{r},ROUND({hours_used}{r}*'
                    'CostPerHour,2)))',
            note="Real timesheet cost where you have one, otherwise "
                 "the hours times the true cost of a cleaner hour."),
        Col("supply_logged", "Supplies logged", 11, "money",
            formula='=IF({job}{r}="","",SUMIFS([[Supply Usage.cost]],'
                    '[[Supply Usage.job]],{job}{r}))',
            hidden=True),
        Col("supplies", "Supplies", 11, "money",
            formula='=IF({job}{r}="","",IF({supply_logged}{r}>0,'
                    '{supply_logged}{r},ROUND({hours_used}{r}*'
                    'SuppliesPerHour,2)))',
            note="What you actually used if you logged it on the "
                 "Supply Usage tab, otherwise the per hour allowance "
                 "from Settings."),
        Col("travel", "Travel", 10, "money",
            formula='=IF({job}{r}="","",ROUND(IFERROR(INDEX('
                    '[[Jobs.travel]],MATCH({job}{r},[[Jobs.id]],0)),0)*'
                    'MileageRate,2))'),
        Col("other", "Other cost", 10, "money",
            validate=number_check(0, 1000000),
            note="The only cell you ever type on this tab: parking, "
                 "a skip, a replacement item, a subcontractor's "
                 "invoice."),
        Col("cost", "Total cost", 12, "money",
            formula=blank("job",
                          '{labour}{r}+{supplies}{r}+{travel}{r}+'
                          'IF({other}{r}="",0,{other}{r})')),
        Col("profit", "Profit", 12, "money",
            formula=blank("job", '{revenue}{r}-{cost}{r}'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bold_bad"},
                {"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "good"}]),
        Col("margin", "Margin", 9, "percent0",
            formula='=IF(OR({job}{r}="",{revenue}{r}=0),"",'
                    '{profit}{r}/{revenue}{r})',
            cf=[{"type": "cell", "criteria": "<", "value": "=MarginFloor",
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">=", "value": "=TargetMargin",
                 "fmt": "good"}]),
        Col("per_hour", "Profit an hour", 11, "money",
            formula='=IF(OR({job}{r}="",{hours_used}{r}=0,'
                    '{revenue}{r}=0),"",{profit}{r}/{hours_used}{r})',
            note="The fairest comparison between a two hour office "
                 "and a six hour deep clean.",
            cf=[{"type": "3_color_scale", "min_color": "#FBD9DD",
                 "mid_color": "#FFF0CE", "max_color": "#D7F2E3"}]),
        Col("hours_over", "Hours over estimate", 10, "number1",
            formula='=IF(OR({job}{r}="",{actual_hours}{r}="",'
                    '{actual_hours}{r}=0,{est_hours}{r}=""),"",'
                    '{actual_hours}{r}-{est_hours}{r})',
            hidden=True),
        rank_desc("profit", guard='{job}{r}=""',
                  note="1 = the most profitable job."),
        rank_desc("per_hour", guard='{job}{r}=""',
                  note="1 = the best use of an hour."),
        Col("alert", "Verdict", 28, "status",
            formula=blank(
                "job",
                'IF({revenue}{r}=0,"",'
                'IF({profit}{r}<0,"Loss - this job cost you money",'
                'IF({margin}{r}<MarginFloor,"Below - margin under the '
                'floor",IF({margin}{r}<TargetMargin,'
                '"Check - under target margin",'
                'IF(AND({hours_over}{r}<>"",{hours_over}{r}>0.5),'
                '"Check - ran over the estimate","OK")))))'),
            cf=alert_cf()),
        Col("notes", "Notes", 28, "wrap",
            note="Why this one went the way it did. Worth writing "
                 "down while you remember."),
    ]

    return SheetSpec(
        name=NAME,
        title="Job Costing & Profitability",
        subtitle="Every job, every cost, every margin - worked out "
                 "for you. Only two columns here accept typing.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        help_text="This tab mirrors the Jobs tab row for row. Add any "
                  "extra cost in the 'Other cost' column and read the "
                  "profit, the margin and the profit per hour.",
        extra_notes=[
            "Wages come from the timesheets when you keep them and "
            "from the estimate when you do not - so the numbers are "
            "useful from day one and get more accurate as you fill "
            "the workbook in.",
            "Sort by 'Profit an hour' and you will see something "
            "uncomfortable and useful: the jobs that feel biggest "
            "are rarely the ones that pay best.",
            "A single loss-making job is nothing. The same client "
            "appearing three times in the red is a conversation "
            "about price.",
        ],
    )
