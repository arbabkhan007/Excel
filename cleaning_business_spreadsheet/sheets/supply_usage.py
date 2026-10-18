"""`Supply Usage` - what got used, where.

A short tab with a long reach: it takes stock off the cupboard, puts a
real materials cost onto the job costing, and shows you which client
quietly gets through twice as much as everybody else.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, id_col, month_col

NAME = "Supply Usage"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("supply", '"SU-"', digits=4, width=10),
        Col("date", "Date", 11, "date", validate=date_check()),
        month_col("date"),
        Col("supply", "Item", 24, "text", validate=list_from("SupplyNames"),
            note="From the Supplies tab. The cost comes with it."),
        Col("qty", "Quantity used", 11, "number1",
            validate=number_check(0, 100000),
            note="In the unit on the Supplies tab - bottles, rolls, "
                 "litres."),
        Col("job", "Job ID", 11, "id", validate=list_from("JobIDs"),
            note="Link it to a job and the cost lands on that job's "
                 "profit. Leave blank for stock used round the "
                 "office."),
        Col("employee", "Used by", 16, "text",
            validate=list_from("EmployeeNames")),
        Col("reason", "Why", 20, "text",
            note="'Weekly visit', 'restocked the van', 'spillage'."),
        Col("notes", "Notes", 26, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("client_name", "Client", 20, "text",
            formula='=IF({job}{r}="","",IFERROR(INDEX('
                    '[[Jobs.client_name]],MATCH({job}{r},[[Jobs.id]],0)),'
                    '""))'),
        Col("unit_cost", "Cost a unit", 10, "money",
            formula='=IF({supply}{r}="","",IFERROR(INDEX('
                    '[[Supplies.unit_cost]],MATCH({supply}{r},'
                    '[[Supplies.name]],0)),""))',
            hidden=True),
        Col("cost", "Cost", 11, "money",
            formula=blank(
                "supply",
                'ROUND(IF({qty}{r}="",0,{qty}{r})*IFERROR('
                '{unit_cost}{r}*1,0),2)'),
            note="Quantity times the cost per unit. This is what "
                 "lands on the job costing.",
            cf=[{"type": "data_bar", "bar_color": "#2F8F9D"}]),
        Col("alert", "Check", 22, "status",
            formula=blank(
                "supply",
                'IF({qty}{r}="","Add - how much?",'
                'IF({unit_cost}{r}="","Check - item not in Supplies",'
                'IF({job}{r}="","Not on a job","OK")))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Supply Usage",
        subtitle="What you used, on which job, and what it cost.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_soft,
        freeze_cols=2,
        samples=sample(cfg).supply_usage,
        help_text="Log the big items - the chemicals, the pads, the "
                  "bags. Nobody needs to log a single cloth, and a "
                  "tab nobody fills in is worse than no tab at all.",
        extra_notes=[
            "Every row here comes off the stock on the Supplies tab "
            "and goes onto the cost of the job you linked it to.",
            "If you do not log usage, the job costing falls back to "
            "the per-hour supplies allowance on the Settings tab - "
            "so your numbers still work, they are just rougher.",
        ],
    )
