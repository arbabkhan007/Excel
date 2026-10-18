"""`Batches` - what a production run actually cost, against what you estimated.

Estimates are where pricing starts; batches are where it gets tested.
Log a run - how many you planned, what the materials really came to,
how long it really took, how many came out wrong - and this tab tells
you the true cost of a good unit and how far your estimate was out.

The seconds are not free. A batch of twenty with three failures is a
batch of seventeen, and all twenty pieces of wax were paid for.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, loss_cf, lookup_product,
                     status_cf)

NAME = "Batches"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("product", "BatchPrefix", digits=4, width=10),
        Col("product", "Product", 12, "id",
            validate=list_from("ProductIDs")),
        Col("product_name", "Product name", 24, "text",
            formula=lookup_product("name")),
        Col("started", "Started", 11, "date", validate=date_check()),
        Col("finished", "Finished", 11, "date", validate=date_check()),
        Col("planned", "Planned", 9, "int",
            validate=number_check(0, 100000, decimal=False)),
        Col("made", "Made", 8, "int",
            validate=number_check(0, 100000, decimal=False)),
        Col("seconds", "Seconds / failures", 13, "int",
            validate=number_check(0, 100000, decimal=False),
            note="Pieces that cannot be sold at full price. Their "
                 "cost belongs to the ones that can."),
        Col("sold_seconds", "Seconds sold for", 13, "money",
            validate=number_check(0, 1000000),
            note="Total recovered from selling the seconds as "
                 "seconds. Credited back to the batch."),
        Col("material_actual", "Materials actually", 14, "money",
            validate=number_check(0, 1000000),
            note="What the run really consumed. Leave blank to use "
                 "the estimate from the catalog."),
        Col("hours_actual", "Hours actually", 12, "number",
            validate=number_check(0, 10000),
            note="Bench time for the whole run, setup included."),
        Col("extra_cost", "Other costs", 11, "money",
            validate=number_check(0, 1000000),
            note="A courier, a rush delivery of supplies, studio "
                 "hire for the day."),
        Col("status", "Status", 15, "status",
            validate=list_from("BatchStatusList"), cf=status_cf()),
        Col("notes", "Notes", 26, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("good", "Good units", 10, "int",
            formula=blank(
                "product",
                'MAX(0,IF({made}{r}="",0,{made}{r})-'
                'IF({seconds}{r}="",0,{seconds}{r}))')),
        Col("defect_rate", "Failure rate", 10, "percent",
            formula=blank(
                "product",
                'IF(OR({made}{r}="",{made}{r}=0),"",'
                'IF({seconds}{r}="",0,{seconds}{r})/{made}{r})'),
            cf=[{"type": "cell", "criteria": ">", "value": "=DefaultDefect",
                 "fmt": "warn"}]),
        Col("est_mat_unit", "Materials estimated", 13, "money",
            formula=lookup_product("mat_cost", missing=""),
            hidden=True),
        Col("est_minutes", "Minutes estimated", 12, "number1",
            formula=lookup_product("minutes", missing=""),
            hidden=True),
        Col("mat_total", "Materials cost", 12, "money",
            formula=blank(
                "product",
                'ROUND(IF({material_actual}{r}<>"",{material_actual}{r},'
                'IF({est_mat_unit}{r}="",0,{est_mat_unit}{r})*'
                'IF({made}{r}="",0,{made}{r})),2)')),
        Col("labour_total", "Labour cost", 11, "money",
            formula=blank(
                "product",
                'ROUND(IF({hours_actual}{r}<>"",{hours_actual}{r},'
                'IF({est_minutes}{r}="",0,{est_minutes}{r})*'
                'IF({made}{r}="",0,{made}{r})/60)*LaborCostHour,2)')),
        Col("batch_cost", "Batch cost", 11, "money",
            formula=blank(
                "product",
                'ROUND({mat_total}{r}+{labour_total}{r}+'
                'IF({extra_cost}{r}="",0,{extra_cost}{r})-'
                'IF({sold_seconds}{r}="",0,{sold_seconds}{r}),2)'),
            note="Everything the run cost, less whatever the "
                 "seconds brought back in."),
        Col("cost_good", "Cost a good unit", 12, "money",
            formula=blank(
                "product",
                'IF({good}{r}=0,"",ROUND({batch_cost}{r}/{good}{r},2))'),
            note="The honest cost of a sellable piece - the "
                 "failures are carried by the survivors."),
        Col("est_unit", "Estimated a unit", 12, "money",
            formula=blank(
                "product",
                'IFERROR(INDEX([[Products.cost_base]],MATCH({product}{r},'
                '[[Products.id]],0)),"")'),
            hidden=True),
        Col("variance", "Versus estimate", 11, "money",
            formula=blank(
                "product",
                'IF(OR({cost_good}{r}="",{est_unit}{r}=""),"",'
                'ROUND({cost_good}{r}-{est_unit}{r},2))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0, "fmt": "bad"},
                {"type": "cell", "criteria": "<=", "value": 0,
                 "fmt": "good"}],
            note="Positive means the real thing costs more than "
                 "the catalog thinks. Two batches saying the same "
                 "means the catalog is wrong."),
        Col("variance_pct", "Out by", 9, "percent",
            formula=blank(
                "product",
                'IF(OR({variance}{r}="",{est_unit}{r}="",'
                '{est_unit}{r}=0),"",{variance}{r}/{est_unit}{r})'),
            hidden=True),
        Col("mins_actual", "Minutes a unit", 11, "number1",
            formula=blank(
                "product",
                'IF(OR({hours_actual}{r}="",{made}{r}="",{made}{r}=0),"",'
                'ROUND({hours_actual}{r}*60/{made}{r},1))')),
        Col("time_variance", "Minutes out by", 12, "number1",
            formula=blank(
                "product",
                'IF(OR({mins_actual}{r}="",{est_minutes}{r}=""),"",'
                'ROUND({mins_actual}{r}-{est_minutes}{r},1))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}],
            note="The number that re-prices a whole product line. "
                 "If every batch takes longer than the estimate, "
                 "the estimate is a wish."),
        Col("value", "Batch retail value", 13, "money",
            formula=blank(
                "product",
                'IFERROR(ROUND({good}{r}*INDEX([[Products.price]],'
                'MATCH({product}{r},[[Products.id]],0)),2),"")')),
        Col("batch_profit", "Profit if it all sells", 14, "money",
            formula=blank(
                "product",
                'IF({value}{r}="","",ROUND({value}{r}-{batch_cost}{r},2))'),
            cf=loss_cf()),
        Col("days", "Days taken", 10, "int",
            formula=blank(
                "product",
                'IF(OR({started}{r}="",{finished}{r}=""),"",'
                '{finished}{r}-{started}{r})'),
            hidden=True),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "product",
                'IF({made}{r}="","Planned - nothing made yet",'
                'IF({good}{r}=0,"Check - no good units",'
                'IF(AND({variance_pct}{r}<>"",{variance_pct}{r}>0.15),'
                '"Check - costing 15% more than the estimate",'
                'IF(AND({defect_rate}{r}<>"",'
                '{defect_rate}{r}>DefaultDefect*2),'
                '"Check - high failure rate",'
                'IF(AND({time_variance}{r}<>"",{time_variance}{r}>5),'
                '"Check - taking longer than estimated",'
                'IF({batch_profit}{r}<0,"Loss - this run lost money",'
                '"OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Batch Production",
        subtitle="What a run really cost, what it really took, and "
                 "how far the estimate was out.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).batches,
        help_text="Log the run as it happens. Leave the actuals "
                  "blank and the batch is costed from the catalog "
                  "estimate instead.",
        extra_notes=[
            "This is the tab that makes the rest of the workbook "
            "honest. Estimates drift; batches do not lie.",
            "Failures are paid for. Twenty poured candles with "
            "three sink holes is seventeen sellable candles and "
            "twenty jars of wax - the cost a good unit column does "
            "that arithmetic for you.",
            "Two batches in a row costing more than the estimate "
            "is not bad luck. Go and update the bill of materials "
            "or the labour minutes.",
        ],
    )
