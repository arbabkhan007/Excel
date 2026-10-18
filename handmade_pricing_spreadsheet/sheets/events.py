"""`Events` - craft fairs, markets and whether the stall paid.

A stall is a fixed cost you pay before you sell anything: the pitch,
the petrol, the parking, the two days you were not making. This tab
works out what the day has to take before it breaks even, and
compares that with what it actually took.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, loss_cf, status_cf

NAME = "Events"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("name", "Event", 26, "text",
            note="The name sales rows pick from, so keep it "
                 "short and distinct."),
        Col("kind", "Type", 18, "text", validate=list_from("EventTypeList")),
        Col("date", "Date", 11, "date", validate=date_check()),
        Col("days", "Days", 7, "number1",
            validate=number_check(0, 60)),
        Col("stall_fee", "Stall fee", 11, "money",
            validate=number_check(0, 100000)),
        Col("travel", "Travel & parking", 13, "money",
            validate=number_check(0, 100000)),
        Col("accommodation", "Accommodation", 13, "money",
            validate=number_check(0, 100000)),
        Col("other_cost", "Stock, display, other", 15, "money",
            validate=number_check(0, 100000),
            note="Banners, a new table cover, card reader fees, "
                 "the coffee."),
        Col("prep_hours", "Prep hours", 11, "number",
            validate=number_check(0, 1000),
            note="Making extra stock, pricing, packing the car."),
        Col("stall_hours", "Hours on the stall", 14, "number",
            validate=number_check(0, 1000),
            note="Including setting up and packing down. They "
                 "count."),
        Col("target", "Sales target", 11, "money",
            validate=number_check(0, 1000000)),
        Col("footfall", "Footfall", 10, "int",
            validate=number_check(0, 10000000, decimal=False)),
        Col("notes", "Notes", 28, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("cash_cost", "Money out", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({stall_fee}{r}="",0,{stall_fee}{r})+'
                'IF({travel}{r}="",0,{travel}{r})+'
                'IF({accommodation}{r}="",0,{accommodation}{r})+'
                'IF({other_cost}{r}="",0,{other_cost}{r}),2)')),
        Col("hours", "Hours", 9, "number",
            formula=blank(
                "name",
                'IF({prep_hours}{r}="",0,{prep_hours}{r})+'
                'IF({stall_hours}{r}="",0,{stall_hours}{r})'),
            hidden=True),
        Col("time_cost", "Your time", 11, "money",
            formula=blank(
                "name", 'ROUND({hours}{r}*LaborCostHour,2)'),
            note="At your own hourly rate. A stall is two days you "
                 "did not spend making, and that is the real cost "
                 "of a craft fair."),
        Col("total_cost", "Total cost", 11, "money",
            formula=blank(
                "name", 'ROUND({cash_cost}{r}+{time_cost}{r},2)')),
        Col("break_even", "Sales to break even", 15, "money",
            formula=blank(
                "name",
                'IF(TargetRetailMargin<=0,"",ROUND({total_cost}{r}/'
                'TargetRetailMargin,2))'),
            note="At your target margin. Taking this much is not a "
                 "good day - it is the day costing you nothing."),
        Col("be_cash", "Cash break-even", 13, "money",
            formula=blank(
                "name",
                'IF(TargetRetailMargin<=0,"",ROUND({cash_cost}{r}/'
                'TargetRetailMargin,2))'),
            hidden=True,
            note="Ignoring your time - the number most stallholders "
                 "use, and the reason craft fairs feel better than "
                 "they are."),
        Col("sales", "Actually took", 11, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.revenue]],[[Sales.event]],'
                '{name}{r}),2)')),
        Col("orders", "Orders", 8, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Sales.event]],{name}{r})')),
        Col("gross_profit", "Profit on the goods", 14, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.profit]],[[Sales.event]],'
                '{name}{r}),2)'),
            hidden=True),
        Col("net", "Profit after the stall", 15, "money",
            formula=blank(
                "name",
                'IF({sales}{r}=0,"",ROUND({gross_profit}{r}-'
                '{total_cost}{r},2))'),
            cf=loss_cf(),
            note="What the day really made once the pitch, the "
                 "petrol and your hours are paid."),
        Col("per_hour", "Earned an hour", 11, "money",
            formula=blank(
                "name",
                'IF(OR({hours}{r}=0,{sales}{r}=0),"",'
                'ROUND(({gross_profit}{r}-{cash_cost}{r})/'
                '{hours}{r},2))'),
            cf=[{"type": "cell", "criteria": "<", "value": "=MinimumWage",
                 "fmt": "bad"}],
            note="Profit on what you sold, less the cash costs, "
                 "over every hour the event took."),
        Col("avg_order", "Average sale", 11, "money",
            formula=blank(
                "name",
                'IF({orders}{r}=0,"",ROUND({sales}{r}/{orders}{r},2))'),
            hidden=True),
        Col("conversion", "Buyers per 100 passing", 15, "percent",
            formula=blank(
                "name",
                'IF(OR({footfall}{r}="",{footfall}{r}=0),"",'
                '{orders}{r}/{footfall}{r})'),
            hidden=True),
        Col("status", "Status", 13, "status",
            formula=blank(
                "name",
                'IF(AND({date}{r}<>"",{date}{r}>TodayDate),"Booked",'
                'IF({sales}{r}=0,"No sales logged","Done"))'),
            cf=status_cf()),
        Col("verdict", "Verdict", 30, "status",
            formula=blank(
                "name",
                'IF({status}{r}="Booked","Coming up - take enough stock",'
                'IF({sales}{r}=0,"Add - log the sales against this event",'
                'IF({net}{r}<0,"Loss - the stall cost more than it made",'
                'IF(AND({per_hour}{r}<>"",{per_hour}{r}<MinimumWage),'
                '"Below - paid under minimum wage for the day",'
                'IF({net}{r}>{total_cost}{r},'
                '"Worth repeating - book it again","OK")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Craft Fairs & Events",
        subtitle="What the stall costs, what it has to take, and "
                 "what it actually did.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).events,
        help_text="Enter the costs before the event and the sales "
                  "after it - tag the Sales rows with the event "
                  "name and they total themselves.",
        extra_notes=[
            "Count your hours. A $45 pitch is not the cost of a "
            "craft fair; a $45 pitch plus fourteen hours is.",
            "Work out the break-even figure before you book. If "
            "it is more than you have ever taken at a similar "
            "event, you already know the answer.",
            "A fair that breaks even but brings three wholesale "
            "enquiries was worth it. Write that in the notes - "
            "the money column cannot see it.",
        ],
    )
