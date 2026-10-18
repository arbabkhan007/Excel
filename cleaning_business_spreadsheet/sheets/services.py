"""`Services` - the price list.

Every job you sell, what it costs you to deliver and what it earns.
The quote calculator, the quote lines and the invoice lines all read
their prices from here, so putting your rates up is a one-cell job.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import (alert_cf, blank, flag_yes_no_cf, id_col, rank_desc)

NAME = "Services"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "ServicePrefix", digits=3, width=9,
               note="Reference only - quotes and jobs pick the service "
                    "by name."),
        Col("name", "Service", 26, "text",
            note="What the client sees on the quote. Keep the names "
                 "short and obvious."),
        Col("category", "Category", 20, "text",
            validate=list_from("ServiceCategoryList")),
        Col("model", "Priced by", 18, "text",
            validate=list_from("PricingModelList"),
            note="How this service is priced. The quote calculator "
                 "asks for a different number for each model."),
        Col("unit", "Unit", 10, "text",
            note="hour, visit, room, oven, sq ft - whatever the price "
                 "below is per."),
        Col("price", "Price per unit", 12, "money",
            validate=number_check(0, 1000000),
            note="Leave blank on an hourly service to use the standard "
                 "rate from the Settings tab."),
        Col("hours", "Hours (one cleaner)", 12, "number1",
            validate=number_check(0, 200),
            note="Typical time for one cleaner. Used to cost the job "
                 "and to block the diary out."),
        Col("uplift", "Uplift", 9, "percent0",
            validate=number_check(0, 5),
            note="Extra percentage for the harder jobs - deep cleans "
                 "and end of tenancy work."),
        Col("supplies", "Supplies cost", 11, "money",
            validate=number_check(0, 100000),
            note="What the products for this service cost you. Blank "
                 "falls back to the per hour figure on Settings."),
        Col("minimum", "Minimum charge", 12, "money",
            validate=number_check(0, 100000),
            note="Blank uses the business minimum from Settings."),
        Col("popular", "Popular?", 9, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="Shows up first on the quote calculator's list of "
                 "ideas."),
        Col("active", "Offered?", 9, "flag", validate=list_from("YesNoList"),
            cf=flag_yes_no_cf(),
            note="Set to No to retire a service without losing its "
                 "history."),
        Col("skills", "Who can do it", 20, "text",
            note="Training or kit the job needs. The scheduler cannot "
                 "enforce it, but you can read it before you assign."),
        Col("equipment", "Equipment needed", 20, "text"),
        Col("description", "What it includes", 46, "wrap",
            note="Paste this straight onto a quote. Being specific "
                 "about what is and is not included prevents most "
                 "arguments."),
        # ------------------------------------------------- the economics --
        Col("standard_price", "Standard price", 13, "money",
            formula=blank(
                "name",
                'ROUND(IF({model}{r}="Per hour",'
                'IF({price}{r}="",HourlyRate,{price}{r})*'
                'IF({hours}{r}="",1,{hours}{r}),'
                'IF({price}{r}="",0,{price}{r}))*'
                '(1+IF({uplift}{r}="",0,{uplift}{r})),2)'),
            note="What a standard booking of this service comes to "
                 "before any property specific adjustment.",
            cf=[{"type": "data_bar", "bar_color": "#2F8F9D"}]),
        Col("labour_cost", "Labour cost", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({hours}{r}="",0,{hours}{r})*CostPerHour,2)'),
            note="Hours times the true cost of a cleaner hour (pay "
                 "plus on-costs) from the Settings tab."),
        Col("supply_cost", "Supply cost", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({supplies}{r}="",IF({hours}{r}="",0,'
                '{hours}{r})*SuppliesPerHour,{supplies}{r}),2)'),
            hidden=True),
        Col("total_cost", "Cost to deliver", 12, "money",
            formula=blank("name",
                          '{labour_cost}{r}+{supply_cost}{r}')),
        Col("profit", "Profit", 11, "money",
            formula=blank("name",
                          '{standard_price}{r}-{total_cost}{r}'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "good"}]),
        Col("margin", "Margin", 9, "percent0",
            formula='=IF(OR({name}{r}="",{standard_price}{r}=0),"",'
                    '{profit}{r}/{standard_price}{r})',
            cf=[{"type": "cell", "criteria": "<", "value": "=MarginFloor",
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">=", "value": "=TargetMargin",
                 "fmt": "good"}],
            note="The honest test of a price. Anything under the floor "
                 "on the Settings tab is work you are doing for "
                 "nothing."),
        Col("hourly_yield", "Earns an hour", 11, "money",
            formula='=IF(OR({name}{r}="",{hours}{r}="",{hours}{r}=0),"",'
                    '{profit}{r}/{hours}{r})',
            note="Profit per hour on site. Sort on this column and you "
                 "know which work to chase and which to drop.",
            cf=[{"type": "3_color_scale", "min_color": "#FBD9DD",
                 "mid_color": "#FFF0CE", "max_color": "#D7F2E3"}]),
        Col("target_price", "Price for target margin", 13, "money",
            formula='=IF(OR({name}{r}="",TargetMargin>=1),"",'
                    'ROUND({total_cost}{r}/(1-TargetMargin),2))',
            note="What you would have to charge to hit the target "
                 "margin from the Settings tab."),
        # --------------------------------------------------- how it sells --
        Col("times_quoted", "Times quoted", 11, "int",
            formula='=IF({name}{r}="","",COUNTIFS('
                    '[[Quote Lines.service]],{name}{r}))'),
        Col("times_booked", "Times booked", 11, "int",
            formula='=IF({name}{r}="","",COUNTIFS([[Jobs.service]],'
                    '{name}{r}))'),
        Col("revenue", "Revenue", 12, "money",
            formula='=IF({name}{r}="","",SUMIFS([[Jobs.price]],'
                    '[[Jobs.service]],{name}{r},[[Jobs.status]],'
                    '"Completed")+SUMIFS([[Jobs.price]],[[Jobs.service]],'
                    '{name}{r},[[Jobs.status]],"Invoiced"))'),
        rank_desc("revenue", guard='{name}{r}=""',
                  note="1 = your biggest earning service."),
        rank_desc("hourly_yield", guard='{name}{r}=""',
                  note="1 = the most profitable use of an hour."),
        Col("alert", "Check", 26, "status",
            formula=blank(
                "name",
                'IF(AND({model}{r}<>"Per hour",{price}{r}=""),'
                '"Add - no price set",'
                'IF({hours}{r}="","Add - how long does it take?",'
                'IF(AND({margin}{r}<>"",{margin}{r}<MarginFloor),'
                '"Below - margin under the floor",'
                'IF(AND({margin}{r}<>"",{margin}{r}<TargetMargin),'
                '"Check - under target margin",'
                'IF({active}{r}="No","Not offered","OK")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Cleaning Service Catalogue",
        subtitle="Your price list, with the margin on every service "
                 "worked out for you.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).services,
        help_text="Set a price and the hours it takes, and the workbook "
                  "tells you what the job costs you, what it makes and "
                  "what you would need to charge to hit your target "
                  "margin.",
        extra_notes=[
            "'Earns an hour' is the column that changes how people "
            "run a cleaning business. A $85 oven clean that takes 90 "
            "minutes beats a $150 half day every time.",
            "Prices here are the default. A client on a special rate "
            "overrides it on the Clients tab, and any quote line can "
            "be typed over by hand.",
            "Retire a service with 'Offered? = No' rather than "
            "deleting the row, or the history on old quotes and jobs "
            "loses its name.",
        ],
    )
