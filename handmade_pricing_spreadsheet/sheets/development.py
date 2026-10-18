"""`Development` - what a new design cost before it earned anything.

Every product started as a sample that did not sell, an afternoon of
sketching and a photo shoot. That time is real and it has to come
back out of the pieces you go on to sell - or you have to admit the
design was a hobby.

This tab tracks development cost per idea and tells you how many
sales recover it.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, id_col, lookup_product, status_cf

NAME = "Development"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "DevPrefix", digits=4, width=10),
        Col("name", "Design", 28, "text",
            note="The idea, whether or not it ever became a "
                 "product."),
        Col("product", "Became product", 12, "id",
            validate=list_from("ProductIDs"),
            note="Link it once the design is listed."),
        Col("stage", "Stage", 16, "text", validate=list_from("DevStageList"),
            cf=status_cf()),
        Col("started", "Started", 11, "date", validate=date_check()),
        Col("design_hours", "Design hours", 11, "number",
            validate=number_check(0, 10000)),
        Col("sample_hours", "Sample hours", 11, "number",
            validate=number_check(0, 10000),
            note="Making the prototypes, including the ones that "
                 "went in the bin."),
        Col("photo_hours", "Photo & listing hours", 15, "number",
            validate=number_check(0, 10000),
            note="Shooting, editing, writing the listing, "
                 "keywords. Hours that never feel like work and "
                 "always are."),
        Col("sample_cost", "Sample materials", 13, "money",
            validate=number_check(0, 1000000)),
        Col("other_cost", "Other costs", 11, "money",
            validate=number_check(0, 1000000),
            note="A pattern you bought to learn from, a tool you "
                 "needed, a model you paid."),
        Col("samples_sold", "Samples sold for", 13, "money",
            validate=number_check(0, 1000000),
            note="Prototypes often sell as seconds. Credited "
                 "back."),
        Col("decided", "Decision date", 11, "date", validate=date_check()),
        Col("notes", "Notes", 28, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("hours", "Hours", 9, "number",
            formula=blank(
                "name",
                'IF({design_hours}{r}="",0,{design_hours}{r})+'
                'IF({sample_hours}{r}="",0,{sample_hours}{r})+'
                'IF({photo_hours}{r}="",0,{photo_hours}{r})'),
            hidden=True),
        Col("dev_cost", "Development cost", 13, "money",
            formula=blank(
                "name",
                'ROUND({hours}{r}*LaborCostHour+'
                'IF({sample_cost}{r}="",0,{sample_cost}{r})+'
                'IF({other_cost}{r}="",0,{other_cost}{r})-'
                'IF({samples_sold}{r}="",0,{samples_sold}{r}),2)'),
            note="Hours at your rate, plus materials and costs, "
                 "less whatever the samples sold for."),
        Col("unit_profit", "Profit a sale", 11, "money",
            formula=lookup_product("profit", missing=""),
            hidden=True),
        Col("sold", "Sold since", 10, "int",
            formula=blank(
                "name",
                'IF({product}{r}="",0,SUMIFS([[Sales.qty]],'
                '[[Sales.product]],{product}{r}))')),
        Col("recovered", "Recovered", 11, "money",
            formula=blank(
                "name",
                'IF({unit_profit}{r}="",0,ROUND({sold}{r}*'
                '{unit_profit}{r},2))')),
        Col("recovery", "Recovered", 10, "percent",
            formula=blank(
                "name",
                'IF({dev_cost}{r}<=0,"",{recovered}{r}/{dev_cost}{r})'),
            cf=[{"type": "cell", "criteria": ">=", "value": 1,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0.5,
                 "fmt": "warn"}]),
        Col("to_recover", "Sales to recover it", 14, "int",
            formula=blank(
                "name",
                'IF(OR({unit_profit}{r}="",{unit_profit}{r}<=0),'
                '"Needs a profitable price",'
                'ROUNDUP(MAX(0,{dev_cost}{r}-{recovered}{r})/'
                '{unit_profit}{r},0))'),
            note="How many more you have to sell before the "
                 "design has paid for its own development."),
        Col("per_unit", "Development a sale", 13, "money",
            formula=blank(
                "name",
                'IF(OR({sold}{r}=0,{dev_cost}{r}<=0),"",'
                'ROUND({dev_cost}{r}/{sold}{r},2))'),
            hidden=True,
            note="Spread over what has sold so far. It falls "
                 "every time one goes out of the door."),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "name",
                'IF({dev_cost}{r}<=0,"Add - hours and sample costs",'
                'IF({product}{r}="","Idea - not listed yet",'
                'IF({recovery}{r}>=1,"Recovered - this one paid for itself",'
                'IF(AND({sold}{r}=0,{started}{r}<>"",'
                'TodayDate-{started}{r}>90),'
                '"Check - three months, no sales",'
                'IF({recovery}{r}<0.5,'
                '"Watch - under half recovered","OK")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Product Development",
        subtitle="What each design cost to create, and how many "
                 "sales pay that back.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).development,
        help_text="Log the hours while you are doing them. Link "
                  "the design to its product and the recovery "
                  "figure follows the sales.",
        extra_notes=[
            "Development is a real cost and it belongs somewhere. "
            "Either recover it through the price of that design "
            "or accept it as the cost of having a range.",
            "A design that has not recovered its development in "
            "six months is telling you something - usually about "
            "the photographs or the price rather than the design.",
            "Count the listing and photography hours. For a "
            "digital product they are most of the cost.",
        ],
    )
