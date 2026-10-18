"""`Quote Lines` - what is actually on each quote.

One row per line. Pick the quote, pick the service and type a quantity;
the price, the hours and the cost come from the Services tab, and the
line total feeds straight back into the quote.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import alert_cf, blank, flag_yes_no_cf, id_col, lookup_service

NAME = "Quote Lines"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("quote", '"QL-"', digits=4, width=10),
        Col("quote", "Quote ID", 11, "id", validate=list_from("QuoteIDs"),
            note="Which quote this line belongs to. Pick it from the "
                 "drop down."),
        Col("service", "Service", 24, "text",
            validate=list_from("ServiceNames"),
            note="From the Services tab. The unit price and the hours "
                 "come with it."),
        Col("description", "Description on the quote", 40, "wrap",
            note="Leave blank to use the standard wording from the "
                 "Services tab."),
        Col("qty", "Quantity", 10, "number1",
            validate=number_check(0, 100000),
            note="Hours, rooms, ovens, visits - whatever the service "
                 "is priced per."),
        Col("price_input", "Price each (override)", 12, "money",
            validate=number_check(0, 1000000),
            note="Blank uses the catalogue price."),
        Col("discount", "Discount", 9, "percent0",
            validate=number_check(0, 1)),
        Col("optional", "Optional extra?", 10, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="Shown on the quote but not counted in the headline "
                 "total until the client says yes. Optional lines "
                 "still add to the subtotal here - take them off the "
                 "quote if they decline."),
        Col("notes", "Line notes", 26, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("catalogue_price", "Catalogue price", 12, "money",
            formula=lookup_service("standard_price", missing="0"),
            hidden=True),
        Col("unit", "Per", 9, "text",
            formula=lookup_service("unit", missing="")),
        Col("price_used", "Price each", 11, "money",
            formula=blank("quote",
                          'IF({price_input}{r}<>"",{price_input}{r},'
                          'IFERROR({catalogue_price}{r}*1,0))')),
        Col("wording", "Wording used", 40, "wrap",
            formula=blank("quote",
                          'IF({description}{r}<>"",{description}{r},'
                          'IFERROR(INDEX([[Services.description]],'
                          'MATCH({service}{r},[[Services.name]],0)),""))'),
            hidden=True,
            note="What the printed quote says for this line."),
        Col("hours", "Hours", 9, "number1",
            formula=blank(
                "quote",
                'IFERROR(INDEX([[Services.hours]],MATCH({service}{r},'
                '[[Services.name]],0)),0)*IF({qty}{r}="",1,'
                'IF(IFERROR(INDEX([[Services.model]],MATCH({service}{r},'
                '[[Services.name]],0)),"")="Per hour",1,{qty}{r}))'),
            note="Hours for the service, multiplied by the quantity "
                 "unless it is already priced by the hour."),
        Col("gross", "Line value", 12, "money",
            formula=blank("quote",
                          'ROUND(IF({qty}{r}="",1,{qty}{r})*'
                          '{price_used}{r},2)'),
            hidden=True),
        Col("total", "Line total", 12, "money",
            formula=blank("quote",
                          'ROUND({gross}{r}*(1-IF({discount}{r}="",0,'
                          '{discount}{r})),2)'),
            cf=[{"type": "data_bar", "bar_color": "#E8A33D"}]),
        Col("cost", "Cost", 11, "money",
            formula=blank(
                "quote",
                'ROUND({hours}{r}*CostPerHour+IF({qty}{r}="",1,'
                '{qty}{r})*IFERROR(INDEX([[Services.supply_cost]],'
                'MATCH({service}{r},[[Services.name]],0)),0),2)'),
            hidden=True),
        Col("margin", "Margin", 9, "percent0",
            formula='=IF(OR({quote}{r}="",{total}{r}=0),"",'
                    '({total}{r}-{cost}{r})/{total}{r})',
            cf=[{"type": "cell", "criteria": "<", "value": "=MarginFloor",
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">=", "value": "=TargetMargin",
                 "fmt": "good"}]),
        Col("seq", "Line number", 9, "int",
            formula='=IF(OR({quote}{r}="",{service}{r}=""),"",'
                    'COUNTIFS({quote_first}:{quote}{r},'
                    '{quote}{r}))',
            hidden=True,
            note="1, 2, 3 within each quote. The printable form "
                 "reads it."),
        Col("key", "Lookup key", 12, "text",
            formula='=IF({seq}{r}="","",{quote}{r}&"|"&{seq}{r})',
            hidden=True),
        Col("alert", "Check", 24, "status",
            formula=blank(
                "quote",
                'IF({service}{r}="","Add - pick a service",'
                'IF({qty}{r}="","Add - how many?",'
                'IF({price_used}{r}=0,"Add - no price",'
                'IF(AND({margin}{r}<>"",{margin}{r}<MarginFloor),'
                '"Below - thin margin","OK"))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Quote Lines",
        subtitle="The detail behind every quote - one row per thing "
                 "you are charging for.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).quote_lines,
        help_text="Pick the quote ID, pick a service, type a quantity. "
                  "Everything else is filled in from the Services tab "
                  "and totalled back onto the quote.",
        extra_notes=[
            "Quoting a job that is not in your catalogue? Leave the "
            "service blank, write your own description and type the "
            "price - the line still totals correctly.",
            "The margin on each line is worth a glance. A quote can "
            "look healthy overall while one line of it is being done "
            "at a loss.",
        ],
    )
