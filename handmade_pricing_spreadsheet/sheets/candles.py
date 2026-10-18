"""`Candles` - wax by the gram, worked out from the vessel.

Tell the workbook how big the jar is and how full you pour it and it
works out the wax, the fragrance load on top of that, the wick, the
dye and the vessel. Melt loss is included, because the wax left in
the pouring pitcher was still bought and paid for.

Quantities are in base units: grams of wax and fragrance, millilitres
of vessel capacity.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import blank, craft_check, craft_tail

NAME = "Candles"

_BASE = ('IFERROR(INDEX([[Materials.base_cost]],MATCH({key}{r},'
         '[[Materials.id]],0)),0)')


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("product", "Product ID", 12, "id",
            validate=list_from("ProductIDs")),
        Col("name", "Candle", 28, "text",
            note="'Fig & cedar, 220 ml amber jar'."),
        Col("wax_type", "Wax", 16, "text", validate=list_from("WaxList")),
        Col("wax_id", "Wax material", 12, "id",
            validate=list_from("MaterialIDs")),
        Col("vessel", "Vessel", 16, "text",
            validate=list_from("VesselList")),
        Col("vessel_ml", "Vessel size (ml)", 13, "number1",
            validate=number_check(0, 100000),
            note="The capacity to the brim. The fill level below "
                 "turns it into the wax you actually pour."),
        Col("fill", "Fill level", 10, "percent",
            validate=number_check(0, 1),
            note="How full you pour it. 85% is typical - leave the "
                 "headroom or the lid sticks."),
        Col("density", "Wax density", 10, "number",
            validate=number_check(0.1, 2),
            note="Grams per millilitre. Soy is about 0.9. Blank "
                 "uses 0.9."),
        Col("wax_input", "Wax override (g)", 13, "number1",
            validate=number_check(0, 100000),
            note="If you already know the pour weight, type it and "
                 "the vessel maths is ignored."),
        Col("melt_loss", "Melt loss", 10, "percent",
            validate=number_check(0, 1),
            note="Wax left in the pitcher and on the thermometer. "
                 "Blank uses the default waste on Settings."),
        Col("fragrance_id", "Fragrance material", 13, "id",
            validate=list_from("MaterialIDs")),
        Col("fragrance_load", "Fragrance load", 12, "percent",
            validate=number_check(0, 0.3),
            note="As a share of the WAX weight, which is how "
                 "fragrance oil is specified. 8-10% is typical."),
        Col("wick", "Wick", 14, "text", validate=list_from("WickList")),
        Col("wick_count", "Wicks", 8, "int",
            validate=number_check(0, 100, decimal=False)),
        Col("wick_cost", "Cost a wick", 10, "money",
            validate=number_check(0, 1000)),
        Col("vessel_cost", "Vessel cost", 11, "money",
            validate=number_check(0, 10000),
            note="Including the lid. Vessels are usually the "
                 "biggest single cost in a candle."),
        Col("dye_cost", "Dye & additives", 11, "money",
            validate=number_check(0, 10000)),
        Col("label_cost", "Labels & warning", 12, "money",
            validate=number_check(0, 10000),
            note="Front label, base warning label. The warning "
                 "label is a legal requirement in most places."),
        Col("batch_size", "Pour batch", 10, "int",
            validate=number_check(1, 10000, decimal=False),
            note="How many you pour at once. The setup and "
                 "clean-down time is shared across them."),
        Col("batch_mins", "Batch minutes", 11, "int",
            validate=number_check(0, 10000, decimal=False),
            note="Melting, measuring, cleaning the pitcher - once "
                 "per batch, not once per candle."),
        Col("item_mins", "Minutes a candle", 13, "int",
            validate=number_check(0, 10000, decimal=False),
            note="Wicking, pouring, trimming, labelling."),
        Col("cure_days", "Cure days", 10, "int",
            validate=number_check(0, 365, decimal=False),
            note="Not a cost, but it is working capital sitting on "
                 "a shelf. Two weeks of cure is two weeks you "
                 "cannot sell it."),
        Col("rate_input", "Rate override", 11, "money",
            validate=number_check(0, 1000)),
        Col("price_input", "Your price", 11, "money",
            validate=number_check(0, 100000)),
        Col("notes", "Notes", 22, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("wax_g", "Wax poured (g)", 12, "number1",
            formula=blank(
                "name",
                'IF({wax_input}{r}<>"",{wax_input}{r},'
                'ROUND(IF({vessel_ml}{r}="",0,{vessel_ml}{r})*'
                'IF({fill}{r}="",0.85,{fill}{r})*'
                'IF({density}{r}="",0.9,{density}{r}),1))'),
            note="Vessel size x fill level x density, or your own "
                 "pour weight."),
        Col("wax_total", "Wax bought (g)", 12, "number1",
            formula=blank(
                "name",
                'ROUND({wax_g}{r}*(1+IF({melt_loss}{r}="",DefaultWaste,'
                '{melt_loss}{r})),1)'),
            hidden=True),
        Col("wax_cost", "Wax cost", 11, "money",
            formula=blank(
                "name",
                'ROUND({wax_total}{r}*' + _BASE.replace("{key}", "{wax_id}") +
                ',2)')),
        Col("fragrance_g", "Fragrance (g)", 11, "number1",
            formula=blank(
                "name",
                'ROUND({wax_g}{r}*IF({fragrance_load}{r}="",0,'
                '{fragrance_load}{r}),2)'),
            hidden=True),
        Col("fragrance_cost", "Fragrance cost", 12, "money",
            formula=blank(
                "name",
                'ROUND({fragrance_g}{r}*' +
                _BASE.replace("{key}", "{fragrance_id}") + ',2)'),
            note="Fragrance oil is the expensive bottle on the "
                 "shelf, and the load is a percentage of the wax, "
                 "so a bigger jar costs more twice over."),
        Col("wicks_cost", "Wicks cost", 10, "money",
            formula=blank(
                "name",
                'ROUND(IF({wick_count}{r}="",0,{wick_count}{r})*'
                'IF({wick_cost}{r}="",0,{wick_cost}{r}),2)')),
        Col("mat_cost", "Materials", 11, "money",
            formula=blank(
                "name",
                'ROUND({wax_cost}{r}+{fragrance_cost}{r}+{wicks_cost}{r}+'
                'IF({vessel_cost}{r}="",0,{vessel_cost}{r})+'
                'IF({dye_cost}{r}="",0,{dye_cost}{r})+'
                'IF({label_cost}{r}="",0,{label_cost}{r}),2)')),
        Col("minutes", "Minutes", 10, "number1",
            formula=blank(
                "name",
                'ROUND(IF({item_mins}{r}="",0,{item_mins}{r})+'
                'IF({batch_mins}{r}="",0,{batch_mins}{r})/'
                'MAX(1,IF({batch_size}{r}="",DefaultBatchSize,'
                '{batch_size}{r})),1)'),
            note="Your per candle time plus this candle's share of "
                 "the batch setup."),
        Col("rate", "Hourly rate", 11, "money",
            formula=blank(
                "name",
                'IF({rate_input}{r}<>"",{rate_input}{r},LaborCostHour)'),
            hidden=True),
        Col("labor_cost", "Labour", 11, "money",
            formula=blank("name", 'ROUND({minutes}{r}/60*{rate}{r},2)')),
        Col("other_cost", "Extras", 10, "money",
            formula=blank(
                "name", 'ROUND({minutes}{r}/60*SuppliesPerHour,2)'),
            hidden=True),
        Col("cost_per_100g", "Cost per 100 g", 12, "money",
            formula=blank(
                "name",
                'IF(OR({wax_g}{r}="",{wax_g}{r}=0),"",'
                'ROUND({mat_cost}{r}/{wax_g}{r}*100,2))'),
            hidden=True),
        Col("burn_hours", "Burn hours", 10, "int",
            formula=blank(
                "name",
                'IF({wax_g}{r}="","",ROUND({wax_g}{r}/7,0))'),
            hidden=True,
            note="Roughly, at 7 g an hour. Buyers compare candles "
                 "on burn time, so it is worth knowing yours."),
    ] + craft_tail() + [craft_check()]

    return SheetSpec(
        name=NAME,
        title="Candle Calculator",
        subtitle="Wax from the vessel size, fragrance load on top, "
                 "wicks, dye, vessel and the pour.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).candles,
        help_text="Enter the vessel size and the fill level - the "
                  "wax, the fragrance and the cost work themselves "
                  "out from there.",
        extra_notes=[
            "Fragrance load is a percentage of the wax, not of the "
            "finished candle. 10% of 180 g of wax is 18 g of oil, "
            "and at $0.09 a gram that is $1.62 before the jar.",
            "The vessel is usually the single biggest cost. A "
            "beautiful ceramic pot can be more than the wax, the "
            "oil and the wick together - price accordingly.",
            "Cure time is working capital. Two weeks on a shelf "
            "before a candle can be sold is two weeks of money you "
            "have already spent.",
        ],
    )
