"""`Woodworking` - timber with the offcuts in, hardware, finish, machine time.

Wood is sold in lengths and used in pieces, and the difference is
sawdust and offcuts you paid for. This tab costs the board, the
hardware, the oil and the abrasives, and it charges for machine time
separately - because a planer and a sander wear out, use power and
have to earn their keep.

Timber quantities are in base units: centimetres of length (the Lists
table converts metres, feet and yards).
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import blank, craft_check, craft_tail

NAME = "Woodworking"

_BASE = ('IFERROR(INDEX([[Materials.base_cost]],MATCH({key}{r},'
         '[[Materials.id]],0)),0)')


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("product", "Product ID", 12, "id",
            validate=list_from("ProductIDs")),
        Col("name", "Piece", 28, "text",
            note="'Walnut serving board, large', 'Oak floating "
                 "shelf 60 cm'."),
        Col("species", "Species", 16, "text",
            validate=list_from("WoodList")),
        Col("wood_id", "Timber material", 12, "id",
            validate=list_from("MaterialIDs")),
        Col("wood_qty", "Timber used (cm)", 13, "number1",
            validate=number_check(0, 1000000),
            note="Length of board at the width and thickness you "
                 "buy it in. Measure the cut list, not the "
                 "finished piece."),
        Col("offcut_waste", "Offcut waste", 11, "percent",
            validate=number_check(0, 1),
            note="Saw kerf, snipe, the end you cut off because it "
                 "split. 15-30% is normal on small work."),
        Col("second_id", "Second timber", 12, "id",
            validate=list_from("MaterialIDs")),
        Col("second_qty", "Second timber (cm)", 13, "number1",
            validate=number_check(0, 1000000)),
        Col("hardware", "Hardware", 11, "money",
            validate=number_check(0, 100000),
            note="Hinges, screws, feet, magnets, brackets."),
        Col("glue", "Glue & fixings", 11, "money",
            validate=number_check(0, 10000)),
        Col("finish_type", "Finish", 16, "text",
            validate=list_from("FinishList")),
        Col("finish_ml", "Finish used (ml)", 13, "number1",
            validate=number_check(0, 100000)),
        Col("finish_rate", "Finish cost a ml", 12, "money4",
            validate=number_check(0, 1000),
            note="A 500 ml tin of hard wax oil at $32 is $0.064 a "
                 "millilitre."),
        Col("abrasives", "Abrasives & blades", 12, "money",
            validate=number_check(0, 10000),
            note="Sanding discs die on hardwood. They are a cost "
                 "per piece, not an annual surprise."),
        Col("machine_mins", "Machine minutes", 12, "int",
            validate=number_check(0, 10000, decimal=False),
            note="Time under power: planer, bandsaw, router, "
                 "sander. Charged at the machine rate as well as "
                 "your time."),
        Col("machine_rate", "Machine rate an hour", 14, "money",
            validate=number_check(0, 1000),
            note="Power, dust extraction, blade and belt wear, "
                 "tool depreciation. $6-$12 an hour is a sane "
                 "starting point for a small shop."),
        Col("mill_mins", "Milling minutes", 12, "int",
            validate=number_check(0, 10000, decimal=False)),
        Col("build_mins", "Assembly minutes", 13, "int",
            validate=number_check(0, 10000, decimal=False)),
        Col("sand_mins", "Sanding minutes", 12, "int",
            validate=number_check(0, 10000, decimal=False),
            note="Always longer than anybody admits."),
        Col("finish_mins", "Finishing minutes", 13, "int",
            validate=number_check(0, 10000, decimal=False),
            note="Coats, waiting, de-nibbing, another coat."),
        Col("dry_days", "Drying days", 10, "int",
            validate=number_check(0, 365, decimal=False)),
        Col("rate_input", "Rate override", 11, "money",
            validate=number_check(0, 1000)),
        Col("price_input", "Your price", 11, "money",
            validate=number_check(0, 100000)),
        Col("notes", "Notes", 22, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("waste_rate", "Waste used", 9, "percent",
            formula=blank(
                "name",
                'IF({offcut_waste}{r}="",DefaultWaste,'
                '{offcut_waste}{r})'),
            hidden=True),
        Col("wood_total", "Timber bought (cm)", 13, "number1",
            formula=blank(
                "name",
                'ROUND((IF({wood_qty}{r}="",0,{wood_qty}{r})+'
                'IF({second_qty}{r}="",0,{second_qty}{r}))*'
                '(1+{waste_rate}{r}),1)'),
            hidden=True),
        Col("wood_cost", "Timber cost", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({wood_qty}{r}="",0,{wood_qty}{r})*'
                '(1+{waste_rate}{r})*' + _BASE.replace("{key}", "{wood_id}") +
                '+IF({second_qty}{r}="",0,{second_qty}{r})*'
                '(1+{waste_rate}{r})*' + _BASE.replace("{key}", "{second_id}") +
                ',2)')),
        Col("finish_cost", "Finish cost", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({finish_ml}{r}="",0,{finish_ml}{r})*'
                'IF({finish_rate}{r}="",0,{finish_rate}{r}),2)')),
        Col("mat_cost", "Materials", 11, "money",
            formula=blank(
                "name",
                'ROUND({wood_cost}{r}+{finish_cost}{r}+'
                'IF({hardware}{r}="",0,{hardware}{r})+'
                'IF({glue}{r}="",0,{glue}{r})+'
                'IF({abrasives}{r}="",0,{abrasives}{r}),2)')),
        Col("minutes", "Minutes", 10, "number1",
            formula=blank(
                "name",
                'IF({mill_mins}{r}="",0,{mill_mins}{r})+'
                'IF({build_mins}{r}="",0,{build_mins}{r})+'
                'IF({sand_mins}{r}="",0,{sand_mins}{r})+'
                'IF({finish_mins}{r}="",0,{finish_mins}{r})')),
        Col("rate", "Hourly rate", 11, "money",
            formula=blank(
                "name",
                'IF({rate_input}{r}<>"",{rate_input}{r},LaborCostHour)'),
            hidden=True),
        Col("labor_cost", "Labour", 11, "money",
            formula=blank("name", 'ROUND({minutes}{r}/60*{rate}{r},2)')),
        Col("machine_cost", "Machine time", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({machine_mins}{r}="",0,{machine_mins}{r})/60*'
                'IF({machine_rate}{r}="",0,{machine_rate}{r}),2)'),
            note="The tools earning their replacement. Leave it out "
                 "and the lathe is a hobby that happens to make "
                 "things."),
        Col("other_cost", "Extras", 10, "money",
            formula=blank(
                "name",
                'ROUND({machine_cost}{r}+{minutes}{r}/60*'
                'SuppliesPerHour,2)'),
            hidden=True),
        Col("waste_cost", "Cost of the offcuts", 13, "money",
            formula=blank(
                "name",
                'ROUND((IF({wood_qty}{r}="",0,{wood_qty}{r})*' +
                _BASE.replace("{key}", "{wood_id}") +
                '+IF({second_qty}{r}="",0,{second_qty}{r})*' +
                _BASE.replace("{key}", "{second_id}") + ')*{waste_rate}{r},2)'),
            hidden=True),
    ] + craft_tail() + [craft_check()]

    return SheetSpec(
        name=NAME,
        title="Woodworking Calculator",
        subtitle="Timber with the offcuts in, hardware, finish, "
                 "abrasives and machine time.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).woodworking,
        help_text="Cost the cut list, not the finished piece, and "
                  "add an honest offcut allowance. Machine time is "
                  "charged on top of your hours.",
        extra_notes=[
            "Offcuts are paid-for wood. A 25% allowance on $0.42 a "
            "centimetre walnut is real money on every board, and "
            "pretending otherwise is why small woodworkers run out "
            "of timber budget.",
            "Charge machine time. Blades, belts, bits and dust "
            "bags are consumables, and a thicknesser that is not "
            "paying for its own replacement is costing you later.",
            "Drying days are not free either - a finish that needs "
            "three days between coats is three days the bench is "
            "occupied.",
        ],
    )
