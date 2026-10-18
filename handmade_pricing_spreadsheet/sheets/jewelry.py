"""`Jewelry` - metal by the gram, stones by the piece, scrap included.

Silver is priced by weight and the price moves, findings are pennies
that add up, and the bench time on a soldered piece is most of what
you are selling. This tab costs all of it, including the metal that
ends up as scrap on the bench mat.

Quantities are in the material's base unit: grams of metal,
centimetres of chain and wire.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import blank, craft_check, craft_tail

NAME = "Jewelry"

_BASE = ('IFERROR(INDEX([[Materials.base_cost]],MATCH({key}{r},'
         '[[Materials.id]],0)),0)')


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("product", "Product ID", 12, "id",
            validate=list_from("ProductIDs")),
        Col("name", "Piece", 28, "text",
            note="'Hammered silver bangle', 'Moonstone drop "
                 "earrings'."),
        Col("metal", "Metal", 16, "text", validate=list_from("MetalList")),
        Col("metal_id", "Metal material", 12, "id",
            validate=list_from("MaterialIDs"),
            note="The material ID you buy the metal as - sheet, "
                 "wire, casting grain."),
        Col("metal_g", "Metal used (g)", 12, "number1",
            validate=number_check(0, 100000),
            note="Weigh the finished piece. Add the sprue, the "
                 "filings and the bits that became scrap in the "
                 "scrap column."),
        Col("metal_rate_input", "Price a gram", 11, "money4",
            validate=number_check(0, 10000),
            note="Override the material price - useful when the "
                 "metal price has moved since you bought it."),
        Col("scrap", "Scrap / loss", 10, "percent",
            validate=number_check(0, 1),
            note="Filings, cut-offs, failed solders. Real on every "
                 "soldered piece, and worth recovering."),
        Col("recovery", "Scrap recovered", 11, "percent",
            validate=number_check(0, 1),
            note="How much of that scrap you actually send back or "
                 "re-melt. Credited against the cost."),
        Col("stone", "Stone", 16, "text", validate=list_from("StoneList")),
        Col("stone_count", "Stones", 9, "int",
            validate=number_check(0, 10000, decimal=False)),
        Col("stone_cost", "Cost a stone", 11, "money",
            validate=number_check(0, 100000)),
        Col("chain_cm", "Chain (cm)", 10, "number1",
            validate=number_check(0, 10000)),
        Col("chain_id", "Chain material", 12, "id",
            validate=list_from("MaterialIDs")),
        Col("wire_cm", "Wire (cm)", 10, "number1",
            validate=number_check(0, 10000)),
        Col("wire_id", "Wire material", 12, "id",
            validate=list_from("MaterialIDs")),
        Col("findings", "Findings", 11, "money",
            validate=number_check(0, 10000),
            note="Clasps, jump rings, ear wires, posts and backs. "
                 "Count them - they are 60 cents each and they are "
                 "on every piece."),
        Col("solder", "Solder & consumables", 13, "money",
            validate=number_check(0, 10000),
            note="Solder, flux, pickle, polishing compound, saw "
                 "blades."),
        Col("plating", "Plating / finishing", 12, "money",
            validate=number_check(0, 10000)),
        Col("hallmark", "Hallmarking", 11, "money",
            validate=number_check(0, 10000),
            note="Per piece, including the courier. Required over "
                 "the weight threshold in several countries."),
        Col("design_mins", "Design minutes", 12, "int",
            validate=number_check(0, 10000, decimal=False)),
        Col("bench_mins", "Bench minutes", 12, "int",
            validate=number_check(0, 10000, decimal=False),
            note="Sawing, forming, soldering, setting."),
        Col("polish_mins", "Polishing minutes", 13, "int",
            validate=number_check(0, 10000, decimal=False)),
        Col("rate_input", "Rate override", 11, "money",
            validate=number_check(0, 1000)),
        Col("price_input", "Your price", 11, "money",
            validate=number_check(0, 100000)),
        Col("notes", "Notes", 24, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("metal_rate", "Metal a gram", 11, "money4",
            formula=blank(
                "name",
                'IF({metal_rate_input}{r}<>"",{metal_rate_input}{r},' +
                _BASE.replace("{key}", "{metal_id}") + ')'),
            hidden=True),
        Col("metal_gross", "Metal with scrap", 11, "number1",
            formula=blank(
                "name",
                'ROUND(IF({metal_g}{r}="",0,{metal_g}{r})*'
                '(1+IF({scrap}{r}="",0,{scrap}{r})),2)'),
            hidden=True),
        Col("metal_cost", "Metal cost", 11, "money",
            formula=blank(
                "name",
                'ROUND({metal_gross}{r}*{metal_rate}{r}-'
                'IF({metal_g}{r}="",0,{metal_g}{r})*'
                'IF({scrap}{r}="",0,{scrap}{r})*'
                'IF({recovery}{r}="",0,{recovery}{r})*{metal_rate}{r},2)'),
            note="Metal used plus the scrap it makes, less whatever "
                 "of that scrap you recover."),
        Col("stones_cost", "Stones cost", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({stone_count}{r}="",0,{stone_count}{r})*'
                'IF({stone_cost}{r}="",0,{stone_cost}{r}),2)')),
        Col("chain_cost", "Chain cost", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({chain_cm}{r}="",0,{chain_cm}{r})*' +
                _BASE.replace("{key}", "{chain_id}") + ',2)')),
        Col("wire_cost", "Wire cost", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({wire_cm}{r}="",0,{wire_cm}{r})*' +
                _BASE.replace("{key}", "{wire_id}") + ',2)')),
        Col("mat_cost", "Materials", 11, "money",
            formula=blank(
                "name",
                'ROUND({metal_cost}{r}+{stones_cost}{r}+{chain_cost}{r}+'
                '{wire_cost}{r}+IF({findings}{r}="",0,{findings}{r}),2)')),
        Col("minutes", "Minutes", 10, "number1",
            formula=blank(
                "name",
                'IF({design_mins}{r}="",0,{design_mins}{r})+'
                'IF({bench_mins}{r}="",0,{bench_mins}{r})+'
                'IF({polish_mins}{r}="",0,{polish_mins}{r})')),
        Col("rate", "Hourly rate", 11, "money",
            formula=blank(
                "name",
                'IF({rate_input}{r}<>"",{rate_input}{r},LaborCostHour)'),
            hidden=True),
        Col("labor_cost", "Labour", 11, "money",
            formula=blank(
                "name", 'ROUND({minutes}{r}/60*{rate}{r},2)')),
        Col("other_cost", "Extras", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({solder}{r}="",0,{solder}{r})+'
                'IF({plating}{r}="",0,{plating}{r})+'
                'IF({hallmark}{r}="",0,{hallmark}{r})+'
                '{minutes}{r}/60*SuppliesPerHour,2)'),
            note="Solder, plating, hallmarking and the bench "
                 "consumables charged per hour."),
        Col("metal_share", "Metal share", 10, "percent",
            formula=blank(
                "name",
                'IF(OR({mat_cost}{r}="",{mat_cost}{r}=0),"",'
                '{metal_cost}{r}/{mat_cost}{r})'),
            hidden=True,
            note="When this is most of your material cost, a silver "
                 "price rise is a re-pricing job, not a rounding "
                 "error."),
    ] + craft_tail() + [craft_check()]

    return SheetSpec(
        name=NAME,
        title="Jewelry Calculator",
        subtitle="Metal by the gram with scrap, stones, findings, "
                 "chain and bench time.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).jewelry,
        help_text="Weigh the metal, count the findings, time the "
                  "bench work. Link the row to a product ID and the "
                  "catalog prices it.",
        extra_notes=[
            "Metal prices move. Keep the price a gram current - a "
            "25% rise in silver on a piece that is 60% metal is a "
            "15% cost rise you will not feel until the quarter is "
            "over.",
            "Findings are the quiet killer. Ear wires, jump rings, "
            "a clasp and a tag can be $1.80 on a pair of earrings "
            "that cost $2.10 in silver.",
            "Recovering scrap is real money. 8% scrap on 12 g of "
            "sterling, recovered, is most of a cup of coffee on "
            "every single piece.",
        ],
    )
