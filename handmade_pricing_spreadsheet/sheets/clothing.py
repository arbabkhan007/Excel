"""`Clothing` - fabric with the layout waste in, notions, and the sewing.

Cutting a garment out of a rectangle always wastes fabric, and the
bigger the size the more of everything it takes. This tab costs a
sewn piece honestly: fabric by the centimetre with a layout
allowance, lining, interfacing, thread, zips, labels, and the four
separate jobs that make up "sewing".

Fabric quantities are in centimetres of length, which is the base
unit the Lists table converts metres and yards into.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import blank, craft_check, craft_tail

NAME = "Clothing"

_BASE = ('IFERROR(INDEX([[Materials.base_cost]],MATCH({key}{r},'
         '[[Materials.id]],0)),0)')


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("product", "Product ID", 12, "id",
            validate=list_from("ProductIDs")),
        Col("name", "Garment", 28, "text",
            note="'Linen pinafore', 'Toddler bucket hat'."),
        Col("garment_type", "Type", 18, "text",
            validate=list_from("FabricList"),
            note="The main fabric. The list is on the Lists tab and "
                 "you can add your own."),
        Col("size", "Size", 14, "text",
            validate=list_from("GarmentSizeList")),
        Col("size_factor", "Size factor", 10, "number",
            validate=number_check(0.1, 5),
            note="How much more fabric and time this size takes "
                 "than your standard one. Blank is 1."),
        Col("fabric_id", "Fabric material", 12, "id",
            validate=list_from("MaterialIDs")),
        Col("fabric_cm", "Fabric (cm)", 11, "number1",
            validate=number_check(0, 100000),
            note="Length of fabric at its usual width. Measure the "
                 "pattern layout, do not guess."),
        Col("layout_waste", "Layout waste", 11, "percent",
            validate=number_check(0, 1),
            note="What the cutting layout wastes - 10 to 25% is "
                 "normal, more on a directional print or a bias "
                 "cut. Blank uses the Settings default."),
        Col("lining_id", "Lining material", 12, "id",
            validate=list_from("MaterialIDs")),
        Col("lining_cm", "Lining (cm)", 11, "number1",
            validate=number_check(0, 100000)),
        Col("interfacing", "Interfacing", 11, "money",
            validate=number_check(0, 10000)),
        Col("thread", "Thread", 10, "money",
            validate=number_check(0, 10000),
            note="A reel is $3 and does ten garments. 30 cents is "
                 "not nothing when the margin is thin."),
        Col("notions", "Zips, buttons, elastic", 14, "money",
            validate=number_check(0, 10000)),
        Col("labels", "Labels & care tag", 12, "money",
            validate=number_check(0, 10000)),
        Col("cut_mins", "Cutting minutes", 12, "int",
            validate=number_check(0, 10000, decimal=False),
            note="Laying out, pinning, cutting. On a one-off this "
                 "is often longer than the sewing."),
        Col("sew_mins", "Sewing minutes", 12, "int",
            validate=number_check(0, 10000, decimal=False)),
        Col("finish_mins", "Finishing minutes", 13, "int",
            validate=number_check(0, 10000, decimal=False),
            note="Hems, buttonholes, hand finishing, labels."),
        Col("press_mins", "Pressing minutes", 13, "int",
            validate=number_check(0, 10000, decimal=False),
            note="Pressing is sewing. Count it."),
        Col("fit_mins", "Fitting & alterations", 14, "int",
            validate=number_check(0, 10000, decimal=False),
            note="Made to measure work lives or dies on this "
                 "number."),
        Col("rate_input", "Rate override", 11, "money",
            validate=number_check(0, 1000)),
        Col("price_input", "Your price", 11, "money",
            validate=number_check(0, 100000)),
        Col("notes", "Notes", 22, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("factor", "Size factor used", 10, "number",
            formula=blank(
                "name",
                'IF({size_factor}{r}="",1,{size_factor}{r})'),
            hidden=True),
        Col("waste_rate", "Waste used", 9, "percent",
            formula=blank(
                "name",
                'IF({layout_waste}{r}="",DefaultWaste,'
                '{layout_waste}{r})'),
            hidden=True),
        Col("fabric_total", "Fabric bought (cm)", 13, "number1",
            formula=blank(
                "name",
                'ROUND(IF({fabric_cm}{r}="",0,{fabric_cm}{r})*'
                '{factor}{r}*(1+{waste_rate}{r}),1)'),
            note="What you have to buy, layout waste and size "
                 "included."),
        Col("fabric_cost", "Fabric cost", 11, "money",
            formula=blank(
                "name",
                'ROUND({fabric_total}{r}*' +
                _BASE.replace("{key}", "{fabric_id}") + ',2)')),
        Col("lining_cost", "Lining cost", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({lining_cm}{r}="",0,{lining_cm}{r})*'
                '{factor}{r}*(1+{waste_rate}{r})*' +
                _BASE.replace("{key}", "{lining_id}") + ',2)')),
        Col("mat_cost", "Materials", 11, "money",
            formula=blank(
                "name",
                'ROUND({fabric_cost}{r}+{lining_cost}{r}+'
                'IF({interfacing}{r}="",0,{interfacing}{r})+'
                'IF({thread}{r}="",0,{thread}{r})+'
                'IF({notions}{r}="",0,{notions}{r})+'
                'IF({labels}{r}="",0,{labels}{r}),2)')),
        Col("minutes", "Minutes", 10, "number1",
            formula=blank(
                "name",
                'ROUND((IF({cut_mins}{r}="",0,{cut_mins}{r})+'
                'IF({sew_mins}{r}="",0,{sew_mins}{r})+'
                'IF({finish_mins}{r}="",0,{finish_mins}{r})+'
                'IF({press_mins}{r}="",0,{press_mins}{r})+'
                'IF({fit_mins}{r}="",0,{fit_mins}{r}))*{factor}{r},1)')),
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
        Col("waste_cost", "Cost of the waste", 12, "money",
            formula=blank(
                "name",
                'ROUND(IF({fabric_cm}{r}="",0,{fabric_cm}{r})*'
                '{factor}{r}*{waste_rate}{r}*' +
                _BASE.replace("{key}", "{fabric_id}") + ',2)'),
            hidden=True,
            note="What the offcuts cost. Worth knowing before you "
                 "decide whether a scrap-busting product is free "
                 "money or not."),
    ] + craft_tail() + [craft_check()]

    return SheetSpec(
        name=NAME,
        title="Clothing & Sewing Calculator",
        subtitle="Fabric with the layout waste in, notions, and all "
                 "four jobs that make up sewing.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).clothing,
        help_text="Enter the fabric at its usual width in "
                  "centimetres and a layout allowance. Size factors "
                  "stretch both the fabric and the time.",
        extra_notes=[
            "Cutting and pressing are sewing. A maker who only "
            "counts time at the machine undercharges by a third on "
            "every garment.",
            "Layout waste is real. 20% on $18 a metre linen is "
            "$3.60 a metre you paid for and did not use - price it "
            "in, then look for a layout that beats it.",
            "Made to measure needs a fitting allowance. If you do "
            "not charge for alterations you are working for free "
            "exactly when the customer is most demanding.",
        ],
    )
