"""`Material Usage` - the bill of materials: what goes into what.

One row per material in a product. 38 g of cream worsted, two wooden
buttons, 40 cm of ribbon. Each line is costed from the Materials tab,
the waste allowance is added here (and only here), and the product
catalog adds the lines up.

If a product is costed on one of the six specialist calculator tabs
instead, leave it out of this tab - use the bill of materials *or* a
specialist sheet, not both, or the material cost is counted twice.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import alert_cf, blank, lookup_material, lookup_product

NAME = "Material Usage"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("product", "Product", 12, "id",
            validate=list_from("ProductIDs"),
            note="The product ID from the Products tab. Type several "
                 "rows with the same ID - one per material."),
        Col("product_name", "Product name", 26, "text",
            formula=lookup_product("name"),
            note="Automatic, so you can see at a glance that the ID "
                 "you typed is the product you meant."),
        Col("material", "Material", 12, "id",
            validate=list_from("MaterialIDs")),
        Col("material_name", "Material name", 26, "text",
            formula=lookup_material("name")),
        Col("unit", "Unit", 9, "text", formula=lookup_material("unit")),
        Col("qty", "Quantity used", 12, "number1",
            validate=number_check(0, 1000000),
            note="In the unit shown. Weigh the yarn left over and "
                 "subtract - guessing is what makes a price wrong."),
        Col("waste", "Waste %", 9, "percent",
            validate=number_check(0, 1),
            note="Blank uses the material's own allowance, or the "
                 "default on Settings. Fabric cut on the bias wastes "
                 "far more than yarn does."),
        Col("optional", "Optional?", 10, "flag",
            validate=list_from("YesNoList"),
            note="Yes for an upgrade or an add-on that not every "
                 "piece gets. Optional lines are still costed - the "
                 "flag is there so you can see what is core."),
        Col("notes", "Notes", 26, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("waste_rate", "Waste used", 9, "percent",
            formula=blank(
                "material",
                'IF({waste}{r}<>"",{waste}{r},'
                'IFERROR(INDEX([[Materials.waste_rate]],'
                'MATCH({material}{r},[[Materials.id]],0)),DefaultWaste))'),
            hidden=True),
        Col("qty_eff", "Quantity with waste", 13, "number1",
            formula=blank(
                "material",
                'IF({qty}{r}="","",ROUND({qty}{r}*'
                '(1+{waste_rate}{r}),4))'),
            note="What you actually have to buy to make one."),
        Col("unit_cost", "Cost a unit", 11, "money4",
            formula=lookup_material("unit_cost", missing="")),
        Col("line_cost", "Line cost", 11, "money4",
            formula=blank(
                "material",
                'IF(OR({qty_eff}{r}="",{unit_cost}{r}=""),"",'
                'ROUND({qty_eff}{r}*{unit_cost}{r},4))'),
            note="Quantity, plus waste, times the cost a unit. This "
                 "is what the catalog adds up."),
        Col("product_total", "Materials in this product", 14, "money",
            formula=blank(
                "product",
                'ROUND(SUMIFS([[Material Usage.line_cost]],'
                '[[Material Usage.product]],{product}{r}),2)'),
            hidden=True),
        Col("share", "Share of the product", 11, "percent",
            formula=blank(
                "product",
                'IF(OR({line_cost}{r}="",{product_total}{r}="",'
                '{product_total}{r}=0),"",{line_cost}{r}/'
                '{product_total}{r})'),
            note="Which material is driving the cost. Usually one "
                 "line is most of it - that is the one to shop "
                 "around for."),
        Col("seq", "Line", 7, "int",
            formula=blank(
                "product",
                'COUNTIFS({product_first}:{product}{r},{product}{r})'),
            hidden=True),
        Col("key", "Key", 14, "text",
            formula=blank("product", '{product}{r}&"|"&{seq}{r}'),
            hidden=True,
            note="Used by the printable costing sheet to pull the "
                 "lines of one product in order."),
        Col("check", "Check", 24, "status",
            formula=(
                '=IF(AND({product}{r}="",{material}{r}=""),"",'
                'IF({product}{r}="","Add - which product?",'
                'IF({material}{r}="","Add - which material?",'
                'IF({product_name}{r}="Not found",'
                '"Check - no such product ID",'
                'IF({material_name}{r}="Add to Materials",'
                '"Check - no such material ID",'
                'IF(OR({qty}{r}="",{qty}{r}=0),"Add - how much?",'
                'IF({unit_cost}{r}="",'
                '"Check - that material has no price","OK")))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Bill of Materials",
        subtitle="What goes into each product, how much of it, and what "
                 "that costs.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).material_usage,
        help_text="One row per material per product. The catalog adds "
                  "the lines up by product ID, so the order of the "
                  "rows does not matter.",
        extra_notes=[
            "Weigh, do not guess. Make one, weigh what you used, and "
            "the price of everything you make from then on is built "
            "on a real number.",
            "Use this tab OR a specialist calculator tab for a given "
            "product, never both - the catalog adds both together, "
            "which is the one way to double count in this workbook.",
            "The share column tells you where the money goes. If one "
            "material is 70% of the cost, a 10% better price on it "
            "is worth more than every other saving combined.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Summary strip under the bill of materials."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Bill of materials summary", last)
    sheet.table_head(row, 0, ["Lines entered", "Products covered",
                              "Materials used", "Total material cost",
                              "Average per line", "Lines needing a look"],
                     [14, 17, 16, 19, 16, 20])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Material Usage.product]],"<>")',
        '=COUNTIFS([[Material Usage.seq]],1)',
        '=COUNTIF([[Materials.lines]],">0")',
        '=ROUND(SUM([[Material Usage.line_cost]]),2)',
        '=IFERROR(ROUND(SUM([[Material Usage.line_cost]])/'
        'COUNTIF([[Material Usage.product]],"<>"),2),"")',
        '=COUNTIFS([[Material Usage.check]],"Check*")+'
        'COUNTIFS([[Material Usage.check]],"Add*")',
    ], ["int", "int", "int", "money", "money", "int"])
