"""`Materials` - every supply you buy, costed down to the gram.

A 100 g skein that cost $8.50 is not an $8.50 cost on a product that
used 38 g of it. This tab turns what you *bought* into what one gram,
one bead, one inch of it costs, and the bill of materials on the next
tab multiplies that by what you actually used.

Waste is deliberately NOT added here - it belongs on the usage line, so
nothing ever gets counted twice. The "with waste" column is a reference
figure only.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, blank, id_col, rank_desc, status_cf

NAME = "Materials"

_USE = "[[Material Usage.material]]"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "MaterialPrefix", digits=4, width=10,
               note="Generated automatically. The bill of materials "
                    "picks from these IDs, so filter the list rather "
                    "than re-sorting it."),
        Col("name", "Material", 26, "text",
            note="What you call it when you reorder. Be specific - "
                 "'Cascade 220 worsted, cream' beats 'wool'."),
        Col("type", "Type", 18, "text",
            validate=list_from("MaterialTypeList")),
        Col("detail", "Colour / size / detail", 20, "text"),
        Col("supplier", "Supplier", 18, "text"),
        Col("code", "Supplier code", 14, "text",
            note="Their SKU. Saves ten minutes every time you reorder."),
        Col("pack_qty", "Quantity in a pack", 12, "number1",
            validate=number_check(0, 1000000),
            note="How much one purchase contains: 100 (g in a skein), "
                 "144 (beads in a gross), 2.4 (metres on a board)."),
        Col("unit", "Unit", 11, "text", validate=list_from("UnitList"),
            note="The unit that quantity is measured in. The unit "
                 "table on the Lists tab converts it to a base unit."),
        Col("pack_cost", "Pack price", 11, "money",
            validate=number_check(0, 1000000),
            note="What the supplier charges for one pack, before any "
                 "trade discount and before delivery."),
        Col("pack_ship", "Delivery per pack", 12, "money",
            validate=number_check(0, 100000),
            note="Postage, divided by how many packs were on the "
                 "order. It is part of what the material costs you."),
        Col("discount", "Trade discount", 11, "percent",
            validate=number_check(0, 1)),
        Col("waste", "Waste allowance", 11, "percent",
            validate=number_check(0, 1),
            note="Leave blank to use the default on Settings. This is "
                 "the share that becomes offcuts and mistakes - it is "
                 "applied on the usage line, never here."),
        Col("stock", "Stock on hand", 11, "number1",
            validate=number_check(0, 10000000),
            note="In the unit above. Count the shelf now and again."),
        Col("reorder", "Reorder at", 10, "number1",
            validate=number_check(0, 10000000)),
        Col("last_cost", "Previous pack price", 12, "money",
            validate=number_check(0, 1000000),
            note="The price before the last rise. The workbook shows "
                 "you how far it has moved."),
        Col("priced_on", "Price checked", 11, "date", validate=date_check(),
            note="When you last confirmed this price. Material costs "
                 "drift upwards quietly."),
        Col("notes", "Notes", 26, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("net_pack", "Pack cost to you", 12, "money",
            formula=blank(
                "name",
                'IF({pack_cost}{r}="","",ROUND({pack_cost}{r}*'
                '(1-IF({discount}{r}="",0,{discount}{r}))+'
                'IF({pack_ship}{r}="",0,{pack_ship}{r}),2))'),
            note="Pack price less the trade discount, plus delivery."),
        Col("unit_cost", "Cost a unit", 11, "money4",
            formula=blank(
                "name",
                'IF(OR({net_pack}{r}="",{pack_qty}{r}="",'
                '{pack_qty}{r}=0),"",ROUND({net_pack}{r}/'
                '{pack_qty}{r},4))'),
            note="The number everything else multiplies: what one "
                 "gram, bead or inch costs you."),
        Col("base_unit", "Counted in", 10, "text",
            formula=blank(
                "name",
                'IFERROR(INDEX(UnitBases,MATCH({unit}{r},UnitNames,0)),"")'),
            hidden=True),
        Col("base_factor", "Unit factor", 9, "number",
            formula=blank(
                "name",
                'IFERROR(INDEX(UnitFactors,MATCH({unit}{r},UnitNames,0)),"")'),
            hidden=True),
        Col("base_cost", "Cost a base unit", 13, "money4",
            formula=blank(
                "name",
                'IF(OR({unit_cost}{r}="",{base_factor}{r}="",'
                '{base_factor}{r}=0),"",ROUND({unit_cost}{r}/'
                '{base_factor}{r},4))'),
            note="The same cost expressed per gram, per millilitre or "
                 "per centimetre, so two suppliers selling different "
                 "pack sizes can be compared honestly."),
        Col("waste_rate", "Waste used", 9, "percent",
            formula=blank(
                "name",
                'IF({waste}{r}="",DefaultWaste,{waste}{r})'),
            hidden=True),
        Col("cost_waste", "Cost a unit with waste", 13, "money4",
            formula=blank(
                "name",
                'IF({unit_cost}{r}="","",ROUND({unit_cost}{r}*'
                '(1+{waste_rate}{r}),4))'),
            note="Reference only. The waste is added on the bill of "
                 "materials line, so it is not double counted here."),
        Col("stock_value", "Stock value", 11, "money",
            formula=blank(
                "name",
                'IF(OR({stock}{r}="",{unit_cost}{r}=""),"",'
                'ROUND({stock}{r}*{unit_cost}{r},2))'),
            note="Add this column up. Most makers are holding far more "
                 "money on the shelf than they think."),
        Col("lines", "Used in", 9, "int",
            formula=blank("name", f'COUNTIFS({_USE},{{id}}{{r}})'),
            note="How many product recipes call for this material."),
        Col("qty_used", "Quantity committed", 12, "number1",
            formula=blank(
                "name",
                f'SUMIFS([[Material Usage.qty_eff]],{_USE},{{id}}{{r}})'),
            note="One of each product that uses it would consume this "
                 "much, waste included."),
        Col("cost_used", "Cost committed", 12, "money",
            formula=blank(
                "name",
                f'ROUND(SUMIFS([[Material Usage.line_cost]],{_USE},'
                f'{{id}}{{r}}),2)')),
        Col("cover", "Stock covers", 9, "int",
            formula=blank(
                "name",
                'IF(OR({stock}{r}="",{qty_used}{r}="",'
                '{qty_used}{r}<=0),"",ROUNDDOWN({stock}{r}/'
                '{qty_used}{r},0))'),
            note="How many complete sets of the products that use it "
                 "you could make from the stock on hand."),
        Col("change", "Price move", 10, "percent",
            formula=blank(
                "name",
                'IF(OR({last_cost}{r}="",{last_cost}{r}=0,'
                '{pack_cost}{r}=""),"",({pack_cost}{r}-{last_cost}{r})/'
                '{last_cost}{r})'),
            cf=[{"type": "cell", "criteria": ">", "value": 0.075,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "good"}],
            note="A material up 10% is a product that quietly stopped "
                 "making money. Re-price it."),
        Col("status", "Stock", 12, "status",
            formula=blank(
                "name",
                'IF({stock}{r}="","Not tracked",'
                'IF({stock}{r}<=0,"Out of stock",'
                'IF(AND({reorder}{r}<>"",{stock}{r}<={reorder}{r}),'
                '"Low stock","In stock")))'),
            cf=status_cf()),
        rank_desc("cost_used", guard='{name}{r}=""',
                  header="Spend rank",
                  note="1 = the material you spend the most on. The "
                       "Reports tab reads this."),
        Col("alert", "Action", 28, "status",
            formula=blank(
                "name",
                'IF({pack_cost}{r}="","Add - no price yet",'
                'IF(OR({pack_qty}{r}="",{pack_qty}{r}=0),'
                '"Add - how much is in a pack?",'
                'IF({status}{r}="Out of stock","Order - none left",'
                'IF({status}{r}="Low stock","Order - running low",'
                'IF(AND({change}{r}<>"",{change}{r}>0.075),'
                '"Check - price jumped, re-price the products",'
                'IF(AND({priced_on}{r}<>"",TodayDate-{priced_on}{r}>365),'
                '"Check - price is over a year old","OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Materials & Supplies",
        subtitle="What you buy, and what one unit of it actually costs "
                 "once discount and delivery are in.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).materials,
        help_text="Enter what a pack costs and how much is in it. The "
                  "workbook works out the cost per unit and every "
                  "product that uses it re-prices itself.",
        extra_notes=[
            "Delivery is part of the cost. A $6 pack with $4 postage "
            "is a $10 pack, and pretending otherwise is how makers "
            "end up working for nothing.",
            "Keep the previous price in its column when you update "
            "one. The price move column is an early warning that a "
            "product needs re-pricing.",
            "Waste belongs on the bill of materials line, not here - "
            "that way a material used carefully in one product and "
            "wastefully in another is costed correctly in both.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Summary strip under the material list."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1        # 0 based row under the data
    last = spec_.last_col_index
    row = sheet.section(row, "Materials summary", last)
    sheet.table_head(row, 0, ["Materials listed", "Priced", "Missing a price",
                              "Value on the shelf", "Low or out of stock",
                              "Prices up since last check",
                              "Not checked in a year"],
                     [17, 12, 16, 18, 19, 22, 20])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Materials.name]],"<>")',
        '=COUNTIFS([[Materials.name]],"<>",[[Materials.pack_cost]],"<>")',
        '=COUNTIFS([[Materials.name]],"<>",[[Materials.pack_cost]],"")',
        '=SUM([[Materials.stock_value]])',
        '=COUNTIFS([[Materials.status]],"Low stock")+'
        'COUNTIFS([[Materials.status]],"Out of stock")',
        '=COUNTIFS([[Materials.change]],">0")',
        '=COUNTIFS([[Materials.name]],"<>",[[Materials.priced_on]],'
        '"<"&TodayDate-365)',
    ], ["int", "int", "int", "money", "int", "int", "int"])
