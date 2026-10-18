"""`Labor` - your hours, stage by stage, priced like the work they are.

One row per stage per product: 35 minutes cutting, 240 minutes making,
20 minutes finishing, 10 minutes packing. The cost of a stage is
minutes / 60 x the hourly rate of whoever does it - the formula every
pricing guide quotes and almost no maker actually applies.

Stages marked "per batch" are divided by the batch size, so the hour
you spend setting up a run of twenty is twenty three-minute shares,
not twenty hours.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import alert_cf, blank, lookup_product

NAME = "Labor"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("product", "Product", 12, "id",
            validate=list_from("ProductIDs"),
            note="The product ID. Add as many stages as the piece "
                 "really has - the total is what the catalog uses."),
        Col("product_name", "Product name", 26, "text",
            formula=lookup_product("name")),
        Col("stage", "Stage", 20, "text",
            validate=list_from("LaborStageList"),
            note="Making is never the whole story. Photography, "
                 "listing and packing are work too, and they are "
                 "where small items lose money."),
        Col("who", "Who does it", 14, "text",
            validate=list_from("LaborWhoList"),
            note="Sets the hourly rate: you, an employee or a "
                 "contractor. Override it in the next column if this "
                 "particular job pays differently."),
        Col("rate_input", "Rate override", 11, "money",
            validate=number_check(0, 1000),
            note="Leave blank to use the rate on the Settings tab for "
                 "whoever is doing it."),
        Col("minutes", "Minutes", 10, "number1",
            validate=number_check(0, 100000),
            note="Time it once with a timer, not from memory. Makers "
                 "underestimate their own hours by about a third."),
        Col("basis", "Per item or batch", 13, "text",
            validate=list_from("PerBasisList"),
            note="'Per batch' for anything you do once for a whole "
                 "run - setting up, pouring a pot of wax, a post "
                 "office trip."),
        Col("batch_qty", "Items in that batch", 12, "int",
            validate=number_check(1, 100000, decimal=False),
            note="Only used when the stage is per batch. Blank uses "
                 "the default batch size on Settings."),
        Col("skilled", "Skilled work?", 11, "flag",
            validate=list_from("YesNoList"),
            note="Flag the stages that took years to learn. The "
                 "reports show how much of your making time is "
                 "skilled - that is the part worth protecting."),
        Col("notes", "Notes", 24, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("rate", "Hourly rate", 11, "money",
            formula=blank(
                "product",
                'IF({rate_input}{r}<>"",{rate_input}{r},'
                'IF({who}{r}="Employee",ROUND(EmployeeRate*'
                '(1+PayrollBurden),2),'
                'IF(OR({who}{r}="Contractor",{who}{r}="Outsourced"),'
                'ROUND(ContractorRate*(1+PayrollBurden),2),'
                'LaborCostHour)))'),
            note="Your own hours carry the on-costs from Settings "
                 "too - holiday and tax do not stop existing because "
                 "you are self employed."),
        Col("hours", "Hours", 9, "number",
            formula=blank(
                "product",
                'IF({minutes}{r}="","",ROUND({minutes}{r}/60,3))'),
            note="Minutes / 60. The textbook formula."),
        Col("batch_size", "Batch used", 10, "int",
            formula=blank(
                "product",
                'IF({basis}{r}="Per batch",IF(OR({batch_qty}{r}="",'
                '{batch_qty}{r}=0),DefaultBatchSize,{batch_qty}{r}),1)'),
            hidden=True),
        Col("stage_cost", "Cost of the stage", 12, "money",
            formula=blank(
                "product",
                'IF(OR({hours}{r}="",{rate}{r}=""),"",'
                'ROUND({hours}{r}*{rate}{r},2))'),
            note="Hours x rate, before it is shared over a batch."),
        Col("cost", "Cost per item", 11, "money",
            formula=blank(
                "product",
                'IF({stage_cost}{r}="","",ROUND({stage_cost}{r}/'
                'MAX(1,{batch_size}{r}),2))'),
            note="What this stage adds to one finished piece. This is "
                 "the column the catalog sums."),
        Col("mins_item", "Minutes per item", 12, "number1",
            formula=blank(
                "product",
                'IF({minutes}{r}="","",ROUND({minutes}{r}/'
                'MAX(1,{batch_size}{r}),2))')),
        Col("consumables", "Consumables", 11, "money",
            formula=blank(
                "product",
                'IF({hours}{r}="","",ROUND({hours}{r}/'
                'MAX(1,{batch_size}{r})*SuppliesPerHour,2))'),
            note="Thread, glue, sandpaper, the sponge you ruin. "
                 "Charged per hour worked from the Settings rate."),
        Col("product_total", "Labour in this product", 13, "money",
            formula=blank(
                "product",
                'ROUND(SUMIFS([[Labor.cost]],[[Labor.product]],'
                '{product}{r}),2)'),
            hidden=True),
        Col("share", "Share of the labour", 11, "percent",
            formula=blank(
                "product",
                'IF(OR({cost}{r}="",{product_total}{r}="",'
                '{product_total}{r}=0),"",{cost}{r}/'
                '{product_total}{r})')),
        Col("seq", "Line", 7, "int",
            formula=blank(
                "product",
                'COUNTIFS({product_first}:{product}{r},{product}{r})'),
            hidden=True),
        Col("key", "Key", 14, "text",
            formula=blank("product", '{product}{r}&"|"&{seq}{r}'),
            hidden=True),
        Col("check", "Check", 24, "status",
            formula=(
                '=IF({product}{r}="","",'
                'IF({product_name}{r}="Not found",'
                '"Check - no such product ID",'
                'IF(OR({minutes}{r}="",{minutes}{r}=0),'
                '"Add - how many minutes?",'
                'IF({stage}{r}="","Add - which stage?",'
                'IF(AND({basis}{r}="Per batch",{batch_qty}{r}=""),'
                '"Check - using the default batch size",'
                'IF({rate}{r}<MinimumWage,'
                '"Check - below minimum wage","OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Labor & Time",
        subtitle="Every stage of making it, timed and priced at a real "
                 "hourly rate.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).labor,
        help_text="One row per stage. Cost = minutes / 60 x hourly "
                  "rate, and batch stages are shared across the batch.",
        extra_notes=[
            "Time yourself once, properly, with a timer running. "
            "Nearly every maker who does this discovers the piece "
            "takes half again as long as they believed.",
            "Count the invisible stages. Photographing, listing, "
            "answering messages and standing in the post office "
            "queue are hours you are not paid for unless you put "
            "them here.",
            "If a product's price cannot carry an honest hourly "
            "rate, the answer is a different product or a different "
            "price - not a quieter rate.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Summary strip under the labour lines."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Labour summary", last)
    sheet.table_head(row, 0, ["Stages entered", "Products covered",
                              "Total minutes per item", "Total labour cost",
                              "Average rate", "Skilled stages",
                              "Lines needing a look"],
                     [15, 17, 20, 18, 14, 15, 20])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Labor.product]],"<>")',
        '=COUNTIFS([[Labor.seq]],1)',
        '=ROUND(SUM([[Labor.mins_item]]),0)',
        '=ROUND(SUM([[Labor.cost]]),2)',
        '=IFERROR(ROUND(AVERAGE([[Labor.rate]]),2),"")',
        '=COUNTIFS([[Labor.skilled]],"Yes")',
        '=COUNTIFS([[Labor.check]],"Check*")+'
        'COUNTIFS([[Labor.check]],"Add*")',
    ], ["int", "int", "int", "money", "money", "int", "int"])
