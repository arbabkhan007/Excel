"""`Waste Log` - the material that never made it into a product.

The frogged shawl, the failed pour, the board that split, the dye lot
that did not match. None of it is free, all of it was bought, and
most makers have no idea what it adds up to in a year.

Logging waste does two things: it tells you what your real material
cost is, and it shows you which mistakes are worth engineering out.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, flag_yes_no_cf, lookup_material,
                     month_col)

NAME = "Waste Log"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        Col("date", "Date", 11, "date", validate=date_check()),
        Col("material", "Material", 12, "id",
            validate=list_from("MaterialIDs")),
        Col("material_name", "Material name", 24, "text",
            formula=lookup_material("name")),
        Col("qty", "Quantity lost", 12, "number1",
            validate=number_check(0, 1000000),
            note="In the material's own unit."),
        Col("reason", "What happened", 22, "text",
            validate=list_from("WasteReasonList")),
        Col("product", "On which product", 12, "id",
            validate=list_from("ProductIDs"),
            note="Optional. Useful when one design keeps going "
                 "wrong."),
        Col("avoidable", "Avoidable?", 11, "flag",
            validate=list_from("YesNoList"), cf=flag_yes_no_cf(),
            note="Trim waste is not avoidable. A miscut is. Be "
                 "honest and the totals become useful."),
        Col("notes", "Notes", 28, "wrap"),
        # ------------------------------------------------- the automatics --
        month_col("date"),
        Col("unit", "Unit", 9, "text", formula=lookup_material("unit"),
            hidden=True),
        Col("unit_cost", "Cost a unit", 11, "money4",
            formula=lookup_material("unit_cost", missing=""),
            hidden=True),
        Col("cost", "Cost", 11, "money",
            formula=blank(
                "date",
                'IF(OR({qty}{r}="",{unit_cost}{r}=""),"",'
                'ROUND({qty}{r}*{unit_cost}{r},2))'),
            note="What walked out of the door. Add the column up "
                 "before you decide a mistake does not matter."),
        Col("share", "Share of that material", 13, "percent",
            formula=blank(
                "date",
                'IFERROR({cost}{r}/SUMIFS([[Waste Log.cost]],'
                '[[Waste Log.material]],{material}{r}),"")'),
            hidden=True),
        Col("check", "Note", 28, "status",
            formula=blank(
                "date",
                'IF({material}{r}="","Add - which material?",'
                'IF({qty}{r}="","Add - how much?",'
                'IF({unit_cost}{r}="",'
                '"Check - that material has no price",'
                'IF({avoidable}{r}="Yes",'
                '"Avoidable - worth a second look","Logged"))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Material Waste Log",
        subtitle="What got wasted, what it cost, and how much of "
                 "it you could stop.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).waste,
        help_text="Log it when it happens - nobody remembers at "
                  "the end of the month. The totals tell you "
                  "whether your waste allowance is right.",
        extra_notes=[
            "Compare the total here with the waste allowance in "
            "your prices. If you allow 5% and waste 12%, every "
            "product is under-priced.",
            "Separate avoidable from unavoidable. Offcuts are "
            "physics; a miscut because you did not check the "
            "pattern twice is a process.",
            "Wasted material is bought material. A $9.40 skein "
            "frogged and unusable is $9.40 off the month, not a "
            "learning experience.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Summary strip under the waste log."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Waste summary", last)
    sheet.table_head(row, 0, ["Entries", "Cost of the waste",
                              "Avoidable", "This month",
                              "Waste as % of material spend"],
                     [12, 18, 14, 14, 26])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Waste Log.date]],"<>")',
        '=ROUND(SUM([[Waste Log.cost]]),2)',
        '=ROUND(SUMIFS([[Waste Log.cost]],[[Waste Log.avoidable]],'
        '"Yes"),2)',
        '=ROUND(SUMIFS([[Waste Log.cost]],[[Waste Log.date]],'
        '">="&MonthStart),2)',
        '=IFERROR(SUM([[Waste Log.cost]])/SUM([[Materials.cost_used]]),"")',
    ], ["int", "money", "money", "money", "percent"])
