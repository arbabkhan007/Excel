"""`Overhead` - the bills that exist whether you sell anything or not.

Studio rent, insurance, the shop subscription, the website, the
accountant, the ring light you bought in January. None of it belongs
to a single product, and all of it has to be paid out of the margin on
the products you sell.

Type the bills here with how often they come. The workbook turns them
into a monthly figure, then into an hourly and a per-item rate - and
the Settings tab decides which of those every product carries.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import alert_cf, bar_cf, blank, id_col

NAME = "Overhead"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", '"OH-"', digits=3, width=9),
        Col("name", "Cost", 26, "text",
            note="What the bill is called on your bank statement."),
        Col("type", "Type", 20, "text",
            validate=list_from("OverheadTypeList")),
        Col("supplier", "Paid to", 18, "text"),
        Col("amount", "Amount", 11, "money",
            validate=number_check(0, 1000000)),
        Col("frequency", "How often", 14, "text",
            validate=list_from("FrequencyList"),
            note="One-off costs are spread over twelve months, which "
                 "is the honest way to carry a kiln or a camera."),
        Col("start", "Started", 11, "date", validate=date_check()),
        Col("ends", "Ends", 11, "date", validate=date_check(),
            note="Leave blank for an ongoing cost. A finished cost "
                 "stops counting against your products."),
        Col("essential", "Essential?", 11, "flag",
            validate=list_from("YesNoList"),
            note="Be honest. The reports total what you could stop "
                 "paying tomorrow."),
        Col("notes", "Notes", 26, "wrap"),
        # ------------------------------------------------- the automatics --
        Col("active", "Active", 8, "flag",
            formula=blank(
                "name",
                'IF(AND({ends}{r}<>"",{ends}{r}<TodayDate),"No","Yes")'),
            hidden=True),
        Col("monthly", "A month", 11, "money",
            formula=blank(
                "name",
                'IF(OR({amount}{r}="",{active}{r}="No"),0,'
                'ROUND({amount}{r}*IF({frequency}{r}="Monthly",1,'
                'IF({frequency}{r}="Quarterly",1/3,'
                'IF({frequency}{r}="Twice a year",1/6,'
                'IF({frequency}{r}="Yearly",1/12,'
                'IF({frequency}{r}="One-off",1/12,0))))),2))'),
            cf=bar_cf(),
            note="Everything converted to what it costs you each "
                 "month. The Settings tab adds this column up."),
        Col("yearly", "A year", 11, "money",
            formula=blank("name", 'ROUND({monthly}{r}*12,2)')),
        Col("share", "Share of overhead", 12, "percent",
            formula=blank(
                "name",
                'IF(OverheadMonthly=0,"",{monthly}{r}/OverheadMonthly)'),
            cf=bar_cf()),
        Col("per_hour", "Per hour worked", 12, "money",
            formula=blank(
                "name",
                'IF(ProductiveHours<=0,"",ROUND({monthly}{r}/'
                'ProductiveHours,2))')),
        Col("per_unit", "Per item made", 11, "money",
            formula=blank(
                "name",
                'IF(UnitsPerMonth<=0,"",ROUND({monthly}{r}/'
                'UnitsPerMonth,2))')),
        Col("units_to_cover", "Items to cover it", 13, "int",
            formula=blank(
                "name",
                'IF(OR(OverheadMonthly=0,{monthly}{r}=0),"",'
                'ROUNDUP({monthly}{r}/MAX(0.01,'
                'IFERROR(AVERAGE([[Products.profit]]),0.01)),0))'),
            note="How many average products you have to sell each "
                 "month just to pay this one bill."),
        Col("check", "Check", 24, "status",
            formula=blank(
                "name",
                'IF({amount}{r}="","Add - how much is it?",'
                'IF({frequency}{r}="","Add - how often?",'
                'IF({active}{r}="No","Ended - no longer counted",'
                'IF(AND({essential}{r}="No",{monthly}{r}>'
                'OverheadMonthly*0.15),'
                '"Check - a big bill you called optional","OK"))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Overhead Costs",
        subtitle="The bills that are not about any one product, spread "
                 "across everything you make.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).overhead,
        help_text="Enter the bill and how often it arrives. The "
                  "monthly column is what the Settings tab divides "
                  "across your hours or your items.",
        extra_notes=[
            "Overhead is the quiet difference between a hobby that "
            "breaks even and a business that pays you. A $420 month "
            "of overhead spread over 55 items is $7.64 a product - "
            "on a $20 item that is more than a third of the margin.",
            "Spread one-off purchases over a year rather than "
            "dropping the whole cost on one month. The kiln does "
            "not stop being useful in February.",
            "Review the 'not essential' lines twice a year. Most "
            "makers find a subscription they stopped using.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Summary strip under the bills."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Overhead summary", last)
    sheet.table_head(row, 0, ["Costs listed", "Total a month",
                              "Total a year", "Per hour worked",
                              "Per item made", "Could be cut",
                              "Items to break even"],
                     [13, 15, 15, 16, 15, 14, 19])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Overhead.name]],"<>")',
        '=ROUND(SUM([[Overhead.monthly]]),2)',
        '=ROUND(SUM([[Overhead.yearly]]),2)',
        '=OverheadPerHour',
        '=OverheadPerUnit',
        '=ROUND(SUMIFS([[Overhead.monthly]],[[Overhead.essential]],"No"),2)',
        '=IFERROR(ROUNDUP(OverheadMonthly/MAX(0.01,'
        'AVERAGE([[Products.profit]])),0),"")',
    ], ["int", "money", "money", "money", "money", "money", "int"])
