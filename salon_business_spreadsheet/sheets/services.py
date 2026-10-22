"""`Services` - the menu, with what each treatment really earns.

A salon price list is usually a list of numbers somebody typed once.
This one costs every service: the backbar product it burns, the
commission it pays, the chair time it occupies - and then says what it
makes an hour, which is the only number that compares a $35 blow dry
with a $220 balayage honestly.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check, date_check
from .common import (alert_cf, blank, id_col, loss_cf, rank_desc, round_price,
                     status_cf, yesno)

NAME = "Services"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "ServicePrefix", digits=3, width=10,
               note="Automatic. Bookings and sales point at this ID, so "
                    "renaming a service never breaks your history."),
        Col("name", "Service", 30, "text",
            note="Exactly as it appears on your menu and on the booking "
                 "page."),
        Col("category", "Category", 22, "text",
            validate=list_from("ServiceCategoryList"),
            note="Drives the revenue-by-category figures on the "
                 "dashboard and the analytics tab."),
        Col("description", "Description", 34, "wrap",
            note="What the client is buying. Used on the price list "
                 "print-out."),
        Col("duration", "Minutes", 9, "int",
            validate=number_check(0, 1440, decimal=False),
            note="Chair time, start to finish. Include the time you "
                 "cannot sell to anybody else."),
        Col("buffer", "Turnaround", 10, "int",
            validate=number_check(0, 240, decimal=False),
            note="Clean down and reset minutes after the client leaves. "
                 "Counted in the chair time, not charged to them."),
        Col("price", "Price", 11, "money",
            validate=number_check(0, 100000),
            note="What you charge today. The suggested price column "
                 "says what the maths would ask for."),
        Col("product_cost", "Product cost", 11, "money",
            validate=number_check(0, 100000),
            note="Backbar cost of one service - colour, developer, "
                 "wax, lash adhesive. Leave blank and the workbook "
                 "estimates it from the percentage on Settings."),
        Col("addon", "Add-on?", 9, "flag", validate=yesno(),
            note="Yes for treatments, toners and finishes sold "
                 "alongside something else."),
        Col("commissionable", "Pays commission?", 13, "flag",
            validate=yesno(),
            note="No for anything the owner keeps whole - a patch "
                 "test, a fringe trim, a redo."),
        Col("online", "Bookable online?", 13, "flag", validate=yesno()),
        Col("level", "Stylist level", 14, "text",
            validate=list_from("StylistLevelList"),
            note="Who is allowed to do it. Staff specific pricing lives "
                 "in the price override on the appointment."),
        Col("status", "Status", 12, "text",
            validate=list_from("ServiceStatusList"), cf=status_cf()),
        Col("reviewed", "Price set on", 11, "date", validate=date_check(),
            note="The last time this price changed. The alert column "
                 "nags when it is older than the review period."),
        Col("notes", "Notes", 26, "wrap"),
        # ------------------------------------------------- what it costs --
        Col("hours", "Chair hours", 10, "number",
            formula=blank(
                "name",
                'ROUND((IF({duration}{r}="",0,{duration}{r})+'
                'IF({buffer}{r}="",0,{buffer}{r}))/60,3)'),
            hidden=True),
        Col("cost_product", "Product cost used", 12, "money",
            formula=blank(
                "name",
                'IF({product_cost}{r}<>"",{product_cost}{r},'
                'IF({price}{r}="",0,'
                'ROUND({price}{r}*ProductCostPercent,2)))'),
            note="Yours if you typed one, otherwise the percentage "
                 "estimate from Settings."),
        Col("cost_commission", "Commission", 11, "money",
            formula=blank(
                "name",
                'IF({commissionable}{r}="No",0,IF({price}{r}="",0,'
                'ROUND({price}{r}*ServiceCommission,2)))'),
            note="At the default service rate on Settings. The real "
                 "figure per stylist is on the Commissions tab."),
        Col("cost_total", "Direct cost", 11, "money",
            formula=blank(
                "name",
                'ROUND({cost_product}{r}+{cost_commission}{r},2)'),
            note="Product plus commission. Rent, power and the front "
                 "desk are overheads - they are in the profit and loss, "
                 "not here."),
        Col("profit", "Gross profit", 11, "money",
            formula=blank("name",
                          'IF({price}{r}="","",'
                          'ROUND({price}{r}-{cost_total}{r},2))'),
            cf=loss_cf(),
            note="Price less the cost of delivering it. This is the "
                 "money that has to cover the rent."),
        Col("margin", "Margin", 9, "percent",
            formula=blank(
                "name",
                'IF({price}{r}="","",IF({price}{r}=0,"",'
                '{profit}{r}/{price}{r}))'),
            cf=[{"type": "cell", "criteria": "<", "value": 0, "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": 0.4,
                 "fmt": "warn"},
                {"type": "cell", "criteria": ">=", "value": 0.4,
                 "fmt": "good"}]),
        Col("price_hour", "Price an hour", 11, "money",
            formula=blank(
                "name",
                'IF({price}{r}="","",IF({hours}{r}=0,"",'
                'ROUND({price}{r}/{hours}{r},2)))'),
            cf=[{"type": "cell", "criteria": "<", "value": "=TargetChairRate",
                 "fmt": "warn"},
                {"type": "cell", "criteria": ">=",
                 "value": "=TargetChairRate", "fmt": "good"}],
            note="The number that compares services fairly. A cheap "
                 "service in half the time can beat an expensive one."),
        Col("profit_hour", "Profit an hour", 11, "money",
            formula=blank(
                "name",
                'IF({profit}{r}="","",IF({hours}{r}=0,"",'
                'ROUND({profit}{r}/{hours}{r},2)))'),
            cf=loss_cf()),
        Col("suggested_raw", "Suggested, raw", 11, "money",
            formula=blank(
                "name",
                'IF({hours}{r}=0,"",IF(ServiceCommission>=1,"",'
                '({cost_product}{r}+TargetChairRate*{hours}{r})/'
                'IF({commissionable}{r}="No",1,1-ServiceCommission)))'),
            hidden=True),
        Col("suggested", "Suggested price", 12, "money",
            formula=blank(
                "name",
                'IF({suggested_raw}{r}="","",ROUND(' +
                round_price("{suggested_raw}{r}") + ',2))'),
            note="(Product cost + your target chair rate x the hours) / "
                 "(1 - commission), rounded by the rule on Settings. "
                 "Change the target, not the formula."),
        Col("gap", "Price gap", 10, "money",
            formula=blank(
                "name",
                'IF({suggested}{r}="","",IF({price}{r}="","",'
                'ROUND({price}{r}-{suggested}{r},2)))'),
            cf=loss_cf(),
            note="Positive means you charge more than the maths asks "
                 "for. Negative is money you are leaving behind."),
        # ----------------------------------------------- how it is doing --
        Col("bookings", "Booked", 9, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Appointments.service]],{id}{r})'),
            note="Every appointment that names this service, whatever "
                 "happened to it."),
        Col("done", "Completed", 10, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Appointments.service]],{id}{r},'
                '[[Appointments.status]],"Completed")')),
        Col("units", "Sold", 8, "int",
            formula=blank(
                "name",
                'SUMIFS([[Sales.qty]],[[Sales.item]],{id}{r},'
                '[[Sales.kind]],"Service")'),
            hidden=True),
        Col("revenue", "Revenue", 12, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.net]],[[Sales.item]],{id}{r},'
                '[[Sales.kind]],"Service"),2)'),
            cf=[{"type": "data_bar", "bar_color": Palette.primary_mid}],
            note="Net of discounts, from the Sales tab."),
        Col("revenue_share", "Share of sales", 11, "percent",
            formula=blank(
                "name",
                'IF(ServiceRevenueAll=0,"",{revenue}{r}/ServiceRevenueAll)'),
            hidden=True),
        Col("profit_total", "Profit earned", 12, "money",
            formula=blank(
                "name",
                'ROUND({units}{r}*{profit}{r},2)'),
            hidden=True),
        Col("hours_total", "Chair hours used", 12, "number1",
            formula=blank("name", 'ROUND({units}{r}*{hours}{r},1)'),
            hidden=True),
        Col("retail_attach", "Retail attached", 11, "percent",
            formula=blank(
                "name",
                'IF({done}{r}=0,"",COUNTIFS([[Sales.item]],{id}{r},'
                '[[Sales.retail_with]],"Yes")/{done}{r})'),
            note="How often a client who had this service also bought "
                 "something to take home."),
        Col("last_done", "Last done", 11, "date",
            formula=blank(
                "name",
                'IFERROR(IF(COUNTIFS([[Appointments.service]],{id}{r})=0,'
                '"",SUMPRODUCT(MAX(([[Appointments.service]]={id}{r})*'
                'IFERROR([[Appointments.date]]*1,0)))),"")')),
        rank_desc("revenue", header="Revenue rank"),
        rank_desc("profit_hour", header="Rate rank"),
        rank_desc("done", header="Popularity"),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "name",
                'IF({price}{r}="","Add - no price yet",'
                'IF({duration}{r}="","Add - how long does it take?",'
                'IF({profit}{r}<0,"Loss - this service costs you money",'
                'IF({price_hour}{r}<TargetChairRate*0.7,'
                '"Check - well under your chair rate",'
                'IF(IF({reviewed}{r}="",FALSE,'
                'TodayDate-{reviewed}{r}>PriceReviewDays),'
                '"Due - price has not been reviewed",'
                'IF(IF({status}{r}="Active",{done}{r}=0,FALSE),'
                '"Check - on the menu but never booked","OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Services & Pricing",
        subtitle="The menu, the chair time, the product cost and what "
                 "each treatment actually earns an hour.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).services,
        help_text="Price, minutes and product cost are the three things "
                  "to get right. Everything to the right of them is "
                  "worked out for you.",
        extra_notes=[
            "Minutes matter as much as the price. Two hours of colour at "
            "$150 earns less an hour than a $45 cut in forty minutes, "
            "and the price an hour column is where that shows up.",
            "Leave the product cost blank and the workbook estimates it "
            "from the percentage on Settings - good enough to start, "
            "worth replacing with real numbers for your colour lines.",
            "The suggested price is built from your target chair rate, "
            "not from a markup. Raise the target on Settings and every "
            "service re-prices itself.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Menu summary under the grid."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1        # 0 based row under the data
    last = spec_.last_col_index
    row = sheet.section(row, "The menu at a glance", last)
    sheet.table_head(row, 0, [
        "Services listed", "Active", "Average price", "Average minutes",
        "Average price an hour", "Best rate an hour", "Never booked",
        "Priced under target"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Services.name]],"<>")',
        '=COUNTIFS([[Services.status]],"Active")',
        '=IFERROR(ROUND(AVERAGEIF([[Services.price]],">0"),2),0)',
        '=IFERROR(ROUND(AVERAGEIF([[Services.duration]],">0"),0),0)',
        '=IFERROR(ROUND(SUMPRODUCT(IFERROR([[Services.price_hour]]*1,0))/'
        'MAX(1,COUNTIF([[Services.name]],"<>")),2),0)',
        '=ROUND(MAX(IFERROR([[Services.price_hour]]*1,0)),2)',
        '=COUNTIFS([[Services.name]],"<>",[[Services.done]],0)',
        '=SUMPRODUCT((IFERROR([[Services.price_hour]]*1,0)>0)*'
        '(IFERROR([[Services.price_hour]]*1,0)<TargetChairRate))',
    ], ["int", "int", "money", "int", "money", "money", "int", "int"])
