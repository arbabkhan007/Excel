"""`Occupancy & Performance` - one row per property per month.

This is the tab a short-term rental investor opens first.  Type a
property and the first of a month; everything else - available nights,
booked nights, occupancy, ADR, RevPAR, average stay, the cleaning cost
per booking, the operating cost per night and the profit - is worked
out from the bookings and the costs you have already entered.

Three definitions are worth stating plainly, because half the
spreadsheets sold online get them wrong:

* **Occupancy** is booked nights over *available* nights, and a night
  you blocked for yourself was never available.
* **ADR** is accommodation revenue over booked nights.  Cleaning fees
  are not accommodation revenue.
* **RevPAR** is accommodation revenue over available nights, which is
  the only one of the three that cannot be flattered by taking fewer
  bookings at a higher price.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from
from .common import (alert_cf, bar_cf, blank, id_col, loss_cf, 
                    lookup_property, money_cf, target_cf)

NAME = "Occupancy & Performance"


def spec(cfg: Config) -> SheetSpec:
    bk_prop = "[[Booking Tracker.property]]"
    bk_month = "[[Booking Tracker.month]]"
    bk_earns = "[[Booking Tracker.earns]]"
    bk_block = "[[Booking Tracker.is_block]]"
    bk_cancel = "[[Booking Tracker.is_cancelled]]"
    bk_nights = "[[Booking Tracker.nights]]"
    bk_accom = "[[Booking Tracker.accommodation]]"
    bk_gross = "[[Booking Tracker.gross]]"
    bk_net = "[[Booking Tracker.net_revenue]]"
    bk_guests = "[[Booking Tracker.guests_total]]"
    bk_fees = "[[Booking Tracker.platform_fee_total]]"
    cl_prop = "[[Cleaning Schedule.property]]"
    cl_month = "[[Cleaning Schedule.month]]"
    cl_cost = "[[Cleaning Schedule.total_cost]]"
    ex_prop = "[[Expense Tracker.property]]"
    ex_month = "[[Expense Tracker.month]]"
    ex_pl = "[[Expense Tracker.pl_amount]]"
    mt_prop = "[[Maintenance Log.property]]"
    mt_month = "[[Maintenance Log.month]]"
    mt_cost = "[[Maintenance Log.cost_used]]"
    rv_prop = "[[Reviews Tracker.property]]"
    rv_month = "[[Reviews Tracker.month]]"
    rv_score = "[[Reviews Tracker.rate_overall]]"
    rv_has = "[[Reviews Tracker.has_score]]"

    def by_month(value: str, prop_range: str, month_range: str,
                 *extra: str) -> str:
        tail = "".join(f",{c}" for c in extra)
        return (f'SUMIFS({value},{prop_range},{{property}}{{r}},'
                f'{month_range},{{month}}{{r}}{tail})')

    def count_month(prop_range: str, month_range: str, *extra: str) -> str:
        tail = "".join(f",{c}" for c in extra)
        return (f'COUNTIFS({prop_range},{{property}}{{r}},'
                f'{month_range},{{month}}{{r}}{tail})')

    # a COUNTIFS/SUMIFS criterion is TWO arguments - the range and the
    # thing it is compared against.  Collapsing them into one argument
    # as `range="Yes"` makes Excel answer #VALUE! on every row.
    _scored = (rv_has, '"Yes"')
    _rv_count = count_month(rv_prop, rv_month, *_scored)
    _rv_total = by_month(rv_score, rv_prop, rv_month, *_scored)

    columns = [
        id_col("period", "PeriodPrefix", digits=cfg.digits("period"),
               width=10, header="Row ID"),
        Col("property", "Property", 11, "text",
            validate=list_from("PropertyList"),
            note="Which property this row measures.",
            group="The period"),
        Col("property_name", "Property name", 19, "text",
            formula=lookup_property("name"), group="The period"),
        Col("period", "Month", 12, "date",
            validate=date_check(),
            note="The first of the month, e.g. 01 Mar 2026. Every row "
                 "on this tab starts here. One row per property per "
                 "month.",
            group="The period"),
        Col("month", "Month key", 10, "text",
            formula=blank("period", 'TEXT({period}{r},"yyyy-mm")'),
            hidden=True, group="The period"),
        Col("month_label", "Period", 13, "text",
            formula=blank("period", 'TEXT({period}{r},"mmm yyyy")'),
            group="The period"),
        Col("days_in_month", "Days", 7, "int",
            formula=blank(
                "period",
                'DAY(DATE(YEAR({period}{r}),MONTH({period}{r})+1,1)-1)'),
            hidden=True, group="The period"),

        # ------------------------------------------------- the nights --
        Col("nights_blocked", "Blocked", 10, "int",
            formula=blank(
                "period",
                by_month(bk_nights, bk_prop, bk_month,
                         f'{bk_block},"Yes"')),
            note="Owner stays and maintenance closures that started in "
                 "this month.",
            group="Nights"),
        Col("nights_available", "Available", 11, "int",
            formula=blank(
                "period",
                'MAX(0,{days_in_month}{r}-{nights_blocked}{r})'),
            note="Days in the month less the nights you took off the "
                 "market yourself.",
            group="Nights"),
        Col("nights_booked", "Booked", 10, "int",
            formula=blank(
                "period",
                by_month(bk_nights, bk_prop, bk_month,
                         f'{bk_earns},"Yes"')),
            cf=bar_cf(), group="Nights"),
        Col("occupancy", "Occupancy", 11, "percent",
            formula=blank(
                "period",
                'IF({nights_available}{r}<=0,"",'
                '{nights_booked}{r}/{nights_available}{r})'),
            cf=target_cf("TargetOccupancy"), group="Nights"),
        Col("bookings", "Bookings", 10, "int",
            formula=blank("period",
                          count_month(bk_prop, bk_month,
                                      f'{bk_earns},"Yes"')),
            group="Nights"),
        Col("cancellations", "Cancelled", 10, "int",
            formula=blank("period",
                          count_month(bk_prop, bk_month,
                                      f'{bk_cancel},"Yes"')),
            group="Nights"),
        Col("cancel_rate", "Cancel rate", 11, "percent",
            formula=blank(
                "period",
                'IF({bookings}{r}+{cancellations}{r}<=0,"",'
                '{cancellations}{r}/({bookings}{r}+{cancellations}{r}))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0.1,
                 "fmt": "warn"}],
            hidden=True, group="Nights"),
        Col("guests", "Guests", 9, "int",
            formula=blank("period",
                          by_month(bk_guests, bk_prop, bk_month,
                                   f'{bk_earns},"Yes"')),
            hidden=True, group="Nights"),
        Col("alos", "Avg stay", 9, "number1",
            formula=blank(
                "period",
                'IF({bookings}{r}<=0,"",'
                'ROUND({nights_booked}{r}/{bookings}{r},1))'),
            group="Nights"),

        # ------------------------------------------------ the headline --
        Col("accommodation", "Room revenue", 13, "money",
            formula=blank("period",
                          'ROUND(' + by_month(bk_accom, bk_prop, bk_month,
                                              f'{bk_earns},"Yes"') + ',2)'),
            group="Rates"),
        Col("adr", "ADR", 11, "money",
            formula=blank(
                "period",
                'IF({nights_booked}{r}<=0,"",'
                'ROUND({accommodation}{r}/{nights_booked}{r},2))'),
            cf=target_cf("TargetADR"),
            note="Average daily rate. Accommodation only - the cleaning "
                 "fee is a cost recovery and stays out of it.",
            group="Rates"),
        Col("revpar", "RevPAR", 11, "money",
            formula=blank(
                "period",
                'IF({nights_available}{r}<=0,"",'
                'ROUND({accommodation}{r}/{nights_available}{r},2))'),
            cf=target_cf("TargetRevPAR"),
            note="Revenue per available night. The honest one: you "
                 "cannot improve it by taking fewer bookings at a "
                 "higher price.",
            group="Rates"),

        # ---------------------------------------------------- the money --
        Col("gross", "Gross", 12, "money",
            formula=blank("period",
                          'ROUND(' + by_month(bk_gross, bk_prop, bk_month,
                                              f'{bk_earns},"Yes"') + ',2)'),
            group="Money"),
        Col("platform_fees", "Platform fees", 12, "money",
            formula=blank("period",
                          'ROUND(' + by_month(bk_fees, bk_prop, bk_month,
                                              f'{bk_earns},"Yes"') + ',2)'),
            group="Money"),
        Col("net_revenue", "Net revenue", 12, "money",
            formula=blank("period",
                          'ROUND(' + by_month(bk_net, bk_prop, bk_month,
                                              f'{bk_earns},"Yes"') + ',2)'),
            cf=money_cf(), group="Money"),
        Col("cleaning_cost", "Cleaning", 11, "money",
            formula=blank("period",
                          'ROUND(' + by_month(cl_cost, cl_prop,
                                              cl_month) + ',2)'),
            group="Money"),
        Col("maintenance_cost", "Maintenance", 12, "money",
            formula=blank("period",
                          'ROUND(' + by_month(mt_cost, mt_prop,
                                              mt_month) + ',2)'),
            group="Money"),
        Col("other_cost", "Other costs", 12, "money",
            formula=blank("period",
                          'ROUND(' + by_month(ex_pl, ex_prop,
                                              ex_month) + ',2)'),
            group="Money"),
        Col("total_cost", "Total costs", 12, "money",
            formula=blank("period",
                          'ROUND({cleaning_cost}{r}+'
                          '{maintenance_cost}{r}+{other_cost}{r},2)'),
            group="Money"),
        Col("profit", "Operating profit", 14, "money",
            formula=blank("period",
                          'ROUND({net_revenue}{r}-{total_cost}{r},2)'),
            cf=loss_cf(), group="Money"),
        Col("margin", "Margin", 9, "percent",
            formula=blank("period",
                          'IF({net_revenue}{r}<=0,"",'
                          '{profit}{r}/{net_revenue}{r})'),
            cf=target_cf("TargetMargin"), group="Money"),

        # -------------------------------------------------- unit costs --
        Col("cost_per_night", "Cost a night", 12, "money",
            formula=blank(
                "period",
                'IF({nights_booked}{r}<=0,"",'
                'ROUND({total_cost}{r}/{nights_booked}{r},2))'),
            note="What a single booked night costs you to deliver. "
                 "Compare it with ADR - the gap is your real margin.",
            group="Unit costs"),
        Col("clean_per_booking", "Cleaning a booking", 14, "money",
            formula=blank(
                "period",
                'IF({bookings}{r}<=0,"",'
                'ROUND({cleaning_cost}{r}/{bookings}{r},2))'),
            note="The reason a run of one-night stays can lose money at "
                 "a perfectly good nightly rate.",
            group="Unit costs"),
        Col("profit_night", "Profit a night", 12, "money",
            formula=blank(
                "period",
                'IF({nights_booked}{r}<=0,"",'
                'ROUND({profit}{r}/{nights_booked}{r},2))'),
            cf=loss_cf(), group="Unit costs"),
        Col("rating", "Rating", 9, "number",
            formula=blank(
                "period",
                f'IF({_rv_count}<=0,"",'
                f'ROUND({_rv_total}/{_rv_count},2))'),
            note="Average review score for stays at this property that "
                 "were reviewed in this month.",
            hidden=True, group="Unit costs"),

        Col("alert", "Action", 28, "status",
            formula=blank(
                "period",
                'IF({property}{r}="","Add - no property on this row",'
                'IF({nights_available}{r}<=0,"Check - no available '
                'nights in this month",'
                'IF({nights_booked}{r}<=0,"Check - nothing booked at '
                'all",'
                'IF({profit}{r}<0,"Loss - costs beat net revenue",'
                'IF({occupancy}{r}<TargetOccupancy*0.7,"Action - '
                'occupancy well under target",'
                'IF({adr}{r}<TargetADR*0.8,"Check - rate well under '
                'target",'
                'IF({occupancy}{r}>0.95,"Check - almost full, your '
                'rate is probably too low","OK")))))))'),
            cf=alert_cf(), group="Unit costs"),
        Col("notes", "Notes", 28, "wrap",
            note="What happened that month - a heatwave, a festival, a "
                 "boiler, a price experiment. In twelve months this is "
                 "the column that explains the numbers.",
            group="Unit costs"),
    ]

    return SheetSpec(
        name=NAME,
        title="Occupancy & Performance",
        subtitle="One row per property per month: occupancy, ADR, "
                 "RevPAR, costs and profit",
        columns=columns,
        rows=cfg.row_count(NAME),
        freeze_cols=3,
        samples=sample("performance", cfg),
        help_text="Type a property and the first of a month. Everything "
                  "else is read from the bookings and costs you have "
                  "already entered, so a row takes two cells and gives "
                  "you the full picture of that property in that "
                  "month. Keep one row per property per month and the "
                  "year builds itself.",
        extra_notes=[
            "A stay is counted in the month it checks in, so a "
            "Hogmanay booking lands in December. That keeps one "
            "booking in one month and makes the totals add up; it is "
            "the same convention the platforms use on their earnings "
            "summaries.",
            "An occupancy figure above 95% is not a triumph, it is a "
            "pricing error, and the Action column says so.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index

    row = sheet.section(row, "Every period on this tab", last)
    row = sheet.table_head(row, 0, [
        "Rows", "Available nights", "Booked nights", "Occupancy",
        "Room revenue", "ADR", "RevPAR", "Net revenue", "Costs", "Profit"],
        [8, 16, 14, 11, 13, 11, 11, 12, 11, 12])
    avail = 'SUM([[Occupancy & Performance.nights_available]])'
    booked = 'SUM([[Occupancy & Performance.nights_booked]])'
    accom = 'SUM([[Occupancy & Performance.accommodation]])'
    sheet.calc_row(row, 0, [
        '=COUNT([[Occupancy & Performance.period]])',
        f"={avail}",
        f"={booked}",
        f'=IF({avail}<=0,"",{booked}/{avail})',
        f"=ROUND({accom},2)",
        f'=IF({booked}<=0,"",ROUND({accom}/{booked},2))',
        f'=IF({avail}<=0,"",ROUND({accom}/{avail},2))',
        '=ROUND(SUM([[Occupancy & Performance.net_revenue]]),2)',
        '=ROUND(SUM([[Occupancy & Performance.total_cost]]),2)',
        '=ROUND(SUM([[Occupancy & Performance.profit]]),2)',
    ], ["int", "int", "int", "percent", "money", "money", "money",
        "money", "money", "money"])
    row += 2

    sheet.note(row, "These totals cover whatever rows you have typed, "
                    "so they are only a portfolio view if you keep one "
                    "row per property per month. The Monthly and Annual "
                    "Reports tabs read the bookings directly and do not "
                    "depend on this tab being complete.", last)
