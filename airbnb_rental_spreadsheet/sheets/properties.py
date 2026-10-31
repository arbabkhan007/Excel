"""`Property Database` - one row per place you let, and what it earns.

This is the spine of the workbook.  Every other tab points at a property
by its ID, so a portfolio of one studio and a portfolio of twelve
apartments behave exactly the same way: add a row here, and the property
appears in every drop-down, every calendar column, every per-property
total and every comparison table in the file.

The commercial side of a property lives here too - the standard nightly
rate, the cleaning fee you charge, the cleaning cost you pay, the
mortgage or the rent if you are letting someone else's building - so a
blank booking row can price itself from the property and a host can see
which of their places actually makes money.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check, time_check
from .common import (alert_cf, bar_cf, blank, flag_yes_no_cf, id_col, 
                    loss_cf, rating_cf, status_cf, target_cf)

NAME = "Property Database"


def spec(cfg: Config) -> SheetSpec:
    bk_prop = "[[Booking Tracker.property]]"
    bk_earns = "[[Booking Tracker.earns]]"
    bk_holds = "[[Booking Tracker.holds]]"
    bk_nights = "[[Booking Tracker.nights]]"
    bk_accom = "[[Booking Tracker.accommodation]]"
    bk_gross = "[[Booking Tracker.gross]]"
    bk_fees = "[[Booking Tracker.platform_fee_total]]"
    bk_net = "[[Booking Tracker.net_revenue]]"
    bk_guests = "[[Booking Tracker.guests_total]]"
    bk_block = "[[Booking Tracker.is_block]]"
    bk_cancel = "[[Booking Tracker.is_cancelled]]"
    ex_prop = "[[Expense Tracker.property]]"
    ex_pl = "[[Expense Tracker.pl_amount]]"
    cl_prop = "[[Cleaning Schedule.property]]"
    cl_cost = "[[Cleaning Schedule.total_cost]]"
    mt_prop = "[[Maintenance Log.property]]"
    mt_cost = "[[Maintenance Log.cost_used]]"
    mt_open = "[[Maintenance Log.is_open]]"
    rv_prop = "[[Reviews Tracker.property]]"
    rv_score = "[[Reviews Tracker.rate_overall]]"
    rv_has = "[[Reviews Tracker.has_score]]"
    sp_prop = "[[Supplies Inventory.property]]"
    sp_low = "[[Supplies Inventory.needs_order]]"

    def by_property(value: str, prop_range: str, *extra: str) -> str:
        """SUMIFS of a value column for this property."""
        tail = "".join(f",{c}" for c in extra)
        return f'SUMIFS({value},{prop_range},{{id}}{{r}}{tail})'

    columns = [
        # ------------------------------------------------------ the place --
        id_col("name", "PropertyPrefix", digits=cfg.digits("property"),
               width=10, header="Property ID",
               note="Automatic. This is the code every other tab uses to "
                    "point at this property, so keep the rows in place "
                    "and filter rather than re-sort."),
        Col("name", "Property name", 24, "text",
            note="What you call it day to day - 'Harbour Loft', "
                 "'Pine Cabin'. Short names read better in the calendar "
                 "and the reports.",
            group="The place"),
        Col("nickname", "Short code", 9, "text",
            note="Two or three letters for the calendar strip and the "
                 "tight comparison tables. 'HL', 'PC'.",
            group="The place"),
        Col("address", "Address", 32, "text",
            note="Street address. Used on the printable turnover sheet "
                 "so a cleaner knows where they are going.",
            group="The place"),
        Col("city", "City / area", 16, "text", group="The place"),
        Col("country", "Country", 14, "text",
            validate=list_from("CountryList"), group="The place"),
        Col("type", "Property type", 16, "text",
            validate=list_from("PropertyTypeList"), group="The place"),
        Col("bedrooms", "Bedrooms", 9, "int",
            validate=number_check(0, 30, "Bedrooms"), group="The place"),
        Col("bathrooms", "Bathrooms", 9, "number1",
            validate=number_check(0, 30, "Bathrooms"),
            note="Half a bathroom is fine - type 1.5.",
            group="The place"),
        Col("beds", "Beds", 7, "int",
            validate=number_check(0, 60, "Beds"), group="The place"),
        Col("max_guests", "Sleeps", 8, "int",
            validate=number_check(1, 60, "Maximum guests"),
            note="Maximum guests allowed. The booking tab warns you when "
                 "a reservation goes over it.",
            group="The place"),
        Col("status", "Status", 15, "status",
            validate=list_from("PropertyStatusList"),
            cf=status_cf(),
            note="Only Active properties are counted in the portfolio "
                 "occupancy and the dashboard targets.",
            group="The place"),
        Col("listed_from", "Listed since", 12, "date",
            validate=date_check(),
            note="The date it first went live. Used to work out how many "
                 "nights it has genuinely been available this year.",
            group="The place"),

        # --------------------------------------------------- who runs it --
        Col("owner", "Owner", 18, "text", group="Who runs it"),
        Col("manager", "Managed by", 18, "text",
            note="You, a co-host, or an agency. Free text.",
            group="Who runs it"),
        Col("ownership", "Ownership", 20, "text",
            validate=list_from("OwnershipList"),
            note="Rental arbitrage means you pay rent on it and sublet "
                 "it nightly - the monthly rent below is then your "
                 "biggest cost.",
            group="Who runs it"),
        Col("main_platform", "Main platform", 15, "text",
            validate=list_from("PlatformList"), group="Who runs it"),
        Col("listing_url", "Listing link", 26, "link",
            note="Paste the public listing URL. Handy when you are "
                 "checking your own calendar against the platform.",
            group="Who runs it"),
        Col("licence", "Licence / permit", 16, "text",
            note="Short-term let registration number, if your city "
                 "requires one.",
            group="Who runs it"),

        # ------------------------------------------------- house settings --
        Col("check_in_time", "Check-in", 9, "time",
            validate=time_check(),
            note="The standard arrival time. Copied onto new "
                 "arrivals so you only type an exception.",
            group="House settings"),
        Col("check_out_time", "Check-out", 9, "time",
            validate=time_check(), group="House settings"),
        Col("min_nights", "Min nights", 10, "int",
            validate=number_check(1, 365, "Minimum nights"),
            group="House settings"),
        Col("wifi_name", "Wi-Fi network", 16, "text",
            group="House settings"),
        Col("access", "Access method", 18, "text",
            note="Key box, smart lock, meet and greet. Printed on the "
                 "turnover sheet.",
            group="House settings"),

        # ----------------------------------------------------- the money --
        Col("nightly_rate", "Standard rate", 13, "money",
            validate=number_check(0, 100000, "Nightly rate"),
            note="Your usual nightly price. A booking with no rate typed "
                 "falls back to this, so seasonal pricing is an override "
                 "and not a chore.",
            group="Money"),
        Col("cleaning_fee", "Cleaning fee charged", 15, "money",
            validate=number_check(0, 100000, "Cleaning fee"),
            note="What the guest pays. Kept out of ADR and RevPAR, "
                 "because it is a cost recovery and not room revenue.",
            group="Money"),
        Col("cleaning_cost", "Cleaning cost paid", 15, "money",
            validate=number_check(0, 100000, "Cleaning cost"),
            note="What the turnover actually costs you. The default for "
                 "new cleaning jobs on this property.",
            group="Money"),
        Col("deposit", "Security deposit", 14, "money",
            validate=number_check(0, 1000000, "Deposit"), group="Money"),
        Col("monthly_rent", "Monthly rent", 13, "money",
            validate=number_check(0, 1000000, "Monthly rent"),
            note="Rental arbitrage only - what you pay the landlord. "
                 "Leave blank if you own it.",
            group="Money"),
        Col("monthly_mortgage", "Monthly mortgage", 14, "money",
            validate=number_check(0, 1000000, "Mortgage"),
            note="Optional. The full payment; the interest part belongs "
                 "in the expense tab if you track it for tax.",
            group="Money"),
        Col("purchase_price", "Purchase price", 14, "money0",
            validate=number_check(0, 100000000, "Purchase price"),
            note="Optional. Only used for the yield figure.",
            group="Money"),
        Col("fixed_monthly", "Fixed cost a month", 14, "money",
            formula=blank("name", "ROUND({monthly_rent}{r}+"
                                  "{monthly_mortgage}{r},2)"),
            note="Rent plus mortgage. This is what the property costs "
                 "you before a single guest arrives.",
            group="Money"),

        # -------------------------------------------------- what it does --
        Col("bookings", "Bookings", 10, "int",
            formula=blank("name", f'COUNTIFS({bk_prop},{{id}}{{r}},'
                                  f'{bk_earns},"Yes")'),
            note="Confirmed, checked in and completed. Inquiries, "
                 "cancellations and blocks are not counted.",
            group="Performance"),
        Col("nights_booked", "Nights booked", 12, "int",
            formula=blank("name", by_property(bk_nights, bk_prop,
                                              f'{bk_earns},"Yes"')),
            cf=bar_cf(), group="Performance"),
        Col("nights_blocked", "Nights blocked", 12, "int",
            formula=blank("name", by_property(bk_nights, bk_prop,
                                              f'{bk_block},"Yes"')),
            note="Owner stays and maintenance closures. These come off "
                 "the available nights, so blocking a week does not "
                 "damage your occupancy rate.",
            group="Performance"),
        Col("guests_hosted", "Guests hosted", 12, "int",
            formula=blank("name", by_property(bk_guests, bk_prop,
                                              f'{bk_earns},"Yes"')),
            group="Performance"),
        Col("cancellations", "Cancellations", 12, "int",
            formula=blank("name", f'COUNTIFS({bk_prop},{{id}}{{r}},'
                                  f'{bk_cancel},"Yes")'),
            group="Performance"),
        Col("nights_available", "Nights available", 13, "int",
            formula=blank(
                "name",
                'MAX(0,MIN(TodayDate,DATE(YEAR(TodayDate),12,31))-'
                'MAX(DATE(YEAR(TodayDate),1,1),'
                'IF({listed_from}{r}="",DATE(YEAR(TodayDate),1,1),'
                '{listed_from}{r}))+1-{nights_blocked}{r})'),
            note="This calendar year so far, from the day it was listed, "
                 "less the nights you blocked. This is the denominator "
                 "of the occupancy figure.",
            group="Performance"),
        Col("occupancy", "Occupancy", 10, "percent",
            formula=blank(
                "name",
                'IF({nights_available}{r}<=0,"",'
                '{nights_booked}{r}/{nights_available}{r})'),
            cf=target_cf("TargetOccupancy"),
            note="Booked nights divided by the nights the place was "
                 "actually on the market this year.",
            group="Performance"),
        Col("adr", "ADR", 11, "money",
            formula=blank(
                "name",
                'IF({nights_booked}{r}<=0,"",'
                'ROUND({accommodation}{r}/{nights_booked}{r},2))'),
            cf=target_cf("TargetADR"),
            note="Average daily rate: accommodation revenue per booked "
                 "night. Cleaning fees are deliberately excluded.",
            group="Performance"),
        Col("revpar", "RevPAR", 11, "money",
            formula=blank(
                "name",
                'IF({nights_available}{r}<=0,"",'
                'ROUND({accommodation}{r}/{nights_available}{r},2))'),
            cf=target_cf("TargetRevPAR"),
            note="Revenue per available night - occupancy and rate in "
                 "one number. The figure to compare two properties on.",
            group="Performance"),
        Col("alos", "Avg stay", 9, "number1",
            formula=blank(
                "name",
                'IF({bookings}{r}<=0,"",'
                'ROUND({nights_booked}{r}/{bookings}{r},1))'),
            note="Average length of stay, in nights.",
            group="Performance"),
        Col("rating", "Rating", 9, "number",
            formula=blank(
                "name",
                f'IF(COUNTIFS({rv_prop},{{id}}{{r}},{rv_has},"Yes")<=0,"",'
                f'ROUND(SUMIFS({rv_score},{rv_prop},{{id}}{{r}},'
                f'{rv_has},"Yes")/COUNTIFS({rv_prop},{{id}}{{r}},'
                f'{rv_has},"Yes"),2))'),
            cf=rating_cf(), group="Performance"),
        Col("reviews", "Reviews", 9, "int",
            formula=blank("name", f'COUNTIFS({rv_prop},{{id}}{{r}},'
                                  f'{rv_has},"Yes")'),
            hidden=True, group="Performance"),

        # ---------------------------------------------------- the result --
        Col("accommodation", "Room revenue", 13, "money",
            formula=blank("name", by_property(bk_accom, bk_prop,
                                              f'{bk_earns},"Yes"')),
            note="Nightly accommodation only - no cleaning fee, no extra "
                 "guest charges. ADR and RevPAR are built on this.",
            group="Result"),
        Col("gross", "Gross revenue", 13, "money",
            formula=blank("name", by_property(bk_gross, bk_prop,
                                              f'{bk_earns},"Yes"')),
            note="Everything the guest paid, before platform fees.",
            group="Result"),
        Col("platform_fees", "Platform fees", 12, "money",
            formula=blank("name", by_property(bk_fees, bk_prop,
                                              f'{bk_earns},"Yes"')),
            group="Result"),
        Col("net_revenue", "Net revenue", 13, "money",
            formula=blank("name", by_property(bk_net, bk_prop,
                                              f'{bk_earns},"Yes"')),
            note="What lands in your account after the platform takes "
                 "its cut.",
            group="Result"),
        Col("cleaning_spend", "Cleaning paid out", 13, "money",
            formula=blank("name", by_property(cl_cost, cl_prop)),
            group="Result"),
        Col("maintenance_spend", "Maintenance", 12, "money",
            formula=blank("name", by_property(mt_cost, mt_prop)),
            group="Result"),
        Col("other_expenses", "Other costs", 12, "money",
            formula=blank("name", by_property(ex_pl, ex_prop)),
            note="Everything on the expense tab for this property that "
                 "counts against profit. Cleaning typed in as an expense "
                 "is excluded automatically so it is never counted twice.",
            group="Result"),
        Col("total_costs", "Total costs", 13, "money",
            formula=blank("name",
                          "ROUND({cleaning_spend}{r}+"
                          "{maintenance_spend}{r}+{other_expenses}{r},2)"),
            group="Result"),
        Col("profit", "Operating profit", 14, "money",
            formula=blank("name",
                          "ROUND({net_revenue}{r}-{total_costs}{r},2)"),
            cf=loss_cf(),
            note="Net revenue less everything it cost to run. Rent and "
                 "mortgage are shown separately so an arbitrage host and "
                 "an owner can read the same column.",
            group="Result"),
        Col("margin", "Margin", 9, "percent",
            formula=blank("name",
                          'IF({net_revenue}{r}<=0,"",'
                          '{profit}{r}/{net_revenue}{r})'),
            cf=target_cf("TargetMargin"), group="Result"),
        Col("profit_night", "Profit a night", 12, "money",
            formula=blank("name",
                          'IF({nights_booked}{r}<=0,"",'
                          'ROUND({profit}{r}/{nights_booked}{r},2))'),
            group="Result"),
        Col("yield_pct", "Gross yield", 10, "percent",
            formula=blank("name",
                          'IF({purchase_price}{r}<=0,"",'
                          '{gross}{r}/{purchase_price}{r})'),
            note="Gross revenue against what you paid for the place. "
                 "Blank unless you filled the purchase price in.",
            hidden=True, group="Result"),
        Col("profit_rank", "Rank", 7, "int",
            formula=('=IF({name}{r}="","",COUNTIFS({profit_first}:'
                     '{profit_last},">"&{profit}{r})+'
                     'COUNTIFS({profit_first}:{profit}{r},{profit}{r}))'),
            note="1 = your best earning property.",
            group="Result"),

        # ------------------------------------------------- what to do now --
        Col("open_jobs", "Open jobs", 10, "int",
            formula=blank("name", f'COUNTIFS({mt_prop},{{id}}{{r}},'
                                  f'{mt_open},"Yes")'),
            note="Maintenance reported and not finished.",
            group="Right now"),
        Col("low_stock", "Low stock", 10, "int",
            formula=blank("name", f'COUNTIFS({sp_prop},{{id}}{{r}},'
                                  f'{sp_low},"Yes")'),
            group="Right now"),
        Col("guest_in", "Guest in tonight", 13, "flag",
            formula=blank(
                "name",
                f'IF(SUMPRODUCT(({bk_prop}={{id}}{{r}})*'
                f'({bk_holds}="Yes")*'
                f'([[Booking Tracker.ci_num]]<=TodayDate)*'
                f'([[Booking Tracker.co_num]]>TodayDate))>0,"Yes","No")'),
            cf=flag_yes_no_cf(),
            note="Is someone staying tonight? Reads the booking tab, so "
                 "it is right the moment you type a reservation.",
            group="Right now"),
        Col("next_arrival", "Next arrival", 12, "date",
            formula=blank(
                "name",
                f'IF(SUMPRODUCT(MIN(([[Booking Tracker.ci_num]]>TodayDate)*'
                f'({bk_prop}={{id}}{{r}})*({bk_holds}="Yes")*'
                f'[[Booking Tracker.ci_num]]+'
                f'(1-([[Booking Tracker.ci_num]]>TodayDate)*'
                f'({bk_prop}={{id}}{{r}})*({bk_holds}="Yes"))*2958465))'
                f'>=2958465,"",'
                f'SUMPRODUCT(MIN(([[Booking Tracker.ci_num]]>TodayDate)*'
                f'({bk_prop}={{id}}{{r}})*({bk_holds}="Yes")*'
                f'[[Booking Tracker.ci_num]]+'
                f'(1-([[Booking Tracker.ci_num]]>TodayDate)*'
                f'({bk_prop}={{id}}{{r}})*({bk_holds}="Yes"))*2958465)))'),
            note="The next check-in date after today.",
            group="Right now"),
        Col("alert", "Action", 26, "status",
            formula=blank(
                "name",
                'IF({status}{r}="Under maintenance","Check - '
                'marked under maintenance",'
                'IF({open_jobs}{r}>0,"Action - "&{open_jobs}{r}&'
                '" open job(s)",'
                'IF({low_stock}{r}>0,"Order - "&{low_stock}{r}&'
                '" supply line(s) low",'
                'IF({nightly_rate}{r}<=0,"Add - no standard nightly rate",'
                'IF({status}{r}<>"Active","",'
                'IF({nights_available}{r}<=0,"",'
                'IF({occupancy}{r}<TargetOccupancy*0.7,'
                '"Check - occupancy well under target",'
                'IF({margin}{r}<0,"Loss - costs are above net revenue",'
                '"OK"))))))))'),
            cf=alert_cf(), group="Right now"),
        Col("notes", "Notes", 34, "wrap",
            note="Anything that does not fit a column - parking, quirks, "
                 "the neighbour to warn before a party.",
            group="Right now"),
    ]

    return SheetSpec(
        name=NAME,
        title="Property Database",
        subtitle="Every place you let, what it costs to run and what it "
                 "actually earns",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=None,
        freeze_cols=2,
        samples=sample("properties", cfg),
        help_text="Add one row per property. The Property ID on the left "
                  "is what every other tab points at, so add a property "
                  "here first and it appears in every drop-down in the "
                  "workbook. White cells are yours to type in; the "
                  "shaded ones work themselves out.",
        extra_notes=[
            "Nights available counts this calendar year from the day you "
            "listed the property, less any nights you blocked - so "
            "blocking a fortnight for yourself does not dent the "
            "occupancy rate.",
            "Cleaning fees charged to guests are deliberately kept out "
            "of ADR and RevPAR. They are a cost recovery, not room "
            "revenue, and mixing them in flatters both figures.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Portfolio totals under the register."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index

    row = sheet.section(row, "The whole portfolio", last)
    row = sheet.table_head(row, 0, [
        "Properties", "Active", "Bedrooms", "Sleeps", "Bookings",
        "Nights booked", "Nights blocked", "Occupancy", "ADR", "RevPAR"],
        [11, 9, 10, 9, 10, 13, 13, 11, 11, 11])
    sheet.calc_row(row, 0, [
        '=COUNTIF([[Property Database.name]],"?*")',
        '=COUNTIFS([[Property Database.status]],"Active")',
        '=SUM([[Property Database.bedrooms]])',
        '=SUM([[Property Database.max_guests]])',
        '=SUM([[Property Database.bookings]])',
        '=SUM([[Property Database.nights_booked]])',
        '=SUM([[Property Database.nights_blocked]])',
        '=IF(SUM([[Property Database.nights_available]])<=0,"",'
        'SUM([[Property Database.nights_booked]])/'
        'SUM([[Property Database.nights_available]]))',
        '=IF(SUM([[Property Database.nights_booked]])<=0,"",'
        'ROUND(SUM([[Property Database.accommodation]])/'
        'SUM([[Property Database.nights_booked]]),2))',
        '=IF(SUM([[Property Database.nights_available]])<=0,"",'
        'ROUND(SUM([[Property Database.accommodation]])/'
        'SUM([[Property Database.nights_available]]),2))',
    ], ["int", "int", "int", "int", "int", "int", "int", "percent",
        "money", "money"])
    row += 2

    row = sheet.table_head(row, 0, [
        "Gross revenue", "Platform fees", "Net revenue", "Cleaning",
        "Maintenance", "Other costs", "Operating profit", "Margin",
        "Rent & mortgage", "After finance"],
        [13, 12, 13, 12, 12, 12, 14, 10, 14, 13])
    sheet.calc_row(row, 0, [
        '=ROUND(SUM([[Property Database.gross]]),2)',
        '=ROUND(SUM([[Property Database.platform_fees]]),2)',
        '=ROUND(SUM([[Property Database.net_revenue]]),2)',
        '=ROUND(SUM([[Property Database.cleaning_spend]]),2)',
        '=ROUND(SUM([[Property Database.maintenance_spend]]),2)',
        '=ROUND(SUM([[Property Database.other_expenses]]),2)',
        '=ROUND(SUM([[Property Database.profit]]),2)',
        '=IF(SUM([[Property Database.net_revenue]])<=0,"",'
        'SUM([[Property Database.profit]])/'
        'SUM([[Property Database.net_revenue]]))',
        '=ROUND(SUM([[Property Database.fixed_monthly]])*12,2)',
        '=ROUND(SUM([[Property Database.profit]])-'
        'SUM([[Property Database.fixed_monthly]])*12,2)',
    ], ["money", "money", "money", "money", "money", "money", "money",
        "percent", "money", "money"])
    row += 2

    sheet.note(row, "Rent and mortgage are shown as a yearly figure "
                    "(monthly x 12) and kept out of operating profit, so "
                    "the profit column compares a property you own with "
                    "one you rent on equal terms. 'After finance' is the "
                    "number an arbitrage host cares about.", last)
