"""`Booking Tracker` - every reservation, block and enquiry, priced.

One register holds the lot: confirmed stays, cancellations, enquiries
that never happened, owner stays and maintenance closures.  A block is
simply a reservation that earns nothing - it takes the dates off the
calendar and comes out of the available nights, which is why blocking a
fortnight for yourself does not wreck your occupancy figure.

Three things are worked out here that a host would otherwise do in their
head at eleven o'clock at night: the nights in a stay, what the platform
keeps, and whether this reservation overlaps one you already took.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, date_num_col, flag_yes_no_cf,
                     id_col, is_block, is_earning, is_holding, lodging_tax,
                     lookup_property,
                     money_cf, month_col, platform_rate, seq_num_col,
                     status_cf, weekday_col, year_col, zero)

NAME = "Booking Tracker"


def spec(cfg: Config) -> SheetSpec:
    cl_booking = "[[Cleaning Schedule.booking]]"
    ar_booking = "[[Check-in & Check-out.booking]]"
    rv_booking = "[[Reviews Tracker.booking]]"

    columns = [
        # ------------------------------------------------- the booking --
        id_col("check_in", "BookingPrefix", digits=cfg.digits("booking"),
               width=11, header="Booking ID",
               note="Automatic, from the row number. Cleaning jobs, "
                    "arrivals, payouts and reviews all point back at "
                    "this reference."),
        Col("property", "Property", 11, "text",
            validate=list_from("PropertyList"),
            note="Pick the property ID. Add the property on the Property "
                 "Database tab first and it appears in this list.",
            group="The booking"),
        Col("property_name", "Property name", 20, "text",
            formula=lookup_property("name"), group="The booking"),
        Col("guest", "Guest", 11, "text",
            validate=list_from("GuestList"),
            note="Pick the guest ID from the Guest Database. Leave it "
                 "blank for a block - the Action column will remind you "
                 "if a real stay has no guest record.",
            group="The booking"),
        Col("guest_name", "Guest name", 20, "text",
            formula=('=IF({check_in}{r}="","",IF({guest}{r}="",'
                     'IF({is_block}{r}="Yes","(blocked)",""),'
                     'IFERROR(INDEX([[Guest Database.name]],'
                     'MATCH({guest}{r},[[Guest Database.id]],0)),'
                     '"Not in Guest Database")))'),
            group="The booking"),
        Col("platform", "Platform", 15, "text",
            validate=list_from("PlatformList"),
            note="Sets the commission and card fee for this booking. "
                 "Every rate is editable on Settings - platform fees "
                 "differ by country, account type and plan.",
            group="The booking"),
        Col("reference", "Platform ref", 14, "text",
            note="The confirmation code the platform gave you, e.g. "
                 "HM4XY29ABC. Useful when you reconcile a payout.",
            group="The booking"),
        Col("status", "Status", 16, "status",
            validate=list_from("BookingStatusList"),
            cf=status_cf(),
            note="Owner block and Maintenance block hold the dates "
                 "without earning anything. Cancelled frees them up.",
            group="The booking"),

        # ---------------------------------------------------- the dates --
        Col("booked_date", "Booked on", 12, "date",
            validate=date_check(),
            note="When the reservation came in - gives you the lead time.",
            group="Dates"),
        Col("check_in", "Check-in", 12, "date",
            validate=date_check(),
            note="Arrival date. Every row on this tab starts here: no "
                 "check-in date, no booking.",
            group="Dates"),
        Col("check_out", "Check-out", 12, "date",
            validate=date_check(), group="Dates"),
        weekday_col("check_in", "ci_day", "In", 6, hidden=True),
        weekday_col("check_out", "co_day", "Out", 6, hidden=True),
        Col("nights", "Nights", 8, "int",
            formula=('=IF({check_in}{r}="","",IF({check_out}{r}="","",'
                     'MAX(0,{check_out}{r}-{check_in}{r})))'),
            cf=bar_cf(),
            note="Worked out for you. The arrival night counts, the "
                 "departure night does not - which is how every "
                 "platform counts it too.",
            group="Dates"),
        date_num_col("check_in", "ci_num", "Check-in (number)"),
        date_num_col("check_out", "co_num", "Check-out (number)"),
        seq_num_col("check_in"),
        month_col("check_in"),
        year_col("check_in"),
        Col("season", "Season", 9, "text",
            formula=blank(
                "check_in",
                'IF(MONTH({check_in}{r})<=2,"Winter",'
                'IF(MONTH({check_in}{r})<=5,"Spring",'
                'IF(MONTH({check_in}{r})<=8,"Summer",'
                'IF(MONTH({check_in}{r})<=11,"Autumn","Winter"))))'),
            hidden=True, group="Dates"),
        Col("lead_time", "Lead time", 10, "int",
            formula=('=IF({check_in}{r}="","",IF({booked_date}{r}="","",'
                     'MAX(0,{check_in}{r}-{booked_date}{r})))'),
            note="Days between the booking coming in and the guest "
                 "arriving. A string of one-day lead times means your "
                 "price is too low.",
            hidden=True, group="Dates"),

        # --------------------------------------------------- the guests --
        Col("adults", "Adults", 8, "int",
            validate=number_check(0, 40, "Adults"), group="Guests"),
        Col("children", "Children", 9, "int",
            validate=number_check(0, 40, "Children"), group="Guests"),
        Col("pets", "Pets", 6, "int",
            validate=number_check(0, 20, "Pets"), group="Guests"),
        Col("guests_total", "Guests", 8, "int",
            formula=blank("check_in", "{adults}{r}+{children}{r}"),
            group="Guests"),
        Col("over_capacity", "Over capacity", 12, "flag",
            formula=blank(
                "check_in",
                'IF({property}{r}="","",'
                'IF({guests_total}{r}>'
                'IFERROR(INDEX([[Property Database.max_guests]],'
                'MATCH({property}{r},[[Property Database.id]],0))*1,'
                '999),"Yes","No"))'),
            cf=[{"type": "text", "criteria": "containing", "value": "Yes",
                 "fmt": "bad"}],
            hidden=True, group="Guests"),

        # ----------------------------------------------------- the money --
        Col("rate_in", "Nightly rate", 12, "money",
            validate=number_check(0, 100000, "Nightly rate"),
            note="Leave blank to use the property's standard rate. Type "
                 "a number here to override it for this stay only.",
            group="Money"),
        Col("rate_used", "Rate used", 11, "money",
            formula=blank(
                "check_in",
                'IF({rate_in}{r}<>"",{rate_in}{r},'
                'IF({property}{r}="",0,'
                'IFERROR(INDEX([[Property Database.nightly_rate]],'
                'MATCH({property}{r},[[Property Database.id]],0))*1,0)))'),
            note="The rate this booking is actually priced at: what you "
                 "typed, or the property's standard rate when you left "
                 "it blank.",
            group="Money"),
        Col("accommodation", "Room revenue", 13, "money",
            formula=blank(
                "check_in",
                'IF({earns}{r}<>"Yes",0,'
                'ROUND({nights}{r}*{rate_used}{r},2))'),
            note="Nights times the rate. This is the only figure ADR and "
                 "RevPAR are built from - cleaning and extras stay out "
                 "of it on purpose.",
            group="Money"),
        Col("cleaning_in", "Cleaning fee", 12, "money",
            validate=number_check(0, 100000, "Cleaning fee"),
            note="Blank uses the property's standard cleaning fee.",
            group="Money"),
        Col("cleaning_fee", "Cleaning charged", 13, "money",
            formula=blank(
                "check_in",
                'IF({earns}{r}<>"Yes",0,IF({cleaning_in}{r}<>"",'
                '{cleaning_in}{r},'
                'IFERROR(INDEX([[Property Database.cleaning_fee]],'
                'MATCH({property}{r},[[Property Database.id]],0))*1,0)))'),
            group="Money"),
        Col("extra_in", "Extra fees", 11, "money",
            validate=number_check(0, 100000, "Extra fees"),
            note="Blank works out the extra guest charge for you from "
                 "the Settings rate. Type a number to override, or 0 to "
                 "waive it.",
            group="Money"),
        Col("extra_fee", "Extra charged", 12, "money",
            formula=blank(
                "check_in",
                'IF({earns}{r}<>"Yes",0,IF({extra_in}{r}<>"",'
                '{extra_in}{r},'
                'ROUND(MAX(0,{guests_total}{r}-ExtraGuestAfter)*'
                'ExtraGuestFee*{nights}{r}+{pets}{r}*PetFee,2)))'),
            note="Extra guests beyond the free allowance, per night, "
                 "plus the pet fee. All three numbers live on Settings.",
            group="Money"),
        Col("discount", "Discount", 11, "money",
            validate=number_check(0, 100000, "Discount"),
            note="Weekly or monthly discount, or the goodwill you gave "
                 "back. Enter it as a positive number.",
            group="Money"),
        Col("gross", "Gross", 12, "money",
            formula=blank(
                "check_in",
                'ROUND(MAX(0,{accommodation}{r}+{cleaning_fee}{r}+'
                '{extra_fee}{r}-{discount}{r}),2)'),
            cf=money_cf(),
            note="Everything the guest paid, before the platform takes "
                 "its cut.",
            group="Money"),
        Col("tax", "Lodging tax", 11, "money",
            formula=blank("check_in", lodging_tax("{accommodation}{r}")),
            note="Zero when the platform collects and remits it for you, "
                 "which is the default. Switch it on Settings.",
            hidden=True, group="Money"),
        Col("fee_pct", "Commission %", 11, "percent",
            formula=blank("check_in", platform_rate("PlatformCommission",
                                                    "{platform}{r}")),
            hidden=True, group="Money"),
        Col("pay_pct", "Card fee %", 10, "percent",
            formula=blank("check_in", platform_rate("PlatformPayment",
                                                    "{platform}{r}")),
            hidden=True, group="Money"),
        Col("platform_fee", "Platform fee", 12, "money",
            formula=blank(
                "check_in",
                'IF({earns}{r}<>"Yes",0,'
                'ROUND({gross}{r}*{fee_pct}{r},2))'),
            group="Money"),
        Col("payment_fee", "Card fee", 10, "money",
            formula=blank(
                "check_in",
                'IF({earns}{r}<>"Yes",0,'
                'ROUND({gross}{r}*{pay_pct}{r},2))'),
            group="Money"),
        Col("platform_fee_total", "Fees total", 11, "money",
            formula=blank("check_in",
                          "ROUND({platform_fee}{r}+{payment_fee}{r},2)"),
            group="Money"),
        Col("net_revenue", "Net revenue", 13, "money",
            formula=blank(
                "check_in",
                'ROUND({gross}{r}-{platform_fee_total}{r}-{tax}{r},2)'),
            cf=money_cf(),
            note="What you actually keep: gross, less the platform fee, "
                 "less the card fee, less any lodging tax you have to "
                 "remit yourself.",
            group="Money"),
        Col("net_night", "Net a night", 11, "money",
            formula=blank("check_in",
                          'IF({nights}{r}<=0,"",'
                          'ROUND({net_revenue}{r}/{nights}{r},2))'),
            hidden=True, group="Money"),
        Col("payment_status", "Payment", 14, "status",
            validate=list_from("PaymentStatusList"),
            cf=status_cf(), group="Money"),

        # ------------------------------------------------------- flags --
        Col("holds", "Holds dates", 10, "flag",
            formula=('=IF({check_in}{r}="","No",'
                     f'IF({is_holding("{status}{r}")},"Yes","No"))'),
            hidden=True,
            note="Yes for anything that takes the dates off the market - "
                 "a real stay or a block. The calendar and the clash "
                 "check read this."),
        Col("earns", "Earns", 8, "flag",
            formula=('=IF({check_in}{r}="","No",'
                     f'IF({is_earning("{status}{r}")},"Yes","No"))'),
            hidden=True),
        Col("is_block", "Is a block", 10, "flag",
            formula=('=IF({check_in}{r}="","No",'
                     f'IF({is_block("{status}{r}")},"Yes","No"))'),
            hidden=True),
        Col("is_cancelled", "Cancelled", 10, "flag",
            formula=('=IF({check_in}{r}="","No",'
                     'IF({status}{r}="Cancelled","Yes","No"))'),
            hidden=True),
        Col("stay_band", "Stay length", 12, "text",
            formula=blank(
                "check_in",
                'IF({nights}{r}<=0,"",'
                'IF({nights}{r}=1,"1 night",'
                'IF({nights}{r}=2,"2 nights",'
                'IF({nights}{r}<=4,"3-4 nights",'
                'IF({nights}{r}<=6,"5-6 nights",'
                'IF({nights}{r}<=9,"1 week",'
                'IF({nights}{r}<=27,"10+ nights","Monthly stay")))))))'),
            hidden=True, group="Flags"),

        # ------------------------------------------------- the warnings --
        Col("clash", "Clash", 9, "int",
            formula=zero(
                "check_in",
                'IF({holds}{r}<>"Yes",0,'
                'COUNTIFS({property_first}:{property_last},{property}{r},'
                '{ci_num_first}:{ci_num_last},"<"&{co_num}{r},'
                '{co_num_first}:{co_num_last},">"&{ci_num}{r},'
                '{holds_first}:{holds_last},"Yes")-1)'),
            hidden=True,
            note="How many other reservations on this property overlap "
                 "these dates. Anything above zero is a double booking."),
        Col("same_day", "Same-day turn", 12, "flag",
            formula=blank(
                "check_in",
                'IF({holds}{r}<>"Yes","No",'
                'IF({nights}{r}<=0,"No",'
                'IF(COUNTIFS({property_first}:{property_last},'
                '{property}{r},{ci_num_first}:{ci_num_last},{co_num}{r},'
                '{holds_first}:{holds_last},"Yes")>0,"Yes","No")))'),
            cf=[{"type": "text", "criteria": "containing", "value": "Yes",
                 "fmt": "warn"}],
            note="Somebody else checks in on the day this guest leaves. "
                 "The cleaner has hours, not a day.",
            group="Warnings"),
        Col("has_cleaning", "Cleaning booked", 13, "flag",
            formula=blank(
                "check_in",
                f'IF({{holds}}{{r}}<>"Yes","No",'
                f'IF(COUNTIFS({cl_booking},{{id}}{{r}})>0,"Yes","No"))'),
            cf=flag_yes_no_cf(), hidden=True, group="Warnings"),
        Col("has_arrival", "Arrival logged", 13, "flag",
            formula=blank(
                "check_in",
                f'IF({{earns}}{{r}}<>"Yes","No",'
                f'IF(COUNTIFS({ar_booking},{{id}}{{r}})>0,"Yes","No"))'),
            hidden=True, group="Warnings"),
        Col("has_review", "Review logged", 13, "flag",
            formula=blank(
                "check_in",
                f'IF({{earns}}{{r}}<>"Yes","No",'
                f'IF(COUNTIFS({rv_booking},{{id}}{{r}})>0,"Yes","No"))'),
            hidden=True, group="Warnings"),
        Col("stay_state", "Where it is", 13, "status",
            formula=blank(
                "check_in",
                'IF({holds}{r}<>"Yes","",'
                'IF({check_out}{r}="","",'
                'IF(TodayDate<{check_in}{r},'
                'IF({check_in}{r}-TodayDate<=UpcomingDays,"Arriving soon",'
                '"Upcoming"),'
                'IF(TodayDate>={check_out}{r},"Past",'
                'IF(TodayDate={check_in}{r},"Arriving today","Staying")))))'),
            cf=alert_cf(), group="Warnings"),
        Col("arr_key", "Arrival order key", 10, "number",
            formula=('=IF({check_in}{r}="",9999999999,'
                     'IF({earns}{r}<>"Yes",9999999999,'
                     'IF({check_in}{r}<TodayDate,9999999999,'
                     '{ci_num}{r}*10000+ROW()-4)))'),
            hidden=True, group="Flags"),
        Col("arrival_rank", "Arriving next", 10, "int",
            formula=('=IF({arr_key}{r}>=9999999999,"",'
                     'COUNTIFS({arr_key_first}:{arr_key_last},"<"&'
                     '{arr_key}{r})+1)'),
            note="1 is the next guest through the door. The dashboard "
                 "reads this to list who is arriving.",
            hidden=True, group="Flags"),
        Col("alert", "Action", 28, "status",
            formula=blank(
                "check_in",
                'IF({property}{r}="","Add - no property on this row",'
                'IF({clash}{r}>0,"Double booked - '
                'overlapping reservation",'
                'IF({check_out}{r}="","Add - no check-out date",'
                'IF({nights}{r}<=0,"Check - check-out is not after '
                'check-in",'
                'IF({status}{r}="Cancelled","Cancelled",'
                'IF({is_block}{r}="Yes","Blocked - '
                'no guest, no revenue",'
                'IF({over_capacity}{r}="Yes","Check - over the sleeps '
                'limit",'
                'IF({status}{r}="Inquiry",'
                'IF(TodayDate>{check_in}{r},"Chase - enquiry went past '
                'the dates","Follow up - still an enquiry"),'
                'IF({guest}{r}="","Add - no guest record",'
                'IF({has_cleaning}{r}="No","Schedule - no cleaning job '
                'for the check-out",'
                'IF({same_day}{r}="Yes","Same day - tight turnaround",'
                'IF({rate_used}{r}<=0,"Add - no nightly rate",'
                'IF(TodayDate>={check_out}{r},'
                'IF({payment_status}{r}="Paid in full",'
                'IF({has_review}{r}="No","Chase - no review logged","OK"),'
                '"Chase - payout not marked paid"),'
                '"OK")))))))))))))'),
            cf=alert_cf(), group="Warnings"),

        # ------------------------------------------------------- notes --
        Col("requests", "Guest requests", 30, "wrap",
            note="Early check-in, travel cot, late flight. Printed on "
                 "the arrival sheet.",
            group="Notes"),
        Col("notes", "Notes", 30, "wrap", group="Notes"),
    ]

    return SheetSpec(
        name=NAME,
        title="Booking Tracker",
        subtitle="Every reservation, enquiry and block - priced, counted "
                 "and checked for clashes",
        columns=columns,
        rows=cfg.row_count(NAME),
        freeze_cols=3,
        samples=sample("bookings", cfg),
        help_text="One row per reservation. Type the property, the guest, "
                  "the platform and the two dates and the rest works "
                  "itself out: nights, revenue, the platform's cut, what "
                  "you keep, and whether the dates clash with a booking "
                  "you already took. Use the Owner block and Maintenance "
                  "block statuses to take dates off the market.",
        extra_notes=[
            "Leave the nightly rate blank to use the property's standard "
            "rate - type one only when this stay is priced differently.",
            "A Clash warning means two reservations that hold the dates "
            "overlap on the same property. Cancel one or change the "
            "dates; the calendar will keep shouting until you do.",
            "Cleaning fees are charged to the guest and recovered as a "
            "cost on the Cleaning Schedule tab. They are never counted "
            "as room revenue, so ADR and RevPAR stay honest.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Booking totals and a platform breakdown under the register."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index

    row = sheet.section(row, "Every booking on this tab", last)
    row = sheet.table_head(row, 0, [
        "Bookings", "Enquiries", "Cancelled", "Blocks", "Nights booked",
        "Nights blocked", "Guests", "Avg stay", "Clashes", "Same-day turns"],
        [11, 11, 10, 9, 13, 13, 9, 10, 9, 14])
    sheet.calc_row(row, 0, [
        '=COUNTIFS([[Booking Tracker.earns]],"Yes")',
        '=COUNTIFS([[Booking Tracker.status]],"Inquiry")',
        '=COUNTIFS([[Booking Tracker.is_cancelled]],"Yes")',
        '=COUNTIFS([[Booking Tracker.is_block]],"Yes")',
        '=SUMIFS([[Booking Tracker.nights]],'
        '[[Booking Tracker.earns]],"Yes")',
        '=SUMIFS([[Booking Tracker.nights]],'
        '[[Booking Tracker.is_block]],"Yes")',
        '=SUMIFS([[Booking Tracker.guests_total]],'
        '[[Booking Tracker.earns]],"Yes")',
        '=IF(COUNTIFS([[Booking Tracker.earns]],"Yes")<=0,"",'
        'ROUND(SUMIFS([[Booking Tracker.nights]],'
        '[[Booking Tracker.earns]],"Yes")/'
        'COUNTIFS([[Booking Tracker.earns]],"Yes"),1))',
        '=COUNTIFS([[Booking Tracker.clash]],">0")',
        '=COUNTIFS([[Booking Tracker.same_day]],"Yes")',
    ], ["int", "int", "int", "int", "int", "int", "int", "number1",
        "int", "int"])
    row += 2

    row = sheet.table_head(row, 0, [
        "Room revenue", "Cleaning charged", "Extras", "Discounts",
        "Gross", "Platform fees", "Card fees", "Lodging tax",
        "Net revenue", "Net a night"],
        [13, 14, 11, 11, 12, 13, 11, 11, 13, 12])
    sheet.calc_row(row, 0, [
        '=ROUND(SUMIFS([[Booking Tracker.accommodation]],'
        '[[Booking Tracker.earns]],"Yes"),2)',
        '=ROUND(SUMIFS([[Booking Tracker.cleaning_fee]],'
        '[[Booking Tracker.earns]],"Yes"),2)',
        '=ROUND(SUMIFS([[Booking Tracker.extra_fee]],'
        '[[Booking Tracker.earns]],"Yes"),2)',
        '=ROUND(SUMIFS([[Booking Tracker.discount]],'
        '[[Booking Tracker.earns]],"Yes"),2)',
        '=ROUND(SUMIFS([[Booking Tracker.gross]],'
        '[[Booking Tracker.earns]],"Yes"),2)',
        '=ROUND(SUMIFS([[Booking Tracker.platform_fee]],'
        '[[Booking Tracker.earns]],"Yes"),2)',
        '=ROUND(SUMIFS([[Booking Tracker.payment_fee]],'
        '[[Booking Tracker.earns]],"Yes"),2)',
        '=ROUND(SUMIFS([[Booking Tracker.tax]],'
        '[[Booking Tracker.earns]],"Yes"),2)',
        '=ROUND(SUMIFS([[Booking Tracker.net_revenue]],'
        '[[Booking Tracker.earns]],"Yes"),2)',
        '=IF(SUMIFS([[Booking Tracker.nights]],'
        '[[Booking Tracker.earns]],"Yes")<=0,"",'
        'ROUND(SUMIFS([[Booking Tracker.net_revenue]],'
        '[[Booking Tracker.earns]],"Yes")/'
        'SUMIFS([[Booking Tracker.nights]],'
        '[[Booking Tracker.earns]],"Yes"),2))',
    ], ["money", "money", "money", "money", "money", "money", "money",
        "money", "money", "money"])
    row += 2

    row = sheet.section(row, "Which platform is worth having", last)
    row = sheet.table_head(row, 0, [
        "Platform", "Bookings", "Nights", "Gross", "Fees",
        "Net revenue", "Fee rate", "Net a night", "Avg stay",
        "Share of net"],
        [18, 10, 9, 12, 11, 13, 10, 12, 10, 12])
    for name, _commission, _payment, _counts, _note in cfg.platform_table:
        key = f'"{name}"'
        earns = '[[Booking Tracker.earns]],"Yes"'
        plat = f'[[Booking Tracker.platform]],{key}'
        gross = f'SUMIFS([[Booking Tracker.gross]],{plat},{earns})'
        fees = (f'SUMIFS([[Booking Tracker.platform_fee_total]],'
                f'{plat},{earns})')
        net = f'SUMIFS([[Booking Tracker.net_revenue]],{plat},{earns})'
        nights = f'SUMIFS([[Booking Tracker.nights]],{plat},{earns})'
        count = f'COUNTIFS({plat},{earns})'
        total_net = (f'SUMIFS([[Booking Tracker.net_revenue]],{earns})')
        sheet.calc_row(row, 0, [
            name,
            f"={count}",
            f"={nights}",
            f"=ROUND({gross},2)",
            f"=ROUND({fees},2)",
            f"=ROUND({net},2)",
            f'=IF({gross}<=0,"",{fees}/{gross})',
            f'=IF({nights}<=0,"",ROUND({net}/{nights},2))',
            f'=IF({count}<=0,"",ROUND({nights}/{count},1))',
            f'=IF({total_net}<=0,"",{net}/{total_net})',
        ], ["text", "int", "int", "money", "money", "money", "percent",
            "money", "number1", "percent"])
        row += 1

    row += 1
    sheet.note(row, "Fee rate is what the platform actually kept out of "
                    "the money the guest handed over, including card "
                    "processing. Compare it against the net-a-night "
                    "column before you decide a channel is not worth "
                    "the commission - a platform that fills midweek "
                    "nights at a lower rate can still beat one that "
                    "only sends you weekends.", last)
