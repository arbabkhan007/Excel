"""`Guest Database` - who stayed, how often, and what they were worth.

A short-term rental business that only ever talks to strangers is paying
full commission on every single night.  This tab keeps the history: how
many times someone has stayed, which properties they liked, what they
have spent and how they reviewed you - so the message you send before
next summer goes to the right twenty people.

Deliberately modest about what it stores.  A name, a way to contact
them, where they came from and what they asked for is enough to run the
business.  Passport numbers, card details and copies of ID do not
belong in a spreadsheet.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec, max_if, min_if
from ..sample_data import get as sample
from ..writer import list_from
from .common import (alert_cf, bar_cf, blank, flag_yes_no_cf, id_col,
                     money_cf, rating_cf, status_cf)

NAME = "Guest Database"


def spec(cfg: Config) -> SheetSpec:
    bk_guest = "[[Booking Tracker.guest]]"
    bk_earns = "[[Booking Tracker.earns]]"
    bk_nights = "[[Booking Tracker.nights]]"
    bk_gross = "[[Booking Tracker.gross]]"
    bk_net = "[[Booking Tracker.net_revenue]]"
    bk_ci = "[[Booking Tracker.ci_num]]"
    bk_cancel = "[[Booking Tracker.is_cancelled]]"
    rv_guest = "[[Reviews Tracker.guest]]"
    rv_score = "[[Reviews Tracker.rate_overall]]"
    rv_has = "[[Reviews Tracker.has_score]]"

    # MIN and MAX over a condition, without an array formula.  Both read
    # the hidden numeric mirror of the check-in date rather than the date
    # column itself: a formula column that returns "" would be multiplied
    # by a zero somewhere in the SUMPRODUCT and answer #VALUE!.
    _guest_is = bk_guest + "={id}{r}"
    _earning = bk_earns + '="Yes"'
    _earliest = min_if(bk_ci, _guest_is, _earning)
    _latest = max_if(bk_ci, _guest_is, _earning)

    columns = [
        # ---------------------------------------------------- the guest --
        id_col("name", "GuestPrefix", digits=cfg.digits("guest"), width=11,
               header="Guest ID",
               note="Automatic. Bookings point at this, which is how the "
                    "history below builds itself."),
        Col("name", "Guest name", 22, "text",
            note="One row per person, not per stay. If they come back, "
                 "point the new booking at the same guest ID.",
            group="The guest"),
        Col("email", "Email", 26, "text",
            note="Platforms mask this until a booking is confirmed. Add "
                 "it when you get it - it is what makes a direct "
                 "re-booking possible.",
            group="The guest"),
        Col("phone", "Phone", 16, "text", group="The guest"),
        Col("country", "Country", 15, "text",
            validate=list_from("CountryList"), group="The guest"),
        Col("guest_type", "Guest type", 13, "text",
            validate=list_from("GuestTypeList"), group="The guest"),
        Col("purpose", "Usual trip", 14, "text",
            validate=list_from("TripPurposeList"), group="The guest"),
        Col("first_platform", "Found you on", 15, "text",
            validate=list_from("PlatformList"),
            note="Where the first booking came from. A guest who first "
                 "arrived via Airbnb and now books direct is the whole "
                 "point of keeping this tab.",
            group="The guest"),
        Col("marketing_ok", "Happy to be contacted", 16, "flag",
            validate=list_from("YesNoList"),
            cf=flag_yes_no_cf(),
            note="Only mail the people who said yes. Keep this honest - "
                 "it is the column that keeps you out of trouble.",
            group="The guest"),

        # -------------------------------------------------- the history --
        Col("stays", "Stays", 8, "int",
            formula=blank("name", f'COUNTIFS({bk_guest},{{id}}{{r}},'
                                  f'{bk_earns},"Yes")'),
            cf=bar_cf(), group="History"),
        Col("cancellations", "Cancelled", 10, "int",
            formula=blank("name", f'COUNTIFS({bk_guest},{{id}}{{r}},'
                                  f'{bk_cancel},"Yes")'),
            hidden=True, group="History"),
        Col("nights", "Nights", 9, "int",
            formula=blank("name", f'SUMIFS({bk_nights},{bk_guest},'
                                  f'{{id}}{{r}},{bk_earns},"Yes")'),
            group="History"),
        Col("first_stay", "First stay", 12, "date",
            formula=blank(
                "name",
                'IF({stays}{r}<=0,"",' + _earliest + ")"),
            note="The check-in date of their very first stay.",
            group="History"),
        Col("last_stay", "Last stay", 12, "date",
            formula=blank(
                "name",
                'IF({stays}{r}<=0,"",' + _latest + ")"),
            group="History"),
        Col("days_since", "Days since", 11, "int",
            formula=blank("name",
                          'IF({last_stay}{r}="","",'
                          'MAX(0,TodayDate-{last_stay}{r}))'),
            group="History"),
        Col("repeat", "Repeat guest", 12, "flag",
            formula=blank("name", 'IF({stays}{r}>1,"Yes","No")'),
            cf=flag_yes_no_cf(), group="History"),

        # ---------------------------------------------------- the money --
        Col("gross", "Lifetime gross", 13, "money",
            formula=blank("name", f'ROUND(SUMIFS({bk_gross},{bk_guest},'
                                  f'{{id}}{{r}},{bk_earns},"Yes"),2)'),
            cf=money_cf(), group="Value"),
        Col("net", "Lifetime net", 13, "money",
            formula=blank("name", f'ROUND(SUMIFS({bk_net},{bk_guest},'
                                  f'{{id}}{{r}},{bk_earns},"Yes"),2)'),
            note="What you kept after the platform's cut. A direct "
                 "repeat guest is worth noticeably more than the gross "
                 "figure suggests.",
            group="Value"),
        Col("avg_stay_value", "Avg per stay", 12, "money",
            formula=blank("name",
                          'IF({stays}{r}<=0,"",'
                          'ROUND({net}{r}/{stays}{r},2))'),
            group="Value"),
        Col("nights_avg", "Avg nights", 11, "number1",
            formula=blank("name",
                          'IF({stays}{r}<=0,"",'
                          'ROUND({nights}{r}/{stays}{r},1))'),
            hidden=True, group="Value"),
        Col("value_rank", "Rank", 7, "int",
            formula=('=IF({name}{r}="","",COUNTIFS({net_first}:{net_last},'
                     '">"&{net}{r})+COUNTIFS({net_first}:{net}{r},'
                     '{net}{r}))'),
            note="1 = your most valuable guest by net revenue.",
            group="Value"),

        # -------------------------------------------------- the reviews --
        Col("reviews", "Reviews left", 12, "int",
            formula=blank("name", f'COUNTIFS({rv_guest},{{id}}{{r}},'
                                  f'{rv_has},"Yes")'),
            hidden=True, group="Reviews"),
        Col("rating_given", "Rating given", 12, "number",
            formula=blank(
                "name",
                f'IF({{reviews}}{{r}}<=0,"",'
                f'ROUND(SUMIFS({rv_score},{rv_guest},{{id}}{{r}},'
                f'{rv_has},"Yes")/{{reviews}}{{r}},2))'),
            cf=rating_cf(),
            note="The average score this guest has given you.",
            group="Reviews"),
        Col("review_status", "Review status", 14, "status",
            validate=list_from("ReviewStatusList"),
            cf=status_cf(), group="Reviews"),

        # ----------------------------------------------- what they like --
        Col("preferences", "Preferences", 28, "wrap",
            note="Ground floor, early check-in, extra pillows, allergic "
                 "to the cat. The reason they book you again.",
            group="Notes"),
        Col("requests", "Standing requests", 26, "wrap", group="Notes"),
        Col("feedback", "What they said", 30, "wrap", group="Notes"),
        Col("internal", "Internal notes", 28, "wrap",
            note="For you, not for them. Keep it factual - a guest can "
                 "ask to see what you hold about them.",
            group="Notes"),

        # ----------------------------------------------- what to do now --
        Col("alert", "Action", 26, "status",
            formula=blank(
                "name",
                'IF({stays}{r}<=0,"Add - no stays linked to this guest",'
                'IF({email}{r}="","Add - no email, no re-booking",'
                'IF({repeat}{r}="Yes",'
                'IF({marketing_ok}{r}="Yes","OK - repeat guest, '
                'contactable","Check - repeat guest, no consent"),'
                'IF({days_since}{r}="","",'
                'IF({days_since}{r}>330,"Follow up - a year since they '
                'stayed",'
                'IF({review_status}{r}="Requested","Chase - review '
                'requested, none received","OK"))))))'),
            cf=alert_cf(), group="Notes"),
    ]

    return SheetSpec(
        name=NAME,
        title="Guest Database",
        subtitle="Who has stayed with you, how often, and what they were "
                 "worth after commission",
        columns=columns,
        rows=cfg.row_count(NAME),
        freeze_cols=2,
        samples=sample("guests", cfg),
        help_text="One row per person, not per stay. Point every booking "
                  "at a guest ID and the stay count, the nights, the "
                  "lifetime value and the review average all build "
                  "themselves. The repeat-guest column is the one worth "
                  "sorting on.",
        extra_notes=[
            "Keep this tab modest. A name, a contact, where they came "
            "from and what they asked for is enough. Identity documents "
            "and card details do not belong in a spreadsheet.",
            "Only contact the guests whose consent column says Yes.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Who is actually worth marketing to."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index

    row = sheet.section(row, "Your guest book", last)
    row = sheet.table_head(row, 0, [
        "Guests on file", "Repeat guests", "Repeat rate", "Contactable",
        "Total stays", "Total nights", "Lifetime net", "Net per guest",
        "Avg rating given"],
        [13, 13, 11, 12, 11, 12, 13, 13, 14])
    guests = '=COUNTIF([[Guest Database.name]],"?*")'
    sheet.calc_row(row, 0, [
        guests,
        '=COUNTIFS([[Guest Database.repeat]],"Yes")',
        '=IF(COUNTIF([[Guest Database.name]],"?*")<=0,"",'
        'COUNTIFS([[Guest Database.repeat]],"Yes")/'
        'COUNTIF([[Guest Database.name]],"?*"))',
        '=COUNTIFS([[Guest Database.marketing_ok]],"Yes")',
        '=SUM([[Guest Database.stays]])',
        '=SUM([[Guest Database.nights]])',
        '=ROUND(SUM([[Guest Database.net]]),2)',
        '=IF(COUNTIF([[Guest Database.name]],"?*")<=0,"",'
        'ROUND(SUM([[Guest Database.net]])/'
        'COUNTIF([[Guest Database.name]],"?*"),2))',
        '=IF(SUM([[Guest Database.reviews]])<=0,"",'
        'ROUND(SUMPRODUCT([[Guest Database.reviews]],'
        'IFERROR([[Guest Database.rating_given]]*1,0))/'
        'SUM([[Guest Database.reviews]]),2))',
    ], ["int", "int", "percent", "int", "int", "int", "money", "money",
        "number"])
    row += 2

    sheet.note(row, "A repeat guest costs you nothing in commission if "
                    "they book direct. Sort by Rank, take the top twenty "
                    "with consent, and that is your whole marketing "
                    "list.", last)
