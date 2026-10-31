"""`Check-in & Check-out` - the eight messages every guest needs.

Hosting is mostly admin, and the admin is always the same eight things:
the directions, the door code, the Wi-Fi, the house rules, the welcome
message, the check-out instructions, the walk round afterwards and the
photos.  Forget one and you get a one-star review about communication.

So this tab is a tick list with a score.  Each row is one stay, each
column is one of the things that has to be sent or done, and the
readiness column tells you - in one glance down the sheet - which guest
is still waiting to hear from you.
"""

from __future__ import annotations

from typing import List, Tuple

from ..config import Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, time_check
from .common import (alert_cf, blank, flag_yes_no_cf, id_col, 
                    lookup_booking, month_col, weekday_col, yesno)

NAME = "Check-in & Check-out"

#: (column key, header, when it matters, help text)
BEFORE_ARRIVAL: List[Tuple[str, str, str]] = [
    ("sent_directions", "Directions sent",
     "How to find the place, parking, which door."),
    ("sent_access", "Access code sent",
     "Key box code, smart lock PIN or who is meeting them."),
    ("sent_wifi", "Wi-Fi details sent",
     "Network and password. The single most asked question."),
    ("sent_rules", "House rules sent",
     "Quiet hours, smoking, pets, extra visitors."),
    ("sent_welcome", "Welcome message sent",
     "The day before or the morning of arrival."),
]

AFTER_DEPARTURE: List[Tuple[str, str, str]] = [
    ("sent_checkout", "Check-out notes sent",
     "What to do with the keys, the bins and the heating."),
    ("checked_out", "Checked out",
     "They have gone and the keys are back."),
    ("inspected", "Property inspected",
     "Someone has walked round it since they left."),
    ("photographed", "Condition photographed",
     "Your evidence if a deposit claim ever comes up."),
]

ALL_STEPS = BEFORE_ARRIVAL + AFTER_DEPARTURE


def spec(cfg: Config) -> SheetSpec:
    steps_done = "+".join(
        f'({{{key}}}{{r}}="Yes")' for key, _h, _n in ALL_STEPS)
    before_done = "+".join(
        f'({{{key}}}{{r}}="Yes")' for key, _h, _n in BEFORE_ARRIVAL)

    columns = [
        id_col("booking", "ArrivalPrefix", digits=cfg.digits("arrival"),
               width=10, header="Arrival ID"),
        Col("booking", "Booking", 11, "text",
            validate=list_from("BookingList"),
            note="Pick the booking reference. Everything on the left of "
                 "this sheet is pulled from it.",
            group="The stay"),
        Col("property", "Property", 11, "text",
            formula=lookup_booking("property"), hidden=True,
            group="The stay"),
        Col("property_name", "Property", 20, "text",
            formula=lookup_booking("property_name"), group="The stay"),
        Col("guest_name", "Guest", 20, "text",
            formula=lookup_booking("guest_name"), group="The stay"),
        Col("arrival_date", "Arrives", 12, "date",
            formula=lookup_booking("check_in"), group="The stay"),
        weekday_col("arrival_date", "arr_day", "Day", 6, hidden=True),
        Col("departure_date", "Leaves", 12, "date",
            formula=lookup_booking("check_out"), group="The stay"),
        Col("nights", "Nights", 8, "int",
            formula=lookup_booking("nights"), hidden=True,
            group="The stay"),
        month_col("arrival_date"),

        # ----------------------------------------------- what happened --
        Col("arrival_time", "Arrived at", 10, "time",
            validate=time_check(),
            note="The actual arrival time, if you know it. Leave blank "
                 "for a self check-in you never saw.",
            group="On the day"),
        Col("departure_time", "Left at", 10, "time",
            validate=time_check(), group="On the day"),
        Col("late_checkout", "Late check-out agreed", 14, "flag",
            validate=yesno("Did you agree a later departure?"),
            cf=flag_yes_no_cf(), group="On the day"),
        Col("early_checkin", "Early check-in agreed", 14, "flag",
            validate=yesno(), cf=flag_yes_no_cf(), group="On the day"),
        Col("key_method", "Access used", 16, "text",
            note="Key box, smart lock, met in person. Worth recording "
                 "when a code goes wrong.",
            group="On the day"),
    ]

    for key, header, note in BEFORE_ARRIVAL:
        columns.append(
            Col(key, header, 13, "flag", validate=yesno(note),
                cf=flag_yes_no_cf(), note=note, group="Before they arrive"))

    columns.append(
        Col("pre_arrival_done", "Pre-arrival score", 13, "percent",
            formula=blank("booking",
                          f"({before_done})/{len(BEFORE_ARRIVAL)}"),
            cf=[{"type": "cell", "criteria": "<", "value": 1,
                 "fmt": "warn"},
                {"type": "cell", "criteria": ">=", "value": 1,
                 "fmt": "good"}],
            note="Everything the guest needs before they travel.",
            group="Before they arrive"))

    for key, header, note in AFTER_DEPARTURE:
        columns.append(
            Col(key, header, 13, "flag", validate=yesno(note),
                cf=flag_yes_no_cf(), note=note, group="After they leave"))

    columns += [
        Col("steps_done", "Done", 7, "int",
            formula=blank("booking", f"({steps_done})"),
            hidden=True, group="Where it stands"),
        Col("readiness", "Complete", 10, "percent",
            formula=blank("booking", f"({steps_done})/{len(ALL_STEPS)}"),
            cf=[{"type": "data_bar", "bar_color": "#3F97A0"}],
            group="Where it stands"),
        Col("outstanding", "Still to do", 11, "int",
            formula=blank("booking",
                          f"{len(ALL_STEPS)}-({steps_done})"),
            group="Where it stands"),
        Col("state", "Where it is", 14, "status",
            formula=blank(
                "booking",
                'IF({arrival_date}{r}="","",'
                'IF(TodayDate<{arrival_date}{r},"Upcoming",'
                'IF(TodayDate={arrival_date}{r},"Arriving today",'
                'IF({departure_date}{r}="","Staying",'
                'IF(TodayDate<{departure_date}{r},"Staying",'
                'IF(TodayDate={departure_date}{r},"Leaving today",'
                '"Past"))))))'),
            cf=alert_cf(), group="Where it stands"),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "booking",
                'IF({arrival_date}{r}="","Check - booking not found",'
                'IF({state}{r}="Upcoming",'
                'IF({arrival_date}{r}-TodayDate>UpcomingDays,"",'
                'IF({pre_arrival_done}{r}<1,"Send - '
                'arrival details still missing","OK - ready for them")),'
                'IF({state}{r}="Arriving today",'
                'IF({pre_arrival_done}{r}<1,"Send - '
                'they arrive today and have not been told everything",'
                '"OK - ready for them"),'
                'IF({state}{r}="Staying","OK - in residence",'
                'IF({state}{r}="Leaving today",'
                'IF({sent_checkout}{r}<>"Yes","Send - check-out notes",'
                '"OK - leaving today"),'
                'IF({checked_out}{r}<>"Yes","Check - never marked '
                'checked out",'
                'IF({inspected}{r}<>"Yes","Action - nobody has '
                'inspected it",'
                'IF({photographed}{r}<>"Yes","Action - no condition '
                'photos","OK"))))))))'),
            cf=alert_cf(), group="Where it stands"),
        Col("issues", "Anything wrong", 30, "wrap",
            note="Damage, a missing key, a complaint. If it costs money, "
                 "log it on the Maintenance Log too.",
            group="Notes"),
        Col("notes", "Notes", 28, "wrap", group="Notes"),
    ]

    return SheetSpec(
        name=NAME,
        title="Check-in & Check-out",
        subtitle="The eight things every guest needs, and which of them "
                 "you have actually done",
        columns=columns,
        rows=cfg.row_count(NAME),
        freeze_cols=2,
        samples=sample("arrivals", cfg),
        help_text="One row per stay. Pick the booking and the property, "
                  "guest and dates arrive with it. Then tick off the "
                  "nine steps. The Action column reads the dates and "
                  "the ticks together, so it only nags you about the "
                  "guests who are actually coming.",
        extra_notes=[
            "The pre-arrival score is the one to watch. A guest who has "
            "not been sent a door code the day before arrival is "
            "tomorrow's problem, and this is the column that spots it "
            "today.",
            "Condition photographs are worth the thirty seconds. They "
            "are the only evidence you will have if you ever need to "
            "claim against a deposit.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """How on top of the admin you actually are."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index

    row = sheet.section(row, "How on top of it you are", last)
    row = sheet.table_head(row, 0, [
        "Stays logged", "Upcoming", "In residence", "Leaving today",
        "Fully prepared", "Missing something", "Average complete",
        "Not inspected", "No photos"],
        [12, 11, 12, 13, 14, 15, 15, 13, 11])
    sheet.calc_row(row, 0, [
        '=COUNTIF([[Check-in & Check-out.booking]],"?*")',
        '=COUNTIFS([[Check-in & Check-out.state]],"Upcoming")',
        '=COUNTIFS([[Check-in & Check-out.state]],"Staying")',
        '=COUNTIFS([[Check-in & Check-out.state]],"Leaving today")',
        '=COUNTIFS([[Check-in & Check-out.pre_arrival_done]],1)',
        '=COUNTIFS([[Check-in & Check-out.state]],"Upcoming",'
        '[[Check-in & Check-out.pre_arrival_done]],"<1")+'
        'COUNTIFS([[Check-in & Check-out.state]],"Arriving today",'
        '[[Check-in & Check-out.pre_arrival_done]],"<1")',
        '=IF(COUNTIF([[Check-in & Check-out.booking]],"?*")<=0,"",'
        'SUM([[Check-in & Check-out.steps_done]])/'
        f'(COUNTIF([[Check-in & Check-out.booking]],"?*")*{len(ALL_STEPS)}))',
        '=COUNTIFS([[Check-in & Check-out.state]],"Past",'
        '[[Check-in & Check-out.inspected]],"<>Yes")',
        '=COUNTIFS([[Check-in & Check-out.state]],"Past",'
        '[[Check-in & Check-out.photographed]],"<>Yes")',
    ], ["int", "int", "int", "int", "int", "int", "percent", "int", "int"])
