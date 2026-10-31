"""`Cleaning Schedule` - the turnover, who does it and what it costs.

The cleaning is the business.  A host can get the pricing wrong and
survive; get a turnover wrong and the next guest walks into someone
else's weekend.  This tab is built around the two numbers that actually
matter on a changeover day: how long the cleaner has, and whether the
property is finished before the next arrival.

The cleaning date defaults to the booking's check-out date, so a job
created straight off a reservation needs one click, not six.  Nine task
flags drive a completion percentage, and the printable room-by-room
checklist on the next tab reads the same nine flags.
"""

from __future__ import annotations

from ..config import CLEANING_TASKS, Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check, time_check
from .common import (alert_cf, blank, date_num_col, seq_num_col, flag_yes_no_cf, 
                    id_col, lookup_booking, lookup_booking_num, 
                    lookup_vendor, month_col, status_cf, weekday_col, 
                    yesno)

NAME = "Cleaning Schedule"


def spec(cfg: Config) -> SheetSpec:
    bk_prop = "[[Booking Tracker.property]]"
    bk_ci = "[[Booking Tracker.ci_num]]"
    bk_holds = "[[Booking Tracker.holds]]"

    tasks_done = "+".join(
        f'({{{key}}}{{r}}="Yes")' for key, _h in CLEANING_TASKS)

    columns = [
        id_col("booking", "CleanPrefix", digits=cfg.digits("clean"),
               width=10, header="Job ID"),
        Col("booking", "Booking", 11, "text",
            validate=list_from("BookingList"),
            note="The stay that is ending. Pick it and the property, "
                 "the guest and the check-out date come with it.",
            group="The job"),
        Col("property", "Property", 11, "text",
            formula=lookup_booking("property"), hidden=True,
            group="The job"),
        Col("property_name", "Property", 20, "text",
            formula=lookup_booking("property_name"), group="The job"),
        Col("guest_name", "Guest leaving", 18, "text",
            formula=lookup_booking("guest_name"), group="The job"),
        Col("checkout_date", "Check-out", 12, "date",
            formula=lookup_booking("check_out"), group="The job"),
        Col("checkout_num", "Check-out (number)", 10, "number",
            formula=('=IF({booking}{r}="",0,IFERROR(INDEX('
                     '[[Booking Tracker.co_num]],MATCH({booking}{r},'
                     '[[Booking Tracker.id]],0))*1,0))'),
            hidden=True, group="The job"),
        Col("date_in", "Cleaning date", 12, "date",
            validate=date_check(),
            note="Leave it blank and the job happens on the check-out "
                 "date, which is what you want nine times in ten. Type "
                 "a date only when it slips.",
            group="The job"),
        Col("clean_date", "Cleaned on", 12, "date",
            formula=('=IF({booking}{r}="","",IF({date_in}{r}<>"",'
                     '{date_in}{r},{checkout_date}{r}))'),
            group="The job"),
        weekday_col("clean_date", "clean_day", "Day", 6, hidden=True),
        date_num_col("clean_date", "clean_num", "Cleaned on (number)"),
        seq_num_col("booking"),
        month_col("clean_date"),

        # ------------------------------------------------ who and when --
        Col("cleaner", "Cleaner", 11, "text",
            validate=list_from("VendorList"),
            note="Pick them from Staff & Vendors. The rate and the "
                 "contact details come from there.",
            group="Who and when"),
        Col("cleaner_name", "Cleaner name", 18, "text",
            formula=lookup_vendor("name", "cleaner"), group="Who and when"),
        Col("start_time", "Started", 9, "time",
            validate=time_check(), group="Who and when"),
        Col("end_time", "Finished", 9, "time",
            validate=time_check(), group="Who and when"),
        Col("hours", "Hours", 8, "number1",
            formula=('=IF({booking}{r}="","",IF({start_time}{r}="","",'
                     'IF({end_time}{r}="","",'
                     'ROUND(MOD({end_time}{r}-{start_time}{r},1)*24,2))))'),
            note="Worked out for you, and it copes with a clean that "
                 "runs past midnight.",
            group="Who and when"),
        Col("status", "Status", 16, "status",
            validate=list_from("CleaningStatusList"),
            cf=status_cf(), group="Who and when"),
    ]

    for key, header in CLEANING_TASKS:
        columns.append(
            Col(key, header, 11, "flag", validate=yesno(),
                cf=flag_yes_no_cf(),
                note=f"{header} done. The printable turnover checklist "
                     f"reads this same flag.",
                group="The turnover"))

    columns += [
        Col("tasks_done", "Tasks done", 10, "int",
            formula=blank("booking", f"({tasks_done})"),
            hidden=True, group="The turnover"),
        Col("completion", "Complete", 10, "percent",
            formula=blank("booking",
                          f"({tasks_done})/{len(CLEANING_TASKS)}"),
            cf=[{"type": "cell", "criteria": ">=", "value": 1,
                 "fmt": "good"},
                {"type": "cell", "criteria": ">=", "value": 0.5,
                 "fmt": "warn"},
                {"type": "cell", "criteria": "<", "value": 0.5,
                 "fmt": "bad"}],
            note="Nine tasks, one percentage. Green only when every one "
                 "of them is ticked.",
            group="The turnover"),

        # ------------------------------------------------- the pressure --
        Col("next_arrival", "Next arrival", 12, "date",
            formula=blank(
                "booking",
                f'IF({{property}}{{r}}="","",'
                f'IF(SUMPRODUCT(MIN(({bk_prop}={{property}}{{r}})*'
                f'({bk_holds}="Yes")*({bk_ci}>={{checkout_num}}{{r}})*'
                f'{bk_ci}+(1-({bk_prop}={{property}}{{r}})*'
                f'({bk_holds}="Yes")*({bk_ci}>={{checkout_num}}{{r}}))*'
                f'2958465))>=2958465,"",'
                f'SUMPRODUCT(MIN(({bk_prop}={{property}}{{r}})*'
                f'({bk_holds}="Yes")*({bk_ci}>={{checkout_num}}{{r}})*'
                f'{bk_ci}+(1-({bk_prop}={{property}}{{r}})*'
                f'({bk_holds}="Yes")*({bk_ci}>={{checkout_num}}{{r}}))*'
                f'2958465))))'),
            note="The next check-in on this property on or after the "
                 "check-out. This is the deadline.",
            group="The pressure"),
        Col("gap_days", "Days spare", 10, "int",
            formula=blank("booking",
                          'IF({next_arrival}{r}="","",'
                          'MAX(0,{next_arrival}{r}-{checkout_date}{r}))'),
            group="The pressure"),
        Col("same_day", "Same-day turn", 12, "flag",
            formula=blank("booking",
                          'IF({next_arrival}{r}="","No",'
                          'IF({gap_days}{r}=0,"Yes","No"))'),
            cf=[{"type": "text", "criteria": "containing", "value": "Yes",
                 "fmt": "warn"}],
            note="Somebody checks in the same day. The cleaner has the "
                 "window between the two times and nothing more.",
            group="The pressure"),
        Col("window_hours", "Hours in the window", 14, "number1",
            formula=blank(
                "booking",
                'IF({same_day}{r}<>"Yes","",'
                'IF({property}{r}="","",'
                'ROUND(MAX(0,'
                'IFERROR(INDEX([[Property Database.check_in_time]],'
                'MATCH({property}{r},[[Property Database.id]],0))*1,0.625)-'
                'IFERROR(INDEX([[Property Database.check_out_time]],'
                'MATCH({property}{r},[[Property Database.id]],0))*1,'
                '0.4583))*24,1)))'),
            cf=[{"type": "cell", "criteria": "<",
                 "value": "=TightTurnHours", "fmt": "bad"}],
            note="Check-out time to check-in time on a same-day turn. "
                 "Under the warning hours on Settings and it goes red.",
            group="The pressure"),
        Col("ready", "Ready for the next guest", 15, "status",
            formula=blank(
                "booking",
                'IF({completion}{r}>=1,"Ready",'
                'IF({next_arrival}{r}="","Not cleaned",'
                'IF(TodayDate>{next_arrival}{r},"Ready",'
                'IF({next_arrival}{r}-TodayDate<=1,'
                '"Not cleaned - guest due","Not cleaned"))))'),
            cf=alert_cf(), group="The pressure"),

        # ----------------------------------------------------- the cost --
        Col("cost_in", "Cleaning cost", 12, "money",
            validate=number_check(0, 100000, "Cleaning cost"),
            note="Blank uses the property's standard cleaning cost.",
            group="The cost"),
        Col("clean_cost", "Cost used", 11, "money",
            formula=blank(
                "booking",
                'IF({cost_in}{r}<>"",{cost_in}{r},'
                'IF({property}{r}="",0,'
                'IFERROR(INDEX([[Property Database.cleaning_cost]],'
                'MATCH({property}{r},[[Property Database.id]],0))*1,0)))'),
            group="The cost"),
        Col("laundry_cost", "Laundry", 10, "money",
            validate=number_check(0, 100000, "Laundry"),
            note="Blank uses the standard laundry cost on Settings.",
            group="The cost"),
        Col("supplies_cost", "Consumables", 11, "money",
            validate=number_check(0, 100000, "Consumables"),
            group="The cost"),
        Col("total_cost", "Turnover cost", 13, "money",
            formula=blank(
                "booking",
                'ROUND({clean_cost}{r}+'
                'IF({laundry_cost}{r}="",LaundryPerTurn,{laundry_cost}{r})+'
                'IF({supplies_cost}{r}="",SuppliesPerTurn,'
                '{supplies_cost}{r}),2)'),
            note="Cleaning plus laundry plus consumables. This is the "
                 "real cost of a changeover, and it is what the profit "
                 "per booking is built on.",
            group="The cost"),
        Col("fee_charged", "Fee charged to guest", 14, "money",
            formula=lookup_booking_num("cleaning_fee"),
            hidden=True, group="The cost"),
        Col("cost_gap", "Over or under recovered", 15, "money",
            formula=blank("booking",
                          'ROUND({fee_charged}{r}-{total_cost}{r},2)'),
            cf=[{"type": "cell", "criteria": "<", "value": -0.004,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 0.004,
                 "fmt": "good"}],
            note="The cleaning fee you charged the guest less what the "
                 "turnover actually cost you. Red means you are "
                 "subsidising the clean out of the room rate.",
            group="The cost"),
        Col("paid", "Cleaner paid", 11, "flag",
            validate=yesno(), cf=flag_yes_no_cf(), group="The cost"),

        # ------------------------------------------------ what to do now --
        Col("alert", "Action", 28, "status",
            formula=blank(
                "booking",
                'IF({checkout_date}{r}="","Check - booking not found",'
                'IF({status}{r}="Cancelled","",'
                'IF({cleaner}{r}="","Add - nobody assigned",'
                'IF({clean_date}{r}>TodayDate,'
                'IF({same_day}{r}="Yes","Same day - tight turnaround '
                'coming","OK - scheduled"),'
                'IF({completion}{r}>=1,'
                'IF({paid}{r}<>"Yes","Unpaid - cleaner not paid yet",'
                '"OK"),'
                'IF({next_arrival}{r}="","Action - not finished",'
                'IF({next_arrival}{r}<=TodayDate+1,'
                '"Not cleaned - guest due","Action - not finished")))))))'),
            cf=alert_cf(), group="What to do"),
        Col("notes", "Notes", 30, "wrap",
            note="Damage found, items left behind, anything the next "
                 "cleaner should know.",
            group="What to do"),
    ]

    return SheetSpec(
        name=NAME,
        title="Cleaning Schedule",
        subtitle="Every turnover: who is doing it, how long they have, "
                 "and what it really costs",
        columns=columns,
        rows=cfg.row_count(NAME),
        freeze_cols=2,
        samples=sample("cleaning", cfg),
        help_text="One row per turnover. Pick the booking that is ending "
                  "and the property, guest and check-out date arrive "
                  "with it - leave the cleaning date blank and the job "
                  "lands on the check-out day by itself. Tick the nine "
                  "tasks as they are done; the printable checklist on "
                  "the next tab reads the same ticks.",
        extra_notes=[
            "The turnover cost is cleaning plus laundry plus "
            "consumables, not just the cleaner's fee. Compare it with "
            "the cleaning fee you charge the guest - the 'over or "
            "under recovered' column is one of the quietest money "
            "leaks in this business.",
            "Hours in the window only appears on a same-day turn. It is "
            "the gap between your check-out and check-in times, so if "
            "it is too tight the answer is usually to move one of them "
            "on the Property Database tab.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Turnover totals and the cleaners' league table."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index

    row = sheet.section(row, "Every turnover on this tab", last)
    row = sheet.table_head(row, 0, [
        "Turnovers", "Complete", "Outstanding", "Same-day turns",
        "Cleaning paid", "Laundry & consumables", "Total cost",
        "Fees charged", "Recovered", "Avg cost"],
        [11, 10, 12, 14, 13, 18, 12, 13, 11, 10])
    total = '=ROUND(SUM([[Cleaning Schedule.total_cost]]),2)'
    jobs = '=COUNTIF([[Cleaning Schedule.booking]],"?*")'
    sheet.calc_row(row, 0, [
        jobs,
        '=COUNTIFS([[Cleaning Schedule.completion]],1)',
        '=COUNTIFS([[Cleaning Schedule.completion]],"<1")',
        '=COUNTIFS([[Cleaning Schedule.same_day]],"Yes")',
        '=ROUND(SUM([[Cleaning Schedule.clean_cost]]),2)',
        '=ROUND(SUM([[Cleaning Schedule.total_cost]])-'
        'SUM([[Cleaning Schedule.clean_cost]]),2)',
        total,
        '=ROUND(SUM([[Cleaning Schedule.fee_charged]]),2)',
        '=IF(SUM([[Cleaning Schedule.total_cost]])<=0,"",'
        'SUM([[Cleaning Schedule.fee_charged]])/'
        'SUM([[Cleaning Schedule.total_cost]]))',
        '=IF(COUNTIF([[Cleaning Schedule.booking]],"?*")<=0,"",'
        'ROUND(SUM([[Cleaning Schedule.total_cost]])/'
        'COUNTIF([[Cleaning Schedule.booking]],"?*"),2))',
    ], ["int", "int", "int", "int", "money", "money", "money", "money",
        "percent", "money"])
    row += 2

    sheet.note(row, "'Recovered' is the cleaning fee you charge guests "
                    "as a percentage of what turnovers actually cost "
                    "you. Under 100% means the room rate is quietly "
                    "paying for the cleaning.", last)
