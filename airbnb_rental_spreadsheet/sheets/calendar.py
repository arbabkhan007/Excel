"""`Booking Calendar` - a month at a glance, and the next fortnight.

Two views on one page.

The **month grid** is one property at a time: pick the property and
the month at the top and every day fills itself in - who is staying,
which platform, and whether the day is free, occupied, an arrival, a
departure, blocked by you or closed for maintenance.  Colour does the
rest.

The **fortnight strip** is every property at once for the next
fourteen days, which is the view you actually want on a Friday
morning.  Underneath it sits the Turnover Command Center: today's
departures, tomorrow's arrivals, which flats are clean and which are
not.

One thing this is not, and the workbook never pretends otherwise: it
does not talk to Airbnb.  There is no sync, no iCal, no API.  It draws
what is in your Booking Tracker, which means it is right exactly as
often as the Booking Tracker is.
"""

from __future__ import annotations

from ..config import CALENDAR_STRIP_DAYS, Config, Palette
from ..layout import Registry, col_letter
from ..writer import SheetWriter, list_from

NAME = "Booking Calendar"
LAST_COL = 7
WEEK_ROWS = 6
DAY_LINES = 3
STRIP_PROPERTIES = 12


def _occupied(day: str, prop: str = "CalProperty") -> str:
    """COUNTIFS for a paying stay covering `day` at `prop`."""
    return (f'COUNTIFS([[Booking Tracker.property]],{prop},'
            f'[[Booking Tracker.check_in]],"<="&{day},'
            f'[[Booking Tracker.check_out]],">"&{day},'
            f'[[Booking Tracker.earns]],"Yes")')


def _blocked(day: str, status: str, prop: str = "CalProperty") -> str:
    return (f'COUNTIFS([[Booking Tracker.property]],{prop},'
            f'[[Booking Tracker.check_in]],"<="&{day},'
            f'[[Booking Tracker.check_out]],">"&{day},'
            f'[[Booking Tracker.status]],"{status}")')


def _arrives(day: str, prop: str = "CalProperty") -> str:
    return (f'COUNTIFS([[Booking Tracker.property]],{prop},'
            f'[[Booking Tracker.check_in]],{day},'
            f'[[Booking Tracker.earns]],"Yes")')


def _departs(day: str, prop: str = "CalProperty") -> str:
    return (f'COUNTIFS([[Booking Tracker.property]],{prop},'
            f'[[Booking Tracker.check_out]],{day},'
            f'[[Booking Tracker.earns]],"Yes")')


def _row_on(day: str, prop: str = "CalProperty") -> str:
    """Row number of the stay covering `day`, or 0. Numeric ranges only."""
    return (f'SUMPRODUCT(MAX(([[Booking Tracker.property]]={prop})*'
            f'([[Booking Tracker.ci_num]]<={day})*'
            f'([[Booking Tracker.co_num]]>{day})*'
            f'([[Booking Tracker.earns]]="Yes")*'
            f'[[Booking Tracker.seq_num]]))')


def _guest_on(day: str, prop: str = "CalProperty") -> str:
    """Name of whoever is in on `day`. No array formula, no CSE."""
    found = _row_on(day, prop)
    return (f'IF({found}<=0,"",IFERROR(INDEX('
            f'[[Booking Tracker.guest_name]],{found})&"",""))')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_mid)
    ws = s.ws

    s.banner("Booking Calendar",
             "One property a month, and every property for the next "
             "fortnight. Drawn from the Booking Tracker, not from "
             "Airbnb.", LAST_COL)
    s.set_col(0, 0, 20)
    for column in range(1, 8):
        s.set_col(column, column, 19)
    s.set_col(8, 20, 13)

    # ------------------------------------------------------- the picker --
    row = 4
    row = s.note(
        row,
        "This calendar is not connected to Airbnb, Vrbo or any other "
        "site and cannot be. It draws whatever is in your Booking "
        "Tracker, so keep that tab honest and this one looks after "
        "itself.",
        LAST_COL, height=26)
    row += 1

    s.merge(row, 0, 0, "Property", fmts.chrome("label_bold"))
    s.merge(row, 1, 2, None, fmts.chrome("form_box", locked=False))
    ws.data_validation(row, 1, row, 1, list_from("PropertyList"))
    book.define_name("CalProperty", f"='{NAME}'!$B${row + 1}")
    s.merge(row, 3, 4, None, fmts.calc("text"))
    s.formula(
        row, 3,
        '=IFERROR(INDEX([[Property Database.name]],MATCH(CalProperty,'
        '[[Property Database.id]],0)),"Pick a property ID")',
        fmts.calc("text"))
    s.merge(row, 5, LAST_COL, "Type a property ID, e.g. PRP-01",
            fmts.chrome("note"))
    picker_row = row
    row += 1

    s.merge(row, 0, 0, "Month", fmts.chrome("label_bold"))
    s.write(row, 1, cfg.today.replace(day=1),
            fmts.chrome("form_box", locked=False,
                        num_format=cfg.date_format))
    book.define_name("CalMonth", f"='{NAME}'!$B${row + 1}")
    s.merge(row, 2, 2, None, fmts.calc("text"))
    s.formula(row, 2, '=TEXT(CalStart,"mmmm yyyy")',
                     fmts.calc("text"))
    s.merge(row, 3, LAST_COL,
            "Any date in the month you want to see - it snaps to the "
            "first.", fmts.chrome("note"))
    row += 1

    # Hidden helpers: first of month, and the Monday the grid starts on.
    helper = row
    s.write(helper, 20, None, fmts.calc("date"))
    s.formula(
        helper, 20,
        '=IF(CalMonth="",DATE(YEAR(TodayDate),MONTH(TodayDate),1),'
        'DATE(YEAR(CalMonth),MONTH(CalMonth),1))', fmts.calc("date"))
    book.define_name("CalStart", f"='{NAME}'!$U${helper + 1}")
    s.write(helper + 1, 20, None, fmts.calc("date"))
    s.formula(helper + 1, 20,
                     "=CalStart-WEEKDAY(CalStart,3)", fmts.calc("date"))
    book.define_name("CalGridStart", f"='{NAME}'!$U${helper + 2}")
    ws.set_column(20, 20, None, None, {"hidden": True})
    row += 1

    # --------------------------------------------------- month summary --
    days_in = ('DAY(DATE(YEAR(CalStart),MONTH(CalStart)+1,1)-1)')
    nights_sold = ('SUMIFS([[Booking Tracker.nights]],'
                   '[[Booking Tracker.property]],CalProperty,'
                   '[[Booking Tracker.month]],TEXT(CalStart,"yyyy-mm"),'
                   '[[Booking Tracker.earns]],"Yes")')
    accom = ('SUMIFS([[Booking Tracker.accommodation]],'
             '[[Booking Tracker.property]],CalProperty,'
             '[[Booking Tracker.month]],TEXT(CalStart,"yyyy-mm"),'
             '[[Booking Tracker.earns]],"Yes")')
    blocked = ('SUMIFS([[Booking Tracker.nights]],'
               '[[Booking Tracker.property]],CalProperty,'
               '[[Booking Tracker.month]],TEXT(CalStart,"yyyy-mm"),'
               '[[Booking Tracker.is_block]],"Yes")')

    s.kpi(row, 0, "Nights sold", f'=IF(CalProperty="","",{nights_sold})',
          "int")
    s.kpi(row, 1, "Nights free",
          f'=IF(CalProperty="","",MAX(0,{days_in}-{nights_sold}-'
          f'{blocked}))', "int", span=1)
    s.kpi(row, 2, "Blocked", f'=IF(CalProperty="","",{blocked})', "int",
          span=1)
    s.kpi(row, 3, "Occupancy",
          f'=IF(CalProperty="","",IF({days_in}-{blocked}<=0,"",'
          f'{nights_sold}/({days_in}-{blocked})))', "percent", span=1)
    s.kpi(row, 4, "Arrivals",
          '=IF(CalProperty="","",COUNTIFS('
          '[[Booking Tracker.property]],CalProperty,'
          '[[Booking Tracker.month]],TEXT(CalStart,"yyyy-mm"),'
          '[[Booking Tracker.earns]],"Yes"))', "int", span=1)
    s.kpi(row, 5, "Room revenue",
          f'=IF(CalProperty="","",ROUND({accom},2))', "money0", span=1)
    s.kpi(row, 6, "ADR",
          f'=IF(CalProperty="","",IF({nights_sold}<=0,"",'
          f'ROUND({accom}/{nights_sold},2)))', "money", span=1)
    row = s.kpi(row, 7, "Same-day turns",
                '=IF(CalProperty="","",COUNTIFS('
                '[[Cleaning Schedule.property]],CalProperty,'
                '[[Cleaning Schedule.month]],TEXT(CalStart,"yyyy-mm"),'
                '[[Cleaning Schedule.same_day]],"Yes"))', "int", span=1)
    row += 1

    # ------------------------------------------------------ month grid --
    row = s.section(row, "The month", LAST_COL)
    for index, day in enumerate(("Monday", "Tuesday", "Wednesday",
                                 "Thursday", "Friday", "Saturday",
                                 "Sunday")):
        s.write(row, index, day, fmts.chrome("header"))
    s.row_height(row, 22)
    row += 1

    grid_top = row
    for week in range(WEEK_ROWS):
        base = row + week * DAY_LINES
        s.row_height(base, 18)
        s.row_height(base + 1, 16)
        s.row_height(base + 2, 16)
        for weekday in range(7):
            offset = week * 7 + weekday
            day = f"(CalGridStart+{offset})"
            in_month = f'MONTH({day})=MONTH(CalStart)'
            # line 1 - the date
            s.formula(
                base, weekday,
                f'=IF({in_month},DAY({day}),"")',
                fmts.chrome("day_number"))
            # line 2 - who is in
            s.formula(
                base + 1, weekday,
                f'=IF(NOT({in_month}),"",IF(CalProperty="","",'
                f'IF({_blocked(day, "Maintenance block")}>0,'
                f'"Maintenance",'
                f'IF({_blocked(day, "Owner block")}>0,"Owner block",'
                f'IF({_occupied(day)}>0,{_guest_on(day)},"Available")'
                f'))))',
                fmts.chrome("day_item"))
            # line 3 - what happens that day
            s.formula(
                base + 2, weekday,
                f'=IF(NOT({in_month}),"",IF(CalProperty="","",'
                f'IF({_departs(day)}>0,'
                f'IF({_arrives(day)}>0,"Turnover - out and in",'
                f'"Check-out"),'
                f'IF({_arrives(day)}>0,"Check-in",'
                f'IF({_occupied(day)}>1,"DOUBLE BOOKED",'
                f'IF({_occupied(day)}>0,"Staying",""))))))',
                fmts.chrome("day_item"))
    grid_bottom = grid_top + WEEK_ROWS * DAY_LINES - 1

    for text, style in (("DOUBLE BOOKED", "bold_bad"),
                        ("Turnover", "warn"),
                        ("Check-out", "info"),
                        ("Check-in", "good"),
                        ("Available", "muted"),
                        ("Owner block", "muted"),
                        ("Maintenance", "bad")):
        ws.conditional_format(grid_top, 0, grid_bottom, 6, {
            "type": "text", "criteria": "containing", "value": text,
            "format": fmts.cf(style)})
    row = grid_bottom + 2

    row = s.note(
        row,
        "A day shows the guest who is sleeping there that night, so a "
        "check-out day belongs to whoever arrives, not to whoever "
        "left. 'DOUBLE BOOKED' means two paying stays claim the same "
        "night in the same property - fix it on the Booking Tracker, "
        "where the clash column will tell you which rows.",
        LAST_COL, height=36)
    row += 1

    # --------------------------------------- the fortnight strip -------
    row = s.section(row, "The next fortnight, every property",
                    STRIP_PROPERTIES + 2)
    row = s.note(
        row,
        "Fourteen days across, every property down. This is the view "
        "worth printing on a Friday: O is occupied, A is an arrival, "
        "D is a departure, T is a same-day turnover, M is closed for "
        "maintenance, B is blocked by you, and a blank cell is a "
        "night you are not selling.",
        STRIP_PROPERTIES + 2, height=38)
    row += 1

    s.write(row, 0, "Property", fmts.chrome("header_left"))
    s.set_col(0, 0, 22)
    for offset in range(CALENDAR_STRIP_DAYS):
        s.formula(
            row, offset + 1,
            f'=TEXT(TodayDate+{offset},"ddd")&" "&'
            f'TEXT(TodayDate+{offset},"d mmm")',
            fmts.chrome("header"))
        s.set_col(offset + 1, offset + 1, 7)
    s.row_height(row, 24)
    row += 1

    strip_top = row
    for n in range(1, STRIP_PROPERTIES + 1):
        pid = f'INDEX([[Property Database.id]],{n})'
        s.formula(
            row, 0,
            f'=IFERROR(IF(INDEX([[Property Database.name]],{n})="","",'
            f'INDEX([[Property Database.name]],{n})),"")',
            fmts.calc("text"))
        for offset in range(CALENDAR_STRIP_DAYS):
            day = f"(TodayDate+{offset})"
            s.formula(
                row, offset + 1,
                f'=IFERROR(IF(INDEX([[Property Database.name]],{n})="",'
                f'"",'
                f'IF({_blocked(day, "Maintenance block", pid)}>0,"M",'
                f'IF({_blocked(day, "Owner block", pid)}>0,"B",'
                f'IF({_departs(day, pid)}+{_arrives(day, pid)}>1,"T",'
                f'IF({_arrives(day, pid)}>0,"A",'
                f'IF({_departs(day, pid)}>0,"D",'
                f'IF({_occupied(day, pid)}>0,"O","")'
                f')))))),"")',
                fmts.calc("text"))
        row += 1
    strip_bottom = row - 1
    for text, style in (("T", "bold_bad"), ("A", "good"),
                        ("D", "info"), ("O", "muted"),
                        ("M", "bad"), ("B", "warn")):
        ws.conditional_format(
            strip_top, 1, strip_bottom, CALENDAR_STRIP_DAYS, {
                "type": "cell", "criteria": "equal to",
                "value": f'"{text}"', "format": fmts.cf(style)})
    row += 1

    # ----------------------------------- turnover command center -------
    row = s.section(row, "Turnover command centre", STRIP_PROPERTIES + 2)
    row = s.note(
        row,
        "Everything happening in the next twenty-four hours, in one "
        "place. If these four numbers are all accounted for, nobody "
        "arrives to a dirty flat.",
        STRIP_PROPERTIES + 2, height=26)

    s.kpi(row, 0, "Leaving today",
          '=COUNTIFS([[Booking Tracker.check_out]],TodayDate,'
          '[[Booking Tracker.earns]],"Yes")', "int")
    s.kpi(row, 2, "Arriving today",
          '=COUNTIFS([[Booking Tracker.check_in]],TodayDate,'
          '[[Booking Tracker.earns]],"Yes")', "int")
    s.kpi(row, 4, "Arriving tomorrow",
          '=COUNTIFS([[Booking Tracker.check_in]],TodayDate+1,'
          '[[Booking Tracker.earns]],"Yes")', "int")
    s.kpi(row, 6, "Cleans booked today",
          '=COUNTIFS([[Cleaning Schedule.clean_date]],TodayDate)',
          "int")
    s.kpi(row, 8, "Cleans still open",
          '=COUNTIFS([[Cleaning Schedule.clean_date]],"<="&TodayDate,'
          '[[Cleaning Schedule.status]],"<>Completed")', "int",
          note="Due today or earlier and not finished")
    row = s.kpi(row, 10, "Ready for a guest",
                '=COUNTIFS([[Property Database.alert]],"OK*")', "int")
    row += 1

    row = s.table_head(row, 0, [
        "Property", "Out today", "In today", "Clean booked",
        "Cleaner", "Turn", "Hours", "State"],
        [22, 10, 10, 13, 18, 8, 9, 26])
    cc_top = row
    for n in range(1, STRIP_PROPERTIES + 1):
        pid = f'INDEX([[Property Database.id]],{n})'
        guard = f'INDEX([[Property Database.name]],{n})=""'
        today = "TodayDate"
        clean_here = ('COUNTIFS([[Cleaning Schedule.property]],'
                      f'{pid},[[Cleaning Schedule.clean_date]],'
                      f'{today})')
        clean_idx = ('SUMPRODUCT(MAX(([[Cleaning Schedule.property]]='
                     f'{pid})*([[Cleaning Schedule.clean_num]]='
                     f'{today})*[[Cleaning Schedule.seq_num]]))')
        s.formula(
            row, 0,
            f'=IFERROR(IF({guard},"",INDEX('
            f'[[Property Database.name]],{n})),"")', fmts.calc("text"))
        s.formula(
            row, 1,
            f'=IFERROR(IF({guard},"",{_departs(today, pid)}),"")',
            fmts.calc("int"))
        s.formula(
            row, 2,
            f'=IFERROR(IF({guard},"",{_arrives(today, pid)}),"")',
            fmts.calc("int"))
        s.formula(
            row, 3,
            f'=IFERROR(IF({guard},"",{clean_here}),"")',
            fmts.calc("int"))
        s.formula(
            row, 4,
            f'=IFERROR(IF({guard},"",IF({clean_idx}<=0,"",'
            f'INDEX([[Cleaning Schedule.cleaner_name]],{clean_idx})&"")),'
            f'"")',
            fmts.calc("text"))
        s.formula(
            row, 5,
            f'=IFERROR(IF({guard},"",IF({clean_idx}<=0,"",'
            f'INDEX([[Cleaning Schedule.same_day]],{clean_idx})&"")),'
            f'"")',
            fmts.calc("text"))
        s.formula(
            row, 6,
            f'=IFERROR(IF({guard},"",IF({clean_idx}<=0,"",'
            f'INDEX([[Cleaning Schedule.window_hours]],{clean_idx})*1)),'
            f'"")',
            fmts.calc("number1"))
        s.formula(
            row, 7,
            f'=IFERROR(IF({guard},"",'
            f'IF({_departs(today, pid)}+{_arrives(today, pid)}>1,'
            f'IF({clean_here}<=0,"Urgent - turnover with no clean '
            f'booked","Same day - out and in"),'
            f'IF({_arrives(today, pid)}>0,"Check-in - guest arriving",'
            f'IF({_departs(today, pid)}>0,'
            f'IF({clean_here}<=0,"Action - nobody booked to clean",'
            f'"Check-out - clean booked"),'
            f'IF({_occupied(today, pid)}>0,"Guest in","Empty"))))),"")',
            fmts.calc("text"))
        row += 1
    for text, style in (("Urgent", "bold_bad"), ("Action", "bad"),
                        ("Same day", "warn"), ("Check-in", "good"),
                        ("Check-out", "info"), ("Empty", "muted")):
        ws.conditional_format(cc_top, 7, row - 1, 7, {
            "type": "text", "criteria": "containing", "value": text,
            "format": fmts.cf(style)})
    row += 1

    row = s.note(
        row,
        "'Hours' is the gap between the departing guest's check-out "
        "time and the arriving guest's check-in time. Anything under "
        "the tight-turnaround setting goes red on the Cleaning "
        "Schedule, because that is the day a cleaner needs telling in "
        "advance rather than at eleven o'clock.",
        STRIP_PROPERTIES + 2, height=34)

    s.finish(freeze_row=picker_row + 3, gridlines=False, fit_width=True,
             print_area=f"A1:{col_letter(LAST_COL)}{grid_bottom + 3}")
    s.protect()
    return s
