"""`Dashboard` - the page that answers "how is the portfolio doing?"

Twenty numbers across the top, the jobs that need doing before you put
the kettle on, every property side by side on revenue, occupancy,
costs and profit, and twelve months of income against target.
Everything reads from the registers, so nothing here is ever out of
date and there is nothing to refresh.

Five questions, in the order a host actually asks them: who is in
tonight, who arrives tomorrow and is the flat clean, are we full, are
we making money, and which property is quietly losing it.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, col_letter
from ..writer import SheetWriter

NAME = "Dashboard"
LAST_COL = 12
MONTHS = 12
PROPERTY_ROWS = 8
LIST_ROWS = 6


def _nth_property(field: str, n: int) -> str:
    """Row n of the Property Database, blank when the row is empty."""
    return (f'IFERROR(IF(INDEX([[Property Database.name]],{n})="","",'
            f'INDEX([[Property Database.{field}]],{n})),"")')


def _by_rank(sheet: str, rank_key: str, field: str, n: int) -> str:
    """The nth row of a register by one of its hidden rank columns."""
    return (f'IFERROR(INDEX([[{sheet}.{field}]],MATCH({n},'
            f'[[{sheet}.{rank_key}]],0)),"")')


def _month_key(n: int) -> str:
    """The `yyyy-mm` key for n months before this one."""
    if n == 0:
        return 'TEXT(TodayDate,"yyyy-mm")'
    return (f'TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{n},1),'
            f'"yyyy-mm")')


def _month_sum(value: str, months: str, n: int, *extra: str) -> str:
    """One month's total out of a register, by its month key column."""
    tail = "".join(f",{c}" for c in extra)
    return f'SUMIFS({value},{months},{_month_key(n)}{tail})'


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws

    s.banner("Short-Term Rental Dashboard",
             "Bookings, turnovers, occupancy and profit - all worked "
             "out from the registers you have already filled in.",
             LAST_COL)
    s.set_col(0, 0, 24)
    for column in range(1, LAST_COL + 1):
        s.set_col(column, column, 13)

    bk_earns = '[[Booking Tracker.earns]],"Yes"'
    bk_month = "[[Booking Tracker.month]]"

    # ---------------------------------------------------------- header --
    row = 4
    s.merge(row, 0, 0, "Business", fmts.chrome("label_bold"))
    s.merge(row, 1, 3, None, fmts.calc("text"))
    s.formula(row, 1, "=HostName", fmts.calc("text"))
    s.merge(row, 4, 4, "Month", fmts.chrome("label_right"))
    s.formula(row, 5, '=TEXT(TodayDate,"mmmm yyyy")',
                     fmts.calc("text"))
    s.merge(row, 6, 6, "Properties", fmts.chrome("label_right"))
    s.formula(row, 7, "=PropertyCount", fmts.calc("int"))
    s.merge(row, 8, 8, "Today", fmts.chrome("label_right"))
    s.formula(row, 9, "=TodayDate", fmts.calc("date"))
    s.merge(row, 10, LAST_COL, None, fmts.chrome("note"))
    s.formula(
        row, 10,
        '="Figures update the moment you type. Nothing to refresh."',
        fmts.chrome("note"))
    row += 2

    # ------------------------------------------------------- tonight ----
    row = s.section(row, "Tonight and the next few days", LAST_COL)
    in_house = ('COUNTIFS([[Booking Tracker.check_in]],"<="&TodayDate,'
                '[[Booking Tracker.check_out]],">"&TodayDate,'
                f'{bk_earns})')
    arriving = ('COUNTIFS([[Booking Tracker.check_in]],TodayDate,'
                f'{bk_earns})')
    leaving = ('COUNTIFS([[Booking Tracker.check_out]],TodayDate,'
               f'{bk_earns})')
    upcoming = ('COUNTIFS([[Booking Tracker.check_in]],">"&TodayDate,'
                '[[Booking Tracker.check_in]],"<="&TodayDate+'
                f'UpcomingDays,{bk_earns})')

    s.kpi(row, 0, "Guests in tonight", f"={in_house}", "int",
          note="Stays that started on or before today and have not "
               "checked out")
    s.kpi(row, 2, "Arriving today", f"={arriving}", "int")
    s.kpi(row, 4, "Departing today", f"={leaving}", "int")
    s.kpi(row, 6, "Same-day turnovers",
          '=COUNTIFS([[Cleaning Schedule.same_day]],"Yes",'
          '[[Cleaning Schedule.clean_date]],TodayDate)', "int",
          note="Out and in on the same day - the ones that go wrong")
    s.kpi(row, 8, "Cleans due today",
          '=COUNTIFS([[Cleaning Schedule.clean_date]],TodayDate,'
          '[[Cleaning Schedule.status]],"<>Completed")', "int")
    row = s.kpi(row, 10, f"Arriving in {cfg.upcoming_window_days} days",
                f"={upcoming}", "int")

    s.kpi(row, 0, "Properties guest-ready",
          '=COUNTIFS([[Property Database.alert]],"OK*")', "int",
          note="Nothing outstanding on the property at all")
    s.kpi(row, 2, "Turnovers unfinished",
          '=COUNTIFS([[Cleaning Schedule.clean_date]],"<="&TodayDate,'
          '[[Cleaning Schedule.status]],"<>Completed")', "int")
    s.kpi(row, 4, "Pre-arrival jobs open",
          '=COUNTIFS([[Check-in & Check-out.arrival_date]],'
          '">="&TodayDate,[[Check-in & Check-out.pre_arrival_done]],'
          '"No")', "int",
          note="Access codes, directions, wifi, house rules, welcome "
               "message")
    s.kpi(row, 6, "Double bookings",
          '=COUNTIFS([[Booking Tracker.clash]],"Yes")', "int",
          note="Two paying stays overlapping in one property")
    s.kpi(row, 8, "Urgent repairs",
          '=COUNTIFS([[Maintenance Log.is_open]],"Yes",'
          '[[Maintenance Log.priority]],"Urgent")', "int")
    row = s.kpi(row, 10, "Blocking a letting",
                '=COUNTIFS([[Maintenance Log.is_open]],"Yes",'
                '[[Maintenance Log.blocks_calendar]],"Yes")', "int")
    row += 1

    # ------------------------------------------------- this month ----
    row = s.section(row, "This month", LAST_COL)
    nights_m = _month_sum("[[Booking Tracker.nights]]", bk_month, 0,
                          bk_earns)
    accom_m = _month_sum("[[Booking Tracker.accommodation]]", bk_month,
                         0, bk_earns)
    gross_m = _month_sum("[[Booking Tracker.gross]]", bk_month, 0,
                         bk_earns)
    net_m = _month_sum("[[Booking Tracker.net_revenue]]", bk_month, 0,
                       bk_earns)
    fees_m = _month_sum("[[Booking Tracker.platform_fee_total]]",
                        bk_month, 0, bk_earns)
    clean_m = _month_sum("[[Cleaning Schedule.total_cost]]",
                         "[[Cleaning Schedule.month]]", 0)
    maint_m = _month_sum("[[Maintenance Log.cost_used]]",
                         "[[Maintenance Log.month]]", 0)
    other_m = _month_sum("[[Expense Tracker.pl_amount]]",
                         "[[Expense Tracker.month]]", 0)
    costs_m = f"({clean_m}+{maint_m}+{other_m})"
    profit_m = f"({net_m}-{costs_m})"
    avail_m = ("MAX(1,PropertyCount*DAY(DATE(YEAR(TodayDate),"
               "MONTH(TodayDate)+1,1)-1))")

    s.kpi(row, 0, "Bookings",
          f'=COUNTIFS({bk_month},{_month_key(0)},{bk_earns})', "int")
    s.kpi(row, 2, "Nights sold", f"={nights_m}", "int")
    s.kpi(row, 4, "Occupancy",
          f'=IF({avail_m}<=0,"",{nights_m}/{avail_m})', "percent",
          note="Nights sold over every night every property could have "
               "sold")
    s.kpi(row, 6, "ADR",
          f'=IF({nights_m}<=0,"",ROUND({accom_m}/{nights_m},2))',
          "money", note="Accommodation only - cleaning fees excluded")
    s.kpi(row, 8, "RevPAR",
          f'=IF({avail_m}<=0,"",ROUND({accom_m}/{avail_m},2))',
          "money")
    row = s.kpi(row, 10, "Average stay",
                f'=IF(COUNTIFS({bk_month},{_month_key(0)},{bk_earns})'
                f'<=0,"",ROUND({nights_m}/COUNTIFS({bk_month},'
                f'{_month_key(0)},{bk_earns}),1))', "number1")

    s.kpi(row, 0, "Gross booked", f"=ROUND({gross_m},2)", "money0")
    s.kpi(row, 2, "Platform fees", f"=ROUND({fees_m},2)", "money0",
          note="What the platforms kept")
    s.kpi(row, 4, "Net revenue", f"=ROUND({net_m},2)", "money0")
    s.kpi(row, 6, "Running costs", f"=ROUND({costs_m},2)", "money0",
          note="Cleaning, repairs and everything on the expense tab")
    s.kpi(row, 8, "Profit", f"=ROUND({profit_m},2)", "money0")
    row = s.kpi(row, 10, "Margin",
                f'=IF({net_m}<=0,"",{profit_m}/{net_m})', "percent")
    row += 1

    # ---------------------------------------------------- the money ----
    row = s.section(row, "Money still moving", LAST_COL)
    s.kpi(row, 0, "Unpaid bookings",
          '=ROUND(SUMIFS([[Booking Tracker.gross]],'
          '[[Booking Tracker.payment_status]],"Unpaid",'
          f'{bk_earns}),2)', "money0")
    s.kpi(row, 2, "Payouts owed",
          '=ROUND(SUMIFS([[Platform Revenue & Fees.exp_net]],'
          '[[Platform Revenue & Fees.is_paid]],"No"),2)', "money0",
          note="Stays that have ended but have not landed in the bank")
    s.kpi(row, 4, "Payouts overdue",
          '=COUNTIFS([[Platform Revenue & Fees.alert]],"Chase*")',
          "int")
    s.kpi(row, 6, "Owed to cleaners",
          '=ROUND(SUMIFS([[Cleaning Schedule.clean_cost]],'
          '[[Cleaning Schedule.paid]],"<>Yes",'
          '[[Cleaning Schedule.status]],"Completed"),2)', "money0")
    s.kpi(row, 8, "Fee variance",
          '=ROUND(SUM([[Platform Revenue & Fees.variance]]),2)',
          "money0",
          note="Real payouts against the fees you assumed. A big "
               "number means a rate on Settings is wrong")
    row = s.kpi(row, 10, "Security deposits held",
                '=ROUND(SUMIFS([[Property Database.deposit]],'
                '[[Property Database.guest_in]],"Yes"),2)', "money0")
    row += 1

    # --------------------------------------------- what needs doing ----
    row = s.section(row, "What needs doing", LAST_COL)
    row = s.note(
        row,
        "Every Action column in the workbook, counted. These are the "
        "only numbers on this page you should want to be zero.",
        LAST_COL, height=24)
    row = s.table_head(row, 0, [
        "Where", "Needs doing", "Of which urgent", "The usual cause"],
        [26, 13, 15, 56])
    jobs = [
        ("Property Database", '[[Property Database.alert]]',
         "Missing", "A property with no rate, no cleaning fee or an "
                    "open job on it"),
        ("Booking Tracker", '[[Booking Tracker.alert]]',
         "Duplicate", "Overlapping dates, a missing rate, or an "
                      "unconfirmed arrival"),
        ("Check-in & Check-out", '[[Check-in & Check-out.alert]]',
         "Send", "Access codes and directions not sent yet"),
        ("Cleaning Schedule", '[[Cleaning Schedule.alert]]',
         "Not cleaned", "A turnover due today or overdue"),
        ("Platform Revenue & Fees", '[[Platform Revenue & Fees.alert]]',
         "Chase", "A payout later than your platform's usual terms"),
        ("Supplies Inventory", '[[Supplies Inventory.alert]]',
         "Out of stock", "Below the minimum, or not enough to cover "
                         "the turnovers booked"),
        ("Maintenance Log", '[[Maintenance Log.alert]]',
         "Urgent", "An open job, or one blocking a letting"),
        ("Reviews Tracker", '[[Reviews Tracker.alert]]',
         "Action", "A poor review with no reply"),
        ("Staff & Vendors", '[[Staff & Vendors.alert]]',
         "Expired", "Lapsed insurance or an unpaid cleaner"),
    ]
    for label, column, urgent, cause in jobs:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.formula(
            row, 1,
            f'=COUNTIF({column},"?*")-COUNTIF({column},"OK*")',
            fmts.calc("int"))
        s.formula(row, 2, f'=COUNTIF({column},"{urgent}*")',
                         fmts.calc("int"))
        s.merge(row, 3, LAST_COL, cause, fmts.chrome("note"))
        row += 1
    row += 1

    # -------------------------------------------- property comparison --
    row = s.section(row, "Every property side by side", LAST_COL)
    row = s.note(
        row,
        "The whole portfolio on one line each, for the life of the "
        "file. Sort your attention by the last column: a property can "
        "be busy, well reviewed and still be the one losing money.",
        LAST_COL, height=26)
    row = s.table_head(row, 0, [
        "Property", "Nights", "Occupancy", "ADR", "RevPAR", "Net revenue",
        "Cleaning", "Repairs", "Other", "Profit", "Margin",
        "Profit a night", "Rating"],
        [24, 9, 10, 10, 10, 12, 10, 9, 9, 11, 9, 12, 8])
    for n in range(1, PROPERTY_ROWS + 1):
        s.write(row, 0, None, fmts.calc("text"))
        s.formula(row, 0, "=" + _nth_property("name", n),
                         fmts.calc("text"))
        for offset, (field, style) in enumerate((
                ("nights_booked", "int"), ("occupancy", "percent"),
                ("adr", "money"), ("revpar", "money"),
                ("net_revenue", "money0"), ("cleaning_spend", "money0"),
                ("maintenance_spend", "money0"),
                ("other_expenses", "money0"), ("profit", "money0"),
                ("margin", "percent"), ("profit_night", "money"),
                ("rating", "number"))):
            s.formula(row, offset + 1,
                             "=" + _nth_property(field, n),
                             fmts.calc(style))
        row += 1
    s.write(row, 0, "Portfolio", fmts.chrome("total_label"))
    total_nights = 'SUM([[Property Database.nights_booked]])'
    total_avail = 'SUM([[Property Database.nights_available]])'
    total_accom = 'SUM([[Property Database.accommodation]])'
    total_net = 'SUM([[Property Database.net_revenue]])'
    for offset, formula in enumerate((
            f"={total_nights}",
            f'=IF({total_avail}<=0,"",{total_nights}/{total_avail})',
            f'=IF({total_nights}<=0,"",ROUND({total_accom}/'
            f'{total_nights},2))',
            f'=IF({total_avail}<=0,"",ROUND({total_accom}/'
            f'{total_avail},2))',
            f"=ROUND({total_net},2)",
            '=ROUND(SUM([[Property Database.cleaning_spend]]),2)',
            '=ROUND(SUM([[Property Database.maintenance_spend]]),2)',
            '=ROUND(SUM([[Property Database.other_expenses]]),2)',
            '=ROUND(SUM([[Property Database.profit]]),2)',
            f'=IF({total_net}<=0,"",SUM([[Property Database.profit]])/'
            f'{total_net})',
            f'=IF({total_nights}<=0,"",ROUND(SUM('
            f'[[Property Database.profit]])/{total_nights},2))',
            '=IF(COUNTIFS([[Reviews Tracker.has_score]],"Yes")<=0,"",'
            'ROUND(SUMIFS([[Reviews Tracker.rate_overall]],'
            '[[Reviews Tracker.has_score]],"Yes")/'
            'COUNTIFS([[Reviews Tracker.has_score]],"Yes"),2))')):
        style = ("int", "percent", "money", "money", "money0", "money0",
                 "money0", "money0", "money0", "percent", "money",
                 "number")[offset]
        s.formula(row, offset + 1, formula,
                         fmts.chrome("total_value",
                                     num_format=s.num_format(style)))
    row += 2

    # ---------------------------------------------- twelve months ------
    row = s.section(row, "The last twelve months", LAST_COL)
    row = s.table_head(row, 0, [
        "Month", "Bookings", "Nights", "Occupancy", "ADR", "Net revenue",
        "Costs", "Profit", "Margin", "Against target"],
        [12, 10, 9, 10, 10, 12, 11, 11, 9, 13])
    first_month_row = row
    for back in range(MONTHS - 1, -1, -1):
        key = _month_key(back)
        nights = _month_sum("[[Booking Tracker.nights]]", bk_month,
                            back, bk_earns)
        accom = _month_sum("[[Booking Tracker.accommodation]]",
                           bk_month, back, bk_earns)
        net = _month_sum("[[Booking Tracker.net_revenue]]", bk_month,
                         back, bk_earns)
        costs = ("(" + _month_sum("[[Cleaning Schedule.total_cost]]",
                                  "[[Cleaning Schedule.month]]", back)
                 + "+" + _month_sum("[[Maintenance Log.cost_used]]",
                                    "[[Maintenance Log.month]]", back)
                 + "+" + _month_sum("[[Expense Tracker.pl_amount]]",
                                    "[[Expense Tracker.month]]", back)
                 + ")")
        avail = ("MAX(1,PropertyCount*DAY(DATE(YEAR(TodayDate),"
                 f"MONTH(TodayDate)-{back}+1,1)-1))")
        s.formula(
            row, 0,
            f'=TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{back},1),'
            f'"mmm yy")', fmts.calc("text"))
        for offset, (formula, style) in enumerate((
                (f'=COUNTIFS({bk_month},{key},{bk_earns})', "int"),
                (f"={nights}", "int"),
                (f'=IF({avail}<=0,"",{nights}/{avail})', "percent"),
                (f'=IF({nights}<=0,"",ROUND({accom}/{nights},2))',
                 "money"),
                (f"=ROUND({net},2)", "money0"),
                (f"=ROUND({costs},2)", "money0"),
                (f"=ROUND({net}-{costs},2)", "money0"),
                (f'=IF({net}<=0,"",({net}-{costs})/{net})', "percent"),
                (f'=IF(TargetRevenueMonth<=0,"",{net}/'
                 f'TargetRevenueMonth)', "percent"))):
            s.formula(row, offset + 1, formula, fmts.calc(style))
        row += 1
    last_month_row = row - 1

    ws.conditional_format(
        first_month_row, 3, last_month_row, 3,
        {"type": "cell", "criteria": ">=", "value": "TargetOccupancy",
         "format": fmts.cf("good")})
    ws.conditional_format(
        first_month_row, 7, last_month_row, 7,
        {"type": "cell", "criteria": "<", "value": 0,
         "format": fmts.cf("bad")})
    ws.conditional_format(
        first_month_row, 9, last_month_row, 9,
        {"type": "cell", "criteria": ">=", "value": 1,
         "format": fmts.cf("good")})
    ws.conditional_format(
        first_month_row, 9, last_month_row, 9,
        {"type": "cell", "criteria": "<", "value": 0.8,
         "format": fmts.cf("warn")})
    ws.add_sparkline(last_month_row + 1, 5, {
        "range": (f"'{NAME}'!F{first_month_row + 1}:"
                  f"F{last_month_row + 1}"),
        "type": "column", "style": 12})
    s.write(last_month_row + 1, 0, "Net revenue trend",
            fmts.chrome("label_bold"))
    ws.add_sparkline(last_month_row + 1, 7, {
        "range": (f"'{NAME}'!H{first_month_row + 1}:"
                  f"H{last_month_row + 1}"),
        "type": "column", "style": 12,
        "negative_points": True})
    row = last_month_row + 3

    # ----------------------------------------------- the short lists ---
    row = s.section(row, "Worth knowing", LAST_COL)

    s.write(row, 0, "Best performing properties",
            fmts.chrome("subsection"))
    s.write(row, 4, "Guests worth looking after",
            fmts.chrome("subsection"))
    s.write(row, 8, "Arriving next", fmts.chrome("subsection"))
    row += 1
    s.table_head(row, 0, ["Property", "Profit", "Occupancy"],
                 [22, 11, 10])
    s.table_head(row, 4, ["Guest", "Stays", "Net"], [20, 8, 11])
    s.table_head(row, 8, ["Guest", "Property", "Arrives", "Nights"],
                 [18, 18, 11, 8])
    row += 1
    for n in range(1, LIST_ROWS + 1):
        s.formula(
            row, 0,
            "=" + _by_rank("Property Database", "profit_rank", "name", n),
            fmts.calc("text"))
        s.formula(
            row, 1,
            "=" + _by_rank("Property Database", "profit_rank", "profit",
                           n), fmts.calc("money0"))
        s.formula(
            row, 2,
            "=" + _by_rank("Property Database", "profit_rank",
                           "occupancy", n), fmts.calc("percent"))
        s.formula(
            row, 4,
            "=" + _by_rank("Guest Database", "value_rank", "name", n),
            fmts.calc("text"))
        s.formula(
            row, 5,
            "=" + _by_rank("Guest Database", "value_rank", "stays", n),
            fmts.calc("int"))
        s.formula(
            row, 6,
            "=" + _by_rank("Guest Database", "value_rank", "net", n),
            fmts.calc("money0"))
        for offset, (field, style) in enumerate((
                ("guest_name", "text"), ("property_name", "text"),
                ("check_in", "date"), ("nights", "int"))):
            s.formula(
                row, 8 + offset,
                f'=IFERROR(INDEX([[Booking Tracker.{field}]],MATCH('
                f'{n},[[Booking Tracker.arrival_rank]],0)),"")',
                fmts.calc(style))
        row += 1
    row += 1

    row = s.note(
        row,
        "Nothing on this page is typed. Every figure traces back to a "
        "register through the property ID and the booking ID, which "
        "is why adding a fifth property or a third platform changes "
        "the dashboard without anyone editing a formula.",
        LAST_COL, height=28)

    s.finish(freeze_row=4, gridlines=False, fit_width=True,
             print_area=f"A1:{col_letter(LAST_COL)}{row + 2}")
    s.protect()
    return s
