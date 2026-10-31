"""`Dashboard` - the one page that answers "how is the setting doing?"

Twenty numbers across the top, the jobs that need doing before you
take your coat off, every room with its occupancy and its ratio, the
families who owe you money, and twelve months of income against
target.  Everything reads from the registers, so it is never out of
date and there is nothing to refresh.

Four questions, in the order a childcare owner actually asks them:
are the rooms full, are the children here, are we getting paid, and
is anything about to go wrong.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter

NAME = "Dashboard"
LAST_COL = 12
MONTHS = 12
ROOM_ROWS = 8
TOP_ROWS = 6


def _by_rank(sheet: str, rank_key: str, field: str, n: int) -> str:
    """The nth row of a register by one of its hidden rank columns."""
    return (f'IFERROR(INDEX([[{sheet}.{field}]],MATCH({n},'
            f'[[{sheet}.{rank_key}]],0)),"")')


def _nth(sheet: str, field: str, n: int, style: str = "text") -> str:
    """Row n of a register, blank rather than zero when it is empty."""
    blank = '""' if style == "text" else '""'
    return (f'IFERROR(IF(INDEX([[{sheet}.name]],{n})="",{blank},'
            f'INDEX([[{sheet}.{field}]],{n})),{blank})')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws
    cur = cfg.currency_symbol

    s.banner("Childcare Dashboard",
             "Enrollment, attendance, money and the jobs that need "
             "doing - all worked out from the registers you have "
             "already filled in.", LAST_COL)
    s.set_col(0, 0, 26)
    for column in range(1, LAST_COL + 1):
        s.set_col(column, column, 13)

    # --------------------------------------------------------- the header --
    row = 4
    s.merge(row, 0, 0, "Setting", fmts.chrome("label_bold"))
    s.merge(row, 1, 3, None, fmts.calc("text"))
    ws.write_formula(row, 1, "=CentreName", fmts.calc("text"))
    s.merge(row, 4, 4, "Month", fmts.chrome("label_right"))
    ws.write_formula(row, 5, '=TEXT(TodayDate,"mmmm yyyy")',
                     fmts.calc("text"))
    s.merge(row, 6, 6, "Month key", fmts.chrome("label_right"))
    ws.write_formula(row, 7, '=TEXT(TodayDate,"yyyy-mm")',
                     fmts.calc("text"))
    book.define_name("DashKey", f"={quote_sheet(NAME)}!$H${row + 1}")
    s.merge(row, 8, 9, "Last month", fmts.chrome("label_right"))
    ws.write_formula(row, 10,
                     '=TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-1,1),'
                     '"yyyy-mm")', fmts.calc("text"))
    book.define_name("DashPrevKey", f"={quote_sheet(NAME)}!$K${row + 1}")
    s.merge(row, 11, 11, "Today", fmts.chrome("label_right"))
    s.merge(row, 12, LAST_COL, None, fmts.calc("date"))
    ws.write_formula(row, 12, "=TodayDate", fmts.calc("date"))
    row += 2

    # ----------------------------------------------------------- the KPIs --
    enrollment = [
        ("Children on the books",
         "=ChildrenEnrolled", "int",
         "active, on hold or on notice"),
        ("Occupancy",
         '=IF(RoomCapacity<=0,"",ChildrenEnrolled/RoomCapacity)',
         "percent", f"target {cfg.target_occupancy:.0%}"),
        ("Places free",
         "=MAX(0,RoomCapacity-ChildrenEnrolled)", "int",
         "across every room"),
        ("Waiting for a place",
         '=COUNTIFS([[Waitlist.closed]],"No")', "int",
         "live enquiries, not yet offered or lost"),
    ]
    attendance = [
        ("Here today",
         '=COUNTIFS([[Attendance.date]],TodayDate,'
         '[[Attendance.present]],1)', "int",
         "checked in on the attendance tab"),
        ("Attendance this month",
         '=IF(SUMIFS([[Attendance.session]],[[Attendance.month]],'
         'DashKey)<=0,"",SUMIFS([[Attendance.present]],'
         '[[Attendance.month]],DashKey)/'
         'SUMIFS([[Attendance.session]],[[Attendance.month]],DashKey))',
         "percent", f"target {cfg.target_attendance:.0%}"),
        ("Absences this month",
         '=SUMIFS([[Attendance.absent]],[[Attendance.month]],DashKey)',
         "int", "sick, family, holiday or unexplained"),
        ("Still to check out",
         '=COUNTIF([[Attendance.alert]],"Missing - no check-out*")',
         "int", "a session with no collection time"),
    ]
    financial = [
        ("Billed this month",
         '=ROUND(SUMIFS([[Invoices & Receipts.billed]],'
         '[[Invoices & Receipts.month]],DashKey),2)', "money0",
         f"target {cur}{cfg.revenue_target_month:,.0f}"),
        ("Against target",
         '=IF(RevenueTargetMonth<=0,"",'
         'SUMIFS([[Invoices & Receipts.billed]],'
         '[[Invoices & Receipts.month]],DashKey)/RevenueTargetMonth)',
         "percent0", "100% means the month is made"),
        ("Collected this month",
         '=ROUND(SUMIFS([[Invoices & Receipts.collected]],'
         '[[Invoices & Receipts.month]],DashKey),2)', "money0",
         "money that actually arrived"),
        ("Outstanding, all families",
         '=ROUND(SUM([[Parent Database.balance]]),2)', "money0",
         "billed and not yet paid"),
    ]
    operating = [
        ("Operating profit, month",
         "=ROUND(MonthBilled-MonthWages-MonthExpenses,2)", "money0",
         "billed less wages less costs"),
        ("Profit margin",
         '=IF(MonthBilled<=0,"",(MonthBilled-MonthWages-MonthExpenses)'
         '/MonthBilled)', "percent",
         f"target {cfg.profit_target_month / max(1.0, cfg.revenue_target_month):.0%}"
         " or better"),
        ("Wage cost of income",
         '=IF(MonthBilled<=0,"",MonthWages/MonthBilled)', "percent",
         "most settings land between 55% and 70%"),
        ("Income per child, week",
         '=IF(ChildrenEnrolled<=0,"",'
         'ROUND(SUM([[Child Enrollment.weekly_net]])/ChildrenEnrolled,'
         '2))', "money", "after discounts and subsidies"),
    ]
    people = [
        ("Staff on the rota today",
         '=COUNTIFS([[Staff Schedule.date]],TodayDate,'
         '[[Staff Schedule.counts]],"Yes")', "int",
         "counted towards ratio"),
        ("Rooms over ratio",
         '=COUNTIF([[Classroom Management.ratio_ok]],"No")', "int",
         "get cover into these first"),
        ("Documents expired or missing",
         '=COUNTIF([[Document Tracker.state]],"Expired")+'
         'COUNTIF([[Document Tracker.state]],"Missing")', "int",
         "required paperwork only"),
        ("Jobs waiting on you",
         '=COUNTIFS([[Parent Communication.status]],"Waiting on us")+'
         'COUNTIF([[Parent Communication.overdue]],"Yes")', "int",
         "parent messages you have not answered"),
    ]

    for band in (enrollment, attendance, financial, operating, people):
        column = 0
        bottom = row
        for label, formula, style, note in band:
            bottom = s.kpi(row, column, label, formula, style, span=3,
                           note=note)
            column += 3
        row = bottom
    row += 1

    # ------------------------------------------------- needs doing today ---
    row = s.section(row, "Needs doing today", LAST_COL)
    row = s.table_head(row, 0, ["What", "How many", "What it means",
                                "Where to fix it"],
                       [28, 11, 50, 18])
    alerts = [
        ("Rooms over ratio",
         '=COUNTIF([[Classroom Management.ratio_ok]],"No")',
         '="More children in the room than your own ratio allows. '
         'Move an adult in before anything else today"',
         "Classroom Management"),
        ("Rooms nearly full",
         '=COUNTIF([[Classroom Management.alert]],"Nearly full*")',
         '="At or above "&TEXT(RatioWarning,"0%")&" of licensed '
         'places. Start offering the waitlist"',
         "Classroom Management"),
        ("Children with no check-out",
         '=COUNTIF([[Attendance.alert]],"Missing - no check-out*")',
         '="A past session with a check-in and no check-out"',
         "Attendance"),
        ("Collected by somebody unexpected",
         '=COUNTIF([[Attendance.alert]],"Not authorised*")',
         '="The name in Collected by is not on that child'
         "'" 's authorised list"',
         "Attendance"),
        ("Unexplained absences to chase",
         '=COUNTIF([[Attendance.alert]],"Unexplained*")',
         '="Marked absent with no reason written down"',
         "Attendance"),
        ("Late pickups not charged",
         '=COUNTIF([[Attendance.alert]],"Invoice - late pickup*")',
         '="Collected after the grace period with no fee raised"',
         "Tuition & Fees"),
        ("Charges not on an invoice",
         '=COUNTIFS([[Tuition & Fees.billed]],"No")',
         '="Work you have done that nobody has been billed for"',
         "Tuition & Fees"),
        ("Invoices overdue",
         '=COUNTIFS([[Invoices & Receipts.status]],"Overdue")',
         '="Past the due date and still showing a balance"',
         "Invoices & Receipts"),
        ("Invoices still a draft",
         '=COUNTIFS([[Invoices & Receipts.status]],"Draft")',
         '="Raised but never marked as sent to the parent"',
         "Invoices & Receipts"),
        ("Families with a balance",
         '=COUNTIFS([[Parent Database.balance]],">0.004")',
         '="Owing something, overdue or not"', "Parent Database"),
        ("Parent messages waiting on you",
         '=COUNTIFS([[Parent Communication.status]],"Waiting on us")',
         '="You said you would come back to them"',
         "Parent Communication"),
        ("Follow-ups past their date",
         '=COUNTIF([[Parent Communication.overdue]],"Yes")',
         '="A follow-up date that has already gone"',
         "Parent Communication"),
        ("Incident forms not signed",
         '=COUNTIF([[Incident Reports.alert]],"Chase - parent*")',
         '="Parent told, but the form has not come back signed"',
         "Incident Reports"),
        ("Incidents with no parent told",
         '=COUNTIF([[Incident Reports.alert]],"Not notified*")',
         '="Somebody has to make that call today"',
         "Incident Reports"),
        ("Medication without written consent",
         '=COUNTIF([[Allergy & Medical.alert]],"Not authorised*")',
         '="A dose logged with no signed consent behind it"',
         "Allergy & Medical"),
        ("Children missing paperwork",
         '=COUNTIFS([[Child Enrollment.docs_missing]],">0")',
         '="Required documents not yet on file"',
         "Document Tracker"),
        ("Documents expiring soon",
         '=COUNTIF([[Document Tracker.state]],"Expiring")',
         '="Runs out inside "&DocWarningDays&" days"',
         "Document Tracker"),
        ("Documents already expired",
         '=COUNTIF([[Document Tracker.state]],"Expired")',
         '="Including your own licence, insurance and inspections"',
         "Document Tracker"),
        ("Staff certificates expired",
         '=COUNTIFS([[Staff Database.certs_ok]],"Expired")',
         '="First aid, CPR, safeguarding or background check"',
         "Staff Database"),
        ("Staff certificates expiring",
         '=COUNTIFS([[Staff Database.certs_ok]],"Expiring")',
         '="Book the course before it lapses"', "Staff Database"),
        ("Shifts double booked",
         '=COUNTIF([[Staff Schedule.alert]],"Duplicate*")',
         '="The same person in two places at once"',
         "Staff Schedule"),
        ("Pay runs not paid",
         '=COUNTIFS([[Payroll Tracker.paid]],"<>Yes",'
         '[[Payroll Tracker.period_end]],"<>")',
         '="Calculated and still open"', "Payroll Tracker"),
        ("Supplies out of stock",
         '=COUNTIF([[Supply Inventory.status]],"Out of stock")',
         '="Nothing left in the cupboard"', "Supply Inventory"),
        ("Supplies to reorder",
         '=COUNTIF([[Supply Inventory.status]],"Low stock")',
         '="Below the minimum or under "&StockWarningWeeks&" weeks '
         'of cover"', "Supply Inventory"),
        ("Birthdays coming up",
         '=COUNTIF([[Child Enrollment.alert]],"Birthday*")',
         '="Inside the next "&BirthdayWindow&" days"',
         "Child Enrollment"),
        ("Children moving room soon",
         '=COUNTIF([[Child Enrollment.alert]],"Moves up*")',
         '="They age out of their band within a month - tell the '
         'parents before they notice"', "Child Enrollment"),
        ("Waitlist enquiries going quiet",
         '=COUNTIF([[Waitlist.alert]],"Chase*")+'
         'COUNTIF([[Waitlist.alert]],"Follow up*")',
         '="Nobody has spoken to them in a fortnight"', "Waitlist"),
        ("Seats to offer the waitlist",
         '=COUNTIF([[Waitlist.alert]],"Space*")',
         '="A place has opened in the room they asked for"',
         "Waitlist"),
    ]
    first_alert = row
    for label, count, message, where in alerts:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        ws.write_formula(row, 1, s.reg.render(count), fmts.calc("int"))
        s.merge(row, 2, LAST_COL - 3, None, fmts.calc("text"))
        ws.write_formula(row, 2, s.reg.render(message),
                         fmts.calc("text"))
        s.merge(row, LAST_COL - 2, LAST_COL, where, fmts.chrome("label"))
        row += 1
    ws.conditional_format(first_alert, 1, row - 1, 1,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("warn")})
    ws.conditional_format(first_alert, 1, row - 1, 1,
                          {"type": "cell", "criteria": "=", "value": 0,
                           "format": fmts.cf("good")})
    row += 1

    # ------------------------------------------------------- the rooms ----
    row = s.section(row, "Every room, right now", LAST_COL)
    row = s.table_head(row, 0, [
        "Room", "Age band", "Places", "Enrolled", "Free", "Occupancy",
        "Here today", "Adults today", "Adults needed", "Ratio now",
        "Within ratio", "Weekly income", "Empty seats cost"],
        [18, 15, 9, 10, 8, 11, 11, 12, 13, 11, 12, 13, 14])
    rooms_first = row
    for n in range(1, ROOM_ROWS + 1):
        cells = [
            (0, _nth("Classroom Management", "name", n), "text"),
            (1, _nth("Classroom Management", "age_band", n), "text"),
            (2, _nth("Classroom Management", "capacity", n), "int"),
            (3, _nth("Classroom Management", "enrolled", n), "int"),
            (4, _nth("Classroom Management", "seats", n), "int"),
            (5, _nth("Classroom Management", "occupancy", n),
             "percent"),
            (6, _nth("Classroom Management", "present_today", n),
             "int"),
            (7, _nth("Classroom Management", "staff_today", n),
             "number1"),
            (8, _nth("Classroom Management", "staff_needed", n),
             "number1"),
            (9, _nth("Classroom Management", "ratio_now", n), "text"),
            (10, _nth("Classroom Management", "ratio_ok", n), "text"),
            (11, _nth("Classroom Management", "revenue", n), "money"),
            (12, _nth("Classroom Management", "empty_cost", n),
             "money"),
        ]
        for column, formula, style in cells:
            ws.write_formula(row, column, s.reg.render("=" + formula),
                             fmts.calc(style))
        row += 1
    ws.conditional_format(rooms_first, 5, row - 1, 5,
                          {"type": "cell", "criteria": ">=",
                           "value": "=RatioWarning",
                           "format": fmts.cf("warn")})
    ws.conditional_format(rooms_first, 10, row - 1, 10,
                          {"type": "text", "criteria": "containing",
                           "value": "No", "format": fmts.cf("bad")})
    ws.conditional_format(rooms_first, 10, row - 1, 10,
                          {"type": "text", "criteria": "containing",
                           "value": "Yes", "format": fmts.cf("good")})
    row += 1

    # --------------------------------------------------- money and people --
    row = s.section(row, "Who owes you, and who earns it", LAST_COL)
    row = s.table_head(row, 0, [
        "Family", "Children", "Weekly fees", "Billed", "Paid",
        "Balance", "Overdue", "Oldest bill", "Status",
        "Biggest fee payer", "A year", "Room"],
        [22, 9, 12, 12, 12, 12, 12, 12, 16, 22, 12, 13])
    money_first = row
    for n in range(1, TOP_ROWS + 1):
        cells = [
            (0, _by_rank("Parent Database", "weekly_fees_rank",
                         "family_name", n), "text"),
            (1, _by_rank("Parent Database", "weekly_fees_rank",
                         "children", n), "int"),
            (2, _by_rank("Parent Database", "weekly_fees_rank",
                         "weekly_fees", n), "money"),
            (3, _by_rank("Parent Database", "weekly_fees_rank",
                         "invoiced", n), "money"),
            (4, _by_rank("Parent Database", "weekly_fees_rank",
                         "paid", n), "money"),
            (5, _by_rank("Parent Database", "weekly_fees_rank",
                         "balance", n), "money"),
            (6, _by_rank("Parent Database", "weekly_fees_rank",
                         "overdue", n), "money"),
            (7, _by_rank("Parent Database", "weekly_fees_rank",
                         "oldest_due", n), "date"),
            (8, _by_rank("Parent Database", "weekly_fees_rank",
                         "status", n), "text"),
            (9, _by_rank("Child Enrollment", "annual_net_rank", "name",
                         n), "text"),
            (10, _by_rank("Child Enrollment", "annual_net_rank",
                          "annual_net", n), "money"),
            (11, _by_rank("Child Enrollment", "annual_net_rank", "room",
                          n), "text"),
        ]
        for column, formula, style in cells:
            ws.write_formula(row, column, s.reg.render("=" + formula),
                             fmts.calc(style))
        row += 1
    ws.conditional_format(money_first, 6, row - 1, 6,
                          {"type": "cell", "criteria": ">",
                           "value": 0.004, "format": fmts.cf("bad")})
    row += 1

    # ------------------------------------------------- twelve months ------
    row = s.section(row, "The last twelve months", LAST_COL)
    row = s.table_head(row, 0, [
        "Month", "Billed", "Collected", "Target", "Against target",
        "Wages", "Other costs", "Operating profit", "Margin",
        "Children billed", "Sessions", "Attendance", "Income a child"],
        [12, 13, 13, 12, 13, 12, 12, 15, 10, 13, 10, 11, 13])
    chart_first = row
    for index in range(MONTHS - 1, -1, -1):
        key = (f'TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{index},1),'
               f'"yyyy-mm")')
        label = (f'=TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{index},'
                 f'1),"mmm yy")')
        billed = (f'SUMIFS([[Invoices & Receipts.billed]],'
                  f'[[Invoices & Receipts.month]],{key})')
        wages = (f'SUMIFS([[Payroll Tracker.total_cost]],'
                 f'[[Payroll Tracker.month]],{key})')
        costs = (f'SUMIFS([[Expense Tracker.pl_amount]],'
                 f'[[Expense Tracker.month]],{key})')
        sessions = (f'SUMIFS([[Attendance.session]],'
                    f'[[Attendance.month]],{key})')
        present = (f'SUMIFS([[Attendance.present]],'
                   f'[[Attendance.month]],{key})')
        s.write(row, 0, None, fmts.calc("text"))
        ws.write_formula(row, 0, label, fmts.calc("text"))
        formulas = [
            (1, f"=ROUND({billed},2)", "money"),
            (2, f'=ROUND(SUMIFS([[Invoices & Receipts.collected]],'
                f'[[Invoices & Receipts.month]],{key}),2)', "money"),
            (3, "=ROUND(RevenueTargetMonth,2)", "money"),
            (4, f"=ROUND({billed}-RevenueTargetMonth,2)", "money"),
            (5, f"=ROUND({wages},2)", "money"),
            (6, f"=ROUND({costs},2)", "money"),
            (7, f"=ROUND({billed}-{wages}-{costs},2)", "money"),
            (8, f'=IF({billed}<=0,"",({billed}-{wages}-{costs})/'
                f'{billed})', "percent"),
            (9, f'=COUNTIFS([[Invoices & Receipts.month]],{key},'
                f'[[Invoices & Receipts.type]],"Invoice")', "int"),
            (10, f"={sessions}", "int"),
            (11, f'=IF({sessions}<=0,"",{present}/{sessions})',
             "percent"),
            (12, f'=IF(ChildrenEnrolled<=0,"",ROUND({billed}/'
                 f'ChildrenEnrolled,2))', "money"),
        ]
        for column, formula, style in formulas:
            ws.write_formula(row, column, s.reg.render(formula),
                             fmts.calc(style))
        row += 1
    last_month_row = row - 1
    ws.conditional_format(chart_first, 1, last_month_row, 1,
                          {"type": "data_bar",
                           "bar_color": Palette.primary_mid})
    ws.conditional_format(chart_first, 7, last_month_row, 7,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(chart_first, 4, last_month_row, 4,
                          {"type": "cell", "criteria": ">=", "value": 0,
                           "format": fmts.cf("good")})
    ws.conditional_format(chart_first, 11, last_month_row, 11,
                          {"type": "cell", "criteria": "<",
                           "value": "=TargetAttendance",
                           "format": fmts.cf("warn")})

    # sparklines beat a chart Excel 2010 might refuse to draw
    ws.add_sparkline(row, 1, {
        "range": (f"{quote_sheet(NAME)}!$B${chart_first + 1}:"
                  f"$B${last_month_row + 1}"),
        "type": "column", "series_color": Palette.primary_mid})
    s.merge(row, 2, 3, "Billed, twelve months", fmts.chrome("note"))
    ws.add_sparkline(row, 4, {
        "range": (f"{quote_sheet(NAME)}!$H${chart_first + 1}:"
                  f"$H${last_month_row + 1}"),
        "type": "column", "series_color": Palette.accent,
        "negative_points": True})
    s.merge(row, 5, 7, "Operating profit, twelve months",
            fmts.chrome("note"))
    ws.add_sparkline(row, 8, {
        "range": (f"{quote_sheet(NAME)}!$L${chart_first + 1}:"
                  f"$L${last_month_row + 1}"),
        "type": "line", "series_color": Palette.deep,
        "markers": True})
    s.merge(row, 9, LAST_COL, "Attendance, twelve months",
            fmts.chrome("note"))
    row += 2

    # ---------------------------------------------------- the month in one -
    row = s.section(row, "This month in one line", LAST_COL)
    row = s.table_head(row, 0, [
        "Billed", "Collected", "Still owed", "Wages", "Other costs",
        "Operating profit", "Margin", "Break-even income",
        "Cost a child, month"],
        [14, 14, 13, 13, 13, 15, 10, 16, 16])
    totals_row = row
    ws.write_formula(
        totals_row, 0,
        s.reg.render('=ROUND(SUMIFS([[Invoices & Receipts.billed]],'
                     '[[Invoices & Receipts.month]],DashKey),2)'),
        fmts.calc("money"))
    book.define_name("MonthBilled",
                     f"={quote_sheet(NAME)}!$A${totals_row + 1}")
    ws.write_formula(
        totals_row, 1,
        s.reg.render('=ROUND(SUMIFS([[Invoices & Receipts.collected]],'
                     '[[Invoices & Receipts.month]],DashKey),2)'),
        fmts.calc("money"))
    book.define_name("MonthCollected",
                     f"={quote_sheet(NAME)}!$B${totals_row + 1}")
    ws.write_formula(
        totals_row, 2,
        s.reg.render("=ROUND(SUM([[Parent Database.balance]]),2)"),
        fmts.calc("money"))
    ws.write_formula(
        totals_row, 3,
        s.reg.render('=ROUND(SUMIFS([[Payroll Tracker.total_cost]],'
                     '[[Payroll Tracker.month]],DashKey),2)'),
        fmts.calc("money"))
    book.define_name("MonthWages",
                     f"={quote_sheet(NAME)}!$D${totals_row + 1}")
    ws.write_formula(
        totals_row, 4,
        s.reg.render('=ROUND(SUMIFS([[Expense Tracker.pl_amount]],'
                     '[[Expense Tracker.month]],DashKey),2)'),
        fmts.calc("money"))
    book.define_name("MonthExpenses",
                     f"={quote_sheet(NAME)}!$E${totals_row + 1}")
    ws.write_formula(totals_row, 5,
                     "=ROUND(MonthBilled-MonthWages-MonthExpenses,2)",
                     fmts.calc("money"))
    book.define_name("MonthProfit",
                     f"={quote_sheet(NAME)}!$F${totals_row + 1}")
    ws.write_formula(totals_row, 6,
                     '=IF(MonthBilled<=0,"",MonthProfit/MonthBilled)',
                     fmts.calc("percent"))
    ws.write_formula(totals_row, 7,
                     "=ROUND(MonthWages+MonthExpenses,2)",
                     fmts.calc("money"))
    ws.write_formula(totals_row, 8,
                     '=IF(ChildrenEnrolled<=0,"",'
                     'ROUND((MonthWages+MonthExpenses)/'
                     'ChildrenEnrolled,2))', fmts.calc("money"))
    ws.conditional_format(totals_row, 5, totals_row, 5,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    row = totals_row + 2

    row = s.note(
        row,
        "Operating profit here is what you billed for the month, less "
        "the full cost of the pay runs, less the expenses logged "
        "against it. Wages are counted once - from the Payroll "
        "Tracker - and wage rows on the Expense Tracker are left out "
        "unless you switch that on in Settings. It is not a tax "
        "computation and it knows nothing about depreciation, "
        "drawings or loans: it is the number that tells you whether "
        "the month worked.",
        LAST_COL, height=46)
    row += 1
    row = s.note(
        row,
        "Ratios on this page come from the rota, not from where staff "
        "are usually based: a shift only counts towards a room when "
        "it has a room against it and is not marked Off, Leave, "
        "Training or Admin. The bands and the numbers themselves are "
        "yours to set on the Settings tab - they are not anybody's "
        "licensing rules and this workbook does not make you "
        "compliant with them.",
        LAST_COL, height=42)

    s.finish(freeze_row=6, gridlines=False, fit_width=True)
    s.protect()
    return s
