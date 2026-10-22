"""`Dashboard` - the one page that answers "how is the salon doing?"

Sixteen numbers across the top, the jobs that need doing today, who
is earning what, which services pay best, and twelve months of
revenue against target. Everything reads from the registers, so it is
never out of date and there is nothing to refresh.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter

NAME = "Dashboard"
LAST_COL = 12
MONTHS = 12
TOP_ROWS = 6


def _by_rank(sheet: str, rank_key: str, field: str, n: int) -> str:
    """The nth row of a register by one of its hidden rank columns."""
    return (f'IFERROR(INDEX([[{sheet}.{field}]],MATCH({n},'
            f'[[{sheet}.{rank_key}]],0)),"")')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws
    money = cfg.currency_symbol

    s.banner("Salon Dashboard",
             "Takings, bookings, clients and chairs - worked out from "
             "the diary, the till and the payroll you have already "
             "filled in.", LAST_COL)
    s.set_col(0, 0, 26)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 13)

    # ----------------------------------------------------------- the header
    row = 4
    s.merge(row, 0, 0, "Salon", fmts.chrome("label_bold"))
    s.merge(row, 1, 3, None, fmts.calc("text"))
    ws.write_formula(row, 1, "=SalonName", fmts.calc("text"))
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

    # ------------------------------------------------------------- the KPIs
    tiles_1 = [
        ("Net sales this month",
         '=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],DashKey),2)',
         "money0", f"target {money}{cfg.revenue_target_month:,.0f}"),
        ("Against target",
         '=IFERROR(SUMIFS([[Sales.net]],[[Sales.month]],DashKey)/'
         'MAX(0.01,RevenueTargetMonth),"")', "percent0",
         "100% means the month is made"),
        ("Operating profit, month",
         "=ROUND(MonthNetSales-MonthCostOfSales-MonthWages-"
         "MonthExpenses,2)", "money0",
         "net sales less stock, wages and costs"),
        ("Profit margin",
         '=IF(MonthNetSales<=0,"",(MonthNetSales-MonthCostOfSales-'
         'MonthWages-MonthExpenses)/MonthNetSales)', "percent",
         "what is left of every dollar"),
    ]
    tiles_2 = [
        ("Clients served, month",
         '=COUNTIFS([[Appointments.month]],DashKey,'
         '[[Appointments.status]],"Completed")', "int",
         "completed appointments"),
        ("Average ticket",
         '=IFERROR(SUMIFS([[Sales.net]],[[Sales.month]],DashKey)/'
         'MAX(1,COUNTIFS([[Appointments.month]],DashKey,'
         '[[Appointments.status]],"Completed")),"")', "money",
         f"target {money}{cfg.target_average_ticket:,.0f}"),
        ("Retail share",
         '=IFERROR(SUMIFS([[Sales.net]],[[Sales.month]],DashKey,'
         '[[Sales.kind]],"Retail")/MAX(0.01,SUMIFS([[Sales.net]],'
         '[[Sales.month]],DashKey)),"")', "percent",
         f"target {cfg.target_retail_percent:.0%}"),
        ("Tips passed on",
         '=ROUND(SUMIFS([[Payments.amount]],[[Payments.month]],DashKey,'
         '[[Payments.kind]],"Tip"),2)', "money0",
         "the team's money, not the salon's"),
    ]
    tiles_3 = [
        ("Completion rate",
         '=IFERROR(COUNTIFS([[Appointments.month]],DashKey,'
         '[[Appointments.status]],"Completed")/'
         'MAX(1,COUNTIFS([[Appointments.month]],DashKey,'
         '[[Appointments.is_final]],"Yes")),"")', "percent",
         "of bookings that reached a final answer"),
        ("Cancellation rate",
         '=IFERROR(COUNTIFS([[Appointments.month]],DashKey,'
         '[[Appointments.status]],"Cancelled")/'
         'MAX(1,COUNTIFS([[Appointments.month]],DashKey,'
         '[[Appointments.is_final]],"Yes")),"")', "percent",
         "cancelled, not no-showed"),
        ("No-show rate",
         '=IFERROR(COUNTIFS([[Appointments.month]],DashKey,'
         '[[Appointments.status]],"No show")/'
         'MAX(1,COUNTIFS([[Appointments.month]],DashKey,'
         '[[Appointments.is_final]],"Yes")),"")', "percent",
         "the most expensive hour in the salon"),
        ("Chair utilisation",
         '=IFERROR(SUMIFS([[Appointments.hours]],'
         '[[Appointments.month]],DashKey,[[Appointments.status]],'
         '"Completed")/MAX(1,ChairHoursWeek*4.33),"")', "percent",
         f"target {cfg.target_utilisation:.0%} of sellable hours"),
    ]
    tiles_4 = [
        ("Rebooking rate",
         '=IFERROR(COUNTIFS([[Appointments.month]],DashKey,'
         '[[Appointments.rebooked]],"Yes")/'
         'MAX(1,COUNTIFS([[Appointments.month]],DashKey,'
         '[[Appointments.status]],"Completed")),"")', "percent",
         f"target {cfg.target_rebook_rate:.0%} leaving with the next "
         f"one booked"),
        ("Repeat client rate",
         '=IFERROR(COUNTIFS([[Clients.retained]],"Yes")/'
         'MAX(1,COUNTIFS([[Clients.visits]],">0")),"")', "percent",
         "clients served who came back at least twice"),
        ("Owed to you",
         '=ROUND(SUMIFS([[Sales.balance]],[[Sales.balance]],">0.004"),2)',
         "money0", "unpaid balances across every client"),
        ("Stock on the shelf",
         '=ROUND(SUM([[Inventory.value_cost]]),2)', "money0",
         "at cost, retail and backbar"),
    ]
    for tiles in (tiles_1, tiles_2, tiles_3, tiles_4):
        for i, (label, formula, style, note) in enumerate(tiles):
            s.kpi(row, i * 3, label, formula, style=style, span=3,
                  note=note)
        row += 4

    # ------------------------------------------------- needs attention ----
    row = s.section(row, "Needs doing today", LAST_COL)
    row = s.table_head(row, 0, ["What", "How many", "What it means",
                                "Where to fix it"],
                       [28, 11, 48, 18])
    alerts = [
        ("Clashes in the diary",
         '=COUNTIF([[Appointments.clash]],"Clash")',
         '="Two bookings for the same stylist at the same time"',
         "Appointments"),
        ("Bookings to confirm",
         '=COUNTIF([[Appointments.alert]],"Confirm*")',
         '="Coming up in the next two days and not confirmed"',
         "Appointments"),
        ("Deposits not taken",
         '=COUNTIF([[Appointments.alert]],"Due - deposit*")',
         '="Bookings over "&TEXT(DepositThreshold,"0")&" with no '
         'deposit against them"', "Appointments"),
        ("Visits with nothing rung up",
         '=COUNTIF([[Appointments.alert]],"Unpaid*")',
         '="Marked completed but no sale was logged"', "Sales"),
        ("Unpaid balances",
         '=COUNTIFS([[Sales.balance]],">0.004")',
         '="Sales where money is still owed"', "Payments"),
        ("Clients lapsing",
         '=COUNTIF([[Clients.segment]],"Lapsing")',
         '="Not in for "&LapsingDays&" days. Call them before they '
         'become inactive"', "Clients"),
        ("Clients due a rebook",
         '=COUNTIF([[Clients.alert]],"Rebook*")',
         '="Past their usual gap with nothing in the book"',
         "Clients"),
        ("Patch tests out of date",
         '=COUNTIF([[Clients.alert]],"Due - patch*")',
         '="Colour clients whose patch test is older than "&'
         'PatchTestDays&" days"', "Clients"),
        ("Products out of stock",
         '=COUNTIF([[Inventory.status]],"Out of stock")',
         '="Nothing left on the shelf to sell or use"', "Inventory"),
        ("Products to reorder",
         '=COUNTIF([[Inventory.status]],"Low stock")',
         '="Below the reorder point or under "&StockWarningWeeks&'
         '" weeks of cover"', "Inventory"),
        ("Deliveries late",
         '=COUNTIF([[Purchases.alert]],"Late*")',
         '="Ordered, expected, still not here"', "Purchases"),
        ("Bills overdue",
         '=COUNTIF([[Expenses.status]],"Overdue")',
         '="Past the due date and not marked paid"', "Expenses"),
        ("Commission to approve",
         '=COUNTIFS([[Commissions.approved]],"<>Yes",'
         '[[Commissions.month]],"<>")',
         '="Calculated and waiting for your sign-off"',
         "Commissions"),
        ("Pay runs not paid",
         '=COUNTIFS([[Payroll.paid]],"<>Yes",'
         '[[Payroll.period_end]],"<>")',
         '="Payroll rows still open"', "Payroll"),
        ("Loyalty rewards ready",
         '=COUNTIF([[Loyalty.alert]],"Action - reward*")',
         '="Clients who have earned something. Tell them"',
         "Loyalty"),
        ("Services below your chair rate",
         '=SUMPRODUCT((IFERROR([[Services.price_hour]]*1,0)>0)*'
         '(IFERROR([[Services.price_hour]]*1,0)<TargetChairRate))',
         '="Earning less than "&TEXT(TargetChairRate,"0")&" an hour '
         'before commission"', "Services"),
    ]
    first_alert = row
    for label, count, message, where in alerts:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        ws.write_formula(row, 1, s.reg.render(count), fmts.calc("int"))
        s.merge(row, 2, LAST_COL - 2, None, fmts.calc("text"))
        ws.write_formula(row, 2, s.reg.render(message), fmts.calc("text"))
        s.merge(row, LAST_COL - 1, LAST_COL, where, fmts.chrome("label"))
        row += 1
    ws.conditional_format(first_alert, 1, row - 1, 1,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("warn")})
    ws.conditional_format(first_alert, 1, row - 1, 1,
                          {"type": "cell", "criteria": "=", "value": 0,
                           "format": fmts.cf("good")})
    row += 1

    # ----------------------------------------------------- the team today --
    row = s.section(row, "Who is earning what", LAST_COL)
    row = s.table_head(row, 0, [
        "Stylist", "Sales this month", "Against target", "Completed",
        "Average ticket", "Retail share", "Utilisation",
        "Rebook rate", "Commission", "Pay this month", "Leaves behind"],
        [20, 14, 13, 11, 13, 12, 12, 12, 12, 13, 13])
    team_first = row
    for n in range(1, 9):
        s.write(row, 0, None, fmts.calc("text"))
        ws.write_formula(
            row, 0,
            s.reg.render('=' + _by_rank("Staff", "rev_total_rank",
                                        "name", n)),
            fmts.calc("text"))
        pairs = [
            (1, "rev_month", "money"),
            (2, "target_gap", "money"),
            (3, "done", "int"),
            (4, "avg_ticket", "money"),
            (5, "retail_share", "percent"),
            (6, "utilisation", "percent"),
            (7, "rebook_rate", "percent"),
            (8, "commission_paid", "money"),
            (9, "cost_month", "money"),
            (10, "margin_month", "money"),
        ]
        for col, key, style in pairs:
            ws.write_formula(
                row, col,
                s.reg.render('=' + _by_rank("Staff", "rev_total_rank",
                                            key, n)),
                fmts.calc(style))
        row += 1
    ws.conditional_format(team_first, 10, row - 1, 10,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    row += 1

    # --------------------------------------------------- best services ----
    row = s.section(row, "What the chairs are selling", LAST_COL)
    row = s.table_head(row, 0, [
        "Most booked", "Times", "Biggest earner", "Revenue",
        "Best an hour", "An hour", "Best clients", "Spent"],
        [24, 9, 24, 12, 24, 11, 22, 12])
    for n in range(1, TOP_ROWS + 1):
        cells = [
            (0, _by_rank("Services", "done_rank", "name", n), "text"),
            (1, _by_rank("Services", "done_rank", "done", n), "int"),
            (2, _by_rank("Services", "revenue_rank", "name", n), "text"),
            (3, _by_rank("Services", "revenue_rank", "revenue", n),
             "money"),
            (4, _by_rank("Services", "profit_hour_rank", "name", n),
             "text"),
            (5, _by_rank("Services", "profit_hour_rank", "price_hour", n),
             "money"),
            (6, _by_rank("Clients", "spend_rank", "name", n), "text"),
            (7, _by_rank("Clients", "spend_rank", "spend", n), "money"),
        ]
        for col, formula, style in cells:
            ws.write_formula(row, col, s.reg.render("=" + formula),
                             fmts.calc(style))
        row += 1
    row += 1

    # ------------------------------------------------- twelve months ------
    row = s.section(row, "The last twelve months", LAST_COL)
    row = s.table_head(row, 0, [
        "Month", "Net sales", "Target", "Against target", "Services",
        "Retail", "Cost of sales", "Wages", "Expenses",
        "Operating profit", "Margin", "Clients served",
        "Average ticket"],
        [12, 13, 12, 13, 12, 11, 13, 12, 12, 14, 10, 12, 13])
    chart_first = row
    for index in range(MONTHS - 1, -1, -1):
        key = (f'TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{index},1),'
               f'"yyyy-mm")')
        label = (f'=TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{index},1),'
                 f'"mmm yy")')
        s.write(row, 0, None, fmts.calc("text"))
        ws.write_formula(row, 0, label, fmts.calc("text"))
        formulas = [
            (1, f'=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{key}),2)',
             "money"),
            (2, "=ROUND(RevenueTargetMonth,2)", "money"),
            (3, f'=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{key})-'
                f'RevenueTargetMonth,2)', "money"),
            (4, f'=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{key},'
                f'[[Sales.kind]],"Service"),2)', "money"),
            (5, f'=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{key},'
                f'[[Sales.kind]],"Retail"),2)', "money"),
            (6, f'=ROUND(SUMIFS([[Sales.cost]],[[Sales.month]],{key}),2)',
             "money"),
            (7, f'=ROUND(SUMIFS([[Payroll.total_cost]],'
                f'[[Payroll.month]],{key}),2)', "money"),
            (8, f'=ROUND(SUMIFS([[Expenses.amount]],[[Expenses.month]],'
                f'{key}),2)', "money"),
            (9, f'=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{key})-'
                f'SUMIFS([[Sales.cost]],[[Sales.month]],{key})-'
                f'SUMIFS([[Payroll.total_cost]],[[Payroll.month]],{key})-'
                f'SUMIFS([[Expenses.amount]],[[Expenses.month]],{key}),2)',
             "money"),
            (10, f'=IF(SUMIFS([[Sales.net]],[[Sales.month]],{key})<=0,"",'
                 f'(SUMIFS([[Sales.net]],[[Sales.month]],{key})-'
                 f'SUMIFS([[Sales.cost]],[[Sales.month]],{key})-'
                 f'SUMIFS([[Payroll.total_cost]],[[Payroll.month]],{key})-'
                 f'SUMIFS([[Expenses.amount]],[[Expenses.month]],{key}))/'
                 f'SUMIFS([[Sales.net]],[[Sales.month]],{key}))',
             "percent"),
            (11, f'=COUNTIFS([[Appointments.month]],{key},'
                 f'[[Appointments.status]],"Completed")', "int"),
            (12, f'=IFERROR(SUMIFS([[Sales.net]],[[Sales.month]],{key})/'
                 f'MAX(1,COUNTIFS([[Appointments.month]],{key},'
                 f'[[Appointments.status]],"Completed")),"")', "money"),
        ]
        for col, formula, style in formulas:
            ws.write_formula(row, col, s.reg.render(formula),
                             fmts.calc(style))
        row += 1
    last_month_row = row - 1
    ws.conditional_format(chart_first, 1, last_month_row, 1,
                          {"type": "data_bar",
                           "bar_color": Palette.primary_mid})
    ws.conditional_format(chart_first, 9, last_month_row, 9,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(chart_first, 3, last_month_row, 3,
                          {"type": "cell", "criteria": ">=", "value": 0,
                           "format": fmts.cf("good")})

    # sparklines beat a chart that Excel 2010 might not draw
    ws.add_sparkline(row, 1, {
        "range": (f"{quote_sheet(NAME)}!$B${chart_first + 1}:"
                  f"$B${last_month_row + 1}"),
        "type": "column", "series_color": Palette.primary_mid})
    s.merge(row, 2, 4, "Net sales, twelve months",
            fmts.chrome("note"))
    ws.add_sparkline(row, 5, {
        "range": (f"{quote_sheet(NAME)}!$J${chart_first + 1}:"
                  f"$J${last_month_row + 1}"),
        "type": "column", "series_color": Palette.accent,
        "negative_points": True})
    s.merge(row, 6, 8, "Operating profit, twelve months",
            fmts.chrome("note"))
    row += 2

    # ------------------------------------------------- month definitions --
    row = s.section(row, "The month in one line", LAST_COL)
    row = s.table_head(row, 0, [
        "Net sales", "Cost of sales", "Gross profit", "Wages",
        "Other costs", "Operating profit", "Margin", "Break-even sales"],
        [14, 14, 14, 13, 13, 15, 10, 15])
    totals_row = row
    s.write(row, 0, None, fmts.calc("money"))
    ws.write_formula(
        row, 0, s.reg.render('=ROUND(SUMIFS([[Sales.net]],'
                             '[[Sales.month]],DashKey),2)'),
        fmts.calc("money"))
    book.define_name("MonthNetSales",
                     f"={quote_sheet(NAME)}!$A${totals_row + 1}")
    ws.write_formula(
        row, 1, s.reg.render('=ROUND(SUMIFS([[Sales.cost]],'
                             '[[Sales.month]],DashKey),2)'),
        fmts.calc("money"))
    book.define_name("MonthCostOfSales",
                     f"={quote_sheet(NAME)}!$B${totals_row + 1}")
    ws.write_formula(row, 2, "=ROUND(MonthNetSales-MonthCostOfSales,2)",
                     fmts.calc("money"))
    ws.write_formula(
        row, 3, s.reg.render('=ROUND(SUMIFS([[Payroll.total_cost]],'
                             '[[Payroll.month]],DashKey),2)'),
        fmts.calc("money"))
    book.define_name("MonthWages",
                     f"={quote_sheet(NAME)}!$D${totals_row + 1}")
    ws.write_formula(
        row, 4, s.reg.render('=ROUND(SUMIFS([[Expenses.amount]],'
                             '[[Expenses.month]],DashKey),2)'),
        fmts.calc("money"))
    book.define_name("MonthExpenses",
                     f"={quote_sheet(NAME)}!$E${totals_row + 1}")
    ws.write_formula(row, 5, "=ROUND(MonthNetSales-MonthCostOfSales-"
                             "MonthWages-MonthExpenses,2)",
                     fmts.calc("money"))
    ws.write_formula(row, 6, '=IF(MonthNetSales<=0,"",'
                             '(MonthNetSales-MonthCostOfSales-'
                             'MonthWages-MonthExpenses)/MonthNetSales)',
                     fmts.calc("percent"))
    ws.write_formula(row, 7, '=IF(MonthNetSales<=0,"",'
                             'ROUND((MonthWages+MonthExpenses)/'
                             'MAX(0.05,(MonthNetSales-MonthCostOfSales)/'
                             'MonthNetSales),2))', fmts.calc("money"))
    row += 2

    row = s.note(
        row,
        "Operating profit here is net sales less the cost of what you "
        "sold, less the full cost of the pay runs, less the expenses "
        "logged for the month. It is not a tax computation and it does "
        "not know about depreciation, drawings or loans - it is the "
        "number that tells you whether the month worked.",
        LAST_COL, height=36)

    s.finish(freeze_row=6, gridlines=False, fit_width=True)
    s.protect()
    return s
