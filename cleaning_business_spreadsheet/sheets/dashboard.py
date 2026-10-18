"""`Dashboard` - the one page that answers 'how is the business doing?'

Twelve numbers across the top, a list of things that need doing today,
the next seven days of work, your biggest clients and twelve months of
revenue, cost and profit. Everything reads straight from the registers,
so it is never out of date.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter

NAME = "Dashboard"
LAST_COL = 11
MONTHS = 12

#: revenue counted the same way everywhere: completed or invoiced work
_DONE = ('SUMIFS([[Jobs.price]],[[Jobs.month]],{key},[[Jobs.status]],'
         '"Completed")+SUMIFS([[Jobs.price]],[[Jobs.month]],{key},'
         '[[Jobs.status]],"Invoiced")')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws

    s.banner("Business Dashboard",
             "Today, this week and the year so far - all worked out "
             "from the jobs, invoices and expenses you have already "
             "entered.", LAST_COL)
    s.set_col(0, 0, 22)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 13)

    # ------------------------------------------------------- the header --
    row = 4
    s.merge(row, 0, 0, "Business", fmts.chrome("label_bold"))
    s.merge(row, 1, 3, None, fmts.calc("text"))
    ws.write_formula(row, 1, "=BizName", fmts.calc("text"))
    s.merge(row, 4, 4, "Month", fmts.chrome("label_right"))
    ws.write_formula(row, 5, '=TEXT(TodayDate,"mmmm yyyy")',
                     fmts.calc("text"))
    book.define_name("DashMonth", f"={quote_sheet(NAME)}!$F${row + 1}")
    s.merge(row, 6, 6, "Month key", fmts.chrome("label_right"))
    ws.write_formula(row, 7, '=TEXT(TodayDate,"yyyy-mm")',
                     fmts.calc("text"))
    book.define_name("DashKey", f"={quote_sheet(NAME)}!$H${row + 1}")
    s.merge(row, 8, 9, "Today", fmts.chrome("label_right"))
    s.merge(row, 10, LAST_COL, None, fmts.calc("date"))
    ws.write_formula(row, 10, "=TodayDate", fmts.calc("date"))
    row += 2

    # ---------------------------------------------------------- the KPIs --
    money = cfg.currency_symbol
    tiles_1 = [
        ("Revenue this month",
         '=' + _DONE.format(key="DashKey"), "money0",
         f"target {money}{cfg.monthly_revenue_target:,.0f}"),
        ("Invoiced this month",
         '=SUMIFS([[Invoices.total]],[[Invoices.month]],DashKey)',
         "money0", "what you have billed"),
        ("Collected this month",
         '=SUMIFS([[Payments.amount]],[[Payments.month]],DashKey)',
         "money0", "money actually in"),
        ("Owed to you right now",
         '=SUM([[Invoices.outstanding]])', "money0",
         "every unpaid invoice"),
    ]
    tiles_2 = [
        ("Costs this month",
         '=SUMIFS([[Expenses.total]],[[Expenses.month]],DashKey)',
         "money0", f"budget {money}{cfg.monthly_expense_budget:,.0f}"),
        ("Profit this month",
         '=' + _DONE.format(key="DashKey") +
         '-SUMIFS([[Expenses.total]],[[Expenses.month]],DashKey)',
         "money0", "revenue less expenses"),
        ("Margin this month",
         '=IFERROR((' + _DONE.format(key="DashKey") +
         '-SUMIFS([[Expenses.total]],[[Expenses.month]],DashKey))/(' +
         _DONE.format(key="DashKey") + '),"")', "percent0",
         f"target {cfg.target_margin:.0%}"),
        ("Average job value",
         '=IFERROR(' + _DONE.format(key="DashKey") +
         '/(COUNTIFS([[Jobs.month]],DashKey,[[Jobs.status]],'
         '"Completed")+COUNTIFS([[Jobs.month]],DashKey,'
         '[[Jobs.status]],"Invoiced")),"")', "money",
         "completed work this month"),
    ]
    tiles_3 = [
        ("Jobs this week",
         '=COUNTIFS([[Jobs.date]],">="&WeekStartDate,[[Jobs.date]],'
         '"<="&WeekStartDate+6,[[Jobs.status]],"<>Cancelled")', "int",
         "booked in the next seven days"),
        ("Hours booked this week",
         '=SUMIFS([[Jobs.hours_each]],[[Jobs.date]],">="&WeekStartDate,'
         '[[Jobs.date]],"<="&WeekStartDate+6,[[Jobs.status]],'
         '"<>Cancelled")', "number1", "across the whole team"),
        ("Quotes waiting",
         '=COUNTIFS([[Quotes.status]],"Sent")+'
         'COUNTIFS([[Quotes.status]],"Followed up")', "int",
         "worth chasing today"),
        ("Active clients",
         '=COUNTIFS([[Clients.status]],"Active")+'
         'COUNTIFS([[Clients.status]],"Regular")', "int",
         "on the books right now"),
    ]
    for tiles in (tiles_1, tiles_2, tiles_3):
        for i, (label, formula, style, note) in enumerate(tiles):
            s.kpi(row, i * 3, label, s.reg.render(formula), style=style,
                  span=3, note=note)
        row += 4

    # ----------------------------------------------- needs your attention
    row = s.section(row, "Needs your attention", LAST_COL)
    row = s.table_head(row, 0, ["What", "How many", "What it means",
                                "Where to fix it"],
                       [26, 11, 46, 20])
    alerts = [
        ("Jobs with no cleaner",
         '=COUNTIFS([[Jobs.employee]],"",[[Jobs.client]],"<>",'
         '[[Jobs.date]],">="&TodayDate)',
         '="Booked work nobody is assigned to yet"', "Jobs"),
        ("Visits to confirm",
         '=COUNTIF([[Jobs.alert]],"Confirm*")',
         '="Coming up within "&ReminderDays&" days and not confirmed '
         'with the client"', "Jobs"),
        ("Double bookings",
         '=COUNTIF([[Jobs.clash]],"Clash")',
         '="The same cleaner booked twice at the same time"', "Jobs"),
        ("Jobs done but not invoiced",
         '=COUNTIFS([[Jobs.status]],"Completed",[[Jobs.invoice]],"")',
         '="Work you have done and not yet billed - worth "&'
         'TEXT(SUMIFS([[Jobs.price]],[[Jobs.status]],"Completed",'
         '[[Jobs.invoice]],""),"#,##0")', "Invoices"),
        ("Overdue invoices",
         '=COUNTIF([[Invoices.alert]],"Overdue*")',
         '="Past their due date, totalling "&'
         'TEXT(SUMIFS([[Invoices.outstanding]],[[Invoices.age]],'
         '"1-30 days")+SUMIFS([[Invoices.outstanding]],'
         '[[Invoices.age]],"31-60 days")+SUMIFS('
         '[[Invoices.outstanding]],[[Invoices.age]],"61-90 days")+'
         'SUMIFS([[Invoices.outstanding]],[[Invoices.age]],'
         '"90+ days"),"#,##0")', "Invoices"),
        ("Quotes to chase",
         '=COUNTIF([[Quotes.alert]],"Chase*")',
         '="Sent more than "&QuoteFollowupDays&" days ago with no '
         'answer - this is the fastest money in the workbook"',
         "Quotes"),
        ("Leads with no quote",
         '=COUNTIF([[Leads.alert]],"Action - no quote*")',
         '="Enquiries that have not been priced yet"', "Leads"),
        ("Clients going quiet",
         '=COUNTIF([[Clients.retention]],"At risk")+'
         'COUNTIF([[Clients.retention]],"Slipping")',
         '="No visit booked, or nothing for "&InactiveClientDays&'
         '" days"', "Client Retention"),
        ("Supplies to reorder",
         '=COUNTIF([[Supplies.status]],"Low stock")+'
         'COUNTIF([[Supplies.status]],"Out of stock")',
         '="Below the reorder level or less than "&LowStockDays&'
         '" days of cover"', "Supplies"),
        ("Equipment due a service",
         '=COUNTIF([[Equipment.alert]],"Overdue*")+'
         'COUNTIF([[Equipment.alert]],"Due*")',
         '="A machine that dies mid-job costs you the day"',
         "Equipment"),
        ("Certificates expiring",
         '=COUNTIF([[Safety.status]],"Due soon")+'
         'COUNTIF([[Safety.status]],"Expired")',
         '="Insurance, checks and training within "&CertWarningDays&'
         '" days of running out"', "Safety"),
        ("Open complaints",
         '=COUNTIF([[Complaints.open]],"Yes")',
         '="Every day one of these stays open costs you a client"',
         "Complaints"),
        ("Loss-making jobs",
         '=COUNTIF([[Job Costing.alert]],"Loss*")',
         '="Jobs that cost more than they earned"', "Job Costing"),
        ("Timesheets to approve",
         '=COUNTIF([[Timesheets.alert]],"Due - needs approving")',
         '="Hours entered but not signed off"', "Timesheets"),
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

    # ------------------------------------------------------ the next week
    row = s.section(row, "The next seven days", LAST_COL)
    s.write(row, 0, "Day", fmts.chrome("header_left"))
    for i in range(7):
        ws.write_formula(row, i + 1, f"=TodayDate+{i}",
                         fmts.chrome("header", num_format="ddd d mmm"))
    s.write(row, 8, "Week", fmts.chrome("header"))
    day_row = row
    row += 1
    week = ('[[Jobs.date]],">="&TodayDate,[[Jobs.date]],'
            '"<="&TodayDate+6,[[Jobs.status]],"<>Cancelled")')
    for label, fn, value_range, style in [
            ("Jobs booked", "COUNTIFS", "", "int"),
            ("Hours on site", "SUMIFS", "[[Jobs.hours_each]],",
             "number1"),
            ("Value", "SUMIFS", "[[Jobs.price]],", "money0"),
            ("Cleaners needed", "SUMIFS", "[[Jobs.crew]],", "int")]:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        for i in range(7):
            cell = f"{chr(66 + i)}${day_row + 1}"
            day = (f'={fn}({value_range}[[Jobs.date]],{cell},'
                   f'[[Jobs.status]],"<>Cancelled")')
            ws.write_formula(row, i + 1, s.reg.render(day),
                             fmts.calc(style))
        ws.write_formula(row, 8,
                         s.reg.render(f"={fn}({value_range}{week}"),
                         fmts.chrome("total_value",
                                     num_format=s.num_format(style)))
        row += 1
    ws.conditional_format(day_row + 1, 1, day_row + 1, 7,
                          {"type": "data_bar", "bar_color": "#2F8F9D"})
    row += 1
    s.merge(row, 0, LAST_COL,
            "Nothing in a column means a free day - the best day to "
            "fit in a quote visit, a deep clean or some time off.",
            fmts.chrome("note"))
    row += 2

    # ---------------------------------------------------- biggest clients
    row = s.section(row, "Your five biggest clients", LAST_COL)
    row = s.table_head(row, 0, ["Client", "Type", "Jobs done",
                                "Invoiced", "Profit", "Margin",
                                "Worth a year", "Last clean",
                                "Standing"],
                       [24, 14, 10, 13, 12, 10, 13, 12, 13])
    top = max(3, cfg.top_rows)
    first_client = row
    for n in range(1, top + 1):
        pick = ('MATCH({n},[[Clients.invoiced_rank]],0)'
                .replace("{n}", str(n)))
        s.write(row, 0, None, fmts.calc("text"))
        for col, (key, style) in enumerate([
                ("name", "text"), ("type", "text"),
                ("jobs_done", "int"), ("invoiced", "money"),
                ("profit", "money"), ("margin", "percent0"),
                ("year_value", "money"), ("last_job", "date"),
                ("retention", "status")]):
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Clients.{key}]],{pick}),'
                             f'"")'),
                fmts.calc(style))
        row += 1
    ws.conditional_format(first_client, 3, row - 1, 3,
                          {"type": "data_bar", "bar_color": "#E8A33D"})
    ws.conditional_format(first_client, 4, row - 1, 4,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    row += 1

    # ------------------------------------------------------ the year table
    row = s.section(row, "The year so far", LAST_COL)
    row = s.table_head(row, 0, ["Month", "Jobs", "Revenue", "Invoiced",
                                "Collected", "Costs", "Profit", "Margin",
                                "New clients", "Trend"],
                       [10, 9, 13, 13, 13, 12, 12, 9, 11, 14])
    first_month = row
    for m in range(MONTHS):
        key = f'TEXT(DATE(PlanYear,{m + 1},1),"yyyy-mm")'
        ws.write_formula(row, 0, f'=TEXT(DATE(PlanYear,{m + 1},1),"mmm")',
                         fmts.calc("text"))
        revenue = _DONE.format(key=key)
        s.calc_row(row, 1, [
            f'=COUNTIFS([[Jobs.month]],{key},[[Jobs.status]],'
            f'"<>Cancelled")',
            f'={revenue}',
            f'=SUMIFS([[Invoices.total]],[[Invoices.month]],{key})',
            f'=SUMIFS([[Payments.amount]],[[Payments.month]],{key})',
            f'=SUMIFS([[Expenses.total]],[[Expenses.month]],{key})',
            f'={revenue}-SUMIFS([[Expenses.total]],[[Expenses.month]],'
            f'{key})',
            f'=IFERROR(({revenue}-SUMIFS([[Expenses.total]],'
            f'[[Expenses.month]],{key}))/({revenue}),"")',
            f'=COUNTIFS([[Clients.since]],">="&DATE(PlanYear,{m + 1},1),'
            f'[[Clients.since]],"<"&DATE(PlanYear,{m + 2},1))',
        ], ["int", "money0", "money0", "money0", "money0", "money0",
            "percent0", "int"])
        row += 1
    last_month = row - 1
    ws.conditional_format(first_month, 2, last_month, 2,
                          {"type": "data_bar", "bar_color": "#2F8F9D"})
    ws.conditional_format(first_month, 6, last_month, 6,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(first_month, 7, last_month, 7,
                          {"type": "cell", "criteria": ">=",
                           "value": "=TargetMargin",
                           "format": fmts.cf("good")})

    for idx, (src, colour) in enumerate(
            [(2, "#14616B"), (5, "#C86B6B"), (6, "#4E8A63")]):
        letter = chr(65 + src)
        try:
            ws.add_sparkline(first_month + idx, 9, {
                "range": (f"{quote_sheet(NAME)}!{letter}"
                          f"{first_month + 1}:{letter}{last_month + 1}"),
                "type": "column",
                "series_color": colour,
            })
        except Exception:        # pragma: no cover - library differences
            pass
    row += 1
    s.merge(row, 0, LAST_COL,
            "Revenue counts jobs marked Completed or Invoiced. "
            "Collected is money actually received. The little charts "
            "on the right are, top to bottom: revenue, costs, profit.",
            fmts.chrome("note"))
    row += 2

    # ------------------------------------------------------- the pipeline
    row = s.section(row, "Sales pipeline", LAST_COL)
    row = s.table_head(row, 0, ["Stage", "How many", "Value", "Notes"],
                       [20, 12, 14, 50])
    stages = [
        ("New enquiries", '=COUNTIF([[Leads.stage]],"New enquiry")',
         '=SUMIFS([[Leads.value]],[[Leads.stage]],"New enquiry")',
         "Call them today - speed wins more cleaning work than price"),
        ("Quotes out",
         '=COUNTIF([[Quotes.status]],"Sent")+'
         'COUNTIF([[Quotes.status]],"Followed up")',
         '=SUMIFS([[Quotes.total]],[[Quotes.status]],"Sent")+'
         'SUMIFS([[Quotes.total]],[[Quotes.status]],"Followed up")',
         "Money sitting in somebody else's inbox"),
        ("Accepted, not booked",
         '=COUNTIFS([[Quotes.status]],"Accepted",[[Quotes.job]],"")',
         '=SUMIFS([[Quotes.total]],[[Quotes.status]],"Accepted",'
         '[[Quotes.job]],"")',
         "They said yes. Put it in the diary before they forget"),
        ("Won this year",
         '=COUNTIF([[Quotes.won]],"Yes")',
         '=SUMIFS([[Quotes.total]],[[Quotes.won]],"Yes")',
         "Quotes accepted"),
        ("Lost this year",
         '=COUNTIF([[Quotes.won]],"No")',
         '=SUMIFS([[Quotes.total]],[[Quotes.won]],"No")',
         "Worth reading the reasons on the Leads tab"),
    ]
    for label, count, value, note in stages:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        ws.write_formula(row, 1, s.reg.render(count), fmts.calc("int"))
        ws.write_formula(row, 2, s.reg.render(value), fmts.calc("money0"))
        s.merge(row, 3, LAST_COL, note, fmts.chrome("note"))
        row += 1
    row += 1

    s.note(row,
           "Every figure on this page is calculated. If a number looks "
           "wrong, the fix is on the tab it came from - usually a job "
           "left at the wrong status, or an invoice with no lines.",
           LAST_COL, height=30)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
