"""`Reports` - the week and the month, on one printable page.

Pick a week ending date and a month at the top; everything below
re-reads itself. Print it for the Monday meeting, the accountant or
the folder you keep for the bank.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, date_check, list_from

NAME = "Reports"
LAST_COL = 9
TOP_ROWS = 6


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.deep)
    ws = s.ws
    sheet_ref = quote_sheet(NAME)

    s.banner("Weekly & Monthly Reports",
             "Two pickers at the top, two reports underneath. Nothing "
             "to rebuild and nothing to refresh.", LAST_COL)
    s.set_col(0, 0, 3)
    s.set_col(1, 1, 30)
    s.set_col(2, LAST_COL, 15)

    row = 4
    row = s.field(row, "Week ending (a Sunday works best)",
                  cfg.today, style="date", name="ReportWeekEnd",
                  validate=date_check(),
                  help_text="The report covers the seven days up to "
                            "and including this date.")
    row = s.field(row, "Month", f"{cfg.today.year:04d}-"
                                f"{cfg.today.month:02d}",
                  style="text", name="ReportMonth",
                  validate=list_from("MonthKeyList"),
                  help_text="Pick any month in yyyy-mm form.")
    week_start = "(ReportWeekEnd-6)"
    row += 1

    # ------------------------------------------------------- the week ---
    row = s.section(row, "The week", LAST_COL)
    s.merge(row, 1, 4, None, fmts.calc("text"))
    ws.write_formula(row, 1,
                     '="Week of "&TEXT(ReportWeekEnd-6,"d mmm")&'
                     '" to "&TEXT(ReportWeekEnd,"d mmm yyyy")',
                     fmts.calc("text"))
    row += 1
    s.table_head(row, 1, [
        "Measure", "This week", "Week before", "Change",
        "Change %", "Four week average"],
        [30, 15, 15, 14, 12, 18])
    row += 1
    week_first = row
    measures = [
        ("Net sales", "money",
         '=ROUND(SUMIFS([[Sales.net]],[[Sales.date]],">="&{from},'
         '[[Sales.date]],"<="&{to}),2)'),
        ("Service sales", "money",
         '=ROUND(SUMIFS([[Sales.net]],[[Sales.kind]],"Service",'
         '[[Sales.date]],">="&{from},[[Sales.date]],"<="&{to}),2)'),
        ("Retail sales", "money",
         '=ROUND(SUMIFS([[Sales.net]],[[Sales.kind]],"Retail",'
         '[[Sales.date]],">="&{from},[[Sales.date]],"<="&{to}),2)'),
        ("Clients seen", "int",
         '=COUNTIFS([[Appointments.status]],"Completed",'
         '[[Appointments.date]],">="&{from},[[Appointments.date]],'
         '"<="&{to})'),
        ("New clients", "int",
         '=COUNTIFS([[Clients.first_visit]],">="&{from},'
         '[[Clients.first_visit]],"<="&{to})'),
        ("Average ticket", "money",
         '=IFERROR(ROUND(SUMIFS([[Sales.net]],[[Sales.date]],">="&'
         '{from},[[Sales.date]],"<="&{to})/MAX(1,COUNTIFS('
         '[[Appointments.status]],"Completed",[[Appointments.date]],'
         '">="&{from},[[Appointments.date]],"<="&{to})),2),"")'),
        ("Cancellations", "int",
         '=COUNTIFS([[Appointments.status]],"Cancelled",'
         '[[Appointments.date]],">="&{from},[[Appointments.date]],'
         '"<="&{to})'),
        ("No shows", "int",
         '=COUNTIFS([[Appointments.status]],"No show",'
         '[[Appointments.date]],">="&{from},[[Appointments.date]],'
         '"<="&{to})'),
        ("Chair hours worked", "number1",
         '=ROUND(SUMIFS([[Appointments.hours]],'
         '[[Appointments.status]],"Completed",[[Appointments.date]],'
         '">="&{from},[[Appointments.date]],"<="&{to}),1)'),
        ("Tips taken", "money",
         '=ROUND(SUMIFS([[Payments.amount]],[[Payments.kind]],"Tip",'
         '[[Payments.date]],">="&{from},[[Payments.date]],"<="&'
         '{to}),2)'),
        ("Money into the bank", "money",
         '=ROUND(SUMIFS([[Payments.net_cash]],[[Payments.date]],">="&'
         '{from},[[Payments.date]],"<="&{to}),2)'),
        ("Costs logged", "money",
         '=ROUND(SUMIFS([[Expenses.amount]],[[Expenses.date]],">="&'
         '{from},[[Expenses.date]],"<="&{to}),2)'),
    ]
    for label, style, body in measures:
        s.write(row, 1, label, fmts.chrome("label_bold"))
        this_week = body.format(**{"from": week_start,
                                   "to": "ReportWeekEnd"})
        last_week = body.format(**{"from": "(ReportWeekEnd-13)",
                                   "to": "(ReportWeekEnd-7)"})
        four_week = body.format(**{"from": "(ReportWeekEnd-27)",
                                   "to": "ReportWeekEnd"})
        ws.write_formula(row, 2, s.reg.render(this_week),
                         fmts.calc(style))
        ws.write_formula(row, 3, s.reg.render(last_week),
                         fmts.calc(style))
        ws.write_formula(row, 4, f"=ROUND(C{row + 1}-D{row + 1},2)",
                         fmts.calc(style))
        ws.write_formula(row, 5,
                         f'=IFERROR((C{row + 1}-D{row + 1})/'
                         f'ABS(D{row + 1}),"")', fmts.calc("percent"))
        if style == "int":
            ws.write_formula(
                row, 6, f"=ROUND(({s.reg.render(four_week)[1:]})/4,1)",
                fmts.calc("number1"))
        else:
            ws.write_formula(
                row, 6, f"=ROUND(({s.reg.render(four_week)[1:]})/4,2)",
                fmts.calc(style))
        row += 1
    week_last = row - 1
    ws.conditional_format(week_first, 4, week_last, 4,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("good")})
    ws.conditional_format(week_first, 4, week_last, 4,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("warn")})
    row += 1

    # --------------------------------------------------- week, by day ---
    row = s.subsection(row, "The week, day by day", LAST_COL)
    s.table_head(row, 1, [
        "Day", "Date", "Appointments", "Completed", "Net sales",
        "Chair hours", "No shows"],
        [14, 13, 14, 12, 14, 13, 12])
    row += 1
    day_first = row
    for offset in range(6, -1, -1):
        day_ref = f"(ReportWeekEnd-{offset})"
        s.write(row, 1, None, fmts.calc("text"))
        ws.write_formula(row, 1, f'=TEXT({day_ref},"dddd")',
                         fmts.calc("text"))
        ws.write_formula(row, 2, f"={day_ref}", fmts.calc("date"))
        ws.write_formula(
            row, 3,
            s.reg.render(f'=COUNTIFS([[Appointments.date]],{day_ref})'),
            fmts.calc("int"))
        ws.write_formula(
            row, 4,
            s.reg.render(f'=COUNTIFS([[Appointments.date]],{day_ref},'
                         f'[[Appointments.status]],"Completed")'),
            fmts.calc("int"))
        ws.write_formula(
            row, 5,
            s.reg.render(f'=ROUND(SUMIFS([[Sales.net]],[[Sales.date]],'
                         f'{day_ref}),2)'), fmts.calc("money"))
        ws.write_formula(
            row, 6,
            s.reg.render(f'=ROUND(SUMIFS([[Appointments.hours]],'
                         f'[[Appointments.date]],{day_ref},'
                         f'[[Appointments.status]],"Completed"),1)'),
            fmts.calc("number1"))
        ws.write_formula(
            row, 7,
            s.reg.render(f'=COUNTIFS([[Appointments.date]],{day_ref},'
                         f'[[Appointments.status]],"No show")'),
            fmts.calc("int"))
        row += 1
    day_last = row - 1
    ws.conditional_format(day_first, 5, day_last, 5,
                          {"type": "data_bar",
                           "bar_color": Palette.primary_mid})
    ws.add_sparkline(row, 5, {
        "range": (f"{sheet_ref}!$F${day_first + 1}:$F${day_last + 1}"),
        "type": "column", "series_color": Palette.primary_mid})
    s.merge(row, 1, 4, "Net sales through the week",
            fmts.chrome("note"))
    row += 2

    # ------------------------------------------------------ the month ---
    row = s.section(row, "The month", LAST_COL)
    s.merge(row, 1, 4, None, fmts.calc("text"))
    ws.write_formula(
        row, 1,
        '="Month of "&TEXT(DATE(VALUE(LEFT(ReportMonth,4)),'
        'VALUE(RIGHT(ReportMonth,2)),1),"mmmm yyyy")',
        fmts.calc("text"))
    row += 1
    s.table_head(row, 1, [
        "Measure", "This month", "Month before", "Change",
        "Change %", "Share of sales"],
        [30, 15, 15, 14, 12, 15])
    row += 1
    prev = ('TEXT(DATE(VALUE(LEFT(ReportMonth,4)),'
            'VALUE(RIGHT(ReportMonth,2))-1,1),"yyyy-mm")')
    month_first = row
    month_measures = [
        ("Net sales", "money", "yes",
         '=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{m}),2)'),
        ("Service sales", "money", "yes",
         '=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{m},'
         '[[Sales.kind]],"Service"),2)'),
        ("Retail sales", "money", "yes",
         '=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{m},'
         '[[Sales.kind]],"Retail"),2)'),
        ("Packages and vouchers", "money", "yes",
         '=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{m},'
         '[[Sales.kind]],"Package")+SUMIFS([[Sales.net]],'
         '[[Sales.month]],{m},[[Sales.kind]],"Gift certificate"),2)'),
        ("Discounts given", "money", "yes",
         '=ROUND(SUMIFS([[Sales.discount]],[[Sales.month]],{m}),2)'),
        ("Refunds", "money", "yes",
         '=ROUND(SUMIFS([[Sales.refund]],[[Sales.month]],{m}),2)'),
        ("Cost of sales", "money", "yes",
         '=ROUND(SUMIFS([[Sales.cost]],[[Sales.month]],{m}),2)'),
        ("Staff cost", "money", "yes",
         '=ROUND(SUMIFS([[Payroll.total_cost]],[[Payroll.month]],'
         '{m}),2)'),
        ("Other expenses", "money", "yes",
         '=ROUND(SUMIFS([[Expenses.amount]],[[Expenses.month]],'
         '{m}),2)'),
        ("Operating profit", "money", "yes",
         '=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{m})-'
         'SUMIFS([[Sales.cost]],[[Sales.month]],{m})-'
         'SUMIFS([[Payroll.total_cost]],[[Payroll.month]],{m})-'
         'SUMIFS([[Expenses.amount]],[[Expenses.month]],{m}),2)'),
        ("Clients seen", "int", "",
         '=COUNTIFS([[Appointments.month]],{m},'
         '[[Appointments.status]],"Completed")'),
        ("Cancellations", "int", "",
         '=COUNTIFS([[Appointments.month]],{m},'
         '[[Appointments.status]],"Cancelled")'),
        ("No shows", "int", "",
         '=COUNTIFS([[Appointments.month]],{m},'
         '[[Appointments.status]],"No show")'),
        ("Tips", "money", "",
         '=ROUND(SUMIFS([[Payments.amount]],[[Payments.month]],{m},'
         '[[Payments.kind]],"Tip"),2)'),
        ("Sales tax collected", "money", "",
         '=ROUND(SUMIFS([[Sales.tax]],[[Sales.month]],{m}),2)'),
        ("Deposits taken", "money", "",
         '=ROUND(SUMIFS([[Payments.amount]],[[Payments.month]],{m},'
         '[[Payments.kind]],"Deposit"),2)'),
    ]
    sales_row = month_first                     # net sales row, for shares
    for label, style, share, body in month_measures:
        s.write(row, 1, label, fmts.chrome("label_bold"))
        ws.write_formula(row, 2, s.reg.render(body.format(m="ReportMonth")),
                         fmts.calc(style))
        ws.write_formula(row, 3, s.reg.render(body.format(m=prev)),
                         fmts.calc(style))
        ws.write_formula(row, 4, f"=ROUND(C{row + 1}-D{row + 1},2)",
                         fmts.calc(style))
        ws.write_formula(row, 5,
                         f'=IFERROR((C{row + 1}-D{row + 1})/'
                         f'ABS(D{row + 1}),"")', fmts.calc("percent"))
        if share:
            ws.write_formula(
                row, 6,
                f'=IF(C{sales_row + 1}=0,"",C{row + 1}/'
                f'C{sales_row + 1})', fmts.calc("percent"))
        else:
            s.write(row, 6, None, fmts.calc("percent"))
        row += 1
    month_last = row - 1
    ws.conditional_format(month_first, 4, month_last, 4,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("good")})
    ws.conditional_format(month_first, 4, month_last, 4,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("warn")})
    row += 1

    # ------------------------------------------------- month, by person --
    row = s.subsection(row, "The month, by stylist", LAST_COL)
    s.table_head(row, 1, [
        "Stylist", "Service sales", "Retail sales", "Commission",
        "Tips", "Clients seen", "Average ticket"],
        [20, 15, 14, 14, 12, 13, 14])
    row += 1
    for n in range(1, 9):
        name = (f'IFERROR(INDEX([[Staff.name]],MATCH({n},'
                f'[[Staff.rev_total_rank]],0)),"")')
        who = (f'IFERROR(INDEX([[Staff.id]],MATCH({n},'
               f'[[Staff.rev_total_rank]],0)),"")')
        s.write(row, 1, None, fmts.calc("text"))
        ws.write_formula(row, 1, s.reg.render("=" + name),
                         fmts.calc("text"))
        cells = [
            (2, f'=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],'
                f'ReportMonth,[[Sales.kind]],"Service",'
                f'[[Sales.staff]],{who}),2)', "money"),
            (3, f'=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],'
                f'ReportMonth,[[Sales.kind]],"Retail",'
                f'[[Sales.staff]],{who}),2)', "money"),
            (4, f'=ROUND(SUMIFS([[Commissions.commission]],'
                f'[[Commissions.month]],ReportMonth,'
                f'[[Commissions.staff]],{who}),2)', "money"),
            (5, f'=ROUND(SUMIFS([[Payments.amount]],'
                f'[[Payments.month]],ReportMonth,[[Payments.kind]],'
                f'"Tip",[[Payments.staff]],{who}),2)', "money"),
            (6, f'=COUNTIFS([[Appointments.month]],ReportMonth,'
                f'[[Appointments.status]],"Completed",'
                f'[[Appointments.staff]],{who})', "int"),
        ]
        for col, formula, style in cells:
            ws.write_formula(row, col, s.reg.render(formula),
                             fmts.calc(style))
        ws.write_formula(row, 7,
                         f'=IF(G{row + 1}=0,"",ROUND((C{row + 1}+'
                         f'D{row + 1})/G{row + 1},2))',
                         fmts.calc("money"))
        row += 1
    row += 1

    # ---------------------------------------------- month, top services --
    row = s.subsection(row, "The month's best sellers", LAST_COL)
    s.table_head(row, 1, [
        "Service", "Times", "Revenue", "Client", "Visits", "Spend"],
        [26, 10, 14, 24, 10, 14])
    row += 1
    for n in range(1, TOP_ROWS + 1):
        service = (f'IFERROR(INDEX([[Services.name]],MATCH({n},'
                   f'[[Services.revenue_rank]],0)),"")')
        service_id = (f'IFERROR(INDEX([[Services.id]],MATCH({n},'
                      f'[[Services.revenue_rank]],0)),"")')
        client = (f'IFERROR(INDEX([[Clients.name]],MATCH({n},'
                  f'[[Clients.spend_rank]],0)),"")')
        client_id = (f'IFERROR(INDEX([[Clients.id]],MATCH({n},'
                     f'[[Clients.spend_rank]],0)),"")')
        cells = [
            (1, "=" + service, "text"),
            (2, f'=COUNTIFS([[Sales.month]],ReportMonth,[[Sales.item]],'
                f'{service_id})', "int"),
            (3, f'=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],'
                f'ReportMonth,[[Sales.item]],{service_id}),2)', "money"),
            (4, "=" + client, "text"),
            (5, f'=COUNTIFS([[Appointments.month]],ReportMonth,'
                f'[[Appointments.client]],{client_id},'
                f'[[Appointments.status]],"Completed")', "int"),
            (6, f'=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],'
                f'ReportMonth,[[Sales.client]],{client_id}),2)',
             "money"),
        ]
        for col, formula, style in cells:
            ws.write_formula(row, col, s.reg.render(formula),
                             fmts.calc(style))
        row += 1
    row += 1

    # --------------------------------------------------- year to date ---
    row = s.section(row, "Year to date", LAST_COL)
    s.table_head(row, 1, [
        "Net sales", "Cost of sales", "Staff cost", "Other expenses",
        "Operating profit", "Margin", "Clients seen", "New clients"],
        [16, 15, 14, 15, 16, 10, 13, 13])
    row += 1
    ytd = '">="&DATE(YEAR(TodayDate),1,1)'
    cells = [
        (1, f'=ROUND(SUMIFS([[Sales.net]],[[Sales.date]],{ytd}),2)',
         "money"),
        (2, f'=ROUND(SUMIFS([[Sales.cost]],[[Sales.date]],{ytd}),2)',
         "money"),
        (3, f'=ROUND(SUMIFS([[Payroll.total_cost]],'
            f'[[Payroll.period_end]],{ytd}),2)', "money"),
        (4, f'=ROUND(SUMIFS([[Expenses.amount]],[[Expenses.date]],'
            f'{ytd}),2)', "money"),
        (7, f'=COUNTIFS([[Appointments.status]],"Completed",'
            f'[[Appointments.date]],{ytd})', "int"),
        (8, f'=COUNTIFS([[Clients.first_visit]],{ytd})', "int"),
    ]
    for col, formula, style in cells:
        ws.write_formula(row, col, s.reg.render(formula),
                         fmts.calc(style))
    ws.write_formula(row, 5,
                     f"=ROUND(B{row + 1}-C{row + 1}-D{row + 1}-"
                     f"E{row + 1},2)", fmts.calc("money"))
    ws.write_formula(row, 6,
                     f'=IF(B{row + 1}=0,"",F{row + 1}/B{row + 1})',
                     fmts.calc("percent"))
    row += 2

    row = s.note(
        row,
        "Change the week ending date or the month at the top and every "
        "number on this tab moves with it. Print range is set, so "
        "Ctrl+P gives you something you can hand over.",
        LAST_COL, height=28)

    s.finish(freeze_row=6, gridlines=False, fit_width=True)
    s.protect()
    return s
