"""`Monthly Reports` - one month, on one page you can print.

Pick a month at the top and everything below it moves: enrollment in
and out, attendance room by room, what you billed and what actually
arrived, where the money went, and how the staffing stacked up
against it.

At the bottom sits the printable invoice.  Choose an invoice number
and the whole document builds itself - your details, the family's,
the charge lines behind it, what has been paid and what is left.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, col_letter, quote_sheet
from ..writer import SheetWriter, list_from

NAME = "Monthly Reports"
LAST_COL = 11
ROOM_ROWS = 8
LINE_ROWS = 14
CAT_ROWS = 12


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.deep)
    ws = s.ws
    sheet_ref = quote_sheet(NAME)

    s.banner("Monthly Report",
             "One month of your setting on a single printable page - "
             "children, attendance, money, costs and staffing.",
             LAST_COL)
    s.set_col(0, 0, 24)
    for column in range(1, LAST_COL + 1):
        s.set_col(column, column, 14)

    # ------------------------------------------------------- the picker --
    row = 4
    s.merge(row, 0, 0, "Month", fmts.chrome("label_bold"))
    s.write(row, 1, None, fmts.input("text"))
    ws.write_formula(row, 1, '=TEXT(TODAY(),"yyyy-mm")',
                     fmts.input("text"))
    ws.data_validation(row, 1, row, 1, list_from("MonthKeyList"))
    book.define_name("ReportMonth", f"={sheet_ref}!$B${row + 1}")
    s.merge(row, 2, 2, "Reading", fmts.chrome("label_right"))
    s.merge(row, 3, 4, None, fmts.calc("text"))
    ws.write_formula(
        row, 3,
        '=IFERROR(TEXT(DATEVALUE(ReportMonth&"-01"),"mmmm yyyy"),'
        '"Pick a month")', fmts.calc("text"))
    s.merge(row, 5, 6, "Compared with", fmts.chrome("label_right"))
    s.merge(row, 7, 8, None, fmts.calc("text"))
    ws.write_formula(
        row, 7,
        '=IFERROR(TEXT(DATE(VALUE(LEFT(ReportMonth,4)),'
        'VALUE(RIGHT(ReportMonth,2))-1,1),"yyyy-mm"),"")',
        fmts.calc("text"))
    book.define_name("ReportPrev", f"={sheet_ref}!$H${row + 1}")
    # Two helper cells, off to the right and hidden: the first and the
    # last day of the chosen month. Counting dated rows with COUNTIFS
    # over that window is faster than TEXT() across a whole column and
    # it never trips over an empty cell.
    s.set_col(13, 14, 11, None, hidden=True)
    ws.write_formula(
        row, 13,
        '=IFERROR(DATE(VALUE(LEFT(ReportMonth,4)),'
        'VALUE(RIGHT(ReportMonth,2)),1),TodayDate)',
        fmts.calc("date"))
    book.define_name("ReportFrom", f"={sheet_ref}!$N${row + 1}")
    ws.write_formula(
        row, 14,
        '=IFERROR(DATE(VALUE(LEFT(ReportMonth,4)),'
        'VALUE(RIGHT(ReportMonth,2))+1,1)-1,TodayDate)',
        fmts.calc("date"))
    book.define_name("ReportTo", f"={sheet_ref}!$O${row + 1}")
    s.merge(row, 9, 9, "Printed", fmts.chrome("label_right"))
    s.merge(row, 10, LAST_COL, None, fmts.calc("date"))
    ws.write_formula(row, 10, "=TodayDate", fmts.calc("date"))
    row += 2

    row = s.note(
        row,
        "Change the month in the blue cell and every figure on this "
        "page follows it. Months are plain text in yyyy-mm form, "
        "which is what every monthly total in the workbook groups on.",
        LAST_COL, height=28)
    row += 1

    billed = ('SUMIFS([[Invoices & Receipts.billed]],'
              '[[Invoices & Receipts.month]],ReportMonth)')
    billed_prev = ('SUMIFS([[Invoices & Receipts.billed]],'
                   '[[Invoices & Receipts.month]],ReportPrev)')
    collected = ('SUMIFS([[Invoices & Receipts.collected]],'
                 '[[Invoices & Receipts.month]],ReportMonth)')
    wages = ('SUMIFS([[Payroll Tracker.total_cost]],'
             '[[Payroll Tracker.month]],ReportMonth)')
    costs = ('SUMIFS([[Expense Tracker.pl_amount]],'
             '[[Expense Tracker.month]],ReportMonth)')
    sessions = ('SUMIFS([[Attendance.session]],'
                '[[Attendance.month]],ReportMonth)')
    present = ('SUMIFS([[Attendance.present]],'
               '[[Attendance.month]],ReportMonth)')

    # ----------------------------------------------------- the headlines --
    row = s.section(row, "The month in nine numbers", LAST_COL)
    tiles = [
        ("Billed", f"=ROUND({billed},2)", "money0",
         "invoices less credits and write-offs"),
        ("Collected", f"=ROUND({collected},2)", "money0",
         "payments less refunds"),
        ("Operating profit", f"=ROUND({billed}-{wages}-{costs},2)",
         "money0", "after wages and costs"),
        ("Margin",
         f'=IF({billed}<=0,"",({billed}-{wages}-{costs})/{billed})',
         "percent", "of everything billed"),
        ("Attendance",
         f'=IF({sessions}<=0,"",{present}/{sessions})', "percent",
         f"target {cfg.target_attendance:.0%}"),
        ("Sessions attended", f"={present}", "int",
         "closures and holidays excluded"),
        ("Children billed",
         '=COUNTIFS([[Invoices & Receipts.month]],ReportMonth,'
         '[[Invoices & Receipts.type]],"Invoice")', "int",
         "one invoice per family"),
        ("Collection rate",
         f'=IF({billed}<=0,"",{collected}/{billed})', "percent",
         f"target {cfg.target_collection:.0%}"),
        ("Against last month",
         f'=IF({billed_prev}<=0,"",{billed}/{billed_prev}-1)',
         "percent", "billed, month on month"),
    ]
    column = 0
    bottom = row
    for index, (label, formula, style, note) in enumerate(tiles):
        if index and index % 3 == 0:
            row = bottom
            column = 0
        bottom = s.kpi(row, column, label, formula, style, span=4,
                       note=note)
        column += 4
    row = bottom + 1

    # ---------------------------------------------------- enrollment -----
    row = s.section(row, "Children in and out", LAST_COL)
    row = s.table_head(row, 0, [
        "On the books", "Attending", "Started this month",
        "Left this month", "Net change", "Places free",
        "Occupancy", "On the waitlist", "Offers out",
        "Enquiries this month", "Converted"],
        [14, 12, 15, 14, 12, 12, 11, 14, 11, 16, 12])
    s.calc_row(row, 0, [
        "=ChildrenEnrolled",
        "=ChildrenActive",
        '=COUNTIFS([[Child Enrollment.start_date]],">="&ReportFrom,[[Child Enrollment.start_date]],"<="&ReportTo)',
        '=COUNTIFS([[Child Enrollment.end_date]],">="&ReportFrom,[[Child Enrollment.end_date]],"<="&ReportTo)',
        '=COUNTIFS([[Child Enrollment.start_date]],">="&ReportFrom,[[Child Enrollment.start_date]],"<="&ReportTo)-COUNTIFS([[Child Enrollment.end_date]],">="&ReportFrom,[[Child Enrollment.end_date]],"<="&ReportTo)',
        "=MAX(0,RoomCapacity-ChildrenEnrolled)",
        '=IF(RoomCapacity<=0,"",ChildrenEnrolled/RoomCapacity)',
        '=COUNTIFS([[Waitlist.closed]],"No")',
        '=COUNTIFS([[Waitlist.status]],"Place offered")',
        '=COUNTIFS([[Waitlist.enquiry_date]],">="&ReportFrom,[[Waitlist.enquiry_date]],"<="&ReportTo)',
        '=COUNTIFS([[Waitlist.converted]],"Yes")',
    ], ["int", "int", "int", "int", "int", "int", "percent", "int",
        "int", "int", "int"])
    row += 2

    # ------------------------------------------------- room by room ------
    row = s.section(row, "Room by room", LAST_COL)
    row = s.table_head(row, 0, [
        "Room", "Places", "Enrolled", "Occupancy", "Sessions",
        "Attended", "Absent", "Attendance", "Weekly income",
        "Income a place", "Empty seats cost"],
        [18, 10, 11, 11, 11, 11, 10, 12, 14, 14, 15])
    rooms_first = row
    for n in range(1, ROOM_ROWS + 1):
        name = f'INDEX([[Classroom Management.name]],{n})'
        guard = f'IFERROR(IF({name}="","",{{body}}),"")'
        sess = (f'SUMIFS([[Attendance.session]],[[Attendance.room]],'
                f'{name},[[Attendance.month]],ReportMonth)')
        here = (f'SUMIFS([[Attendance.present]],[[Attendance.room]],'
                f'{name},[[Attendance.month]],ReportMonth)')
        away = (f'SUMIFS([[Attendance.absent]],[[Attendance.room]],'
                f'{name},[[Attendance.month]],ReportMonth)')
        cells = [
            (0, f'IFERROR({name}&"","")', "text"),
            (1, guard.replace(
                "{body}", f'INDEX([[Classroom Management.capacity]],'
                          f'{n})'), "int"),
            (2, guard.replace(
                "{body}", f'INDEX([[Classroom Management.enrolled]],'
                          f'{n})'), "int"),
            (3, guard.replace(
                "{body}", f'INDEX([[Classroom Management.occupancy]],'
                          f'{n})'), "percent"),
            (4, guard.replace("{body}", sess), "int"),
            (5, guard.replace("{body}", here), "int"),
            (6, guard.replace("{body}", away), "int"),
            (7, guard.replace(
                "{body}", f'IF({sess}<=0,"",{here}/{sess})'),
             "percent"),
            (8, guard.replace(
                "{body}", f'INDEX([[Classroom Management.revenue]],'
                          f'{n})'), "money"),
            (9, guard.replace(
                "{body}", f'INDEX([[Classroom Management.revenue_seat]]'
                          f',{n})'), "money"),
            (10, guard.replace(
                "{body}", f'INDEX([[Classroom Management.empty_cost]],'
                          f'{n})'), "money"),
        ]
        for column, formula, style in cells:
            ws.write_formula(row, column, s.reg.render("=" + formula),
                             fmts.calc(style))
        row += 1
    ws.conditional_format(rooms_first, 7, row - 1, 7,
                          {"type": "cell", "criteria": "<",
                           "value": "=TargetAttendance",
                           "format": fmts.cf("warn")})
    ws.conditional_format(rooms_first, 3, row - 1, 3,
                          {"type": "data_bar",
                           "bar_color": Palette.primary_mid})
    row += 1

    # ------------------------------------------------------ the money ----
    row = s.section(row, "Money in", LAST_COL)
    row = s.table_head(row, 0, [
        "Invoiced", "Credit notes", "Written off", "Billed",
        "Payments", "Refunds", "Collected", "Card fees",
        "Still owed, all time", "Overdue invoices", "Collection rate"],
        [13, 13, 12, 13, 12, 11, 12, 11, 17, 15, 13])
    s.calc_row(row, 0, [
        '=ROUND(SUMIFS([[Invoices & Receipts.total]],'
        '[[Invoices & Receipts.month]],ReportMonth,'
        '[[Invoices & Receipts.type]],"Invoice"),2)',
        '=ROUND(SUMIFS([[Invoices & Receipts.total]],'
        '[[Invoices & Receipts.month]],ReportMonth,'
        '[[Invoices & Receipts.type]],"Credit note"),2)',
        '=ROUND(SUMIFS([[Invoices & Receipts.total]],'
        '[[Invoices & Receipts.month]],ReportMonth,'
        '[[Invoices & Receipts.type]],"Write-off"),2)',
        f"=ROUND({billed},2)",
        '=ROUND(SUMIFS([[Invoices & Receipts.total]],'
        '[[Invoices & Receipts.month]],ReportMonth,'
        '[[Invoices & Receipts.type]],"Payment"),2)',
        '=ROUND(SUMIFS([[Invoices & Receipts.total]],'
        '[[Invoices & Receipts.month]],ReportMonth,'
        '[[Invoices & Receipts.type]],"Refund"),2)',
        f"=ROUND({collected},2)",
        '=ROUND(SUMIFS([[Invoices & Receipts.fee]],'
        '[[Invoices & Receipts.month]],ReportMonth),2)',
        "=ROUND(SUM([[Parent Database.balance]]),2)",
        '=COUNTIFS([[Invoices & Receipts.status]],"Overdue")',
        f'=IF({billed}<=0,"",{collected}/{billed})',
    ], ["money", "money", "money", "money", "money", "money", "money",
        "money", "money", "int", "percent"])
    row += 2

    # -------------------------------------------------- what was charged --
    row = s.section(row, "What was charged, by type", LAST_COL)
    row = s.table_head(row, 0, [
        "Tuition", "Registration", "Late pickup", "Meals",
        "Transport", "Trips & activities", "Other charges",
        "Discounts given", "Subsidies", "Net charges",
        "Not yet invoiced"],
        [13, 13, 12, 11, 11, 15, 13, 14, 12, 12, 14])
    charge = ('SUMIFS([[Tuition & Fees.net]],[[Tuition & Fees.month]],'
              'ReportMonth,[[Tuition & Fees.fee_type]],"{t}")')
    others = ("+".join(
        charge.replace("{t}", kind) for kind in
        ("Snack fee", "Activity fee", "Supply fee",
         "Nappy & wipes fee", "Holiday club", "Extra day",
         "Extra hours", "Late payment fee", "Other charge")))
    discounts = ("+".join(
        charge.replace("{t}", kind) for kind in
        ("Sibling discount", "Staff discount", "Promotional discount",
         "Goodwill credit")))
    subsidies = ("+".join(
        charge.replace("{t}", kind) for kind in
        ("Subsidy credit", "Scholarship credit")))
    s.calc_row(row, 0, [
        "=ROUND(" + charge.replace("{t}", "Tuition") + ",2)",
        "=ROUND(" + charge.replace("{t}", "Registration fee")
        + "+" + charge.replace("{t}", "Enrollment deposit") + ",2)",
        "=ROUND(" + charge.replace("{t}", "Late pickup fee") + ",2)",
        "=ROUND(" + charge.replace("{t}", "Meal fee") + ",2)",
        "=ROUND(" + charge.replace("{t}", "Transport fee") + ",2)",
        "=ROUND(" + charge.replace("{t}", "Field trip") + ",2)",
        "=ROUND(" + others + ",2)",
        "=ROUND(" + discounts + ",2)",
        "=ROUND(" + subsidies + ",2)",
        '=ROUND(SUMIFS([[Tuition & Fees.net]],'
        '[[Tuition & Fees.month]],ReportMonth),2)',
        '=ROUND(SUMIFS([[Tuition & Fees.net]],'
        '[[Tuition & Fees.month]],ReportMonth,'
        '[[Tuition & Fees.billed]],"No"),2)',
    ], ["money"] * 11)
    row += 2

    # ----------------------------------------------------- money out -----
    row = s.section(row, "Money out", LAST_COL)
    row = s.table_head(row, 0, [
        "Wages & on-costs", "Rota hours", "Overtime hours",
        "Wage cost of income", "Other costs", "Budget",
        "Over or under", "Cost a child", "Operating profit",
        "Margin", "Profit a child"],
        [16, 12, 14, 17, 13, 12, 14, 12, 15, 10, 13])
    s.calc_row(row, 0, [
        f"=ROUND({wages},2)",
        '=ROUND(SUMIFS([[Payroll Tracker.total_hours]],'
        '[[Payroll Tracker.month]],ReportMonth),1)',
        '=ROUND(SUMIFS([[Payroll Tracker.ot_hours]],'
        '[[Payroll Tracker.month]],ReportMonth),1)',
        f'=IF({billed}<=0,"",{wages}/{billed})',
        f"=ROUND({costs},2)",
        "=ROUND(BudgetMonthlyTotal,2)",
        f"=ROUND({costs}-BudgetMonthlyTotal,2)",
        f'=IF(ChildrenEnrolled<=0,"",ROUND(({wages}+{costs})/'
        f'ChildrenEnrolled,2))',
        f"=ROUND({billed}-{wages}-{costs},2)",
        f'=IF({billed}<=0,"",({billed}-{wages}-{costs})/{billed})',
        f'=IF(ChildrenEnrolled<=0,"",ROUND(({billed}-{wages}-{costs})/'
        f'ChildrenEnrolled,2))',
    ], ["money", "number1", "number1", "percent", "money", "money",
        "money", "money", "money", "percent", "money"])
    row += 2

    # ------------------------------------------- expenses by category ----
    row = s.section(row, "Where the money went", LAST_COL)
    row = s.table_head(row, 0, [
        "Category", "Spent this month", "Budget", "Over or under",
        "Share of costs", "Category", "Spent this month", "Budget",
        "Over or under", "Share of costs"],
        [22, 15, 12, 14, 13, 22, 15, 12, 14, 13])
    cat_first = row
    for n in range(1, CAT_ROWS + 1):
        for half in (0, 1):
            index = n + half * CAT_ROWS
            base = half * 5
            cat = f'INDEX(BudgetCategory,{index})'
            spent = (f'SUMIFS([[Expense Tracker.pl_amount]],'
                     f'[[Expense Tracker.month]],ReportMonth,'
                     f'[[Expense Tracker.category]],{cat})')
            budget = f'INDEX(BudgetAmount,{index})'
            cells = [
                (base + 0, f'IFERROR({cat}&"","")', "text"),
                (base + 1, f'IFERROR(ROUND({spent},2),"")', "money"),
                (base + 2, f'IFERROR(ROUND({budget},2),"")', "money"),
                (base + 3, f'IFERROR(ROUND({spent}-{budget},2),"")',
                 "money"),
                (base + 4, f'IFERROR(IF({costs}<=0,"",{spent}/{costs}),'
                           f'"")', "percent"),
            ]
            for column, formula, style in cells:
                ws.write_formula(row, column,
                                 s.reg.render("=" + formula),
                                 fmts.calc(style))
        row += 1
    for column in (3, 8):
        ws.conditional_format(cat_first, column, row - 1, column,
                              {"type": "cell", "criteria": ">",
                               "value": 0.004,
                               "format": fmts.cf("bad")})
        ws.conditional_format(cat_first, column, row - 1, column,
                              {"type": "cell", "criteria": "<=",
                               "value": 0, "format": fmts.cf("good")})
    row += 1

    # ------------------------------------------------- care and safety ---
    row = s.section(row, "Care, safety and paperwork", LAST_COL)
    row = s.table_head(row, 0, [
        "Daily reports written", "Meals logged", "Meals refused",
        "Medication doses", "Incidents", "Of those, serious",
        "Average minutes to tell a parent", "Forms signed",
        "Parent conversations", "Complaints", "Compliments"],
        [18, 12, 13, 15, 11, 15, 22, 12, 17, 12, 13])
    s.calc_row(row, 0, [
        '=COUNTIFS([[Daily Reports.date]],">="&ReportFrom,[[Daily Reports.date]],"<="&ReportTo)',
        '=COUNTIFS([[Meal Tracker.month]],ReportMonth)',
        '=COUNTIFS([[Meal Tracker.month]],ReportMonth,'
        '[[Meal Tracker.refused]],"Yes")',
        '=COUNTIFS([[Allergy & Medical.given_date]],">="&ReportFrom,[[Allergy & Medical.given_date]],"<="&ReportTo)',
        '=COUNTIFS([[Incident Reports.month]],ReportMonth)',
        '=COUNTIFS([[Incident Reports.month]],ReportMonth,'
        '[[Incident Reports.severity]],"Serious")+'
        'COUNTIFS([[Incident Reports.month]],ReportMonth,'
        '[[Incident Reports.severity]],"Emergency services called")',
        '=IFERROR(ROUND(SUMIFS([[Incident Reports.notify_mins]],'
        '[[Incident Reports.month]],ReportMonth)/'
        'MAX(1,COUNTIFS([[Incident Reports.month]],ReportMonth,'
        '[[Incident Reports.notified]],"Yes")),0),0)',
        '=COUNTIFS([[Incident Reports.month]],ReportMonth,'
        '[[Incident Reports.signed]],"Yes")',
        '=COUNTIFS([[Parent Communication.month]],ReportMonth)',
        '=COUNTIFS([[Parent Communication.month]],ReportMonth,'
        '[[Parent Communication.category]],"Complaint")',
        '=COUNTIFS([[Parent Communication.month]],ReportMonth,'
        '[[Parent Communication.category]],"Compliment")',
    ], ["int", "int", "int", "int", "int", "int", "int", "int", "int",
        "int", "int"])
    row += 2

    # ------------------------------------------------ printable invoice --
    row = s.section(row, "Printable invoice or receipt", LAST_COL)
    row = s.note(
        row,
        "Pick a reference from the drop-down and the document builds "
        "itself from the ledger and the charge lines behind it. Print "
        "this block - the print area is set to end here.",
        LAST_COL, height=28)
    row += 1

    s.merge(row, 0, 0, "Reference", fmts.chrome("label_bold"))
    s.write(row, 1, None, fmts.input("text"))
    ws.data_validation(row, 1, row, 1, list_from("InvoiceList"))
    book.define_name("PrintInvoice", f"={sheet_ref}!$B${row + 1}")
    s.merge(row, 2, 3, "Choose an invoice number", fmts.chrome("note"))
    row += 2

    def look(field: str) -> str:
        return (f'IFERROR(INDEX([[Invoices & Receipts.{field}]],'
                f'MATCH(PrintInvoice,[[Invoices & Receipts.id]],0)),"")')

    s.merge(row, 0, 5, "", fmts.chrome("print_title"))
    ws.write_formula(row, 0, "=CentreName", fmts.chrome("print_title"))
    s.merge(row, 6, LAST_COL, None, fmts.chrome("print_sub"))
    ws.write_formula(
        row, 6,
        s.reg.render('=IF(PrintInvoice="","INVOICE",'
                     'IF(LEFT(PrintInvoice,LEN(ReceiptPrefix))'
                     '=ReceiptPrefix,"RECEIPT","INVOICE"))'),
        fmts.chrome("print_sub"))
    row += 1

    left = [
        ("From", "=CentreName"),
        ("", "=OwnerName"),
        ("Reference", "=PrintInvoice"),
        ("Date", f"={look('date')}"),
        ("Period", f"={look('period')}"),
    ]
    right = [
        ("Bill to", f"={look('family_name')}"),
        ("Family ID", f"={look('family')}"),
        ("Type", f"={look('type')}"),
        ("Due", f"={look('due_date')}"),
        ("Status", f"={look('status')}"),
    ]
    # the ruled lines are filled in by the ledger, not by hand, so they
    # stay locked even though they are drawn like a paper form
    filled = fmts.chrome("form_line", locked=True)
    for index in range(len(left)):
        label, formula = left[index]
        s.write(row, 0, label, fmts.chrome("form_label"))
        s.merge(row, 1, 4, None, filled)
        ws.write_formula(row, 1, s.reg.render(formula), filled)
        label, formula = right[index]
        s.write(row, 5, label, fmts.chrome("form_label"))
        s.merge(row, 6, LAST_COL, None, filled)
        ws.write_formula(row, 6, s.reg.render(formula), filled)
        row += 1
    row += 1

    row = s.table_head(row, 0, [
        "Line", "Child", "Description", "Period", "Quantity", "Rate",
        "Discount", "Tax", "Amount"],
        [7, 18, 30, 20, 10, 12, 11, 10, 13])
    for n in range(1, LINE_ROWS + 1):
        key = f'PrintInvoice&"|"&{n}'

        def line(field: str) -> str:
            return (f'IFERROR(INDEX([[Tuition & Fees.{field}]],'
                    f'MATCH({key},[[Tuition & Fees.key]],0)),"")')

        cells = [
            (0, f'IF({line("description")}="","",{n})', "int"),
            (1, line("child_name"), "text"),
            (2, line("description"), "text"),
            (3, f'IF({line("period_from")}="","",'
                f'TEXT({line("period_from")},"d mmm")&" - "&'
                f'TEXT({line("period_to")},"d mmm"))', "text"),
            (4, line("qty"), "number1"),
            (5, line("rate"), "money"),
            (6, line("discount"), "money"),
            (7, line("tax"), "money"),
            (8, line("net"), "money"),
        ]
        for column, formula, style in cells:
            ws.write_formula(row, column, s.reg.render("=" + formula),
                             fmts.calc(style))
        row += 1
    row += 1

    totals = [
        ("Charge lines", f"={look('charges')}", "money"),
        ("Invoice total", f"={look('total')}", "money"),
        ("Paid", f"={look('paid')}", "money"),
        ("Credited", f"={look('credited')}", "money"),
        ("Balance", f"={look('balance')}", "money"),
    ]
    for label, formula, style in totals:
        s.merge(row, 5, 7, label, fmts.chrome("total_label"))
        s.merge(row, 8, LAST_COL, None,
                fmts.chrome("total_value",
                            num_format=cfg.money_format))
        ws.write_formula(row, 8, s.reg.render(formula),
                         fmts.chrome("total_value",
                                     num_format=cfg.money_format))
        row += 1
    ws.conditional_format(row - 1, 8, row - 1, 8,
                          {"type": "cell", "criteria": ">",
                           "value": 0.004, "format": fmts.cf("bad")})
    row += 1

    s.merge(row, 0, LAST_COL, None, fmts.chrome("note_box"))
    ws.write_formula(
        row, 0,
        s.reg.render('="Payment terms "&PaymentTermsDays&" days. '
                     'Please quote "&IF(PrintInvoice="","the invoice '
                     'number",PrintInvoice)&" with your payment. '
                     'Questions about this bill? Call "&CentreName&'
                     '" and ask for "&OwnerName&"."'),
        fmts.chrome("note_box"))
    s.row_height(row, 26)
    print_last = row
    row += 2

    row = s.note(
        row,
        "Billed and collected are deliberately two different "
        "figures. Billed is revenue - what you charged, less credits "
        "and write-offs. Collected is cash - what actually arrived, "
        "less refunds. A month can be good on one and poor on the "
        "other, and knowing which is the whole point of keeping them "
        "apart.",
        LAST_COL, height=38)

    s.finish(freeze_row=6, gridlines=False, fit_width=True,
             print_area=f"A1:{col_letter(LAST_COL)}{print_last + 1}")
    s.protect()
    return s
