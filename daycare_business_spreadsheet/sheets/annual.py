"""`Annual Reports` - twelve months side by side, and what they add up to.

Pick a year and read across: what you billed and collected each
month, what the staff cost, what everything else cost, what was left,
how full the rooms were and how many children were on the books.

Underneath sits the part that actually changes decisions - the fee
mix for the year, the expense categories ranked by how much they took,
the children and families who contribute most, and a straight
comparison with the year before.
"""

from __future__ import annotations

from ..config import MONTH_NAMES, Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, number_check

NAME = "Annual Reports"
LAST_COL = 12
TOP_ROWS = 10
CAT_ROWS = 15


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.deep)
    ws = s.ws
    sheet_ref = quote_sheet(NAME)

    s.banner("Annual Report",
             "Twelve months of the setting in one table, with the "
             "totals, the averages and the comparison with last "
             "year.", LAST_COL)
    s.set_col(0, 0, 22)
    for column in range(1, LAST_COL + 1):
        s.set_col(column, column, 14)

    # -------------------------------------------------------- the picker --
    row = 4
    s.merge(row, 0, 0, "Year", fmts.chrome("label_bold"))
    s.write(row, 1, None, fmts.input("int"))
    ws.write_formula(row, 1, "=YEAR(TODAY())", fmts.input("int"))
    ws.data_validation(row, 1, row, 1,
                       number_check(2000, 2100, decimal=False))
    book.define_name("ReportYear", f"={sheet_ref}!$B${row + 1}")
    s.merge(row, 2, 3, "Compared with", fmts.chrome("label_right"))
    s.write(row, 4, None, fmts.calc("int"))
    ws.write_formula(row, 4, "=ReportYear-1", fmts.calc("int"))
    book.define_name("ReportYearPrev", f"={sheet_ref}!$E${row + 1}")
    s.merge(row, 5, 6, "Weeks open a year", fmts.chrome("label_right"))
    s.write(row, 7, None, fmts.calc("int"))
    ws.write_formula(row, 7, "=WeeksOpen", fmts.calc("int"))
    s.merge(row, 8, 9, "Printed", fmts.chrome("label_right"))
    s.merge(row, 10, LAST_COL, None, fmts.calc("date"))
    ws.write_formula(row, 10, "=TodayDate", fmts.calc("date"))
    row += 2

    row = s.note(
        row,
        "Type a year in the blue cell. Every figure on this page is "
        "summed from the month keys that begin with it, so a month "
        "you have not logged yet simply reads zero rather than "
        "breaking the total.",
        LAST_COL, height=28)
    row += 1

    def mk(index: int) -> str:
        """The yyyy-mm key for month `index` (1-12) of the chosen year."""
        return f'ReportYear&"-"&TEXT({index},"00")'

    def year_sum(sheet: str, value: str, key: str,
                 year: str = "ReportYear") -> str:
        """Sum one register column across the twelve months of a year.

        The result is wrapped in brackets: these totals get divided by
        a target or a headcount, and Excel divides before it adds.
        """
        return "(" + "+".join(
            f'SUMIFS([[{sheet}.{value}]],[[{sheet}.{key}]],'
            f'{year}&"-"&TEXT({m},"00"))' for m in range(1, 13)) + ")"

    billed_year = year_sum("Invoices & Receipts", "billed", "month")
    billed_prev = year_sum("Invoices & Receipts", "billed", "month",
                           "ReportYearPrev")
    collected_year = year_sum("Invoices & Receipts", "collected",
                              "month")
    wages_year = year_sum("Payroll Tracker", "total_cost", "month")
    wages_prev = year_sum("Payroll Tracker", "total_cost", "month",
                          "ReportYearPrev")
    costs_year = year_sum("Expense Tracker", "pl_amount", "month")
    costs_prev = year_sum("Expense Tracker", "pl_amount", "month",
                          "ReportYearPrev")
    sessions_year = year_sum("Attendance", "session", "month")
    present_year = year_sum("Attendance", "present", "month")

    # ------------------------------------------------------ the headlines -
    row = s.section(row, "The year in eight numbers", LAST_COL)
    tiles = [
        ("Billed", f"=ROUND({billed_year},2)", "money0",
         f"target {cfg.currency_symbol}"
         f"{cfg.revenue_target_year:,.0f}"),
        ("Against target",
         f'=IF(RevenueTargetYear<=0,"",{billed_year}/'
         f'RevenueTargetYear)', "percent0", "100% makes the year"),
        ("Collected", f"=ROUND({collected_year},2)", "money0",
         "cash that arrived"),
        ("Collection rate",
         f'=IF({billed_year}<=0,"",{collected_year}/{billed_year})',
         "percent", f"target {cfg.target_collection:.0%}"),
        ("Wages", f"=ROUND({wages_year},2)", "money0",
         "including employer costs"),
        ("Other costs", f"=ROUND({costs_year},2)", "money0",
         "everything that is not wages"),
        ("Operating profit",
         f"=ROUND({billed_year}-{wages_year}-{costs_year},2)",
         "money0", "what the year left you"),
        ("Margin",
         f'=IF({billed_year}<=0,"",({billed_year}-{wages_year}-'
         f'{costs_year})/{billed_year})', "percent",
         "of everything billed"),
    ]
    column = 0
    bottom = row
    for index, (label, formula, style, note) in enumerate(tiles):
        if index and index % 4 == 0:
            row = bottom
            column = 0
        bottom = s.kpi(row, column, label, formula, style, span=3,
                       note=note)
        column += 3
    row = bottom + 1

    # -------------------------------------------------- month by month ---
    row = s.section(row, "Month by month", LAST_COL)
    row = s.table_head(row, 0, [
        "Month", "Billed", "Collected", "Still to collect", "Wages",
        "Other costs", "Operating profit", "Margin",
        "Children billed", "Sessions", "Attendance", "Incidents",
        "New starters"],
        [12, 13, 13, 14, 12, 12, 15, 10, 13, 10, 11, 10, 12])
    grid_first = row
    for index in range(1, 13):
        key = mk(index)
        billed = (f'SUMIFS([[Invoices & Receipts.billed]],'
                  f'[[Invoices & Receipts.month]],{key})')
        collected = (f'SUMIFS([[Invoices & Receipts.collected]],'
                     f'[[Invoices & Receipts.month]],{key})')
        wages = (f'SUMIFS([[Payroll Tracker.total_cost]],'
                 f'[[Payroll Tracker.month]],{key})')
        costs = (f'SUMIFS([[Expense Tracker.pl_amount]],'
                 f'[[Expense Tracker.month]],{key})')
        sessions = (f'SUMIFS([[Attendance.session]],'
                    f'[[Attendance.month]],{key})')
        present = (f'SUMIFS([[Attendance.present]],'
                   f'[[Attendance.month]],{key})')
        s.write(row, 0, MONTH_NAMES[index - 1], fmts.chrome("label_bold"))
        formulas = [
            (1, f"=ROUND({billed},2)", "money"),
            (2, f"=ROUND({collected},2)", "money"),
            (3, f"=ROUND({billed}-{collected},2)", "money"),
            (4, f"=ROUND({wages},2)", "money"),
            (5, f"=ROUND({costs},2)", "money"),
            (6, f"=ROUND({billed}-{wages}-{costs},2)", "money"),
            (7, f'=IF({billed}<=0,"",({billed}-{wages}-{costs})/'
                f'{billed})', "percent"),
            (8, f'=COUNTIFS([[Invoices & Receipts.month]],{key},'
                f'[[Invoices & Receipts.type]],"Invoice")', "int"),
            (9, f"={sessions}", "int"),
            (10, f'=IF({sessions}<=0,"",{present}/{sessions})',
             "percent"),
            (11, f'=COUNTIFS([[Incident Reports.month]],{key})', "int"),
            (12, f'=COUNTIFS([[Child Enrollment.start_date]],">="&'
                 f'DATE(ReportYear,{index},1),'
                 f'[[Child Enrollment.start_date]],"<"&'
                 f'DATE(ReportYear,{index}+1,1))', "int"),
        ]
        for column, formula, style in formulas:
            ws.write_formula(row, column, s.reg.render(formula),
                             fmts.calc(style))
        row += 1
    grid_last = row - 1

    # the totals line
    s.write(row, 0, "Year", fmts.chrome("total_label"))
    totals = [
        (1, f"=ROUND({billed_year},2)", "money"),
        (2, f"=ROUND({collected_year},2)", "money"),
        (3, f"=ROUND({billed_year}-{collected_year},2)", "money"),
        (4, f"=ROUND({wages_year},2)", "money"),
        (5, f"=ROUND({costs_year},2)", "money"),
        (6, f"=ROUND({billed_year}-{wages_year}-{costs_year},2)",
         "money"),
        (7, f'=IF({billed_year}<=0,"",({billed_year}-{wages_year}-'
            f'{costs_year})/{billed_year})', "percent"),
        (8, f'=SUM({_rng(grid_first, grid_last, "I")})', "int"),
        (9, f"={sessions_year}", "int"),
        (10, f'=IF({sessions_year}<=0,"",{present_year}/'
             f'{sessions_year})', "percent"),
        (11, f'=SUM({_rng(grid_first, grid_last, "L")})', "int"),
        (12, f'=SUM({_rng(grid_first, grid_last, "M")})', "int"),
    ]
    for column, formula, style in totals:
        ws.write_formula(
            row, column, s.reg.render(formula),
            fmts.chrome("total_value",
                        num_format=s.num_format(style)))
    row += 1
    ws.conditional_format(grid_first, 1, grid_last, 1,
                          {"type": "data_bar",
                           "bar_color": Palette.primary_mid})
    ws.conditional_format(grid_first, 6, grid_last, 6,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(grid_first, 10, grid_last, 10,
                          {"type": "cell", "criteria": "<",
                           "value": "=TargetAttendance",
                           "format": fmts.cf("warn")})

    ws.add_sparkline(row, 1, {
        "range": f"{sheet_ref}!$B${grid_first + 1}:$B${grid_last + 1}",
        "type": "column", "series_color": Palette.primary_mid})
    s.merge(row, 2, 3, "Billed by month", fmts.chrome("note"))
    ws.add_sparkline(row, 4, {
        "range": f"{sheet_ref}!$G${grid_first + 1}:$G${grid_last + 1}",
        "type": "column", "series_color": Palette.accent,
        "negative_points": True})
    s.merge(row, 5, 6, "Operating profit by month",
            fmts.chrome("note"))
    ws.add_sparkline(row, 7, {
        "range": f"{sheet_ref}!$K${grid_first + 1}:$K${grid_last + 1}",
        "type": "line", "series_color": Palette.deep, "markers": True})
    s.merge(row, 8, LAST_COL, "Attendance by month",
            fmts.chrome("note"))
    row += 2

    # ------------------------------------------------- this year v last --
    row = s.section(row, "This year against last", LAST_COL)
    row = s.table_head(row, 0, [
        "Figure", "This year", "Last year", "Change", "Change %",
        "What it tells you"],
        [22, 15, 15, 14, 12, 46])
    compare = [
        ("Billed", billed_year, billed_prev,
         "Revenue. Growth here without growth in places means your "
         "fees went up."),
        ("Wages & on-costs", wages_year, wages_prev,
         "The biggest line in any setting. Watch it as a share of "
         "billing, not on its own."),
        ("Other costs", costs_year, costs_prev,
         "Rent, food, nappies, insurance and everything else."),
        ("Operating profit",
         f"({billed_year})-({wages_year})-({costs_year})",
         f"({billed_prev})-({wages_prev})-({costs_prev})",
         "Billed less wages less costs. The number that decides "
         "whether next year happens."),
    ]
    for label, this_year, last_year, meaning in compare:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        ws.write_formula(row, 1, s.reg.render(f"=ROUND({this_year},2)"),
                         fmts.calc("money"))
        ws.write_formula(row, 2, s.reg.render(f"=ROUND({last_year},2)"),
                         fmts.calc("money"))
        ws.write_formula(
            row, 3,
            s.reg.render(f"=ROUND(({this_year})-({last_year}),2)"),
            fmts.calc("money"))
        ws.write_formula(
            row, 4,
            s.reg.render(f'=IF(({last_year})<=0,"",({this_year})/'
                         f'({last_year})-1)'),
            fmts.calc("percent"))
        s.merge(row, 5, LAST_COL, meaning, fmts.chrome("note"))
        s.row_height(row, 28)
        row += 1
    row += 1

    # ------------------------------------------------------- the fee mix --
    row = s.section(row, "What the year was billed for", LAST_COL)
    row = s.table_head(row, 0, [
        "Tuition", "Registration & deposits", "Late pickup",
        "Meals & snacks", "Transport", "Trips & activities",
        "Other charges", "Discounts", "Subsidies", "Net charges",
        "Average a child"],
        [14, 18, 12, 14, 12, 15, 13, 12, 12, 13, 14])

    def fee_year(*kinds: str) -> str:
        parts = []
        for kind in kinds:
            parts.append("+".join(
                f'SUMIFS([[Tuition & Fees.net]],'
                f'[[Tuition & Fees.month]],ReportYear&"-"&'
                f'TEXT({m},"00"),[[Tuition & Fees.fee_type]],'
                f'"{kind}")' for m in range(1, 13)))
        return "+".join(parts)

    net_charges = year_sum("Tuition & Fees", "net", "month")
    s.calc_row(row, 0, [
        "=ROUND(" + fee_year("Tuition") + ",2)",
        "=ROUND(" + fee_year("Registration fee",
                             "Enrollment deposit") + ",2)",
        "=ROUND(" + fee_year("Late pickup fee") + ",2)",
        "=ROUND(" + fee_year("Meal fee", "Snack fee") + ",2)",
        "=ROUND(" + fee_year("Transport fee") + ",2)",
        "=ROUND(" + fee_year("Field trip", "Activity fee") + ",2)",
        "=ROUND(" + fee_year("Extra day", "Extra hours",
                             "Holiday club", "Supply fee",
                             "Nappy & wipes fee", "Late payment fee",
                             "Other charge") + ",2)",
        "=ROUND(" + fee_year("Sibling discount", "Staff discount",
                             "Promotional discount",
                             "Goodwill credit") + ",2)",
        "=ROUND(" + fee_year("Subsidy credit",
                             "Scholarship credit") + ",2)",
        f"=ROUND({net_charges},2)",
        f'=IF(ChildrenEnrolled<=0,"",ROUND({net_charges}/'
        f'ChildrenEnrolled,2))',
    ], ["money"] * 11)
    row += 2

    # --------------------------------------------- costs by category -----
    row = s.section(row, "Where the year's money went", LAST_COL)
    row = s.table_head(row, 0, [
        "Category", "Spent this year", "Budget a year",
        "Over or under", "Share of costs", "A month",
        "Category", "Spent this year", "Budget a year",
        "Over or under", "Share of costs", "A month"],
        [20, 14, 14, 13, 12, 11, 20, 14, 14, 13, 12, 11])
    cat_first = row
    for n in range(1, CAT_ROWS + 1):
        for half in (0, 1):
            index = n + half * CAT_ROWS
            base = half * 6
            cat = f'INDEX(BudgetCategory,{index})'
            spent = "+".join(
                f'SUMIFS([[Expense Tracker.pl_amount]],'
                f'[[Expense Tracker.month]],ReportYear&"-"&'
                f'TEXT({m},"00"),[[Expense Tracker.category]],{cat})'
                for m in range(1, 13))
            budget = f"INDEX(BudgetAmount,{index})*12"
            cells = [
                (base + 0, f'IFERROR({cat}&"","")', "text"),
                (base + 1, f'IFERROR(ROUND({spent},2),"")', "money"),
                (base + 2, f'IFERROR(ROUND({budget},2),"")', "money"),
                (base + 3, f'IFERROR(ROUND(({spent})-({budget}),2),"")',
                 "money"),
                (base + 4, f'IFERROR(IF({costs_year}<=0,"",({spent})/'
                           f'({costs_year})),"")', "percent"),
                (base + 5, f'IFERROR(ROUND(({spent})/12,2),"")',
                 "money"),
            ]
            for column, formula, style in cells:
                ws.write_formula(row, column,
                                 s.reg.render("=" + formula),
                                 fmts.calc(style))
        row += 1
    for column in (3, 9):
        ws.conditional_format(cat_first, column, row - 1, column,
                              {"type": "cell", "criteria": ">",
                               "value": 0.004,
                               "format": fmts.cf("bad")})
        ws.conditional_format(cat_first, column, row - 1, column,
                              {"type": "cell", "criteria": "<=",
                               "value": 0, "format": fmts.cf("good")})
    row += 1

    # ------------------------------------------------- who pays the bills -
    row = s.section(row, "Who pays the bills", LAST_COL)
    row = s.table_head(row, 0, [
        "Family", "Children", "Weekly fees", "A year", "Balance",
        "Child", "Room", "Weekly fee", "A year", "Attendance",
        "With us, months"],
        [22, 9, 13, 13, 12, 20, 14, 12, 13, 12, 14])
    for n in range(1, TOP_ROWS + 1):
        cells = [
            (0, _rank("Parent Database", "weekly_fees_rank",
                      "family_name", n), "text"),
            (1, _rank("Parent Database", "weekly_fees_rank",
                      "children", n), "int"),
            (2, _rank("Parent Database", "weekly_fees_rank",
                      "weekly_fees", n), "money"),
            (3, 'IFERROR(INDEX([[Parent Database.weekly_fees]],'
                f'MATCH({n},[[Parent Database.weekly_fees_rank]],0))'
                '*WeeksOpen,"")', "money"),
            (4, _rank("Parent Database", "weekly_fees_rank",
                      "balance", n), "money"),
            (5, _rank("Child Enrollment", "annual_net_rank", "name", n),
             "text"),
            (6, _rank("Child Enrollment", "annual_net_rank", "room", n),
             "text"),
            (7, _rank("Child Enrollment", "annual_net_rank",
                      "weekly_net", n), "money"),
            (8, _rank("Child Enrollment", "annual_net_rank",
                      "annual_net", n), "money"),
            (9, _rank("Child Enrollment", "annual_net_rank",
                      "attendance_pct", n), "percent"),
            (10, _rank("Child Enrollment", "annual_net_rank",
                       "months_with_us", n), "int"),
        ]
        for column, formula, style in cells:
            ws.write_formula(row, column, s.reg.render("=" + formula),
                             fmts.calc(style))
        row += 1
    row += 1

    # ------------------------------------------------ the year in people --
    row = s.section(row, "The year in people", LAST_COL)
    row = s.table_head(row, 0, [
        "Children started", "Children left", "Net change",
        "On the books now", "Enquiries", "Converted",
        "Conversion rate", "Staff on the books", "Staff hours",
        "Incidents", "Doses given", "Parent conversations"],
        [15, 13, 12, 15, 12, 12, 14, 15, 12, 11, 12, 17])
    starts = 'COUNTIFS([[Child Enrollment.start_date]],">="&DATE(ReportYear,1,1),[[Child Enrollment.start_date]],"<="&DATE(ReportYear+1,1,1)-1)'
    leaves = 'COUNTIFS([[Child Enrollment.end_date]],">="&DATE(ReportYear,1,1),[[Child Enrollment.end_date]],"<="&DATE(ReportYear+1,1,1)-1)'
    enquiries = 'COUNTIFS([[Waitlist.enquiry_date]],">="&DATE(ReportYear,1,1),[[Waitlist.enquiry_date]],"<="&DATE(ReportYear+1,1,1)-1)'
    converted = ('COUNTIFS([[Waitlist.enquiry_date]],">="&DATE(ReportYear,1,1),'
                 '[[Waitlist.enquiry_date]],"<="&DATE(ReportYear+1,1,1)-1,'
                 '[[Waitlist.converted]],"Yes")')
    s.calc_row(row, 0, [
        f"={starts}",
        f"={leaves}",
        f"={starts}-{leaves}",
        "=ChildrenEnrolled",
        f"={enquiries}",
        f"={converted}",
        f'=IF({enquiries}<=0,"",{converted}/{enquiries})',
        "=StaffCount",
        "=ROUND(" + year_sum("Payroll Tracker", "total_hours",
                             "month") + ",1)",
        "=" + "+".join(
            f'COUNTIFS([[Incident Reports.month]],ReportYear&"-"&'
            f'TEXT({m},"00"))' for m in range(1, 13)),
        '=COUNTIFS([[Allergy & Medical.given_date]],">="&DATE(ReportYear,1,1),[[Allergy & Medical.given_date]],"<="&DATE(ReportYear+1,1,1)-1)',
        "=" + "+".join(
            f'COUNTIFS([[Parent Communication.month]],ReportYear&"-"&'
            f'TEXT({m},"00"))' for m in range(1, 13)),
    ], ["int", "int", "int", "int", "int", "int", "percent", "int",
        "number1", "int", "int", "int"])
    row += 2

    row = s.note(
        row,
        "Everything on this page is summed from the month keys, so "
        "the year total and the twelve monthly figures can never "
        "disagree. A month with nothing logged reads zero rather "
        "than an error, which means a part-finished year still adds "
        "up.",
        LAST_COL, height=32)
    row += 1
    row = s.note(
        row,
        "Occupancy, fee income and the per-child figures use the "
        "children on the books today, not an average across the "
        "year. If your numbers moved a lot, read the month-by-month "
        "table rather than the year total.",
        LAST_COL, height=30)

    s.finish(freeze_row=6, gridlines=False, fit_width=True)
    s.protect()
    return s


def _rng(first: int, last: int, letter: str) -> str:
    return f"{quote_sheet(NAME)}!${letter}${first + 1}:${letter}${last + 1}"


def _rank(sheet: str, rank_key: str, field: str, n: int) -> str:
    return (f'IFERROR(INDEX([[{sheet}.{field}]],MATCH({n},'
            f'[[{sheet}.{rank_key}]],0)),"")')
