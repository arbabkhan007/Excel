"""`Goals & Budget` - what you said you would do, against what happened.

Three blocks: the year's targets, a month-by-month plan with the
actuals beside it, and a budget for every expense category. The
budget figures here are what the expense tab compares itself to, so
setting them once makes the whole workbook sharper.
"""

from __future__ import annotations

from ..config import EXPENSE_CATEGORIES, Config, Palette
from ..layout import Registry, col_letter, quote_sheet
from ..writer import SheetWriter, number_check

NAME = "Goals & Budget"
LAST_COL = 10
MONTHS = 12


def _key(index: int) -> str:
    return (f'TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{index},1),'
            f'"yyyy-mm")')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.accent)
    ws = s.ws
    sheet_ref = quote_sheet(NAME)

    s.banner("Goals & Budget",
             "The year you planned, the months as they land, and a "
             "budget for every line of cost.", LAST_COL)
    s.set_col(0, 0, 3)
    s.set_col(1, 1, 32)
    s.set_col(2, LAST_COL, 15)

    row = 4
    row = s.note(
        row,
        "Blue cells are yours. The targets you set here are what the "
        "dashboard, the profit and loss and the expense tab measure "
        "against - change them and every red number in the workbook "
        "changes with them.",
        LAST_COL, height=30)
    row += 1

    # ------------------------------------------------------ the year ----
    row = s.section(row, "This year's goals", LAST_COL)
    s.table_head(row, 1, ["Goal", "Target", "Actual so far",
                          "Difference", "On track?"],
                 [32, 15, 15, 15, 15])
    row += 1
    goals = [
        ("Revenue for the year", "money", "=ROUND(RevenueTargetYear,2)",
         '=ROUND(SUMIFS([[Sales.net]],[[Sales.date]],">="&'
         'DATE(YEAR(TodayDate),1,1)),2)'),
        ("Operating profit for the year", "money",
         "=ROUND(ProfitTargetMonth*12,2)",
         '=ROUND(SUMIFS([[Sales.net]],[[Sales.date]],">="&'
         'DATE(YEAR(TodayDate),1,1))-SUMIFS([[Sales.cost]],'
         '[[Sales.date]],">="&DATE(YEAR(TodayDate),1,1))-'
         'SUMIFS([[Payroll.total_cost]],[[Payroll.period_end]],">="&'
         'DATE(YEAR(TodayDate),1,1))-SUMIFS([[Expenses.amount]],'
         '[[Expenses.date]],">="&DATE(YEAR(TodayDate),1,1)),2)'),
        ("Clients served", "int", "=ROUND(ClientTargetYear,0)",
         '=COUNTIFS([[Appointments.status]],"Completed",'
         '[[Appointments.date]],">="&DATE(YEAR(TodayDate),1,1))'),
        ("New clients won", "int", "=ROUND(NewClientTargetYear,0)",
         '=COUNTIFS([[Clients.first_visit]],">="&'
         'DATE(YEAR(TodayDate),1,1))'),
        ("Average ticket", "money", "=ROUND(TargetTicket,2)",
         '=IFERROR(ROUND(SUMIFS([[Sales.net]],[[Sales.date]],">="&'
         'DATE(YEAR(TodayDate),1,1))/MAX(1,COUNTIFS('
         '[[Appointments.status]],"Completed",[[Appointments.date]],'
         '">="&DATE(YEAR(TodayDate),1,1))),2),0)'),
        ("Retail share of sales", "percent",
         "=ROUND(TargetRetailShare,4)",
         '=IFERROR(SUMIFS([[Sales.net]],[[Sales.kind]],"Retail",'
         '[[Sales.date]],">="&DATE(YEAR(TodayDate),1,1))/'
         'MAX(0.01,SUMIFS([[Sales.net]],[[Sales.date]],">="&'
         'DATE(YEAR(TodayDate),1,1))),0)'),
        ("Rebooking rate", "percent", "=ROUND(TargetRebook,4)",
         '=IFERROR(COUNTIFS([[Appointments.rebooked]],"Yes",'
         '[[Appointments.date]],">="&DATE(YEAR(TodayDate),1,1))/'
         'MAX(1,COUNTIFS([[Appointments.status]],"Completed",'
         '[[Appointments.date]],">="&DATE(YEAR(TodayDate),1,1))),0)'),
        ("Repeat client rate", "percent", "=ROUND(TargetRetention,4)",
         '=IFERROR(COUNTIFS([[Clients.retained]],"Yes")/'
         'MAX(1,COUNTIFS([[Clients.visits]],">0")),0)'),
        ("Chair utilisation", "percent", "=ROUND(TargetUtilisation,4)",
         '=IFERROR(SUMIFS([[Appointments.hours]],'
         '[[Appointments.status]],"Completed",[[Appointments.date]],'
         '">="&TodayDate-28)/MAX(1,ChairHoursWeek*4),0)'),
    ]
    first_goal = row
    for label, style, target, actual in goals:
        s.write(row, 1, label, fmts.chrome("label_bold"))
        ws.write_formula(row, 2, s.reg.render(target), fmts.calc(style))
        ws.write_formula(row, 3, s.reg.render(actual), fmts.calc(style))
        ws.write_formula(row, 4, f"=D{row + 1}-C{row + 1}",
                         fmts.calc(style))
        ws.write_formula(
            row, 5,
            f'=IF(C{row + 1}=0,"Set a target",'
            f'IF(D{row + 1}>=C{row + 1},"On track",'
            f'IF(D{row + 1}>=C{row + 1}*0.8,"Close","Behind")))',
            fmts.calc("status"))
        row += 1
    for kind, text in (("good", "On track"), ("warn", "Close"),
                       ("bad", "Behind")):
        ws.conditional_format(first_goal, 5, row - 1, 5,
                              {"type": "text", "criteria": "containing",
                               "value": text, "format": fmts.cf(kind)})
    ws.conditional_format(first_goal, 4, row - 1, 4,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("warn")})
    row += 1

    # the two yearly counts live here rather than on Settings because
    # they are goals, not settings
    row = s.subsection(row, "Two targets only this tab uses", LAST_COL)
    row = s.field(row, "Clients served, target for the year",
                  int(cfg.revenue_target_year /
                      max(1, cfg.target_average_ticket)),
                  style="int", name="ClientTargetYear",
                  validate=number_check(0, 1000000, decimal=False),
                  help_text="Revenue target divided by the average "
                            "ticket is a good first guess.")
    row = s.field(row, "New clients, target for the year",
                  int(cfg.revenue_target_year /
                      max(1, cfg.target_average_ticket) * 0.18),
                  style="int", name="NewClientTargetYear",
                  validate=number_check(0, 1000000, decimal=False))
    row += 1

    # --------------------------------------------------- month by month --
    row = s.section(row, "Month by month", LAST_COL)
    row = s.note(
        row,
        "Type a revenue target for each month - quiet January, busy "
        "December. Leave one blank and the monthly target from Settings "
        "is used instead.",
        LAST_COL, height=28)
    row += 1
    s.table_head(row, 1, [
        "Month", "Revenue target", "Actual", "Difference",
        "Clients served", "Average ticket", "Profit", "Hit it?"],
        [14, 15, 15, 15, 14, 14, 14, 12])
    row += 1
    month_first = row
    for index in range(MONTHS - 1, -1, -1):
        key = _key(index)
        s.write(row, 1, None, fmts.calc("text"))
        ws.write_formula(
            row, 1,
            f'=TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{index},1),'
            f'"mmm yy")', fmts.calc("text"))
        s.write(row, 2, None, fmts.input("money"))
        ws.data_validation(row, 2, row, 2, number_check(0, 10000000))
        ws.write_formula(
            row, 3,
            s.reg.render(f'=ROUND(SUMIFS([[Sales.net]],'
                         f'[[Sales.month]],{key}),2)'),
            fmts.calc("money"))
        ws.write_formula(
            row, 4,
            f"=ROUND(D{row + 1}-IF(C{row + 1}=\"\",RevenueTargetMonth,"
            f"C{row + 1}),2)", fmts.calc("money"))
        ws.write_formula(
            row, 5,
            s.reg.render(f'=COUNTIFS([[Appointments.month]],{key},'
                         f'[[Appointments.status]],"Completed")'),
            fmts.calc("int"))
        ws.write_formula(
            row, 6,
            f'=IFERROR(ROUND(D{row + 1}/MAX(1,F{row + 1}),2),"")',
            fmts.calc("money"))
        ws.write_formula(
            row, 7,
            s.reg.render(
                f'=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{key})-'
                f'SUMIFS([[Sales.cost]],[[Sales.month]],{key})-'
                f'SUMIFS([[Payroll.total_cost]],[[Payroll.month]],{key})-'
                f'SUMIFS([[Expenses.amount]],[[Expenses.month]],{key}),2)'),
            fmts.calc("money"))
        ws.write_formula(
            row, 8,
            f'=IF(D{row + 1}=0,"",IF(E{row + 1}>=0,"Yes","No"))',
            fmts.calc("flag"))
        row += 1
    month_last = row - 1
    ws.conditional_format(month_first, 4, month_last, 4,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("warn")})
    ws.conditional_format(month_first, 4, month_last, 4,
                          {"type": "cell", "criteria": ">=", "value": 0,
                           "format": fmts.cf("good")})
    ws.conditional_format(month_first, 8, month_last, 8,
                          {"type": "text", "criteria": "containing",
                           "value": "No", "format": fmts.cf("bad")})
    ws.conditional_format(month_first, 8, month_last, 8,
                          {"type": "text", "criteria": "containing",
                           "value": "Yes", "format": fmts.cf("good")})
    s.write(row, 1, "Year", fmts.chrome("label_bold"))
    for col in (2, 3, 4, 5, 7):
        letter = col_letter(col)
        ws.write_formula(
            row, col,
            f"=ROUND(SUM({letter}{month_first + 1}:"
            f"{letter}{month_last + 1}),2)",
            fmts.calc("int" if col == 5 else "money"))
    ws.write_formula(row, 6,
                     f'=IF(F{row + 1}=0,"",ROUND(D{row + 1}/'
                     f'F{row + 1},2))', fmts.calc("money"))
    row += 2

    # ---------------------------------------------------- the budget ----
    row = s.section(row, "Budget by category", LAST_COL)
    row = s.note(
        row,
        "Set a monthly budget for each line of cost. The expense tab "
        "reads these figures, so every bill you log knows whether it is "
        "inside the plan. Variance is actual less budget: positive "
        "means you overspent.",
        LAST_COL, height=30)
    row += 1
    s.table_head(row, 1, [
        "Category", "Monthly budget", "This month", "Variance",
        "Last month", "Year to date", "Budget to date",
        "Variance to date"],
        [26, 15, 14, 14, 14, 14, 15, 15])
    row += 1
    budget_first = row
    defaults = {
        "Rent": 2400, "Utilities": 320, "Product & stock": 0,
        "Backbar supplies": 0, "Insurance": 95,
        "Equipment": 120, "Repairs & maintenance": 90,
        "Marketing & advertising": 200,
        "Software & subscriptions": 80, "Card processing fees": 320,
        "Laundry": 165, "Cleaning": 240, "Training & education": 150,
        "Licences & memberships": 25, "Accounting & legal": 120,
        "Telephone & internet": 65, "Travel & parking": 30,
        "Bank charges": 18, "Loan repayment": 420, "Other": 120,
    }
    categories = [c for c in EXPENSE_CATEGORIES
                  if c not in ("Salaries & wages", "Commission",
                               "Payroll taxes")]
    for category in categories:
        s.write(row, 1, category, fmts.input("text"))
        s.write(row, 2, defaults.get(category, 0)
                if cfg.include_sample_data else None,
                fmts.input("money"))
        ws.data_validation(row, 2, row, 2, number_check(0, 1000000))
        ws.write_formula(
            row, 3,
            s.reg.render('=ROUND(SUMIFS([[Expenses.amount]],'
                         '[[Expenses.month]],ThisMonthKey,'
                         f'[[Expenses.category]],$B${row + 1}),2)'),
            fmts.calc("money"))
        ws.write_formula(row, 4,
                         f'=ROUND(D{row + 1}-IF(C{row + 1}="",0,'
                         f'C{row + 1}),2)', fmts.calc("money"))
        ws.write_formula(
            row, 5,
            s.reg.render('=ROUND(SUMIFS([[Expenses.amount]],'
                         '[[Expenses.month]],TEXT(DATE(YEAR(TodayDate),'
                         'MONTH(TodayDate)-1,1),"yyyy-mm"),'
                         f'[[Expenses.category]],$B${row + 1}),2)'),
            fmts.calc("money"))
        ws.write_formula(
            row, 6,
            s.reg.render('=ROUND(SUMIFS([[Expenses.amount]],'
                         '[[Expenses.date]],">="&DATE(YEAR(TodayDate),'
                         '1,1),[[Expenses.category]],'
                         f'$B${row + 1}),2)'),
            fmts.calc("money"))
        ws.write_formula(row, 7,
                         f'=ROUND(IF(C{row + 1}="",0,C{row + 1})*'
                         f'MONTH(TodayDate),2)', fmts.calc("money"))
        ws.write_formula(row, 8, f"=ROUND(G{row + 1}-H{row + 1},2)",
                         fmts.calc("money"))
        row += 1
    budget_last = row - 1
    s.write(row, 1, "Total", fmts.chrome("label_bold"))
    for col in range(2, 9):
        letter = col_letter(col)
        ws.write_formula(
            row, col,
            f"=ROUND(SUM({letter}{budget_first + 1}:"
            f"{letter}{budget_last + 1}),2)", fmts.calc("money"))
    total_row = row
    for column in (4, 8):
        ws.conditional_format(budget_first, column, total_row, column,
                              {"type": "cell", "criteria": ">", "value": 0,
                               "format": fmts.cf("bad")})
        ws.conditional_format(budget_first, column, total_row, column,
                              {"type": "cell", "criteria": "<=",
                               "value": 0, "format": fmts.cf("good")})
    book.define_name(
        "BudgetCategory",
        f"={sheet_ref}!$B${budget_first + 1}:$B${budget_last + 1}")
    book.define_name(
        "BudgetAmount",
        f"={sheet_ref}!$C${budget_first + 1}:$C${budget_last + 1}")
    book.define_name(
        "BudgetMonthlyTotal", f"={sheet_ref}!$C${total_row + 1}")
    row += 2

    # ------------------------------------------------------ break even --
    row = s.section(row, "What you have to take to break even", LAST_COL)
    row = s.table_head(row, 1, [
        "Fixed costs a month", "Staff cost a month", "Gross margin",
        "Break-even sales", "Sales this month", "Still to go",
        "Days left in the month"],
        [18, 18, 14, 16, 16, 14, 16])
    ws.write_formula(
        row, 1,
        s.reg.render('=ROUND(SUM([[Expenses.monthly]]),2)'),
        fmts.calc("money"))
    ws.write_formula(
        row, 2,
        s.reg.render('=ROUND(SUMIFS([[Payroll.total_cost]],'
                     '[[Payroll.month]],ThisMonthKey),2)'),
        fmts.calc("money"))
    ws.write_formula(
        row, 3,
        s.reg.render('=IF(SUMIFS([[Sales.net]],[[Sales.month]],'
                     'ThisMonthKey)<=0,0.85,(SUMIFS([[Sales.net]],'
                     '[[Sales.month]],ThisMonthKey)-'
                     'SUMIFS([[Sales.cost]],[[Sales.month]],'
                     'ThisMonthKey))/SUMIFS([[Sales.net]],'
                     '[[Sales.month]],ThisMonthKey))'),
        fmts.calc("percent"))
    ws.write_formula(row, 4,
                     f'=IFERROR(ROUND((B{row + 1}+C{row + 1})/'
                     f'MAX(0.05,D{row + 1}),2),"")', fmts.calc("money"))
    ws.write_formula(
        row, 5,
        s.reg.render('=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],'
                     'ThisMonthKey),2)'), fmts.calc("money"))
    ws.write_formula(row, 6, f"=ROUND(MAX(0,E{row + 1}-F{row + 1}),2)",
                     fmts.calc("money"))
    ws.write_formula(row, 7,
                     "=DAY(DATE(YEAR(TodayDate),MONTH(TodayDate)+1,0))-"
                     "DAY(TodayDate)", fmts.calc("int"))
    ws.conditional_format(row, 6, row, 6,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("warn")})
    ws.conditional_format(row, 6, row, 6,
                          {"type": "cell", "criteria": "=", "value": 0,
                           "format": fmts.cf("good")})
    row += 2

    row = s.note(
        row,
        "Break-even is the fixed overhead plus the pay run, divided by "
        "the gross margin you are actually achieving. It is the number "
        "to know before you agree to a discount, a new hire or a bigger "
        "unit.",
        LAST_COL, height=30)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    s.protect()
    return s
