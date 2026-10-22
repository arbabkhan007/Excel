"""`Profit & Loss` - twelve months, in the order an accountant reads it.

Income, cost of sales, gross profit, staff costs, every expense
category, operating profit. Each column is a month, worked out from
the till, the payroll and the expense log - so it agrees with the
dashboard by construction rather than by luck.
"""

from __future__ import annotations

from ..config import EXPENSE_CATEGORIES, Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter

NAME = "Profit & Loss"
MONTHS = 12
LAST_COL = MONTHS + 2          # label + 12 months + year to date


def _key(index: int) -> str:
    """The yyyy-mm key for `index` months before this one."""
    return (f'TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{index},1),'
            f'"yyyy-mm")')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws

    s.banner("Profit & Loss",
             "Twelve months of income, costs and what was left - built "
             "from the sales, payroll and expenses you have already "
             "logged.", LAST_COL)
    s.set_col(0, 0, 30)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 12)

    row = 4
    row = s.note(
        row,
        "Every column is a calendar month. The year-to-date column adds "
        "the twelve together. Nothing here is a tax return: it is a "
        "management view, so drawings, depreciation, loan capital and "
        "stock still on the shelf are deliberately left out.",
        LAST_COL, height=32)
    row += 1

    # ------------------------------------------------------- month header
    s.write(row, 0, "Line", fmts.chrome("header_left"))
    for index in range(MONTHS - 1, -1, -1):
        col = MONTHS - index
        ws.write_formula(
            row, col,
            f'=TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{index},1),'
            f'"mmm yy")', fmts.chrome("header"))
    s.write(row, MONTHS + 1, "12 months", fmts.chrome("header"))
    s.write(row, MONTHS + 2, "Share", fmts.chrome("header"))
    s.row_height(row, 24)
    header_row = row
    row += 1

    def line(label: str, body: str, style: str = "money",
             mode: str = "calc", bold: bool = False,
             share_of: str = "") -> int:
        """One P&L line: the label, twelve months, the year, the share."""
        nonlocal row
        fmt_label = fmts.chrome("label_bold" if bold else "label")
        s.write(row, 0, label, fmt_label)
        for index in range(MONTHS - 1, -1, -1):
            col = MONTHS - index
            formula = body.format(key=_key(index))
            ws.write_formula(row, col, s.reg.render(formula),
                             fmts.calc(style))
        letters = [f"{chr(66 + i)}" for i in range(MONTHS)]
        total = "+".join(f"{letter}{row + 1}" for letter in letters)
        ws.write_formula(row, MONTHS + 1, f"=ROUND({total},2)",
                         fmts.calc(style))
        if share_of:
            ws.write_formula(
                row, MONTHS + 2,
                f'=IF({share_of}=0,"",{chr(66 + MONTHS)}{row + 1}/'
                f'{share_of})', fmts.calc("percent"))
        else:
            s.write(row, MONTHS + 2, None, fmts.calc("percent"))
        row += 1
        return row - 1

    # ------------------------------------------------------------- income
    row = s.section(row, "Income", LAST_COL)
    service_row = line(
        "Service sales",
        '=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{key},'
        '[[Sales.kind]],"Service"),2)')
    retail_row = line(
        "Retail sales",
        '=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{key},'
        '[[Sales.kind]],"Retail"),2)')
    package_row = line(
        "Packages, courses & gift certificates",
        '=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{key},'
        '[[Sales.kind]],"Package")+SUMIFS([[Sales.net]],'
        '[[Sales.month]],{key},[[Sales.kind]],"Gift certificate")+'
        'SUMIFS([[Sales.net]],[[Sales.month]],{key},[[Sales.kind]],'
        '"Course"),2)')
    other_row = line(
        "Other income (chair rent, fees)",
        '=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{key},'
        '[[Sales.kind]],"Other")+SUMIFS([[Appointments.fee_charged]],'
        '[[Appointments.month]],{key}),2)')
    net_row = line("Net sales", "=0", bold=True)
    # the generic {key} substitution cannot build column references, so
    # the total line is rewritten cell by cell
    for index in range(MONTHS):
        col = index + 1
        letter = chr(66 + index)
        ws.write_formula(
            net_row, col,
            f"=ROUND({letter}{service_row + 1}+{letter}{retail_row + 1}+"
            f"{letter}{package_row + 1}+{letter}{other_row + 1},2)",
            fmts.calc("money"))
    book.define_name("PLNetSalesRow",
                     f"={quote_sheet(NAME)}!$N${net_row + 1}")
    net_total = f"$N${net_row + 1}"
    ws.write_formula(net_row, MONTHS + 2, "=1", fmts.calc("percent"))

    row = s.subsection(row, "What came off the top", LAST_COL)
    line("Discounts given (memo)",
         '=ROUND(SUMIFS([[Sales.discount]],[[Sales.month]],{key}),2)',
         share_of=net_total)
    line("Refunds (memo)",
         '=ROUND(SUMIFS([[Sales.refund]],[[Sales.month]],{key}),2)',
         share_of=net_total)
    line("Sales tax collected (not yours)",
         '=ROUND(SUMIFS([[Sales.tax]],[[Sales.month]],{key}),2)',
         share_of=net_total)
    line("Tips passed to staff (not income)",
         '=ROUND(SUMIFS([[Payments.amount]],[[Payments.month]],{key},'
         '[[Payments.kind]],"Tip"),2)', share_of=net_total)
    row += 1

    # ------------------------------------------------------ cost of sales
    row = s.section(row, "Cost of sales", LAST_COL)
    backbar_row = line(
        "Product used in services",
        '=ROUND(SUMIFS([[Sales.cost]],[[Sales.month]],{key},'
        '[[Sales.kind]],"Service"),2)', share_of=net_total)
    stock_row = line(
        "Retail stock sold",
        '=ROUND(SUMIFS([[Sales.cost]],[[Sales.month]],{key},'
        '[[Sales.kind]],"Retail"),2)', share_of=net_total)
    cos_row = line("Total cost of sales", "=0", bold=True,
                   share_of=net_total)
    for index in range(MONTHS):
        col = index + 1
        letter = chr(66 + index)
        ws.write_formula(
            cos_row, col,
            f"=ROUND({letter}{backbar_row + 1}+{letter}{stock_row + 1},2)",
            fmts.calc("money"))
    gross_row = line("Gross profit", "=0", bold=True, share_of=net_total)
    for index in range(MONTHS):
        col = index + 1
        letter = chr(66 + index)
        ws.write_formula(
            gross_row, col,
            f"=ROUND({letter}{net_row + 1}-{letter}{cos_row + 1},2)",
            fmts.calc("money"))
    margin_row = line("Gross margin", "=0", style="percent", bold=True)
    for index in range(MONTHS):
        col = index + 1
        letter = chr(66 + index)
        ws.write_formula(
            margin_row, col,
            f'=IF({letter}{net_row + 1}=0,"",{letter}{gross_row + 1}/'
            f'{letter}{net_row + 1})', fmts.calc("percent"))
    ws.write_formula(margin_row, MONTHS + 1,
                     f'=IF($N${net_row + 1}=0,"",$N${gross_row + 1}/'
                     f'$N${net_row + 1})', fmts.calc("percent"))
    row += 1

    # --------------------------------------------------------- staff cost
    row = s.section(row, "Staff costs", LAST_COL)
    wages_row = line(
        "Hourly pay",
        '=ROUND(SUMIFS([[Payroll.base_pay]],[[Payroll.month]],{key}),2)',
        share_of=net_total)
    commission_row = line(
        "Commission released",
        '=ROUND(SUMIFS([[Payroll.commission]],[[Payroll.month]],'
        '{key}),2)', share_of=net_total)
    extra_row = line(
        "Bonuses and holiday pay",
        '=ROUND(SUMIFS([[Payroll.bonus]],[[Payroll.month]],{key})+'
        'SUMIFS([[Payroll.holiday]],[[Payroll.month]],{key}),2)',
        share_of=net_total)
    employer_row = line(
        "Employer taxes and cover",
        '=ROUND(SUMIFS([[Payroll.employer_cost]],[[Payroll.month]],'
        '{key}),2)', share_of=net_total)
    staff_row = line("Total staff cost", "=0", bold=True,
                     share_of=net_total)
    for index in range(MONTHS):
        col = index + 1
        letter = chr(66 + index)
        ws.write_formula(
            staff_row, col,
            f"=ROUND({letter}{wages_row + 1}+{letter}{commission_row + 1}+"
            f"{letter}{extra_row + 1}+{letter}{employer_row + 1},2)",
            fmts.calc("money"))
    line("Staff cost as a share of service sales", "=0",
         style="percent")
    share_row = row - 1
    for index in range(MONTHS):
        col = index + 1
        letter = chr(66 + index)
        ws.write_formula(
            share_row, col,
            f'=IF({letter}{service_row + 1}=0,"",'
            f'{letter}{staff_row + 1}/{letter}{service_row + 1})',
            fmts.calc("percent"))
    ws.write_formula(share_row, MONTHS + 1,
                     f'=IF($N${service_row + 1}=0,"",$N${staff_row + 1}/'
                     f'$N${service_row + 1})', fmts.calc("percent"))
    row += 1

    # ----------------------------------------------------- the overheads
    row = s.section(row, "Operating expenses", LAST_COL)
    first_expense = row
    for category in EXPENSE_CATEGORIES:
        if category in ("Salaries & wages", "Commission", "Payroll taxes"):
            continue                    # they live on the payroll tab
        safe = category.replace('"', "")
        line(category,
             '=ROUND(SUMIFS([[Expenses.amount]],[[Expenses.month]],'
             '{key},[[Expenses.category]],"' + safe + '"),2)',
             share_of=net_total)
    last_expense = row - 1
    total_exp_row = line("Total operating expenses", "=0", bold=True,
                         share_of=net_total)
    for index in range(MONTHS):
        col = index + 1
        letter = chr(66 + index)
        ws.write_formula(
            total_exp_row, col,
            f"=ROUND(SUM({letter}{first_expense + 1}:"
            f"{letter}{last_expense + 1}),2)", fmts.calc("money"))
    row += 1

    # ------------------------------------------------------------- profit
    row = s.section(row, "The bottom line", LAST_COL)
    profit_row = line("Operating profit", "=0", bold=True,
                      share_of=net_total)
    for index in range(MONTHS):
        col = index + 1
        letter = chr(66 + index)
        ws.write_formula(
            profit_row, col,
            f"=ROUND({letter}{gross_row + 1}-{letter}{staff_row + 1}-"
            f"{letter}{total_exp_row + 1},2)", fmts.calc("money"))
    profit_margin_row = line("Operating margin", "=0", style="percent",
                             bold=True)
    for index in range(MONTHS):
        col = index + 1
        letter = chr(66 + index)
        ws.write_formula(
            profit_margin_row, col,
            f'=IF({letter}{net_row + 1}=0,"",{letter}{profit_row + 1}/'
            f'{letter}{net_row + 1})', fmts.calc("percent"))
    ws.write_formula(profit_margin_row, MONTHS + 1,
                     f'=IF($N${net_row + 1}=0,"",$N${profit_row + 1}/'
                     f'$N${net_row + 1})', fmts.calc("percent"))
    target_row = line("Profit target", "=ROUND(ProfitTargetMonth,2)")
    variance_row = line("Against target", "=0", bold=True)
    for index in range(MONTHS):
        col = index + 1
        letter = chr(66 + index)
        ws.write_formula(
            variance_row, col,
            f"=ROUND({letter}{profit_row + 1}-{letter}{target_row + 1},2)",
            fmts.calc("money"))
    ws.conditional_format(variance_row, 1, variance_row, MONTHS + 1,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(variance_row, 1, variance_row, MONTHS + 1,
                          {"type": "cell", "criteria": ">=", "value": 0,
                           "format": fmts.cf("good")})
    ws.conditional_format(profit_row, 1, profit_row, MONTHS + 1,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(net_row, 1, net_row, MONTHS,
                          {"type": "data_bar",
                           "bar_color": Palette.primary_mid})
    row += 1

    ws.add_sparkline(row, 1, {
        "range": (f"{quote_sheet(NAME)}!$B${profit_row + 1}:"
                  f"$M${profit_row + 1}"),
        "type": "column", "series_color": Palette.accent,
        "negative_points": True})
    s.merge(row, 2, 5, "Operating profit, month by month",
            fmts.chrome("note"))
    row += 2

    row = s.note(
        row,
        "Wages come from the payroll tab, never from the expense log, so "
        "they are counted once. Stock is counted when it is sold, not "
        "when it is bought, which is why a big delivery does not make a "
        "month look like a disaster. Deposits are money received and "
        "appear as revenue only when the sale is rung up.",
        LAST_COL, height=36)

    s.finish(freeze_row=header_row + 1, freeze_col=1, gridlines=False,
             fit_width=True)
    s.protect()
    return s
