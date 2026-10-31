"""`Tax Summary` - your figures, grouped the way an accountant wants them.

Pick a year. This page then adds up gross rental income, what the
platforms kept, every category of cost under its tax heading, and what
is left - with capital items kept separate, because a new boiler is
not a repair and your accountant will say so.

Read the disclaimer. This tab sorts your own numbers into sensible
piles; it does not know your country, your entity, whether you are
trading or letting, or what your local rules allow. It will save your
accountant two hours. It will not replace them.
"""

from __future__ import annotations

from ..config import (CAPITAL_CATEGORIES, Config, DISCLAIMER,
                      EXPENSE_CATEGORIES, OWNERSHIP_CATEGORIES,
                      Palette, TAX_CATEGORIES)
from ..layout import Registry, col_letter
from ..writer import SheetWriter, number_check

NAME = "Tax Summary"
LAST_COL = 8


def _year_window(column: str) -> str:
    """`range,">="&start,range,"<="&end` for the chosen tax year."""
    return (f'{column},">="&TaxYearStart,{column},"<="&TaxYearEnd')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.deep)
    ws = s.ws

    s.banner("Tax & Financial Summary",
             "A year of income and deductible costs, grouped for "
             "whoever does your return.", LAST_COL)
    s.set_col(0, 0, 4)
    s.set_col(1, 1, 34)
    s.set_col(2, 2, 15)
    s.set_col(3, 3, 15)
    s.set_col(4, 4, 12)
    s.set_col(5, LAST_COL, 16)

    # -------------------------------------------------------- the year --
    row = 4
    s.merge(row, 1, 1, "Tax year", fmts.chrome("label_bold"))
    s.write(row, 2, cfg.today.year, fmts.chrome("form_box", locked=False,
                                                num_format="0"))
    ws.data_validation(row, 2, row, 2, number_check(2000, 2100,
                                                    decimal=False))
    book.define_name("TaxYear", f"='{NAME}'!$C${row + 1}")
    s.merge(row, 3, LAST_COL,
            "The calendar year to report on. If your tax year does "
            "not start in January, change the two dates below.",
            fmts.chrome("note"))
    row += 1

    s.merge(row, 1, 1, "Year starts", fmts.chrome("label_bold"))
    s.write(row, 2, None, fmts.calc("date"))
    s.formula(row, 2, "=DATE(TaxYear,1,1)", fmts.calc("date"))
    book.define_name("TaxYearStart", f"='{NAME}'!$C${row + 1}")
    s.merge(row, 3, 3, "Year ends", fmts.chrome("label_right"))
    s.write(row, 4, None, fmts.calc("date"))
    s.formula(row, 4, "=DATE(TaxYear+1,1,1)-1", fmts.calc("date"))
    book.define_name("TaxYearEnd", f"='{NAME}'!$E${row + 1}")
    s.merge(row, 5, LAST_COL,
            "Edit these two formulas for a 6 April or 1 July year.",
            fmts.chrome("note"))
    row += 2

    bk_year = _year_window("[[Booking Tracker.check_in]]")
    earns = '[[Booking Tracker.earns]],"Yes"'
    ex_year = _year_window("[[Expense Tracker.date]]")
    cl_year = _year_window("[[Cleaning Schedule.clean_date]]")
    mt_year = _year_window("[[Maintenance Log.report_date]]")

    gross = f'SUMIFS([[Booking Tracker.gross]],{bk_year},{earns})'
    accom = f'SUMIFS([[Booking Tracker.accommodation]],{bk_year},{earns})'
    cleaning_income = (f'SUMIFS([[Booking Tracker.cleaning_fee]],'
                       f'{bk_year},{earns})')
    extra_income = (f'SUMIFS([[Booking Tracker.extra_fee]],{bk_year},'
                    f'{earns})')
    fees = (f'SUMIFS([[Booking Tracker.platform_fee_total]],{bk_year},'
            f'{earns})')
    lodging = f'SUMIFS([[Booking Tracker.tax]],{bk_year},{earns})'
    net_income = f'({gross}-{fees}-{lodging})'
    clean_cost = f'SUMIFS([[Cleaning Schedule.total_cost]],{cl_year})'
    maint_cost = f'SUMIFS([[Maintenance Log.cost_used]],{mt_year})'
    other_cost = f'SUMIFS([[Expense Tracker.pl_amount]],{ex_year})'
    capital = (f'SUMIFS([[Expense Tracker.amount]],{ex_year},'
               f'[[Expense Tracker.is_capital]],"Yes")')
    total_cost = f'({clean_cost}+{maint_cost}+{other_cost})'

    # ----------------------------------------------------- headline -----
    row = s.section(row, "The headline", LAST_COL)
    s.kpi(row, 1, "Gross rental income", f"=ROUND({gross},2)", "money0",
          note="Everything a guest paid you")
    s.kpi(row, 3, "Deductible costs", f"=ROUND({total_cost}+{fees},2)",
          "money0", note="Including what the platforms kept")
    s.kpi(row, 5, "Taxable profit",
          f"=ROUND({net_income}-{total_cost},2)", "money0")
    row = s.kpi(row, 7, "Capital items",
                f"=ROUND({capital},2)", "money0",
                note="Not deducted here - ask about allowances")
    row += 1

    # ------------------------------------------------------- income -----
    row = s.section(row, "Income", LAST_COL)
    row = s.table_head(row, 1, ["Line", "Amount", "Per property",
                                "What it is"],
                       [34, 15, 15, 12])
    income_lines = [
        ("Accommodation", f"={accom}",
         "Room rate only. This is the number most tax systems call "
         "rental income."),
        ("Cleaning fees charged", f"={cleaning_income}",
         "Income in most jurisdictions, even though it recovers a "
         "cost. The cost itself is deducted below."),
        ("Extra guest and pet fees", f"={extra_income}",
         "Charged on top of the rate."),
        ("Gross received from guests", f"={gross}",
         "The three above, less discounts given."),
        ("Less platform commission and card fees", f"=-{fees}",
         "What Airbnb, Vrbo and the rest kept. Deductible almost "
         "everywhere, and easy to forget because you never see it."),
        ("Less lodging tax collected", f"=-{lodging}",
         "Only non-zero if you collect it yourself. Tax you remit is "
         "not your income."),
        ("Net rental income", f"={net_income}",
         "What actually reached you."),
    ]
    for label, formula, what in income_lines:
        bold = label.startswith(("Gross", "Net"))
        s.write(row, 1, label,
                fmts.chrome("total_label" if bold else "label"))
        fmt = (fmts.chrome("total_value",
                           num_format=s.num_format("money"))
               if bold else fmts.calc("money"))
        s.formula(row, 2, s.reg.render(f"=ROUND({formula[1:]},2)"),
                         fmt)
        s.formula(
            row, 3,
            s.reg.render(f'=IF(PropertyCount<=0,"",ROUND(({formula[1:]})'
                         f'/PropertyCount,2))'),
            fmts.calc("money"))
        s.merge(row, 4, LAST_COL, what, fmts.chrome("note"))
        s.row_height(row, 26)
        row += 1
    row += 1

    # ----------------------------------------------------- expenses -----
    row = s.section(row, "Costs by tax heading", LAST_COL)
    row = s.note(
        row,
        "Every expense row carries a tax heading, which you set on "
        "the Expense Tracker. Cleaning and maintenance are added from "
        "their own tabs so they are counted once and only once.",
        LAST_COL, height=26)
    row = s.table_head(row, 1, ["Tax heading", "Amount", "Share",
                                "Entries"],
                       [34, 15, 15, 12])
    for heading in TAX_CATEGORIES:
        amount = (f'SUMIFS([[Expense Tracker.amount]],{ex_year},'
                  f'[[Expense Tracker.tax_category]],"{heading}")')
        s.write(row, 1, heading, fmts.chrome("label"))
        s.formula(row, 2, s.reg.render(f"=ROUND({amount},2)"),
                         fmts.calc("money"))
        s.formula(
            row, 3,
            s.reg.render(f'=IF({total_cost}<=0,"",{amount}/'
                         f'{total_cost})'),
            fmts.calc("percent"))
        s.formula(
            row, 4,
            s.reg.render(f'=COUNTIFS({ex_year},'
                         f'[[Expense Tracker.tax_category]],'
                         f'"{heading}")'), fmts.calc("int"))
        row += 1
    row += 1

    # -------------------------------------------- the full breakdown ----
    row = s.section(row, "Costs by category", LAST_COL)
    row = s.table_head(row, 1, ["Category", "Amount", "Share",
                                "Entries", "Treatment"],
                       [34, 15, 15, 12, 20])
    for category in EXPENSE_CATEGORIES:
        amount = (f'SUMIFS([[Expense Tracker.amount]],{ex_year},'
                  f'[[Expense Tracker.category]],"{category}")')
        if category in CAPITAL_CATEGORIES:
            treatment = "Capital - not deducted here"
        elif category in OWNERSHIP_CATEGORIES:
            treatment = "Ownership cost - rules vary"
        else:
            treatment = "Running cost"
        s.write(row, 1, category, fmts.chrome("label"))
        s.formula(row, 2, s.reg.render(f"=ROUND({amount},2)"),
                         fmts.calc("money"))
        s.formula(
            row, 3,
            s.reg.render(f'=IF({total_cost}<=0,"",{amount}/'
                         f'{total_cost})'), fmts.calc("percent"))
        s.formula(
            row, 4,
            s.reg.render(f'=COUNTIFS({ex_year},'
                         f'[[Expense Tracker.category]],'
                         f'"{category}")'), fmts.calc("int"))
        s.merge(row, 5, LAST_COL, treatment, fmts.chrome("note"))
        row += 1

    s.write(row, 1, "Cleaning and turnovers", fmts.chrome("label"))
    s.formula(row, 2, s.reg.render(f"=ROUND({clean_cost},2)"),
                     fmts.calc("money"))
    s.formula(
        row, 3,
        s.reg.render(f'=IF({total_cost}<=0,"",{clean_cost}/'
                     f'{total_cost})'), fmts.calc("percent"))
    s.formula(row, 4,
                     s.reg.render(f'=COUNTIFS({cl_year})'),
                     fmts.calc("int"))
    s.merge(row, 5, LAST_COL, "From the Cleaning Schedule",
            fmts.chrome("note"))
    row += 1
    s.write(row, 1, "Repairs and maintenance", fmts.chrome("label"))
    s.formula(row, 2, s.reg.render(f"=ROUND({maint_cost},2)"),
                     fmts.calc("money"))
    s.formula(
        row, 3,
        s.reg.render(f'=IF({total_cost}<=0,"",{maint_cost}/'
                     f'{total_cost})'), fmts.calc("percent"))
    s.formula(row, 4,
                     s.reg.render(f'=COUNTIFS({mt_year})'),
                     fmts.calc("int"))
    s.merge(row, 5, LAST_COL, "From the Maintenance Log",
            fmts.chrome("note"))
    row += 1
    s.write(row, 1, "Total costs", fmts.chrome("total_label"))
    s.formula(row, 2, s.reg.render(f"=ROUND({total_cost},2)"),
                     fmts.chrome("total_value",
                                 num_format=s.num_format("money")))
    row += 2

    # ------------------------------------------------- the bottom line --
    row = s.section(row, "The bottom line", LAST_COL)
    row = s.table_head(row, 1, ["Line", "Amount", "Per property",
                                "Note"], [34, 15, 15, 12])
    bottom = [
        ("Net rental income", f"={net_income}",
         "After platform fees and any lodging tax you remit."),
        ("Less running costs", f"=-{total_cost}",
         "Cleaning, repairs and everything on the expense tab, "
         "excluding capital."),
        ("Taxable profit", f"={net_income}-{total_cost}",
         "Before any allowance, relief or depreciation your "
         "jurisdiction may give you."),
        ("Capital spend in the year", f"={capital}",
         "Furniture, appliances, refurbishment. Usually relieved "
         "over several years, not in one. Not deducted above."),
        ("Mortgage interest paid", '=SUMIFS([[Expense Tracker.amount]],'
         + ex_year + ',[[Expense Tracker.category]],'
                     '"Mortgage interest")',
         "Shown separately because how much of it is deductible is "
         "one of the most country-specific rules there is."),
        ("Rent paid to a landlord (rent-to-rent)",
         '=SUMIFS([[Expense Tracker.amount]],' + ex_year +
         ',[[Expense Tracker.category]],"Rent")',
         "For arbitrage operators, usually the largest single cost."),
    ]
    for label, formula, note in bottom:
        bold = label in ("Taxable profit", "Net rental income")
        s.write(row, 1, label,
                fmts.chrome("total_label" if bold else "label"))
        fmt = (fmts.chrome("total_value",
                           num_format=s.num_format("money"))
               if bold else fmts.calc("money"))
        s.formula(row, 2,
                         s.reg.render(f"=ROUND({formula[1:]},2)"), fmt)
        s.formula(
            row, 3,
            s.reg.render(f'=IF(PropertyCount<=0,"",ROUND(({formula[1:]})'
                         f'/PropertyCount,2))'), fmts.calc("money"))
        s.merge(row, 4, LAST_COL, note, fmts.chrome("note"))
        s.row_height(row, 26)
        row += 1
    row += 1

    # ------------------------------------------------- per property -----
    row = s.section(row, "By property, for the year", LAST_COL)
    row = s.table_head(row, 1, [
        "Property", "Gross", "Platform fees", "Costs", "Profit",
        "Nights sold"], [26, 15, 15, 14, 14, 12])
    for n in range(1, 9):
        pid = f'INDEX([[Property Database.id]],{n})'
        guard = f'INDEX([[Property Database.name]],{n})=""'
        p_gross = (f'SUMIFS([[Booking Tracker.gross]],{bk_year},{earns},'
                   f'[[Booking Tracker.property]],{pid})')
        p_fees = (f'SUMIFS([[Booking Tracker.platform_fee_total]],'
                  f'{bk_year},{earns},[[Booking Tracker.property]],'
                  f'{pid})')
        p_nights = (f'SUMIFS([[Booking Tracker.nights]],{bk_year},'
                    f'{earns},[[Booking Tracker.property]],{pid})')
        p_costs = (f'(SUMIFS([[Cleaning Schedule.total_cost]],{cl_year},'
                   f'[[Cleaning Schedule.property]],{pid})+'
                   f'SUMIFS([[Maintenance Log.cost_used]],{mt_year},'
                   f'[[Maintenance Log.property]],{pid})+'
                   f'SUMIFS([[Expense Tracker.pl_amount]],{ex_year},'
                   f'[[Expense Tracker.property]],{pid}))')
        s.formula(
            row, 1,
            s.reg.render(f'=IFERROR(IF({guard},"",INDEX('
                         f'[[Property Database.name]],{n})),"")'),
            fmts.calc("text"))
        for offset, (formula, style) in enumerate((
                (f"ROUND({p_gross},2)", "money"),
                (f"ROUND({p_fees},2)", "money"),
                (f"ROUND({p_costs},2)", "money"),
                (f"ROUND({p_gross}-{p_fees}-{p_costs},2)", "money"),
                (p_nights, "int"))):
            s.formula(
                row, 2 + offset,
                s.reg.render(f'=IFERROR(IF({guard},"",{formula}),"")'),
                fmts.calc(style))
        row += 1
    row += 1

    row = s.note(
        row,
        "Everything on this page is attributed to the year a stay "
        "checked in and the date a cost was dated. If your accountant "
        "wants revenue spread across the nights of a stay that "
        "straddles new year, tell them which convention this file "
        "uses - it is consistent, which is what matters.",
        LAST_COL, height=32)
    row += 1
    row = s.note(row, DISCLAIMER, LAST_COL, height=112)

    s.finish(freeze_row=4, gridlines=False, fit_width=True,
             print_area=f"A1:{col_letter(LAST_COL)}{row + 2}")
    s.protect()
    return s
