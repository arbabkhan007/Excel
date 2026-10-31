"""`Annual Reports` - twelve months across, and the year that was.

Pick a year. Every month becomes a column: nights, occupancy, ADR,
RevPAR, revenue, costs and profit, with the season visible at a
glance because it is laid out the way a year actually behaves.

Below that, the year in totals, the year against the year before,
every property ranked, and the twelve questions worth asking in
January - which month carried the year, what the quiet months really
cost, and whether your rates moved at all.
"""

from __future__ import annotations

from ..config import Config, Palette, SEASONS
from ..layout import Registry, col_letter
from ..writer import SheetWriter, number_check

NAME = "Annual Reports"
LAST_COL = 14
PROPERTY_ROWS = 8


def _key(month: int, year: str = "ReportYear") -> str:
    """The `yyyy-mm` key for one month of the chosen year."""
    return f'TEXT(DATE({year},{month},1),"yyyy-mm")'


def _year_sum(value: str, months: str, year: str = "ReportYear") -> str:
    """A value summed over all twelve months of a year. Bracketed."""
    parts = "+".join(
        f'SUMIFS({value},{months},{_key(m, year)})' for m in range(1, 13))
    return f"({parts})"


def _year_sum_c(value: str, months: str, extra: str,
                year: str = "ReportYear") -> str:
    parts = "+".join(
        f'SUMIFS({value},{months},{_key(m, year)},{extra})'
        for m in range(1, 13))
    return f"({parts})"


def _year_count(months: str, extra: str,
                year: str = "ReportYear") -> str:
    parts = "+".join(
        f'COUNTIFS({months},{_key(m, year)},{extra})'
        for m in range(1, 13))
    return f"({parts})"


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.deep)
    ws = s.ws

    s.banner("Annual Report",
             "Twelve months across. The year in full, against last "
             "year, property by property.", LAST_COL)
    s.set_col(0, 0, 26)
    for column in range(1, 13):
        s.set_col(column, column, 11)
    s.set_col(13, LAST_COL, 13)

    row = 4
    s.write(row, 0, "Year", fmts.chrome("label_bold"))
    s.write(row, 1, cfg.today.year,
            fmts.chrome("form_box", locked=False, num_format="0"))
    ws.data_validation(row, 1, row, 1,
                       number_check(2000, 2100, decimal=False))
    book.define_name("ReportYear", f"='{NAME}'!$B${row + 1}")
    s.merge(row, 2, LAST_COL,
            "Every figure on this page follows this one cell. A "
            "booking belongs to the year it checks in.",
            fmts.chrome("note"))
    row += 2

    bk_month = "[[Booking Tracker.month]]"
    earns = '[[Booking Tracker.earns]],"Yes"'
    cl_month = "[[Cleaning Schedule.month]]"
    mt_month = "[[Maintenance Log.month]]"
    ex_month = "[[Expense Tracker.month]]"

    y_nights = _year_sum_c("[[Booking Tracker.nights]]", bk_month, earns)
    y_accom = _year_sum_c("[[Booking Tracker.accommodation]]", bk_month,
                          earns)
    y_gross = _year_sum_c("[[Booking Tracker.gross]]", bk_month, earns)
    y_fees = _year_sum_c("[[Booking Tracker.platform_fee_total]]",
                         bk_month, earns)
    y_net = _year_sum_c("[[Booking Tracker.net_revenue]]", bk_month,
                        earns)
    y_guests = _year_sum_c("[[Booking Tracker.guests_total]]", bk_month,
                           earns)
    y_blocked = _year_sum_c("[[Booking Tracker.nights]]", bk_month,
                            '[[Booking Tracker.is_block]],"Yes"')
    y_bookings = _year_count(bk_month, earns)
    y_cancel = _year_count(bk_month,
                           '[[Booking Tracker.is_cancelled]],"Yes"')
    y_clean = _year_sum("[[Cleaning Schedule.total_cost]]", cl_month)
    y_maint = _year_sum("[[Maintenance Log.cost_used]]", mt_month)
    y_other = _year_sum("[[Expense Tracker.pl_amount]]", ex_month)
    y_costs = f"({y_clean}+{y_maint}+{y_other})"
    y_profit = f"({y_net}-{y_costs})"
    y_days = "(DATE(ReportYear+1,1,1)-DATE(ReportYear,1,1))"
    y_avail = f"MAX(1,PropertyCount*{y_days}-{y_blocked})"

    # --------------------------------------------------- the headline --
    row = s.section(row, "The year", LAST_COL)
    s.kpi(row, 0, "Bookings", f"={y_bookings}", "int")
    s.kpi(row, 2, "Nights sold", f"={y_nights}", "int")
    s.kpi(row, 4, "Occupancy",
          f'=IF({y_avail}<=0,"",{y_nights}/{y_avail})', "percent")
    s.kpi(row, 6, "ADR",
          f'=IF({y_nights}<=0,"",ROUND({y_accom}/{y_nights},2))',
          "money")
    s.kpi(row, 8, "RevPAR",
          f'=IF({y_avail}<=0,"",ROUND({y_accom}/{y_avail},2))', "money")
    s.kpi(row, 10, "Guests hosted", f"={y_guests}", "int")
    row = s.kpi(row, 12, "Average stay",
                f'=IF({y_bookings}<=0,"",ROUND({y_nights}/'
                f'{y_bookings},1))', "number1")

    s.kpi(row, 0, "Gross", f"=ROUND({y_gross},2)", "money0")
    s.kpi(row, 2, "Platform fees", f"=ROUND({y_fees},2)", "money0")
    s.kpi(row, 4, "Net revenue", f"=ROUND({y_net},2)", "money0")
    s.kpi(row, 6, "Costs", f"=ROUND({y_costs},2)", "money0")
    s.kpi(row, 8, "Profit", f"=ROUND({y_profit},2)", "money0")
    s.kpi(row, 10, "Margin",
          f'=IF({y_net}<=0,"",{y_profit}/{y_net})', "percent")
    row = s.kpi(row, 12, "Against target",
                f'=IF(TargetRevenueYear<=0,"",{y_net}/'
                f'TargetRevenueYear)', "percent")
    row += 1

    # ---------------------------------------------- the twelve months --
    row = s.section(row, "Month by month", LAST_COL)
    header = ["Measure"] + [
        f"{m}" for m in ("Jan", "Feb", "Mar", "Apr", "May", "Jun",
                         "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    ] + ["Year"]
    row = s.table_head(row, 0, header,
                       [26] + [11] * 12 + [13])

    def month_set(builder):
        """A row of twelve month formulas plus the year total."""
        return [builder(m) for m in range(1, 13)]

    def m_nights(m):
        return (f'SUMIFS([[Booking Tracker.nights]],{bk_month},'
                f'{_key(m)},{earns})')

    def m_accom(m):
        return (f'SUMIFS([[Booking Tracker.accommodation]],{bk_month},'
                f'{_key(m)},{earns})')

    def m_net(m):
        return (f'SUMIFS([[Booking Tracker.net_revenue]],{bk_month},'
                f'{_key(m)},{earns})')

    def m_gross(m):
        return (f'SUMIFS([[Booking Tracker.gross]],{bk_month},'
                f'{_key(m)},{earns})')

    def m_fees(m):
        return (f'SUMIFS([[Booking Tracker.platform_fee_total]],'
                f'{bk_month},{_key(m)},{earns})')

    def m_bookings(m):
        return f'COUNTIFS({bk_month},{_key(m)},{earns})'

    def m_costs(m):
        return (f'(SUMIFS([[Cleaning Schedule.total_cost]],{cl_month},'
                f'{_key(m)})+SUMIFS([[Maintenance Log.cost_used]],'
                f'{mt_month},{_key(m)})+'
                f'SUMIFS([[Expense Tracker.pl_amount]],{ex_month},'
                f'{_key(m)}))')

    def m_days(m):
        return f'DAY(DATE(ReportYear,{m + 1},1)-1)'

    def m_avail(m):
        blocked = (f'SUMIFS([[Booking Tracker.nights]],{bk_month},'
                   f'{_key(m)},[[Booking Tracker.is_block]],"Yes")')
        return f'MAX(1,PropertyCount*{m_days(m)}-{blocked})'

    lines = [
        ("Bookings", month_set(m_bookings), f"={y_bookings}", "int"),
        ("Nights sold", month_set(m_nights), f"={y_nights}", "int"),
        ("Occupancy",
         [f'IF({m_avail(m)}<=0,"",{m_nights(m)}/{m_avail(m)})'
          for m in range(1, 13)],
         f'=IF({y_avail}<=0,"",{y_nights}/{y_avail})', "percent"),
        ("ADR",
         [f'IF({m_nights(m)}<=0,"",ROUND({m_accom(m)}/{m_nights(m)},2))'
          for m in range(1, 13)],
         f'=IF({y_nights}<=0,"",ROUND({y_accom}/{y_nights},2))',
         "money"),
        ("RevPAR",
         [f'IF({m_avail(m)}<=0,"",ROUND({m_accom(m)}/{m_avail(m)},2))'
          for m in range(1, 13)],
         f'=IF({y_avail}<=0,"",ROUND({y_accom}/{y_avail},2))', "money"),
        ("Gross", [f"ROUND({m_gross(m)},2)" for m in range(1, 13)],
         f"=ROUND({y_gross},2)", "money0"),
        ("Platform fees",
         [f"ROUND({m_fees(m)},2)" for m in range(1, 13)],
         f"=ROUND({y_fees},2)", "money0"),
        ("Net revenue", [f"ROUND({m_net(m)},2)" for m in range(1, 13)],
         f"=ROUND({y_net},2)", "money0"),
        ("Costs", [f"ROUND({m_costs(m)},2)" for m in range(1, 13)],
         f"=ROUND({y_costs},2)", "money0"),
        ("Profit",
         [f"ROUND({m_net(m)}-{m_costs(m)},2)" for m in range(1, 13)],
         f"=ROUND({y_profit},2)", "money0"),
        ("Margin",
         [f'IF({m_net(m)}<=0,"",({m_net(m)}-{m_costs(m)})/{m_net(m)})'
          for m in range(1, 13)],
         f'=IF({y_net}<=0,"",{y_profit}/{y_net})', "percent"),
        ("Share of the year",
         [f'IF({y_net}<=0,"",{m_net(m)}/{y_net})'
          for m in range(1, 13)],
         "=1", "percent"),
    ]
    grid_top = row
    sparkline_rows = {}
    for label, cells, total, style in lines:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        for index, formula in enumerate(cells):
            s.formula(row, index + 1,
                             s.reg.render("=" + formula.lstrip("=")),
                             fmts.calc(style))
        s.formula(row, 13, s.reg.render(total),
                         fmts.chrome("total_value",
                                     num_format=s.num_format(style)))
        sparkline_rows[label] = row
        row += 1
    grid_bottom = row - 1

    for label, kind in (("Occupancy", "line"), ("ADR", "line"),
                        ("Net revenue", "column"),
                        ("Profit", "column")):
        line = sparkline_rows[label]
        ws.add_sparkline(line, 14, {
            "range": f"'{NAME}'!B{line + 1}:M{line + 1}",
            "type": kind, "style": 12,
            "negative_points": kind == "column"})
    s.write(grid_top - 1, 14, "Shape", fmts.chrome("header"))

    ws.conditional_format(sparkline_rows["Occupancy"], 1,
                          sparkline_rows["Occupancy"], 12, {
                              "type": "cell", "criteria": ">=",
                              "value": "TargetOccupancy",
                              "format": fmts.cf("good")})
    ws.conditional_format(sparkline_rows["Profit"], 1,
                          sparkline_rows["Profit"], 12, {
                              "type": "cell", "criteria": "<",
                              "value": 0, "format": fmts.cf("bad")})
    row = grid_bottom + 2

    # ------------------------------------------------------ seasons ----
    row = s.section(row, "By season", LAST_COL)
    row = s.note(
        row,
        "Short-term lets are a seasonal business and the averages "
        "hide it. Three strong months paying for nine quiet ones is a "
        "completely different business from twelve steady ones, and "
        "they need different pricing.",
        LAST_COL, height=28)
    row = s.table_head(row, 0, [
        "Season", "Months", "Nights", "Occupancy", "ADR", "RevPAR",
        "Net revenue", "Costs", "Profit", "Share of profit"],
        [20, 18, 10, 11, 11, 11, 13, 12, 12, 14])
    for season, first_month, last_month in SEASONS:
        if first_month <= last_month:
            months = list(range(first_month, last_month + 1))
        else:                       # winter wraps the year end
            months = (list(range(first_month, 13))
                      + list(range(1, last_month + 1)))
        nights = "+".join(m_nights(m) for m in months)
        accom = "+".join(m_accom(m) for m in months)
        net = "+".join(m_net(m) for m in months)
        costs = "+".join(m_costs(m) for m in months)
        avail = "+".join(m_avail(m) for m in months)
        label = ", ".join(
            ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug",
             "Sep", "Oct", "Nov", "Dec")[m - 1] for m in months)
        s.calc_row(row, 0, [
            season, label,
            f"=({nights})",
            f'=IF(({avail})<=0,"",({nights})/({avail}))',
            f'=IF(({nights})<=0,"",ROUND(({accom})/({nights}),2))',
            f'=IF(({avail})<=0,"",ROUND(({accom})/({avail}),2))',
            f"=ROUND({net},2)",
            f"=ROUND({costs},2)",
            f"=ROUND(({net})-({costs}),2)",
            f'=IF({y_profit}<=0,"",(({net})-({costs}))/{y_profit})',
        ], ["text", "text", "int", "percent", "money", "money",
            "money0", "money0", "money0", "percent"])
        row += 1
    row += 1

    # --------------------------------------------- against last year ---
    row = s.section(row, "Against last year", LAST_COL)
    row = s.table_head(row, 0, [
        "Measure", "This year", "Last year", "Change", "Change %"],
        [26, 15, 15, 14, 12])
    ly = "ReportYear-1"
    ly_nights = _year_sum_c("[[Booking Tracker.nights]]", bk_month,
                            earns, ly)
    ly_accom = _year_sum_c("[[Booking Tracker.accommodation]]",
                           bk_month, earns, ly)
    ly_gross = _year_sum_c("[[Booking Tracker.gross]]", bk_month,
                           earns, ly)
    ly_fees = _year_sum_c("[[Booking Tracker.platform_fee_total]]",
                          bk_month, earns, ly)
    ly_net = _year_sum_c("[[Booking Tracker.net_revenue]]", bk_month,
                         earns, ly)
    ly_bookings = _year_count(bk_month, earns, ly)
    ly_costs = ("(" + _year_sum("[[Cleaning Schedule.total_cost]]",
                                cl_month, ly)
                + "+" + _year_sum("[[Maintenance Log.cost_used]]",
                                  mt_month, ly)
                + "+" + _year_sum("[[Expense Tracker.pl_amount]]",
                                  ex_month, ly) + ")")
    ly_profit = f"({ly_net}-{ly_costs})"
    ly_days = "(DATE(ReportYear,1,1)-DATE(ReportYear-1,1,1))"
    ly_avail = f"MAX(1,PropertyCount*{ly_days})"

    pairs = [
        ("Bookings", y_bookings, ly_bookings, "int", False),
        ("Nights sold", y_nights, ly_nights, "int", False),
        ("Occupancy", f'IF({y_avail}<=0,0,{y_nights}/{y_avail})',
         f'IF({ly_avail}<=0,0,{ly_nights}/{ly_avail})', "percent", True),
        ("ADR", f'IF({y_nights}<=0,0,{y_accom}/{y_nights})',
         f'IF({ly_nights}<=0,0,{ly_accom}/{ly_nights})', "money", False),
        ("RevPAR", f'IF({y_avail}<=0,0,{y_accom}/{y_avail})',
         f'IF({ly_avail}<=0,0,{ly_accom}/{ly_avail})', "money", False),
        ("Gross", y_gross, ly_gross, "money0", False),
        ("Platform fees", y_fees, ly_fees, "money0", False),
        ("Net revenue", y_net, ly_net, "money0", False),
        ("Costs", y_costs, ly_costs, "money0", False),
        ("Profit", y_profit, ly_profit, "money0", False),
        ("Margin", f'IF({y_net}<=0,0,{y_profit}/{y_net})',
         f'IF({ly_net}<=0,0,{ly_profit}/{ly_net})', "percent", True),
    ]
    compare_top = row
    for label, this, last, style, absolute in pairs:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.formula(row, 1, (f"=ROUND({this},4)"),
                         fmts.calc(style))
        s.formula(row, 2, s.reg.render(f"=ROUND({last},4)"),
                         fmts.calc(style))
        s.formula(row, 3,
                         s.reg.render(f"=ROUND({this}-{last},4)"),
                         fmts.calc("percent" if absolute else style))
        s.formula(
            row, 4,
            s.reg.render(f'=IF({last}<=0,"",{this}/{last}-1)'),
            fmts.calc("percent"))
        row += 1
    for column in (3, 4):
        ws.conditional_format(compare_top, column, row - 1, column, {
            "type": "cell", "criteria": "<", "value": 0,
            "format": fmts.cf("bad")})
        ws.conditional_format(compare_top, column, row - 1, column, {
            "type": "cell", "criteria": ">", "value": 0,
            "format": fmts.cf("good")})
    row += 1

    # ------------------------------------------------- by property -----
    row = s.section(row, "Every property, for the year", LAST_COL)
    row = s.table_head(row, 0, [
        "Property", "Bookings", "Nights", "Occupancy", "ADR", "RevPAR",
        "Net revenue", "Costs", "Profit", "Margin", "Profit a night",
        "Share of profit"],
        [24, 11, 10, 11, 11, 11, 13, 12, 12, 10, 13, 14])
    for n in range(1, PROPERTY_ROWS + 1):
        pid = f'INDEX([[Property Database.id]],{n})'
        guard = f'INDEX([[Property Database.name]],{n})=""'
        prop = f'[[Booking Tracker.property]],{pid}'
        p_nights = "+".join(
            f'SUMIFS([[Booking Tracker.nights]],{bk_month},{_key(m)},'
            f'{earns},{prop})' for m in range(1, 13))
        p_accom = "+".join(
            f'SUMIFS([[Booking Tracker.accommodation]],{bk_month},'
            f'{_key(m)},{earns},{prop})' for m in range(1, 13))
        p_net = "+".join(
            f'SUMIFS([[Booking Tracker.net_revenue]],{bk_month},'
            f'{_key(m)},{earns},{prop})' for m in range(1, 13))
        p_bookings = "+".join(
            f'COUNTIFS({bk_month},{_key(m)},{earns},{prop})'
            for m in range(1, 13))
        p_costs = "+".join(
            f'SUMIFS([[Cleaning Schedule.total_cost]],{cl_month},'
            f'{_key(m)},[[Cleaning Schedule.property]],{pid})+'
            f'SUMIFS([[Maintenance Log.cost_used]],{mt_month},'
            f'{_key(m)},[[Maintenance Log.property]],{pid})+'
            f'SUMIFS([[Expense Tracker.pl_amount]],{ex_month},'
            f'{_key(m)},[[Expense Tracker.property]],{pid})'
            for m in range(1, 13))
        p_avail = f"MAX(1,{y_days})"
        p_profit = f"(({p_net})-({p_costs}))"
        s.formula(
            row, 0,
            s.reg.render(f'=IFERROR(IF({guard},"",INDEX('
                         f'[[Property Database.name]],{n})),"")'),
            fmts.calc("text"))
        cells = [
            (f"({p_bookings})", "int"),
            (f"({p_nights})", "int"),
            (f'IF({p_avail}<=0,"",({p_nights})/{p_avail})', "percent"),
            (f'IF(({p_nights})<=0,"",ROUND(({p_accom})/({p_nights}),2))',
             "money"),
            (f'IF({p_avail}<=0,"",ROUND(({p_accom})/{p_avail},2))',
             "money"),
            (f"ROUND({p_net},2)", "money0"),
            (f"ROUND({p_costs},2)", "money0"),
            (f"ROUND({p_profit},2)", "money0"),
            (f'IF(({p_net})<=0,"",{p_profit}/({p_net}))', "percent"),
            (f'IF(({p_nights})<=0,"",ROUND({p_profit}/({p_nights}),2))',
             "money"),
            (f'IF({y_profit}<=0,"",{p_profit}/{y_profit})', "percent"),
        ]
        for offset, (formula, style) in enumerate(cells):
            s.formula(
                row, offset + 1,
                s.reg.render(f'=IFERROR(IF({guard},"",{formula}),"")'),
                fmts.calc(style))
        row += 1
    row += 1

    # ----------------------------------------------- the twelve asks ---
    row = s.section(row, "The questions worth asking in January",
                    LAST_COL)
    month_head = f"$B${grid_top}:$M${grid_top}"
    net_band = (f"$B${sparkline_rows['Net revenue'] + 1}:"
                f"$M${sparkline_rows['Net revenue'] + 1}")
    profit_band = (f"$B${sparkline_rows['Profit'] + 1}:"
                   f"$M${sparkline_rows['Profit'] + 1}")
    row = s.table_head(row, 0, ["Question", "Answer", "Why it matters"],
                       [34, 16, 80])
    asks = [
        ("Which month carried the year?",
         f'=IFERROR(INDEX({month_head},MATCH(MAX({net_band}),'
         f'{net_band},0)),"")',
         "text",
         "If one month is more than a fifth of the year, your "
         "pricing the rest of the time is doing very little work."),
        ("Which month was the thinnest?",
         f'=IFERROR(INDEX({month_head},MATCH(MIN({net_band}),'
         f'{net_band},0)),"")',
         "text",
         "The one to attack with a minimum-stay change, a price "
         "drop or a different audience."),
        ("How many months lost money?",
         f'=COUNTIF({profit_band},"<0")',
         "int",
         "A month below zero is not a disaster; three in a row is a "
         "business model."),
        ("Best month against worst",
         f'=IF(MIN({net_band})<=0,"",MAX({net_band})/MIN({net_band}))',
         "number1",
         "How seasonal you really are. Above three and you are "
         "running a summer business with twelve months of costs."),
        ("Nights you never sold",
         f'=MAX(0,{y_avail}-{y_nights})', "int",
         "Each one is an ADR you will never get back."),
        ("What the empty nights cost",
         f'=ROUND(MAX(0,{y_avail}-{y_nights})*IF({y_nights}<=0,0,'
         f'{y_accom}/{y_nights}),2)', "money0",
         "Empty nights valued at your own average rate. Not a real "
         "loss - nobody sells every night - but the size of the "
         "prize."),
        ("Cleaning as a share of net revenue",
         f'=IF({y_net}<=0,"",{y_clean}/{y_net})', "percent",
         "Over about a fifth and either your stays are too short or "
         "your cleaning fee is too low."),
        ("Platform fees as a share of gross",
         f'=IF({y_gross}<=0,"",{y_fees}/{y_gross})', "percent",
         "Worth comparing with the rates on Settings. A gap means "
         "your mix has shifted towards the expensive platforms."),
        ("Cost of a booked night",
         f'=IF({y_nights}<=0,"",ROUND({y_costs}/{y_nights},2))',
         "money",
         "Your break-even nightly rate, near enough."),
        ("Profit a booked night",
         f'=IF({y_nights}<=0,"",ROUND({y_profit}/{y_nights},2))',
         "money", "What one more night sold is worth."),
        ("Profit a property",
         f'=IF(PropertyCount<=0,"",ROUND({y_profit}/PropertyCount,2))',
         "money0",
         "Divide by the hours you put in and decide whether this is "
         "a business or a hobby."),
        ("Cancellation rate",
         f'=IF({y_bookings}+{y_cancel}<=0,"",{y_cancel}/'
         f'({y_bookings}+{y_cancel}))', "percent",
         "Above one in ten, look at your deposit and cancellation "
         "terms."),
        ("Nights you blocked yourself",
         f"={y_blocked}", "int",
         "Your own stays and closures. Valued at your average rate "
         "they are the real cost of keeping the place available to "
         "you."),
        ("Average review score",
         '=IF(' + _year_count('[[Reviews Tracker.month]]',
                              '[[Reviews Tracker.has_score]],"Yes"') +
         '<=0,"",ROUND(' +
         _year_sum_c('[[Reviews Tracker.rate_overall]]',
                     '[[Reviews Tracker.month]]',
                     '[[Reviews Tracker.has_score]],"Yes"') + '/' +
         _year_count('[[Reviews Tracker.month]]',
                     '[[Reviews Tracker.has_score]],"Yes"') + ',2))',
         "number",
         "The number that quietly sets your occupancy next year."),
    ]
    for question, formula, style, why in asks:
        s.write(row, 0, question, fmts.chrome("label_bold"))
        s.formula(row, 1, (formula),
                         fmts.calc(style))
        s.merge(row, 2, LAST_COL, why, fmts.chrome("note"))
        s.row_height(row, 28)
        row += 1
    row += 1

    row = s.note(
        row,
        "Everything on this page attributes a stay to the year it "
        "checks in, and a cost to the date it carries. Change the "
        "year cell at the top and the whole page follows - including "
        "the comparison, which simply reads the year before.",
        LAST_COL, height=28)

    s.finish(freeze_row=6, freeze_col=1, gridlines=False,
             fit_width=True, print_area=f"A1:{col_letter(LAST_COL)}{row + 2}")
    s.protect()
    return s
