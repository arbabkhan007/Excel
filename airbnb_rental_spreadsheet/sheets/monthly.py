"""`Monthly Reports` - any month you choose, in full.

Type a month key at the top - `2026-03` - and the whole page becomes
that month: occupancy and rate, revenue by platform, costs by
category, every property compared, the guests who stayed and how the
month sits against the one before it and the same month a year ago.

It is the page to send an owner you co-host for, and the page to read
yourself on the first of the month.
"""

from __future__ import annotations

from ..config import Config, EXPENSE_CATEGORIES, Palette
from ..layout import Registry, col_letter
from ..writer import SheetWriter, list_from

NAME = "Monthly Reports"
LAST_COL = 9
PROPERTY_ROWS = 8


def _window(key: str) -> str:
    """Date bounds for a `yyyy-mm` key cell, as first and last day."""
    return (f'DATE(VALUE(LEFT({key},4)),VALUE(RIGHT({key},2)),1)',
            f'DATE(VALUE(LEFT({key},4)),VALUE(RIGHT({key},2))+1,1)-1')


def _period(prefix: str, key: str) -> dict:
    """Every total for one month key, as a dict of formula fragments."""
    earns = '[[Booking Tracker.earns]],"Yes"'
    bk = f'[[Booking Tracker.month]],{key}'
    nights = f'SUMIFS([[Booking Tracker.nights]],{bk},{earns})'
    accom = f'SUMIFS([[Booking Tracker.accommodation]],{bk},{earns})'
    gross = f'SUMIFS([[Booking Tracker.gross]],{bk},{earns})'
    fees = f'SUMIFS([[Booking Tracker.platform_fee_total]],{bk},{earns})'
    net = f'SUMIFS([[Booking Tracker.net_revenue]],{bk},{earns})'
    clean = f'SUMIFS([[Cleaning Schedule.total_cost]],' \
            f'[[Cleaning Schedule.month]],{key})'
    maint = f'SUMIFS([[Maintenance Log.cost_used]],' \
            f'[[Maintenance Log.month]],{key})'
    other = f'SUMIFS([[Expense Tracker.pl_amount]],' \
            f'[[Expense Tracker.month]],{key})'
    costs = f'({clean}+{maint}+{other})'
    blocked = f'SUMIFS([[Booking Tracker.nights]],{bk},' \
              f'[[Booking Tracker.is_block]],"Yes")'
    start, end = _window(key)
    days = f'DAY({end})'
    avail = f'MAX(1,PropertyCount*{days}-{blocked})'
    return {
        "bookings": f'COUNTIFS({bk},{earns})',
        "cancellations": f'COUNTIFS({bk},'
                         f'[[Booking Tracker.is_cancelled]],"Yes")',
        "nights": nights, "accom": accom, "gross": gross,
        "fees": fees, "net": net, "clean": clean, "maint": maint,
        "other": other, "costs": costs, "blocked": blocked,
        "avail": avail, "start": start, "end": end,
        "profit": f'({net}-{costs})',
        "occupancy": f'IF({avail}<=0,"",{nights}/{avail})',
        "adr": f'IF({nights}<=0,"",ROUND({accom}/{nights},2))',
        "revpar": f'IF({avail}<=0,"",ROUND({accom}/{avail},2))',
        "alos": f'IF(COUNTIFS({bk},{earns})<=0,"",ROUND({nights}/'
                f'COUNTIFS({bk},{earns}),1))',
        "guests": f'SUMIFS([[Booking Tracker.guests_total]],{bk},'
                  f'{earns})',
    }


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_mid)
    ws = s.ws

    s.banner("Monthly Report",
             "Pick a month. Occupancy, rate, revenue, costs and every "
             "property compared.", LAST_COL)
    s.set_col(0, 0, 26)
    for column in range(1, LAST_COL + 1):
        s.set_col(column, column, 14)

    row = 4
    s.write(row, 0, "Month", fmts.chrome("label_bold"))
    s.write(row, 1, cfg.today.strftime("%Y-%m"),
            fmts.chrome("form_box", locked=False))
    ws.data_validation(row, 1, row, 1, list_from("MonthKeyList"))
    book.define_name("ReportMonth", f"='{NAME}'!$B${row + 1}")
    s.merge(row, 2, 3, None, fmts.calc("text"))
    s.formula(
        row, 2,
        '=IF(ReportMonth="","Pick a month",TEXT(DATE(VALUE(LEFT('
        'ReportMonth,4)),VALUE(RIGHT(ReportMonth,2)),1),"mmmm yyyy"))',
        fmts.calc("text"))
    s.merge(row, 4, LAST_COL,
            "In yyyy-mm form. The list comes from Settings.",
            fmts.chrome("note"))
    row += 1

    # hidden previous-month and same-month-last-year keys
    s.write(row, 20, None, fmts.calc("text"))
    s.formula(
        row, 20,
        '=TEXT(DATE(VALUE(LEFT(ReportMonth,4)),VALUE(RIGHT('
        'ReportMonth,2))-1,1),"yyyy-mm")', fmts.calc("text"))
    book.define_name("PrevMonthKey", f"='{NAME}'!$U${row + 1}")
    s.write(row + 1, 20, None, fmts.calc("text"))
    s.formula(
        row + 1, 20,
        '=TEXT(DATE(VALUE(LEFT(ReportMonth,4))-1,VALUE(RIGHT('
        'ReportMonth,2)),1),"yyyy-mm")', fmts.calc("text"))
    book.define_name("LastYearMonthKey", f"='{NAME}'!$U${row + 2}")
    ws.set_column(20, 20, None, None, {"hidden": True})
    row += 1

    now = _period("now", "ReportMonth")
    prev = _period("prev", "PrevMonthKey")
    year = _period("year", "LastYearMonthKey")

    # ------------------------------------------------------ headline ----
    row = s.section(row, "The month", LAST_COL)
    s.kpi(row, 0, "Bookings", f'={now["bookings"]}', "int")
    s.kpi(row, 2, "Nights sold", f'={now["nights"]}', "int")
    s.kpi(row, 4, "Occupancy", f'={now["occupancy"]}', "percent")
    s.kpi(row, 6, "ADR", f'={now["adr"]}', "money",
          note="Accommodation only")
    row = s.kpi(row, 8, "RevPAR", f'={now["revpar"]}', "money")

    s.kpi(row, 0, "Gross", f'=ROUND({now["gross"]},2)', "money0")
    s.kpi(row, 2, "Platform fees", f'=ROUND({now["fees"]},2)', "money0")
    s.kpi(row, 4, "Net revenue", f'=ROUND({now["net"]},2)', "money0")
    s.kpi(row, 6, "Costs", f'=ROUND({now["costs"]},2)', "money0")
    row = s.kpi(row, 8, "Profit", f'=ROUND({now["profit"]},2)',
                "money0")
    row += 1

    # ------------------------------------------------- the comparison ---
    row = s.section(row, "Against last month and last year", LAST_COL)
    row = s.table_head(row, 0, [
        "Measure", "This month", "Last month", "Change",
        "Same month last year", "Change", "Target", "Against target"],
        [26, 14, 14, 12, 18, 12, 13, 14])
    measures = [
        ("Bookings", "bookings", "int", ""),
        ("Nights sold", "nights", "int", ""),
        ("Occupancy", "occupancy", "percent", "TargetOccupancy"),
        ("ADR", "adr", "money", "TargetADR"),
        ("RevPAR", "revpar", "money", "TargetRevPAR"),
        ("Average stay", "alos", "number1", ""),
        ("Guests hosted", "guests", "int", ""),
        ("Gross", "gross", "money0", ""),
        ("Platform fees", "fees", "money0", ""),
        ("Net revenue", "net", "money0", "TargetRevenueMonth"),
        ("Cleaning", "clean", "money0", ""),
        ("Repairs", "maint", "money0", ""),
        ("Other costs", "other", "money0", ""),
        ("Total costs", "costs", "money0", ""),
        ("Profit", "profit", "money0", "TargetProfitMonth"),
    ]
    compare_top = row
    for label, key, style, target in measures:
        a, b, c = now[key], prev[key], year[key]
        pct = (style in ("percent",))
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.formula(row, 1, (f"={a}"),
                         fmts.calc(style))
        s.formula(row, 2, s.reg.render(f"={b}"),
                         fmts.calc(style))
        if pct:
            s.formula(
                row, 3,
                s.reg.render(f'=IF({b}="","",IF({a}="","",{a}-{b}))'),
                fmts.calc("percent"))
        else:
            s.formula(
                row, 3,
                s.reg.render(f'=IF({b}="","",IF({b}<=0,"",IF({a}="",'
                             f'"",{a}/{b}-1)))'),
                fmts.calc("percent"))
        s.formula(row, 4, s.reg.render(f"={c}"),
                         fmts.calc(style))
        if pct:
            s.formula(
                row, 5,
                s.reg.render(f'=IF({c}="","",IF({a}="","",{a}-{c}))'),
                fmts.calc("percent"))
        else:
            s.formula(
                row, 5,
                s.reg.render(f'=IF({c}="","",IF({c}<=0,"",IF({a}="",'
                             f'"",{a}/{c}-1)))'),
                fmts.calc("percent"))
        if target:
            s.formula(row, 6, f"={target}", fmts.calc(style))
            s.formula(
                row, 7,
                s.reg.render(f'=IF({target}<=0,"",IF({a}="","",{a}/'
                             f'{target}))'), fmts.calc("percent"))
        else:
            s.write(row, 6, None, fmts.calc("text"))
            s.write(row, 7, None, fmts.calc("text"))
        row += 1
    for column in (3, 5):
        ws.conditional_format(compare_top, column, row - 1, column, {
            "type": "cell", "criteria": "<", "value": 0,
            "format": fmts.cf("bad")})
        ws.conditional_format(compare_top, column, row - 1, column, {
            "type": "cell", "criteria": ">", "value": 0,
            "format": fmts.cf("good")})
    ws.conditional_format(compare_top, 7, row - 1, 7, {
        "type": "cell", "criteria": ">=", "value": 1,
        "format": fmts.cf("good")})
    ws.conditional_format(compare_top, 7, row - 1, 7, {
        "type": "cell", "criteria": "<", "value": 0.8,
        "format": fmts.cf("warn")})
    row += 1

    # ---------------------------------------------------- by platform ---
    row = s.section(row, "Where the bookings came from", LAST_COL)
    row = s.table_head(row, 0, [
        "Platform", "Bookings", "Nights", "Gross", "Commission",
        "Card fee", "Net", "Effective fee", "Share of net"],
        [26, 11, 10, 13, 13, 11, 13, 13, 12])
    earns = '[[Booking Tracker.earns]],"Yes"'
    bk = '[[Booking Tracker.month]],ReportMonth'
    for name, _c, _p, _counts, _note in cfg.platform_table:
        plat = f'[[Booking Tracker.platform]],"{name}"'
        p_gross = f'SUMIFS([[Booking Tracker.gross]],{bk},{earns},{plat})'
        p_comm = (f'SUMIFS([[Booking Tracker.platform_fee]],{bk},'
                  f'{earns},{plat})')
        p_card = (f'SUMIFS([[Booking Tracker.payment_fee]],{bk},'
                  f'{earns},{plat})')
        p_net = (f'SUMIFS([[Booking Tracker.net_revenue]],{bk},{earns},'
                 f'{plat})')
        s.calc_row(row, 0, [
            name,
            f"=COUNTIFS({bk},{earns},{plat})",
            f"=SUMIFS([[Booking Tracker.nights]],{bk},{earns},{plat})",
            f"=ROUND({p_gross},2)",
            f"=ROUND({p_comm},2)",
            f"=ROUND({p_card},2)",
            f"=ROUND({p_net},2)",
            f'=IF({p_gross}<=0,"",({p_comm}+{p_card})/{p_gross})',
            f'=IF({now["net"]}<=0,"",{p_net}/{now["net"]})',
        ], ["text", "int", "int", "money", "money", "money", "money",
            "percent", "percent"])
        row += 1
    s.write(row, 0, "All platforms", fmts.chrome("total_label"))
    for offset, (formula, style) in enumerate((
            (f'={now["bookings"]}', "int"),
            (f'={now["nights"]}', "int"),
            (f'=ROUND({now["gross"]},2)', "money"),
            ('=ROUND(SUMIFS([[Booking Tracker.platform_fee]],'
             f'{bk},{earns}),2)', "money"),
            ('=ROUND(SUMIFS([[Booking Tracker.payment_fee]],'
             f'{bk},{earns}),2)', "money"),
            (f'=ROUND({now["net"]},2)', "money"),
            (f'=IF({now["gross"]}<=0,"",{now["fees"]}/{now["gross"]})',
             "percent"),
            ("=1", "percent"))):
        s.formula(row, offset + 1, s.reg.render(formula),
                         fmts.chrome("total_value",
                                     num_format=s.num_format(style)))
    row += 2

    # ---------------------------------------------------- by property ---
    row = s.section(row, "Every property this month", LAST_COL)
    row = s.table_head(row, 0, [
        "Property", "Nights", "Occupancy", "ADR", "Net revenue",
        "Cleaning", "Other costs", "Profit", "Margin", "Rating"],
        [26, 10, 11, 11, 13, 11, 12, 12, 10, 9])
    start, end = _window("ReportMonth")
    for n in range(1, PROPERTY_ROWS + 1):
        pid = f'INDEX([[Property Database.id]],{n})'
        guard = f'INDEX([[Property Database.name]],{n})=""'
        prop = f'[[Booking Tracker.property]],{pid}'
        p_nights = (f'SUMIFS([[Booking Tracker.nights]],{bk},{earns},'
                    f'{prop})')
        p_accom = (f'SUMIFS([[Booking Tracker.accommodation]],{bk},'
                   f'{earns},{prop})')
        p_net = (f'SUMIFS([[Booking Tracker.net_revenue]],{bk},{earns},'
                 f'{prop})')
        p_clean = (f'SUMIFS([[Cleaning Schedule.total_cost]],'
                   f'[[Cleaning Schedule.month]],ReportMonth,'
                   f'[[Cleaning Schedule.property]],{pid})')
        p_other = (f'(SUMIFS([[Maintenance Log.cost_used]],'
                   f'[[Maintenance Log.month]],ReportMonth,'
                   f'[[Maintenance Log.property]],{pid})+'
                   f'SUMIFS([[Expense Tracker.pl_amount]],'
                   f'[[Expense Tracker.month]],ReportMonth,'
                   f'[[Expense Tracker.property]],{pid}))')
        p_blocked = (f'SUMIFS([[Booking Tracker.nights]],{bk},{prop},'
                     f'[[Booking Tracker.is_block]],"Yes")')
        p_avail = f'MAX(1,DAY({end})-{p_blocked})'
        p_profit = f'({p_net}-{p_clean}-{p_other})'
        p_scored = ('COUNTIFS([[Reviews Tracker.property]],' + pid +
                    ',[[Reviews Tracker.month]],ReportMonth,'
                    '[[Reviews Tracker.has_score]],"Yes")')
        p_rating = ('SUMIFS([[Reviews Tracker.rate_overall]],'
                    '[[Reviews Tracker.property]],' + pid +
                    ',[[Reviews Tracker.month]],ReportMonth,'
                    '[[Reviews Tracker.has_score]],"Yes")')
        s.formula(
            row, 0,
            s.reg.render(f'=IFERROR(IF({guard},"",INDEX('
                         f'[[Property Database.name]],{n})),"")'),
            fmts.calc("text"))
        cells = [
            (p_nights, "int"),
            (f'IF({p_avail}<=0,"",{p_nights}/{p_avail})', "percent"),
            (f'IF({p_nights}<=0,"",ROUND({p_accom}/{p_nights},2))',
             "money"),
            (f"ROUND({p_net},2)", "money"),
            (f"ROUND({p_clean},2)", "money"),
            (f"ROUND({p_other},2)", "money"),
            (f"ROUND({p_profit},2)", "money"),
            (f'IF({p_net}<=0,"",{p_profit}/{p_net})', "percent"),
            (f'IF({p_scored}<=0,"",ROUND({p_rating}/{p_scored},2))',
             "number"),
        ]
        for offset, (formula, style) in enumerate(cells):
            s.formula(
                row, offset + 1,
                s.reg.render(f'=IFERROR(IF({guard},"",{formula}),"")'),
                fmts.calc(style))
        row += 1
    row += 1

    # ------------------------------------------------- costs breakdown --
    row = s.section(row, "What the month cost", LAST_COL)
    row = s.table_head(row, 0, [
        "Category", "Amount", "Share of costs", "Share of net revenue",
        "Entries"], [26, 14, 15, 19, 11])
    cost_total = now["costs"]
    s.calc_row(row, 0, [
        "Cleaning and turnovers",
        f'=ROUND({now["clean"]},2)',
        f'=IF({cost_total}<=0,"",{now["clean"]}/{cost_total})',
        f'=IF({now["net"]}<=0,"",{now["clean"]}/{now["net"]})',
        '=COUNTIFS([[Cleaning Schedule.month]],ReportMonth)',
    ], ["text", "money", "percent", "percent", "int"])
    row += 1
    s.calc_row(row, 0, [
        "Repairs and maintenance",
        f'=ROUND({now["maint"]},2)',
        f'=IF({cost_total}<=0,"",{now["maint"]}/{cost_total})',
        f'=IF({now["net"]}<=0,"",{now["maint"]}/{now["net"]})',
        '=COUNTIFS([[Maintenance Log.month]],ReportMonth)',
    ], ["text", "money", "percent", "percent", "int"])
    row += 1
    for category in EXPENSE_CATEGORIES:
        amount = (f'SUMIFS([[Expense Tracker.pl_amount]],'
                  f'[[Expense Tracker.month]],ReportMonth,'
                  f'[[Expense Tracker.category]],"{category}")')
        count = (f'COUNTIFS([[Expense Tracker.month]],ReportMonth,'
                 f'[[Expense Tracker.category]],"{category}")')
        s.calc_row(row, 0, [
            category,
            f"=ROUND({amount},2)",
            f'=IF({cost_total}<=0,"",{amount}/{cost_total})',
            f'=IF({now["net"]}<=0,"",{amount}/{now["net"]})',
            f"={count}",
        ], ["text", "money", "percent", "percent", "int"])
        row += 1
    s.write(row, 0, "Total", fmts.chrome("total_label"))
    s.formula(row, 1, (f'=ROUND({cost_total},2)'),
                     fmts.chrome("total_value",
                                 num_format=s.num_format("money")))
    s.formula(row, 2, "=1",
                     fmts.chrome("total_value", num_format="0.0%"))
    s.formula(
        row, 3,
        s.reg.render(f'=IF({now["net"]}<=0,"",{cost_total}/'
                     f'{now["net"]})'),
        fmts.chrome("total_value", num_format="0.0%"))
    row += 2

    # ------------------------------------------------- operations -------
    row = s.section(row, "How the month ran", LAST_COL)
    row = s.table_head(row, 0, ["Measure", "Value", "What it tells you"],
                       [26, 14, 70])
    ops = [
        ("Turnovers", '=COUNTIFS([[Cleaning Schedule.month]],'
                      'ReportMonth)',
         "int", "Should be close to the number of check-outs."),
        ("Same-day turnovers",
         '=COUNTIFS([[Cleaning Schedule.month]],ReportMonth,'
         '[[Cleaning Schedule.same_day]],"Yes")', "int",
         "The ones that go wrong. Worth pricing for."),
        ("Turnovers left incomplete",
         '=COUNTIFS([[Cleaning Schedule.month]],ReportMonth,'
         '[[Cleaning Schedule.completion]],"<1")', "int",
         "A task unticked is a task not done."),
        ("Cleaning cost a booking",
         f'=IF({now["bookings"]}<=0,"",ROUND({now["clean"]}/'
         f'{now["bookings"]},2))', "money",
         "The number that makes one-night stays look expensive."),
        ("Cost a booked night",
         f'=IF({now["nights"]}<=0,"",ROUND({now["costs"]}/'
         f'{now["nights"]},2))', "money",
         "Compare with ADR. The gap is the real margin."),
        ("Profit a booked night",
         f'=IF({now["nights"]}<=0,"",ROUND({now["profit"]}/'
         f'{now["nights"]},2))', "money",
         "What one more night sold is actually worth."),
        ("Cancellations", f'={now["cancellations"]}', "int",
         "Rows marked cancelled with a check-in in this month."),
        ("Nights you blocked", f'={now["blocked"]}', "int",
         "Owner stays and closures. Taken out of occupancy."),
        ("Reviews received",
         '=COUNTIFS([[Reviews Tracker.month]],ReportMonth)', "int", ""),
        ("Average review score",
         '=IF(COUNTIFS([[Reviews Tracker.month]],ReportMonth,'
         '[[Reviews Tracker.has_score]],"Yes")<=0,"",ROUND(SUMIFS('
         '[[Reviews Tracker.rate_overall]],[[Reviews Tracker.month]],'
         'ReportMonth,[[Reviews Tracker.has_score]],"Yes")/COUNTIFS('
         '[[Reviews Tracker.month]],ReportMonth,'
         '[[Reviews Tracker.has_score]],"Yes"),2))', "number", ""),
        ("Payouts received",
         '=ROUND(SUMIFS([[Platform Revenue & Fees.received]],'
         '[[Platform Revenue & Fees.month]],ReportMonth),2)', "money",
         "What actually landed in the bank this month."),
        ("Fee variance",
         '=ROUND(SUMIFS([[Platform Revenue & Fees.variance]],'
         '[[Platform Revenue & Fees.month]],ReportMonth),2)', "money",
         "Real payouts against the fee rates you assumed."),
    ]
    for label, formula, style, note in ops:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.formula(row, 1, (formula),
                         fmts.calc(style))
        s.merge(row, 2, LAST_COL, note, fmts.chrome("note"))
        row += 1
    row += 1

    row = s.note(
        row,
        "A booking belongs to the month it checks in, which keeps one "
        "stay in one month and makes the columns add up. Costs belong "
        "to the month they are dated. If you want a stay split across "
        "two months, that is an accounting policy rather than a "
        "spreadsheet setting - note it and stay consistent.",
        LAST_COL, height=34)

    s.finish(freeze_row=6, gridlines=False, fit_width=True,
             print_area=f"A1:{col_letter(LAST_COL)}{row + 2}")
    s.protect()
    return s
