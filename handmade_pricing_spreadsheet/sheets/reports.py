"""`Reports` - the numbers you would want if someone asked how the year went.

Six reports, all built from the same registers: profit by product,
where the material money goes, what your hours are worth, the month by
month profit and loss, how your prices sit against the market, and the
year against your targets.

Nothing here needs refreshing and nothing here is typed twice.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter

NAME = "Reports"
LAST_COL = 11
TOP = 12
MATERIAL_ROWS = 12
MONTHS = 12


def _rank(rank_key: str, field: str, n: int) -> str:
    return (f'IFERROR(INDEX([[Products.{field}]],MATCH({n},'
            f'[[Products.{rank_key}]],0)),"")')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.accent)
    ws = s.ws

    s.banner("Reports & Analytics",
             "Profit by product, where the material money goes, what "
             "your hours earn, and the year against your targets.",
             LAST_COL)
    s.set_col(0, 0, 30)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 13)

    row = 4

    # -------------------------------------------- 1. profit by product --
    row = s.section(row, "1. Profitability by product", LAST_COL)
    row = s.note(row, "Ranked by the profit each sale makes. The hourly "
                      "column is the one that should decide what you "
                      "make next week.",
                 LAST_COL, height=26)
    row = s.table_head(row, 0, ["Product", "Category", "Cost", "Price",
                                "Profit", "Margin", "Markup", "Hours",
                                "An hour", "Sold", "Revenue", "Profit made"],
                       [30, 16, 11, 11, 11, 10, 10, 9, 11, 9, 13, 13])
    first = row
    for n in range(1, TOP + 1):
        fields = [("name", "text"), ("category", "text"),
                  ("total_cost", "money"), ("price", "money"),
                  ("profit", "money"), ("margin", "percent"),
                  ("markup", "percent"), ("hours", "number1"),
                  ("profit_hour", "money"), ("sold", "int"),
                  ("revenue", "money0"), ("actual_profit", "money0")]
        for col, (field, style) in enumerate(fields):
            ws.write_formula(row, col,
                             s.reg.render("=" + _rank("profit_rank", field,
                                                      n)),
                             fmts.calc(style))
        row += 1
    ws.conditional_format(first, 4, row - 1, 4,
                          {"type": "data_bar", "bar_color": "#8C9E6E"})
    ws.conditional_format(first, 5, row - 1, 5,
                          {"type": "cell", "criteria": "<",
                           "value": "=MarginWarning",
                           "format": fmts.cf("warn")})
    ws.conditional_format(first, 8, row - 1, 8,
                          {"type": "cell", "criteria": "<",
                           "value": "=MinimumWage",
                           "format": fmts.cf("bad")})
    row += 2

    # ------------------------------------------------ 2. material spend --
    row = s.section(row, "2. Where the material money goes", LAST_COL)
    row = s.table_head(row, 0, ["Material", "Type", "Used in", "Unit cost",
                                "Cost used", "Stock value", "Waste",
                                "Share of material cost"],
                       [30, 18, 11, 13, 13, 13, 11, 20])
    mat_first = row
    for n in range(1, MATERIAL_ROWS + 1):
        name = f'IFERROR(INDEX([[Materials.name]],MATCH({n},' \
               f'[[Materials.cost_used_rank]],0)),"")'
        mat_id = f'IFERROR(INDEX([[Materials.id]],MATCH({n},' \
                 f'[[Materials.cost_used_rank]],0)),"")'
        ws.write_formula(row, 0, s.reg.render(f"={name}"),
                         fmts.calc("text"))
        for col, field, style in ((1, "type", "text"),
                                  (2, "lines", "int"),
                                  (3, "unit_cost", "money4"),
                                  (4, "cost_used", "money"),
                                  (5, "stock_value", "money")):
            body = (f'IFERROR(INDEX([[Materials.{field}]],MATCH({n},'
                    f'[[Materials.cost_used_rank]],0)),"")')
            ws.write_formula(row, col, s.reg.render(f"={body}"),
                             fmts.calc(style))
        ws.write_formula(
            row, 6,
            s.reg.render(f'=IF({mat_id}="","",ROUND(SUMIFS('
                         f'[[Waste Log.cost]],[[Waste Log.material]],'
                         f'{mat_id}),2))'), fmts.calc("money"))
        ws.write_formula(
            row, 7,
            s.reg.render(f'=IFERROR(E{row + 1}/SUM([[Materials.cost_used]]),'
                         f'"")'), fmts.calc("percent"))
        row += 1
    ws.conditional_format(mat_first, 7, row - 1, 7,
                          {"type": "data_bar", "bar_color": "#B5704F"})
    row += 1
    s.merge(row, 0, LAST_COL,
            "If two or three materials dominate this list, that is "
            "where a supplier conversation or a bulk order is worth "
            "the most.", fmts.chrome("note"))
    row += 2

    # --------------------------------------------------- 3. your hours --
    row = s.section(row, "3. What your hours are worth", LAST_COL)
    row = s.table_head(row, 0, ["Measure", "Value", "What it means"],
                       [30, 15, 60])
    hours = [
        ("Your hourly rate", "=HourlyRate", "money",
         "What you decided your time is worth, on the Settings tab."),
        ("Average earned an hour, by product",
         '=IFERROR(AVERAGE([[Products.wage_hour]]),"")', "money",
         "Across the catalog, at your current prices."),
        ("Best paid product an hour",
         '=IFERROR(MAX([[Products.wage_hour]]),"")', "money",
         "Worth knowing what it is and making more of it."),
        ("Worst paid product an hour",
         '=IFERROR(MIN([[Products.wage_hour]]),"")', "money",
         "Worth knowing what it is and re-pricing it."),
        ("Products that pay under minimum wage",
         '=COUNTIFS([[Products.wage_hour]],"<"&MinimumWage,'
         '[[Products.name]],"<>")', "int",
         "Each one of these is a decision to work for less."),
        ("Hours sold in the last twelve months",
         '=ROUND(SUM([[Sales.hours]]),1)', "number1",
         "Making time that actually turned into money."),
        ("Profit an hour, actual",
         '=IFERROR(ROUND(SUM([[Sales.profit]])/SUM([[Sales.hours]]),2),"")',
         "money",
         "Total profit over total making hours - the real figure."),
        ("Productive hours you assumed",
         "=ProductiveHours", "number",
         "What the overhead per hour is based on. Be honest about "
         "it."),
    ]
    hour_first = row
    for label, formula, style, what in hours:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        ws.write_formula(row, 1, s.reg.render(formula), fmts.calc(style))
        s.merge(row, 2, LAST_COL, what, fmts.chrome("note"))
        row += 1
    ws.conditional_format(hour_first + 1, 1, hour_first + 3, 1,
                          {"type": "cell", "criteria": "<",
                           "value": "=MinimumWage",
                           "format": fmts.cf("bad")})
    row += 2

    # --------------------------------------------- 4. profit and loss ----
    row = s.section(row, "4. Month by month", LAST_COL)
    s.write(row, 0, "Month", fmts.chrome("header_left"))
    for i in range(MONTHS):
        back = MONTHS - 1 - i
        ws.write_formula(
            row, i + 1,
            f'=TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{back},1),'
            f'"mmm yy")', fmts.chrome("header"))
        ws.write_formula(
            row + 1, i + 1,
            f'=TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{back},1),'
            f'"yyyy-mm")', fmts.chrome("header", font_size=8))
    s.write(row + 1, 0, "Key", fmts.chrome("label"))
    key_row = row + 1
    row += 2
    band = {}
    pl_lines = [
        ("Revenue", "SUMIFS([[Sales.revenue]],[[Sales.month]],{key})",
         "money0"),
        ("Materials & labour",
         "SUMIFS([[Sales.cogs]],[[Sales.month]],{key})", "money0"),
        ("Selling fees", "SUMIFS([[Sales.fees]],[[Sales.month]],{key})",
         "money0"),
        ("Postage absorbed",
         "SUMIFS([[Sales.ship_absorbed]],[[Sales.month]],{key})",
         "money0"),
    ]
    for label, body, style in pl_lines:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        for i in range(MONTHS):
            cell = f"{chr(66 + i)}${key_row + 1}"
            ws.write_formula(row, i + 1,
                             s.reg.render("=ROUND(" +
                                          body.format(key=cell) + ",2)"),
                             fmts.calc(style))
        band[label] = row
        row += 1
    s.write(row, 0, "Overhead", fmts.chrome("label_bold"))
    for i in range(MONTHS):
        ws.write_formula(row, i + 1, "=ROUND(OverheadMonthly,2)",
                         fmts.calc("money0"))
    band["Overhead"] = row
    row += 1
    s.write(row, 0, "Profit after overhead", fmts.chrome("total_label"))
    for i in range(MONTHS):
        col = chr(66 + i)
        parts = "-".join(f"{col}{band[k] + 1}" for k in
                         ("Materials & labour", "Selling fees",
                          "Postage absorbed", "Overhead"))
        ws.write_formula(
            row, i + 1,
            f'=ROUND({col}{band["Revenue"] + 1}-{parts},2)',
            fmts.chrome("total_value", num_format=s.num_format("money0")))
    profit_row = row
    row += 1
    s.write(row, 0, "Margin after overhead", fmts.chrome("label_bold"))
    for i in range(MONTHS):
        col = chr(66 + i)
        ws.write_formula(
            row, i + 1,
            f'=IFERROR({col}{profit_row + 1}/{col}'
            f'{band["Revenue"] + 1},"")', fmts.calc("percent0"))
    margin_row = row
    row += 1
    ws.add_sparkline(profit_row, MONTHS + 1, {
        "range": (f"{quote_sheet(NAME)}!B{profit_row + 1}:"
                  f"M{profit_row + 1}"),
        "type": "column", "negative_points": True,
        "series_color": "#5E7B5A", "negative_color": "#B3261E",
    })
    ws.conditional_format(profit_row, 1, profit_row, MONTHS,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(margin_row, 1, margin_row, MONTHS,
                          {"type": "cell", "criteria": ">=",
                           "value": "=TargetRetailMargin",
                           "format": fmts.cf("good")})
    s.set_col(MONTHS + 1, MONTHS + 1, 12)
    row += 1
    s.merge(row, 0, LAST_COL,
            "Overhead is charged at the same amount every month "
            "because that is how it arrives. A month that does not "
            "cover it is a month the studio paid for itself out of "
            "another month.", fmts.chrome("note"))
    row += 2

    # --------------------------------------------- 5. against the market
    row = s.section(row, "5. Your prices against the market", LAST_COL)
    row = s.table_head(row, 0, ["Product", "Your price", "Market average",
                                "Difference", "You are", "Research notes"],
                       [30, 12, 15, 12, 14, 34])
    comp_first = row
    for n in range(1, TOP + 1):
        name = _rank("margin_rank", "name", n)
        ws.write_formula(row, 0, s.reg.render("=" + name),
                         fmts.calc("text"))
        ws.write_formula(row, 1,
                         s.reg.render("=" + _rank("margin_rank", "price", n)),
                         fmts.calc("money"))
        ws.write_formula(
            row, 2,
            s.reg.render("=" + _rank("margin_rank", "comp_price", n)),
            fmts.calc("money"))
        ws.write_formula(
            row, 3,
            f'=IF(OR(B{row + 1}="",C{row + 1}="",C{row + 1}=0),"",'
            f'ROUND(B{row + 1}-C{row + 1},2))', fmts.calc("money"))
        ws.write_formula(
            row, 4,
            f'=IF(D{row + 1}="","No research yet",'
            f'IF(D{row + 1}>C{row + 1}*0.15,"Dearer",'
            f'IF(D{row + 1}<-C{row + 1}*0.15,"Cheaper","In line")))',
            fmts.calc("text"))
        ws.write_formula(
            row, 5,
            f'=IF(A{row + 1}="","",IF(E{row + 1}="Cheaper",'
            f'"Worth testing a rise",IF(E{row + 1}="Dearer",'
            f'"Make sure the listing shows why","")))',
            fmts.calc("text"))
        row += 1
    ws.conditional_format(comp_first, 4, row - 1, 4,
                          {"type": "text", "criteria": "containing",
                           "value": "Cheaper", "format": fmts.cf("warn")})
    ws.conditional_format(comp_first, 4, row - 1, 4,
                          {"type": "text", "criteria": "containing",
                           "value": "In line", "format": fmts.cf("good")})
    row += 2

    # ------------------------------------------------ 6. year to date ----
    row = s.section(row, "6. The year against your targets", LAST_COL)
    row = s.table_head(row, 0, ["Measure", "This year so far", "Target",
                                "Difference", "Where you are"],
                       [30, 16, 14, 14, 30])
    targets = [
        ("Revenue",
         '=ROUND(SUMIFS([[Sales.revenue]],[[Sales.date]],">="&YearStart),2)',
         "=RevenueTargetYear", "money0"),
        ("Profit",
         '=ROUND(SUMIFS([[Sales.profit]],[[Sales.date]],">="&YearStart),2)',
         "=ProfitTargetMonth*12", "money0"),
        ("Items sold",
         '=SUMIFS([[Sales.qty]],[[Sales.date]],">="&YearStart)',
         "=UnitsPerMonth*12", "int"),
        ("Earned an hour",
         '=IFERROR(ROUND(SUMIFS([[Sales.profit]],[[Sales.date]],'
         '">="&YearStart)/MAX(0.1,SUMIFS([[Sales.hours]],'
         '[[Sales.date]],">="&YearStart)),2),"")',
         "=TargetHourly", "money"),
    ]
    for label, actual, target, style in targets:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        ws.write_formula(row, 1, s.reg.render(actual), fmts.calc(style))
        ws.write_formula(row, 2, target, fmts.calc(style))
        ws.write_formula(row, 3,
                         f'=IF(B{row + 1}="","",ROUND(B{row + 1}-'
                         f'C{row + 1},2))', fmts.calc(style))
        s.merge(row, 4, LAST_COL, None, fmts.calc("text"))
        ws.write_formula(
            row, 4,
            f'=IF(OR(B{row + 1}="",C{row + 1}=0),"",'
            f'TEXT(B{row + 1}/C{row + 1},"0%")&" of the year\'s '
            f'target")', fmts.calc("text"))
        row += 1
    row += 1

    s.note(row,
           "Everything on this page comes from the Sales, Products, "
           "Materials and Waste tabs. If a report looks wrong, the "
           "fix is on one of those tabs rather than here.",
           LAST_COL, height=28)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
