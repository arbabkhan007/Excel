"""`Dashboard` - one page that answers "am I actually making money?"

Twelve numbers across the top, a list of the things that need fixing
today, your best and worst products by the measures that matter, and
twelve months of revenue and profit. Everything reads from the
registers, so it is never out of date and there is nothing to refresh.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter

NAME = "Dashboard"
LAST_COL = 13
MONTHS = 12
TOP_ROWS = 8


def _by_rank(rank_key: str, field: str, n: int) -> str:
    """The nth product by one of the catalog's hidden rank columns."""
    return (f'IFERROR(INDEX([[Products.{field}]],MATCH({n},'
            f'[[Products.{rank_key}]],0)),"")')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws

    s.banner("Pricing Dashboard",
             "Your range, your margins and your month - worked out "
             "from the products, costs and sales you have already "
             "entered.", LAST_COL)
    s.set_col(0, 0, 24)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 13)

    # ---------------------------------------------------------- the header
    row = 4
    s.merge(row, 0, 0, "Shop", fmts.chrome("label_bold"))
    s.merge(row, 1, 3, None, fmts.calc("text"))
    ws.write_formula(row, 1, "=ShopName", fmts.calc("text"))
    s.merge(row, 4, 4, "Month", fmts.chrome("label_right"))
    ws.write_formula(row, 5, '=TEXT(TodayDate,"mmmm yyyy")',
                     fmts.calc("text"))
    s.merge(row, 6, 6, "Month key", fmts.chrome("label_right"))
    ws.write_formula(row, 7, '=TEXT(TodayDate,"yyyy-mm")',
                     fmts.calc("text"))
    book.define_name("DashKey", f"={quote_sheet(NAME)}!$H${row + 1}")
    s.merge(row, 8, 9, "Today", fmts.chrome("label_right"))
    s.merge(row, 10, LAST_COL, None, fmts.calc("date"))
    ws.write_formula(row, 10, "=TodayDate", fmts.calc("date"))
    row += 2

    # ------------------------------------------------------------ the KPIs
    money = cfg.currency_symbol
    tiles_1 = [
        ("Products priced",
         '=COUNTIF([[Products.name]],"<>")', "int",
         "in the catalog"),
        ("Average margin",
         '=IFERROR(AVERAGE([[Products.margin]]),"")', "percent",
         f"target {cfg.target_retail_margin:.0%}"),
        ("Average you earn an hour",
         '=IFERROR(AVERAGE([[Products.wage_hour]]),"")', "money",
         f"your rate is {money}{cfg.hourly_rate:,.2f}"),
        ("Products below target",
         '=COUNTIFS([[Products.margin]],"<"&MarginWarning,'
         '[[Products.name]],"<>")', "int",
         "margin under your warning level"),
    ]
    tiles_2 = [
        ("Revenue this month",
         '=ROUND(SUMIFS([[Sales.revenue]],[[Sales.month]],DashKey),2)',
         "money0", f"target {money}{cfg.revenue_target_month:,.0f}"),
        ("Profit this month",
         '=ROUND(SUMIFS([[Sales.profit]],[[Sales.month]],DashKey),2)',
         "money0", "after fees and postage"),
        ("Margin this month",
         '=IFERROR(SUMIFS([[Sales.profit]],[[Sales.month]],DashKey)/'
         'SUMIFS([[Sales.revenue]],[[Sales.month]],DashKey),"")',
         "percent", "what you actually kept"),
        ("Items sold this month",
         '=SUMIFS([[Sales.qty]],[[Sales.month]],DashKey)', "int",
         "across every channel"),
    ]
    tiles_3 = [
        ("Overhead a month",
         "=ROUND(OverheadMonthly,2)", "money0",
         "studio, insurance, software"),
        ("Covered by this month",
         '=IFERROR(SUMIFS([[Sales.profit]],[[Sales.month]],DashKey)/'
         'MAX(0.01,OverheadMonthly),"")', "percent0",
         "100% means the studio has paid for itself"),
        ("Value of your materials",
         '=ROUND(SUM([[Materials.stock_value]]),2)', "money0",
         "money sitting on the shelf"),
        ("Hours of work sold",
         '=ROUND(SUMIFS([[Sales.hours]],[[Sales.month]],DashKey),1)',
         "number1", "this month"),
    ]
    for tiles in (tiles_1, tiles_2, tiles_3):
        for i, (label, formula, style, note) in enumerate(tiles):
            s.kpi(row, i * 3, label, formula, style=style, span=3,
                  note=note)
        row += 4

    # --------------------------------------------------- needs attention --
    row = s.section(row, "Needs your attention", LAST_COL)
    row = s.table_head(row, 0, ["What", "How many", "What it means",
                                "Where to fix it"],
                       [26, 11, 46, 20])
    alerts = [
        ("Products with no price",
         '=COUNTIFS([[Products.name]],"<>",[[Products.price]],"")',
         '="Listed in the catalog with nothing to sell them at"',
         "Products"),
        ("Products that lose money",
         '=COUNTIF([[Products.alert]],"Loss*")',
         '="The price does not cover the cost and the fees"',
         "Products"),
        ("Products below minimum wage",
         '=COUNTIF([[Products.alert]],"Below*")',
         '="You would earn less than "&TEXT(MinimumWage,"0.00")&'
         '" an hour making these"', "Products"),
        ("Products under your margin warning",
         '=COUNTIFS([[Products.margin]],"<"&MarginWarning,'
         '[[Products.name]],"<>")',
         '="Thinner than "&TEXT(MarginWarning,"0%")&" - one bad month '
         'away from nothing"', "Retail Calculator"),
        ("Products with no costs entered",
         '=COUNTIF([[Products.alert]],"Add - no costs*")',
         '="Nothing on the Material Usage or Labor tabs for these"',
         "Material Usage"),
        ("Materials with no price",
         '=COUNTIF([[Materials.alert]],"Add*")',
         '="Every product using them is under-costed"', "Materials"),
        ("Materials running low",
         '=COUNTIF([[Materials.status]],"Low stock")+'
         'COUNTIF([[Materials.status]],"Out of stock")',
         '="Below the reorder level or less than "&LowStockDays&'
         '" days of cover"', "Materials"),
        ("Prices going stale",
         '=COUNTIF([[Products.alert]],"Review*")',
         '="Not changed in "&PriceReviewDays&" days, while your costs '
         'moved"', "Price History"),
        ("Fee settings to check",
         '=IF(FeesStale="Yes",1,0)',
         '="Marketplace fees change. Yours were last reviewed "&'
         'TEXT(FeesReviewed,"d mmm yyyy")', "Channels"),
        ("Orders that lost money",
         '=COUNTIFS([[Sales.profit]],"<0")',
         '="Real orders where the money went the wrong way"',
         "Sales"),
        ("Batches costing more than the estimate",
         '=COUNTIF([[Batches.alert]],"Check - costing*")',
         '="The catalog is optimistic about these"', "Batches"),
        ("Wholesale orders below target margin",
         '=COUNTIF([[Wholesale Orders.alert]],"Low margin*")',
         '="The discount is deeper than the product can carry"',
         "Wholesale Orders"),
        ("Unpaid wholesale invoices",
         '=ROUND(SUMIFS([[Wholesale Orders.order_value]],'
         '[[Wholesale Orders.paid]],"No"),0)',
         '="Money shops owe you right now"', "Wholesale Customers"),
        ("Avoidable waste this year",
         '=ROUND(SUMIFS([[Waste Log.cost]],[[Waste Log.avoidable]],'
         '"Yes"),0)',
         '="Material you paid for and did not sell"', "Waste Log"),
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

    # ------------------------------------------------------ best and worst
    row = s.section(row, "What is working", LAST_COL)
    row = s.table_head(row, 0, ["Most profit a sale", "Profit",
                                "Best margin", "Margin",
                                "Best paid an hour", "An hour",
                                "Most sold", "Sold"],
                       [24, 10, 24, 10, 24, 10, 24, 9])
    top_first = row
    for n in range(1, TOP_ROWS + 1):
        pairs = [
            (0, "profit_rank", "name", "text"),
            (1, "profit_rank", "profit", "money"),
            (2, "margin_rank", "name", "text"),
            (3, "margin_rank", "margin", "percent"),
            (4, "profit_hour_rank", "name", "text"),
            (5, "profit_hour_rank", "profit_hour", "money"),
            (6, "revenue_rank", "name", "text"),
            (7, "revenue_rank", "sold", "int"),
        ]
        for col, rank_key, field, style in pairs:
            ws.write_formula(row, col,
                             s.reg.render("=" + _by_rank(rank_key, field, n)),
                             fmts.calc(style))
        row += 1
    ws.conditional_format(top_first, 1, row - 1, 1,
                          {"type": "data_bar", "bar_color": "#8C9E6E"})
    ws.conditional_format(top_first, 3, row - 1, 3,
                          {"type": "data_bar", "bar_color": "#8C9E6E"})
    ws.conditional_format(top_first, 5, row - 1, 5,
                          {"type": "cell", "criteria": "<",
                           "value": "=MinimumWage",
                           "format": fmts.cf("bad")})
    row += 1
    s.merge(row, 0, LAST_COL,
            "The four columns rarely agree. The piece that makes the "
            "most profit is often not the one that pays best per hour "
            "- and the hourly column is the one that decides what you "
            "should be making more of.", fmts.chrome("note"))
    row += 2

    # ------------------------------------------------- by category --------
    row = s.section(row, "By category", LAST_COL)
    row = s.table_head(row, 0, ["Category", "Products", "Average cost",
                                "Average price", "Average margin",
                                "Sold", "Revenue", "Profit"],
                       [22, 10, 13, 13, 14, 10, 13, 13])
    cat_first = row
    for n in range(1, 11):
        cat = f'IFERROR(INDEX(CategoryList,{n}),"")'
        ws.write_formula(row, 0, f"={cat}", fmts.calc("text"))
        ws.write_formula(
            row, 1,
            s.reg.render(f'=IF({cat}="","",COUNTIFS([[Products.category]],'
                         f'{cat}))'), fmts.calc("int"))
        for col, field, style in ((2, "total_cost", "money"),
                                  (3, "price", "money"),
                                  (4, "margin", "percent")):
            ws.write_formula(
                row, col,
                s.reg.render(f'=IF({cat}="","",IFERROR(AVERAGEIFS('
                             f'[[Products.{field}]],[[Products.category]],'
                             f'{cat}),""))'), fmts.calc(style))
        ws.write_formula(
            row, 5,
            s.reg.render(f'=IF({cat}="","",SUMIFS([[Products.sold]],'
                         f'[[Products.category]],{cat}))'),
            fmts.calc("int"))
        ws.write_formula(
            row, 6,
            s.reg.render(f'=IF({cat}="","",ROUND(SUMIFS('
                         f'[[Products.revenue]],[[Products.category]],'
                         f'{cat}),2))'), fmts.calc("money0"))
        ws.write_formula(
            row, 7,
            s.reg.render(f'=IF({cat}="","",ROUND(SUMIFS('
                         f'[[Products.actual_profit]],'
                         f'[[Products.category]],{cat}),2))'),
            fmts.calc("money0"))
        row += 1
    ws.conditional_format(cat_first, 4, row - 1, 4,
                          {"type": "3_color_scale"})
    ws.conditional_format(cat_first, 7, row - 1, 7,
                          {"type": "data_bar", "bar_color": "#B5704F"})
    row += 2

    # ----------------------------------------------------- twelve months --
    row = s.section(row, "The last twelve months", LAST_COL)
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
    lines = [
        ("Revenue", "SUMIFS([[Sales.revenue]],[[Sales.month]],{key})",
         "money0"),
        ("Cost of goods", "SUMIFS([[Sales.cogs]],[[Sales.month]],{key})",
         "money0"),
        ("Fees", "SUMIFS([[Sales.fees]],[[Sales.month]],{key})", "money0"),
        ("Profit", "SUMIFS([[Sales.profit]],[[Sales.month]],{key})",
         "money0"),
        ("Items", "SUMIFS([[Sales.qty]],[[Sales.month]],{key})", "int"),
        ("Hours sold", "SUMIFS([[Sales.hours]],[[Sales.month]],{key})",
         "number1"),
    ]
    band = {}
    for label, body, style in lines:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        for i in range(MONTHS):
            cell = f"{chr(66 + i)}${key_row + 1}"
            ws.write_formula(row, i + 1,
                             s.reg.render("=ROUND(" +
                                          body.format(key=cell) + ",2)"),
                             fmts.calc(style))
        band[label] = row
        row += 1
    s.write(row, 0, "Margin", fmts.chrome("label_bold"))
    for i in range(MONTHS):
        col = chr(66 + i)
        ws.write_formula(
            row, i + 1,
            f'=IFERROR({col}{band["Profit"] + 1}/{col}'
            f'{band["Revenue"] + 1},"")', fmts.calc("percent0"))
    margin_row = row
    row += 1
    for label in ("Revenue", "Profit"):
        ws.add_sparkline(band[label], MONTHS + 1, {
            "range": (f"{quote_sheet(NAME)}!B{band[label] + 1}:"
                      f"M{band[label] + 1}"),
            "type": "column",
            "series_color": "#8C4A32" if label == "Revenue" else "#5E7B5A",
        })
    ws.conditional_format(band["Revenue"], 1, band["Revenue"], MONTHS,
                          {"type": "data_bar", "bar_color": "#B5704F"})
    ws.conditional_format(margin_row, 1, margin_row, MONTHS,
                          {"type": "cell", "criteria": "<",
                           "value": "=MarginWarning",
                           "format": fmts.cf("warn")})
    s.set_col(MONTHS + 1, MONTHS + 1, 12)
    row += 1
    s.merge(row, 0, LAST_COL,
            "Blank months are months with nothing logged on the Sales "
            "tab, not months with no sales.", fmts.chrome("note"))
    row += 2

    # --------------------------------------------------------- channels --
    row = s.section(row, "Where the money comes from", LAST_COL)
    row = s.table_head(row, 0, ["Channel", "Orders", "Revenue", "Fees",
                                "Fee rate", "Profit", "Margin",
                                "Share of revenue"],
                       [22, 10, 13, 12, 11, 13, 10, 16])
    ch_first = row
    for n in range(1, 11):
        name = f'IFERROR(INDEX([[Channels.name]],{n})&"","")'
        ws.write_formula(row, 0, s.reg.render(f"={name}"),
                         fmts.calc("text"))
        ws.write_formula(
            row, 1,
            s.reg.render(f'=IF({name}="","",COUNTIFS([[Sales.channel]],'
                         f'{name}))'), fmts.calc("int"))
        ws.write_formula(
            row, 2,
            s.reg.render(f'=IF({name}="","",ROUND(SUMIFS('
                         f'[[Sales.revenue]],[[Sales.channel]],{name}),2))'),
            fmts.calc("money0"))
        ws.write_formula(
            row, 3,
            s.reg.render(f'=IF({name}="","",ROUND(SUMIFS([[Sales.fees]],'
                         f'[[Sales.channel]],{name}),2))'),
            fmts.calc("money0"))
        ws.write_formula(
            row, 4,
            f'=IFERROR(D{row + 1}/C{row + 1},"")', fmts.calc("percent"))
        ws.write_formula(
            row, 5,
            s.reg.render(f'=IF({name}="","",ROUND(SUMIFS('
                         f'[[Sales.profit]],[[Sales.channel]],{name}),2))'),
            fmts.calc("money0"))
        ws.write_formula(
            row, 6, f'=IFERROR(F{row + 1}/C{row + 1},"")',
            fmts.calc("percent"))
        ws.write_formula(
            row, 7,
            s.reg.render(f'=IFERROR(C{row + 1}/SUM([[Sales.revenue]]),"")'),
            fmts.calc("percent"))
        row += 1
    ws.conditional_format(ch_first, 7, row - 1, 7,
                          {"type": "data_bar", "bar_color": "#8C4A32"})
    ws.conditional_format(ch_first, 4, row - 1, 4,
                          {"type": "cell", "criteria": ">", "value": 0.25,
                           "format": fmts.cf("warn")})
    row += 2

    # ------------------------------------------------------- what to do --
    row = s.section(row, "If you only do one thing this week", LAST_COL)
    s.merge(row, 0, LAST_COL, None, fmts.calc("text"))
    ws.write_formula(
        row, 0,
        s.reg.render(
            '=IF(COUNTIF([[Products.alert]],"Loss*")>0,'
            '"Fix the products that lose money. There are "&'
            'COUNTIF([[Products.alert]],"Loss*")&" of them on the '
            'Product Catalog - every sale of those makes things worse.",'
            'IF(COUNTIF([[Products.alert]],"Below*")>0,'
            '"Re-price the "&COUNTIF([[Products.alert]],"Below*")&'
            '" products that pay you under minimum wage. Start with '
            'the Retail Price Calculator.",'
            'IF(COUNTIFS([[Products.name]],"<>",[[Products.price]],"")>0,'
            '"Price the products that have no price yet.",'
            'IF(FeesStale="Yes","Check your marketplace fee rates on '
            'the Channels tab - they were last reviewed "&'
            'TEXT(FeesReviewed,"d mmm yyyy")&".",'
            'IF(COUNTIF([[Products.alert]],"Review*")>0,'
            '"Review the prices that have not moved in "&'
            'PriceReviewDays&" days. Your costs have.",'
            '"Everything is priced, nothing is losing money and your '
            'fees are current. Log this month\'s sales and look at '
            'what pays you best an hour.")))))'),
        fmts.calc("text"))
    s.row_height(row, 34)
    row += 2

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
