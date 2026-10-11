"""`Annual Dashboard` - the whole year on one page."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, col_letter, quote_sheet
from ..writer import SheetWriter

NAME = "Annual Dashboard"
LAST_COL = 17
FIRST_CALC_ROW = 5
MONTH_ROWS = 12

_YEAR = ('[[Content Calendar.date]],">="&DATE(PlanYear,1,1),'
         '[[Content Calendar.date]],"<="&DATE(PlanYear,12,31)')


def _calc(key: str, row: int) -> str:
    return f"[[Calc.{key}#{row}]]"


def _quarter(metric: str, first_month: int) -> str:
    """Sum three month rows of the Calc month series."""
    rows = [FIRST_CALC_ROW + first_month + i for i in range(3)]
    return "=" + "+".join(f"N({_calc(metric, r)})" for r in rows)


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.ink)
    ws = s.ws

    s.banner("Annual Dashboard",
             f"Everything you planned and published in {cfg.year}, "
             f"quarter by quarter and month by month.",
             LAST_COL)
    for c in range(0, LAST_COL + 1):
        s.set_col(c, c, 12.5)

    # ------------------------------------------------------- year headline
    row = 4
    row = s.section(row, f"{cfg.year} in numbers", LAST_COL)
    s.kpi(row, 0, "Posts planned", f"=COUNTIFS({_YEAR})", "int", 3,
          "Every row on the calendar this year")
    s.kpi(row, 3, "Published",
          f'=COUNTIFS({_YEAR},[[Content Calendar.status]],"Published")',
          "int", 3, "Actually went out")
    book.define_name("Year_Published", f"={quote_sheet(NAME)}!$D${row + 2}")
    s.kpi(row, 6, "Reach", "=SUM([[Calc.m_reach]])", "int", 3,
          "Across every account and month")
    book.define_name("Year_Reach", f"={quote_sheet(NAME)}!$G${row + 2}")
    s.kpi(row, 9, "Engagements", "=SUM([[Calc.m_engagements]])", "int", 3,
          "Likes, comments, shares, saves")
    book.define_name("Year_Eng", f"={quote_sheet(NAME)}!$J${row + 2}")
    s.kpi(row, 12, "Engagement rate", "=IFERROR(Year_Eng/Year_Reach,0)",
          "percent", 3, "The year's average")
    s.kpi(row, 15, "Follower growth", "=SUM([[Calc.m_growth]])", "int", 3,
          "New followers less unfollows")
    row += 4

    s.kpi(row, 0, "Clicks", "=SUM([[Calc.m_clicks]])", "int", 3,
          "Traffic sent out")
    s.kpi(row, 3, "Leads", "=SUM([[Calc.m_leads]])", "int", 3,
          "Enquiries and sign-ups")
    s.kpi(row, 6, "Sales", "=SUM([[Calc.m_conversions]])", "int", 3,
          "Orders traced back to social")
    s.kpi(row, 9, "Revenue", "=SUM([[Calc.m_revenue]])", "money", 3,
          "Money attributed to social")
    book.define_name("Year_Revenue", f"={quote_sheet(NAME)}!$J${row + 2}")
    s.kpi(row, 12, "Spend", "=SUM([[Calc.m_spend]])", "money", 3,
          "Everything on the Budget tab")
    book.define_name("Year_Spend", f"={quote_sheet(NAME)}!$M${row + 2}")
    s.kpi(row, 15, "Return on spend", "=IFERROR(Year_Revenue/Year_Spend,0)",
          "number1", 3, "Revenue divided by spend")
    row += 4

    s.kpi(row, 0, "Posts a week",
          "=IFERROR(Year_Published/52,0)", "number1", 3,
          "Your real publishing rhythm")
    s.kpi(row, 3, "Reach per post",
          "=IFERROR(Year_Reach/Year_Published,0)", "int", 3,
          "How far the average post travelled")
    s.kpi(row, 6, "Revenue per post",
          "=IFERROR(Year_Revenue/Year_Published,0)", "money", 3,
          "Money per published post")
    s.kpi(row, 9, "Best month",
          '=IFERROR(INDEX([[Calc.m_label]],MATCH(MAX([[Calc.m_reach]]),'
          '[[Calc.m_reach]],0)),"")', "text", 3, "By reach")
    s.kpi(row, 12, "Best platform",
          '=IFERROR(INDEX([[Calc.p_name]],MATCH(MAX([[Calc.p_key]]),'
          '[[Calc.p_key]],0)),"")', "text", 3, "By reach this year")
    s.kpi(row, 15, "Best day to post",
          '=IFERROR(INDEX([[Calc.dw_day]],MATCH(MAX([[Calc.dw_key]]),'
          '[[Calc.dw_key]],0)),"")', "text", 3,
          "Highest average engagement rate")
    row += 4

    # ------------------------------------------------------------ quarters
    row = s.section(row, "Quarter by quarter", LAST_COL)
    s.table_head(row, 0, ["Quarter", "Published", "Reach", "Engagements",
                          "Eng. rate", "Followers gained", "Clicks",
                          "Leads", "Sales", "Revenue", "Spend", "Net"])
    first_q = row + 1
    for i, label in enumerate(("Q1", "Q2", "Q3", "Q4")):
        r = first_q + i
        start = i * 3
        pub = _quarter("m_published", start)
        reach = _quarter("m_reach", start)
        eng = _quarter("m_engagements", start)
        rev = _quarter("m_revenue", start)
        spend = _quarter("m_spend", start)
        reach_cell = f"{col_letter(2)}{r + 1}"
        eng_cell = f"{col_letter(3)}{r + 1}"
        rev_cell = f"{col_letter(9)}{r + 1}"
        spend_cell = f"{col_letter(10)}{r + 1}"
        s.calc_row(r, 0, [
            label, pub, reach, eng,
            f"=IFERROR({eng_cell}/{reach_cell},\"\")",
            _quarter("m_growth", start),
            _quarter("m_clicks", start),
            _quarter("m_leads", start),
            _quarter("m_conversions", start),
            rev, spend,
            f"={rev_cell}-{spend_cell}",
        ], ["text", "int", "int", "int", "percent", "int", "int", "int",
            "int", "money", "money", "money"])
    ws.conditional_format(first_q, 2, first_q + 3, 2,
                          {"type": "data_bar", "bar_color": "#6350A8"})
    ws.conditional_format(first_q, 9, first_q + 3, 9,
                          {"type": "data_bar", "bar_color": "#0E6B45"})
    row = first_q + 5

    # ------------------------------------------------------ monthly detail
    row = s.section(row, "Every month in detail", LAST_COL)
    headers = ["Month", "Planned", "Published", "On target?", "Reach",
               "Impressions", "Engagements", "Eng. rate", "Followers",
               "Growth", "Clicks", "Leads", "Sales", "Revenue", "Spend",
               "Net", "Return"]
    s.table_head(row, 0, headers)
    first_month_row = row + 1
    for i in range(MONTH_ROWS):
        r = first_month_row + i
        cr = FIRST_CALC_ROW + i
        rev = f"{col_letter(13)}{r + 1}"
        spend = f"{col_letter(14)}{r + 1}"
        s.calc_row(r, 0, [
            f"={_calc('m_label', cr)}",
            f"={_calc('m_planned', cr)}",
            f"={_calc('m_published', cr)}",
            f"=IFERROR({_calc('m_published', cr)}/"
            f"{_calc('m_target', cr)},\"\")",
            f"={_calc('m_reach', cr)}",
            f"={_calc('m_impressions', cr)}",
            f"={_calc('m_engagements', cr)}",
            f"={_calc('m_er', cr)}",
            f"={_calc('m_followers', cr)}",
            f"={_calc('m_growth', cr)}",
            f"={_calc('m_clicks', cr)}",
            f"={_calc('m_leads', cr)}",
            f"={_calc('m_conversions', cr)}",
            f"={_calc('m_revenue', cr)}",
            f"={_calc('m_spend', cr)}",
            f"={rev}-{spend}",
            f"=IFERROR({rev}/{spend},\"\")",
        ], ["text", "int", "int", "percent", "int", "int", "int",
            "percent", "int", "int", "int", "int", "int", "money",
            "money", "money", "number1"])
    last_month_row = first_month_row + MONTH_ROWS - 1
    for col_idx, colour in ((4, "#6350A8"), (6, "#F2665E"), (8, "#2A4A8B"),
                            (13, "#0E6B45")):
        ws.conditional_format(first_month_row, col_idx, last_month_row,
                              col_idx, {"type": "data_bar",
                                        "bar_color": colour})
    ws.conditional_format(first_month_row, 3, last_month_row, 3, {
        "type": "cell", "criteria": ">=", "value": 1,
        "format": fmts.cf("good")})
    ws.conditional_format(first_month_row, 7, last_month_row, 7, {
        "type": "cell", "criteria": ">=", "value": "=TargetEngagement",
        "format": fmts.cf("good")})

    spark_row = last_month_row + 1
    s.write(spark_row, 0, "Trend", fmts.chrome("legend_label"))
    for col_idx, colour in ((2, "#6350A8"), (4, "#3B2A6B"),
                            (8, "#2A4A8B"), (13, "#0E6B45")):
        letter = col_letter(col_idx)
        try:
            ws.add_sparkline(spark_row, col_idx, {
                "range": (f"{quote_sheet(NAME)}!{letter}"
                          f"{first_month_row + 1}:{letter}"
                          f"{last_month_row + 1}"),
                "type": "column",
                "series_color": colour,
            })
        except Exception:        # pragma: no cover
            pass
    for col_idx in range(1, len(headers)):
        if col_idx not in (2, 4, 8, 13):
            ws.write_blank(spark_row, col_idx, None,
                           fmts.chrome("legend_calc"))
    row = spark_row + 2

    # ------------------------------------------------ platforms and goals
    row = s.section(row, "Platforms, formats and goals", LAST_COL)
    s.merge(row, 0, 6, "Platform totals for the year",
            fmts.chrome("subsection"))
    s.merge(row, 8, LAST_COL, "Goals", fmts.chrome("subsection"))
    row += 1
    s.table_head(row, 0, ["Platform", "Posts", "Reach", "Engagements",
                          "Eng. rate", "Growth", "Share of reach"])
    s.table_head(row, 8, ["Goal", "Target", "Now", "Progress", "On track?",
                          "By when"])
    first_p = row + 1
    platform_rows = 8
    for i in range(platform_rows):
        r = first_p + i
        cr = FIRST_CALC_ROW + i
        guard = f'IF({_calc("p_name", cr)}="","",'
        s.calc_row(r, 0, [
            f"={guard}{_calc('p_name', cr)})",
            f"={guard}{_calc('p_posts_y', cr)})",
            f"={guard}{_calc('p_reach_y', cr)})",
            f"={guard}{_calc('p_eng_y', cr)})",
            f"={guard}{_calc('p_er_y', cr)})",
            f"={guard}{_calc('p_growth_y', cr)})",
            f"=IFERROR({_calc('p_reach_y', cr)}/SUM([[Calc.p_reach_y]]),"
            f'"")',
        ], ["text", "int", "int", "int", "percent", "int", "percent"])
    ws.conditional_format(first_p, 2, first_p + platform_rows - 1, 2,
                          {"type": "data_bar", "bar_color": "#6350A8"})

    goal_rows = min(8, cfg.row_count("Goals"))
    for i in range(goal_rows):
        r = first_p + i
        gr = FIRST_CALC_ROW + i
        guard = f'IF([[Goals.name#{gr}]]="","",'
        s.calc_row(r, 8, [
            f"={guard}[[Goals.name#{gr}]])",
            f"={guard}[[Goals.target#{gr}]])",
            f"={guard}[[Goals.current#{gr}]])",
            f"={guard}[[Goals.progress#{gr}]])",
            f"={guard}[[Goals.pace#{gr}]])",
            f"={guard}[[Goals.target_date#{gr}]])",
        ], ["text", "number", "number", "percent", "status", "date"])
    ws.conditional_format(first_p, 11, first_p + goal_rows - 1, 11,
                          {"type": "data_bar", "bar_color": "#0E6B45"})
    row = first_p + max(platform_rows, goal_rows) + 1

    # ------------------------------------------------------- best of year
    row = s.section(row, f"Best of {cfg.year}", LAST_COL)
    s.merge(row, 0, 7, "Top 10 posts of the year (by engagement rate)",
            fmts.chrome("subsection"))
    s.merge(row, 9, LAST_COL, "Where your reach came from",
            fmts.chrome("subsection"))
    row += 1
    s.table_head(row, 0, ["#", "Post", "Platform", "Pillar", "Date",
                          "Reach", "Eng. rate"])
    s.table_head(row, 9, ["Format", "Posts", "Reach", "Engagements",
                          "Eng. rate"])
    first_top = row + 1
    for i in range(10):
        r = first_top + i
        n = i + 1
        pick = (f"MATCH(LARGE([[Calc.pf_key_y]],{n}),[[Calc.pf_key_y]],0)")
        s.calc_row(r, 0, [
            n,
            f'=IFERROR(INDEX([[Calc.pf_title]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_platform]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_pillar]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_date]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_reach]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_er]],{pick}),"")',
        ], ["int", "text", "text", "text", "date", "int", "percent"])
    ws.conditional_format(first_top, 5, first_top + 9, 5,
                          {"type": "data_bar", "bar_color": "#6350A8"})

    for i in range(10):
        r = first_top + i
        cr = FIRST_CALC_ROW + i
        guard = f'IF({_calc("fm_name", cr)}="","",'
        s.calc_row(r, 9, [
            f"={guard}{_calc('fm_name', cr)})",
            f"={guard}{_calc('fm_posts', cr)})",
            f"={guard}{_calc('fm_reach', cr)})",
            f"={guard}{_calc('fm_eng', cr)})",
            f"={guard}{_calc('fm_er', cr)})",
        ], ["text", "int", "int", "int", "percent"])
    ws.conditional_format(first_top, 11, first_top + 9, 11,
                          {"type": "data_bar", "bar_color": "#F2665E"})
    row = first_top + 11

    # --------------------------------------------------- timing and review
    row = s.section(row, "When to post & what to do next", LAST_COL)
    s.table_head(row, 0, ["Day", "Posts", "Reach", "Engagements",
                          "Eng. rate"])
    s.table_head(row, 6, ["Time slot", "Posts", "Reach", "Engagements",
                          "Eng. rate"])
    first_day = row + 1
    for i in range(7):
        r = first_day + i
        cr = FIRST_CALC_ROW + i
        guard = f'IF({_calc("dw_day", cr)}="","",'
        s.calc_row(r, 0, [
            f"={guard}{_calc('dw_day', cr)})",
            f"={guard}{_calc('dw_posts', cr)})",
            f"={guard}{_calc('dw_reach', cr)})",
            f"={guard}{_calc('dw_eng', cr)})",
            f"={guard}{_calc('dw_er', cr)})",
        ], ["text", "int", "int", "int", "percent"])
        guard2 = f'IF({_calc("sl_slot", cr)}="","",'
        s.calc_row(r, 6, [
            f"={guard2}{_calc('sl_slot', cr)})",
            f"={guard2}{_calc('sl_posts', cr)})",
            f"={guard2}{_calc('sl_reach', cr)})",
            f"={guard2}{_calc('sl_eng', cr)})",
            f"={guard2}{_calc('sl_er', cr)})",
        ], ["text", "int", "int", "int", "percent"])
    ws.conditional_format(first_day, 4, first_day + 6, 4,
                          {"type": "3_color_scale",
                           "min_color": "#FBD9DD", "mid_color": "#FFF1CF",
                           "max_color": "#D9F2E4"})
    ws.conditional_format(first_day, 10, first_day + 6, 10,
                          {"type": "3_color_scale",
                           "min_color": "#FBD9DD", "mid_color": "#FFF1CF",
                           "max_color": "#D9F2E4"})

    s.merge(first_day - 1, 12, LAST_COL, "Campaigns by revenue",
            fmts.chrome("subsection"))
    s.table_head(first_day, 12, ["Campaign", "Revenue", "Spend", "Return"])
    for i in range(6):
        r = first_day + 1 + i
        n = i + 1
        pick = f"MATCH(LARGE([[Calc.cm_key]],{n}),[[Calc.cm_key]],0)"
        s.calc_row(r, 12, [
            f'=IFERROR(INDEX([[Calc.cm_name]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.cm_revenue]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.cm_spend]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.cm_revenue]],{pick})'
            f'/INDEX([[Calc.cm_spend]],{pick}),"")',
        ], ["text", "money", "money", "number1"])
    row = first_day + 8

    s.note(row,
           "Use this page at the end of the year: look at the best month, "
           "the best day, the best format and your top ten posts, then "
           "build next year's pillars around what actually worked.",
           LAST_COL, height=26)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
