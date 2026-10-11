"""`Dashboard` - the one page view of the month."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, col_letter, quote_sheet
from ..writer import SheetWriter, list_from

NAME = "Dashboard"
LAST_COL = 17            # column R - six KPI tiles of three columns each
FIRST_CALC_ROW = 5       # Calc sheet data starts on row 5
MONTH_ROWS = 12

_WINDOW = ('[[Content Calendar.date]],">="&DashStart,'
           '[[Content Calendar.date]],"<="&DashEnd')
_AN = ('[[Analytics.month]],">="&DashStart,'
       '[[Analytics.month]],"<="&DashEnd')


def _calc(key: str, row: int) -> str:
    return f"[[Calc.{key}#{row}]]"


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.ink)
    ws = s.ws

    s.banner("Monthly Dashboard",
             "Pick a month, then read your whole social media presence at "
             "a glance. Every number is live - it updates as you fill in "
             "the calendar and the Analytics tab.",
             LAST_COL)

    for c in range(0, LAST_COL + 1):
        s.set_col(c, c, 12.5)

    # ------------------------------------------------------ period picker
    row = 4
    row = s.section(row, "Which month are you looking at?", LAST_COL)
    s.merge(row, 0, 1, "Month", fmts.chrome("label_bold"))
    s.merge(row, 2, 3, cfg.today.strftime("%B"), fmts.input("text"))
    ws.data_validation(row, 2, row, 2, list_from("MonthList"))
    book.define_name("DashMonthName", f"={quote_sheet(NAME)}!$C${row + 1}")

    s.merge(row, 4, 5, "Period start", fmts.chrome("label_bold"))
    s.merge(row, 6, 7, None, fmts.calc("date"))
    ws.write_formula(
        row, 6,
        '=IF(DashMonthName="All year",DATE(PlanYear,1,1),'
        'DATE(PlanYear,MAX(1,MATCH(DashMonthName,MonthList,0)-1),1))',
        fmts.calc("date"))
    book.define_name("DashStart", f"={quote_sheet(NAME)}!$G${row + 1}")

    s.merge(row, 8, 9, "Period end", fmts.chrome("label_bold"))
    s.merge(row, 10, 11, None, fmts.calc("date"))
    ws.write_formula(
        row, 10,
        '=IF(DashMonthName="All year",DATE(PlanYear,12,31),'
        'EOMONTH(DashStart,0))', fmts.calc("date"))
    book.define_name("DashEnd", f"={quote_sheet(NAME)}!$K${row + 1}")

    s.merge(row, 12, 13, "Month key", fmts.chrome("label_bold"))
    s.merge(row, 14, 15, None, fmts.calc("text"))
    ws.write_formula(row, 14, '=TEXT(DashStart,"yyyy-mm")',
                     fmts.calc("text"))
    book.define_name("DashKey", f"={quote_sheet(NAME)}!$O${row + 1}")

    s.merge(row, 16, 16, "Today", fmts.chrome("label_bold"))
    s.write(row, 17, "=TODAY()", fmts.calc("date"))
    row += 2

    # ------------------------------------------------------- output tiles
    row = s.section(row, "What you published", LAST_COL)
    s.kpi(row, 0, "Posts planned",
          f"=COUNTIFS({_WINDOW})", "int", 3,
          "Everything on the calendar in this period")
    s.kpi(row, 3, "Published",
          f'=COUNTIFS({_WINDOW},[[Content Calendar.status]],"Published")',
          "int", 3, "Actually went out")
    book.define_name("Dash_Published", f"={quote_sheet(NAME)}!$D${row + 2}")
    s.kpi(row, 6, "Still to come",
          f'=COUNTIFS({_WINDOW},[[Content Calendar.status]],"Scheduled")'
          f'+COUNTIFS({_WINDOW},[[Content Calendar.status]],"Approved")',
          "int", 3, "Approved or scheduled, not yet live")
    s.kpi(row, 9, "Target for the month",
          '=IF(DashMonthName="All year",MonthlyPostTarget*12,'
          'MonthlyPostTarget)', "int", 3,
          "From the Settings tab")
    s.kpi(row, 12, "Against target",
          '=IFERROR(Dash_Published/IF(DashMonthName="All year",'
          'MonthlyPostTarget*12,MonthlyPostTarget),0)', "percent", 3,
          "Published divided by your target")
    s.kpi(row, 15, "Needing attention",
          '=COUNTIFS([[Content Calendar.alert]],"<>OK",'
          '[[Content Calendar.title]],"<>")', "int", 3,
          "Rows with an action in the last column")
    row += 4

    # -------------------------------------------------------- reach tiles
    row = s.section(row, "How it performed", LAST_COL)
    s.kpi(row, 0, "Reach", f"=SUMIFS([[Analytics.reach]],{_AN})", "int", 3,
          "People reached across every account")
    book.define_name("Dash_Reach", f"={quote_sheet(NAME)}!$A${row + 2}")
    s.kpi(row, 3, "Impressions",
          f"=SUMIFS([[Analytics.impressions]],{_AN})", "int", 3,
          "Times your content was shown")
    s.kpi(row, 6, "Engagements",
          f"=SUMIFS([[Analytics.engagements]],{_AN})", "int", 3,
          "Likes, comments, shares and saves")
    book.define_name("Dash_Engagements",
                     f"={quote_sheet(NAME)}!$G${row + 2}")
    s.kpi(row, 9, "Engagement rate",
          "=IFERROR(Dash_Engagements/Dash_Reach,0)", "percent", 3,
          "Engagements divided by reach")
    s.kpi(row, 12, "Target rate", "=TargetEngagement", "percent", 3,
          "Your goal, set on the Settings tab")
    s.kpi(row, 15, "Reach per post",
          "=IFERROR(Dash_Reach/Dash_Published,0)", "int", 3,
          "How far the average post travelled")
    row += 4

    # ----------------------------------------------------- audience tiles
    row = s.section(row, "Audience & results", LAST_COL)
    s.kpi(row, 0, "Followers",
          f"=SUMIFS([[Analytics.followers_end]],{_AN})", "int", 3,
          "Total across every account at the end of the period")
    s.kpi(row, 3, "Follower growth",
          f"=SUMIFS([[Analytics.net_growth]],{_AN})", "int", 3,
          "New followers less unfollows")
    s.kpi(row, 6, "Link clicks", f"=SUMIFS([[Analytics.clicks]],{_AN})",
          "int", 3, "Traffic you sent out")
    s.kpi(row, 9, "Leads", f"=SUMIFS([[Analytics.leads]],{_AN})", "int", 3,
          "Sign-ups, enquiries, waitlist joins")
    s.kpi(row, 12, "Sales", f"=SUMIFS([[Analytics.conversions]],{_AN})",
          "int", 3, "Orders you can trace back to social")
    s.kpi(row, 15, "Revenue", f"=SUMIFS([[Analytics.revenue]],{_AN})",
          "money", 3, "Money attributed to social")
    book.define_name("Dash_Revenue", f"={quote_sheet(NAME)}!$P${row + 2}")
    row += 4

    # --------------------------------------------------------- money tiles
    row = s.section(row, "What it cost", LAST_COL)
    s.kpi(row, 0, "Spend",
          '=SUMIFS([[Budget.actual]],[[Budget.month]],">="&DashStart,'
          '[[Budget.month]],"<="&DashEnd)', "money", 3,
          "Everything on the Budget tab for this period")
    book.define_name("Dash_Spend", f"={quote_sheet(NAME)}!$A${row + 2}")
    s.kpi(row, 3, "Budget", '=IF(DashMonthName="All year",AnnualBudget,'
                            'MonthlyBudget)', "money", 3,
          "From the Settings tab")
    s.kpi(row, 6, "Left to spend",
          '=IF(DashMonthName="All year",AnnualBudget,MonthlyBudget)'
          '-Dash_Spend', "money", 3, "Budget less actual spend")
    s.kpi(row, 9, "Return on spend",
          "=IFERROR(Dash_Revenue/Dash_Spend,0)", "number1", 3,
          "Revenue divided by spend")
    s.kpi(row, 12, "Cost per lead",
          f"=IFERROR(Dash_Spend/SUMIFS([[Analytics.leads]],{_AN}),0)",
          "money", 3, "Spend divided by leads")
    s.kpi(row, 15, "Revenue per post",
          "=IFERROR(Dash_Revenue/Dash_Published,0)", "money", 3,
          "Money per published post")
    row += 4

    # ------------------------------------------------- twelve month table
    row = s.section(row, f"The whole of {cfg.year} month by month", LAST_COL)
    headers = ["Month", "Published", "Reach", "Engagements", "Eng. rate",
               "Followers", "Growth", "Clicks", "Leads", "Sales",
               "Revenue", "Spend", "Net"]
    s.table_head(row, 0, headers)
    first_table_row = row + 1
    keys = ["m_label", "m_published", "m_reach", "m_engagements", "m_er",
            "m_followers", "m_growth", "m_clicks", "m_leads",
            "m_conversions", "m_revenue", "m_spend", "m_profit"]
    styles = ["text", "int", "int", "int", "percent", "int", "int", "int",
              "int", "int", "money", "money", "money"]
    for i in range(MONTH_ROWS):
        r = first_table_row + i
        calc_row = FIRST_CALC_ROW + i
        s.calc_row(r, 0, [f"={_calc(key, calc_row)}" for key in keys],
                   styles)
    last_table_row = first_table_row + MONTH_ROWS - 1

    for col_idx, colour in ((2, "#6350A8"), (3, "#F2665E"), (5, "#2A4A8B"),
                            (10, "#0E6B45")):
        ws.conditional_format(first_table_row, col_idx, last_table_row,
                              col_idx, {"type": "data_bar",
                                        "bar_color": colour})
    ws.conditional_format(first_table_row, 4, last_table_row, 4, {
        "type": "cell", "criteria": ">=", "value": "=TargetEngagement",
        "format": fmts.cf("good")})

    # sparkline strip under the table
    spark_row = last_table_row + 1
    s.write(spark_row, 0, "Trend", fmts.chrome("legend_label"))
    for col_idx, colour in ((1, "#6350A8"), (2, "#3B2A6B"), (3, "#F2665E"),
                            (5, "#2A4A8B"), (10, "#0E6B45")):
        letter = col_letter(col_idx)
        try:
            ws.add_sparkline(spark_row, col_idx, {
                "range": (f"{quote_sheet(NAME)}!{letter}"
                          f"{first_table_row + 1}:{letter}"
                          f"{last_table_row + 1}"),
                "type": "column",
                "series_color": colour,
            })
        except Exception:        # pragma: no cover - library differences
            pass
    for col_idx in range(1, len(headers)):
        if col_idx not in (1, 2, 3, 5, 10):
            ws.write_blank(spark_row, col_idx, None,
                           fmts.chrome("legend_calc"))
    row = spark_row + 2

    # ------------------------------------------------ platform comparison
    row = s.section(row, "Platform by platform (selected period)", 8)
    s.table_head(row, 0, ["Platform", "Posts", "Reach", "Engagements",
                          "Eng. rate", "Followers", "Growth", "Clicks",
                          "Revenue"])
    first_platform = row + 1
    platform_rows = 8
    pkeys = ["p_name", "p_published_m", "p_reach_m", "p_eng_m", "p_er_m",
             "p_followers", "p_growth_m", "p_clicks_m", "p_revenue_m"]
    pstyles = ["text", "int", "int", "int", "percent", "int", "int", "int",
               "money"]
    for i in range(platform_rows):
        r = first_platform + i
        calc_row = FIRST_CALC_ROW + i
        s.calc_row(r, 0, [f'=IF({_calc("p_name", calc_row)}="","",'
                          f"{_calc(key, calc_row)})" for key in pkeys],
                   pstyles)
    ws.conditional_format(first_platform, 2, first_platform
                          + platform_rows - 1, 2,
                          {"type": "data_bar", "bar_color": "#6350A8"})
    ws.conditional_format(first_platform, 4, first_platform
                          + platform_rows - 1, 4,
                          {"type": "3_color_scale",
                           "min_color": "#FBD9DD", "mid_color": "#FFF1CF",
                           "max_color": "#D9F2E4"})

    # pillar table alongside
    s.merge(row - 1, 10, LAST_COL, "Content pillars (this year)",
            fmts.chrome("subsection"))
    s.table_head(row, 10, ["Pillar", "Posts", "Reach", "Engagements",
                           "Eng. rate", "Share"])
    pillar_rows = min(platform_rows, cfg.row_count("Content Pillars"))
    for i in range(pillar_rows):
        r = first_platform + i
        calc_row = FIRST_CALC_ROW + i
        s.calc_row(r, 10, [
            f'=IF({_calc("pl_name", calc_row)}="","",'
            f'{_calc("pl_name", calc_row)})',
            f'=IF({_calc("pl_name", calc_row)}="","",'
            f'{_calc("pl_posts", calc_row)})',
            f'=IF({_calc("pl_name", calc_row)}="","",'
            f'{_calc("pl_reach", calc_row)})',
            f'=IF({_calc("pl_name", calc_row)}="","",'
            f'{_calc("pl_eng", calc_row)})',
            f'=IF({_calc("pl_name", calc_row)}="","",'
            f'{_calc("pl_er", calc_row)})',
            f'=IFERROR({_calc("pl_posts", calc_row)}'
            f'/SUM([[Calc.pl_posts]]),"")',
        ], ["text", "int", "int", "int", "percent", "percent"])
    ws.conditional_format(first_platform, 12, first_platform
                          + pillar_rows - 1, 12,
                          {"type": "data_bar", "bar_color": "#F2665E"})
    row = first_platform + platform_rows + 1

    # ------------------------------------------------------- top & alerts
    row = s.section(row, "Your best posts & what needs doing", LAST_COL)
    s.merge(row, 0, 8, "Top 8 posts in this period (by engagement rate)",
            fmts.chrome("subsection"))
    s.merge(row, 10, LAST_COL, "What needs attention right now",
            fmts.chrome("subsection"))
    row += 1
    s.table_head(row, 0, ["#", "Post", "Platform", "Pillar", "Date",
                          "Reach", "Eng. rate"])
    s.table_head(row, 10, ["Action", "How many", "", "", "", "", "", ""])
    first_top = row + 1
    for i in range(8):
        r = first_top + i
        n = i + 1
        pick = (f'MATCH(LARGE([[Calc.pf_key_m]],{n}),'
                f'[[Calc.pf_key_m]],0)')
        s.calc_row(r, 0, [
            n,
            f'=IFERROR(INDEX([[Calc.pf_title]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_platform]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_pillar]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_date]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_reach]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_er]],{pick}),"")',
        ], ["int", "text", "text", "text", "date", "int", "percent"])
    ws.conditional_format(first_top, 5, first_top + 7, 5,
                          {"type": "data_bar", "bar_color": "#6350A8"})

    alerts = [
        ("Posts overdue", '=COUNTIF([[Content Calendar.alert]],'
                          '"Overdue*")'),
        ("Drafts running late", '=COUNTIF([[Content Calendar.alert]],'
                                '"Late*")'),
        ("Waiting for approval",
         '=COUNTIF([[Content Calendar.status]],"Awaiting review")'),
        ("Posts with no caption",
         '=COUNTIFS([[Content Calendar.title]],"<>",'
         '[[Content Calendar.caption]],"")'),
        ("Published with no link",
         '=COUNTIF([[Content Calendar.published]],"Add*")'),
        ("Published with no numbers yet",
         '=COUNTIFS([[Content Calendar.status]],"Published",'
         '[[Content Calendar.reach]],"")'),
        ("Ideas worth scheduling",
         '=COUNTIF([[Idea Bank.alert]],"Plan*")'),
        ("Overdue tasks", '=COUNTIF([[Tasks.alert]],"Overdue*")'),
        ("UGC waiting on permission",
         '=COUNTIF([[UGC Tracker.alert]],"Action*")'),
        ("Goals behind pace", '=COUNTIF([[Goals.pace]],"Behind*")'),
        ("Accounts gone quiet",
         '=COUNTIF([[Accounts.alert]],"Check - account*")'),
        ("Trends about to expire",
         '=COUNTIF([[Trends.alert]],"Expired*")'),
    ]
    for i, (label, formula) in enumerate(alerts):
        r = first_top + i
        s.calc_row(r, 10, [label, formula], ["text", "int"])
        ws.conditional_format(r, 11, r, 11, {
            "type": "cell", "criteria": ">", "value": 0,
            "format": fmts.cf("bad")})
    row = first_top + max(8, len(alerts)) + 1

    # ------------------------------------------------------- what is next
    row = s.section(row, "Coming up & running late", LAST_COL)
    s.merge(row, 0, 8, "Next 8 posts due", fmts.chrome("subsection"))
    s.merge(row, 10, LAST_COL, "Overdue - deal with these first",
            fmts.chrome("subsection"))
    row += 1
    s.table_head(row, 0, ["Date", "Platform", "Post", "Status", "Owner",
                          "Days"])
    s.table_head(row, 10, ["Date", "Platform", "Post", "Status", "Owner",
                           "Days late"])
    first_next = row + 1
    for i in range(8):
        r = first_next + i
        n = i + 1
        up = (f'MATCH(SMALL([[Content Calendar.queue_key]],{n}),'
              f'[[Content Calendar.queue_key]],0)')
        late = (f'MATCH(SMALL([[Content Calendar.late_key]],{n}),'
                f'[[Content Calendar.late_key]],0)')
        s.calc_row(r, 0, [
            f'=IFERROR(INDEX([[Content Calendar.date]],{up}),"")',
            f'=IFERROR(INDEX([[Content Calendar.platform]],{up}),"")',
            f'=IFERROR(INDEX([[Content Calendar.title]],{up}),"")',
            f'=IFERROR(INDEX([[Content Calendar.status]],{up}),"")',
            f'=IFERROR(INDEX([[Content Calendar.owner]],{up}),"")',
            f'=IFERROR(INDEX([[Content Calendar.days_to_go]],{up}),"")',
        ], ["date", "text", "text", "status", "text", "int"])
        s.calc_row(r, 10, [
            f'=IFERROR(INDEX([[Content Calendar.date]],{late}),"")',
            f'=IFERROR(INDEX([[Content Calendar.platform]],{late}),"")',
            f'=IFERROR(INDEX([[Content Calendar.title]],{late}),"")',
            f'=IFERROR(INDEX([[Content Calendar.status]],{late}),"")',
            f'=IFERROR(INDEX([[Content Calendar.owner]],{late}),"")',
            f'=IFERROR(-INDEX([[Content Calendar.days_to_go]],{late}),"")',
        ], ["date", "text", "text", "status", "text", "int"])
    row = first_next + 9

    s.note(row,
           "Tip: pick 'All year' in the month box to see the whole plan. "
           "The Annual Dashboard compares every month side by side, and "
           "the Reports tab turns the same numbers into something you can "
           "print or send to a client.",
           LAST_COL, height=26)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
