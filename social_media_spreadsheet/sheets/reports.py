"""`Reports` - pick two dates, print the result or send it to a client."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, date_check

NAME = "Reports"
LAST_COL = 13
FIRST_CALC_ROW = 5

_CAL = ('[[Content Calendar.date]],">="&RepFrom,'
        '[[Content Calendar.date]],"<="&RepTo')
_AN = '[[Analytics.month]],">="&RepFrom,[[Analytics.month]],"<="&RepTo'
_PF = '[[Performance.date]],">="&RepFrom,[[Performance.date]],"<="&RepTo'


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws

    s.banner("Reports & Exports",
             "Choose a date range, then print this page or copy it into "
             "an email. Everything updates from the calendar, the "
             "Analytics tab and the Performance tab.",
             LAST_COL)
    s.set_col(0, 0, 26)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 14)

    # -------------------------------------------------------- date range
    row = 4
    row = s.section(row, "Report period", LAST_COL)
    s.merge(row, 0, 0, "From", fmts.chrome("label_bold"))
    s.write(row, 1, cfg.today.replace(day=1), fmts.input("date"))
    ws.data_validation(row, 1, row, 1, date_check())
    book.define_name("RepFrom", f"={quote_sheet(NAME)}!$B${row + 1}")

    s.merge(row, 2, 2, "To", fmts.chrome("label_bold"))
    s.write(row, 3, cfg.today, fmts.input("date"))
    ws.data_validation(row, 3, row, 3, date_check())
    book.define_name("RepTo", f"={quote_sheet(NAME)}!$D${row + 1}")

    s.merge(row, 4, 5, "Days in the range", fmts.chrome("label_bold"))
    s.write(row, 6, "=RepTo-RepFrom+1", fmts.calc("int"))
    s.merge(row, 7, 8, "Prepared on", fmts.chrome("label_bold"))
    s.write(row, 9, "=TODAY()", fmts.calc("date"))
    s.merge(row, 10, LAST_COL, "=BrandName", fmts.calc("text"))
    row += 2

    # ------------------------------------------------------------ summary
    row = s.section(row, "Headline numbers", LAST_COL)
    pairs = [
        ("Posts published",
         f'=COUNTIFS({_CAL},[[Content Calendar.status]],"Published")',
         "int"),
        ("Posts planned", f"=COUNTIFS({_CAL})", "int"),
        ("Reach", f"=SUMIFS([[Analytics.reach]],{_AN})", "int"),
        ("Impressions", f"=SUMIFS([[Analytics.impressions]],{_AN})",
         "int"),
        ("Engagements", f"=SUMIFS([[Analytics.engagements]],{_AN})",
         "int"),
        ("Engagement rate",
         f"=IFERROR(SUMIFS([[Analytics.engagements]],{_AN})"
         f"/SUMIFS([[Analytics.reach]],{_AN}),0)", "percent"),
        ("Followers gained", f"=SUMIFS([[Analytics.gained]],{_AN})", "int"),
        ("Unfollows", f"=SUMIFS([[Analytics.lost]],{_AN})", "int"),
        ("Net follower growth",
         f"=SUMIFS([[Analytics.net_growth]],{_AN})", "int"),
        ("Link clicks", f"=SUMIFS([[Analytics.clicks]],{_AN})", "int"),
        ("Leads", f"=SUMIFS([[Analytics.leads]],{_AN})", "int"),
        ("Sales", f"=SUMIFS([[Analytics.conversions]],{_AN})", "int"),
        ("Revenue", f"=SUMIFS([[Analytics.revenue]],{_AN})", "money"),
        ("Spend",
         '=SUMIFS([[Budget.actual]],[[Budget.month]],">="&RepFrom,'
         '[[Budget.month]],"<="&RepTo)', "money"),
        ("Return on spend",
         f'=IFERROR(SUMIFS([[Analytics.revenue]],{_AN})'
         f'/SUMIFS([[Budget.actual]],[[Budget.month]],">="&RepFrom,'
         f'[[Budget.month]],"<="&RepTo),0)', "number1"),
        ("Cost per lead",
         f'=IFERROR(SUMIFS([[Budget.actual]],[[Budget.month]],">="&RepFrom,'
         f'[[Budget.month]],"<="&RepTo)/SUMIFS([[Analytics.leads]],{_AN}),'
         f"0)", "money"),
    ]
    first_summary = row
    for i, (label, formula, style) in enumerate(pairs):
        r = first_summary + i % 8
        col = 0 if i < 8 else 5
        s.merge(r, col, col + 2, label, fmts.chrome("label_bold"))
        ws.merge_range(r, col + 3, r, col + 4, "", fmts.calc(style))
        ws.write_formula(r, col + 3, s.reg.render(formula),
                         fmts.calc(style))
    row = first_summary + 9

    # --------------------------------------------------------- platforms
    row = s.section(row, "By platform", LAST_COL)
    s.table_head(row, 0, ["Platform", "Posts", "Reach", "Impressions",
                          "Engagements", "Eng. rate", "Clicks", "Leads",
                          "Sales", "Revenue", "Followers gained"])
    first_platform = row + 1
    for i in range(8):
        r = first_platform + i
        cr = FIRST_CALC_ROW + i
        name = f"[[Calc.p_name#{cr}]]"
        guard = f'IF({name}="","",'
        s.calc_row(r, 0, [
            f"={guard}{name})",
            f'={guard}COUNTIFS([[Content Calendar.platform]],{name},'
            f'{_CAL},[[Content Calendar.status]],"Published"))',
            f"={guard}SUMIFS([[Analytics.reach]],"
            f"[[Analytics.platform]],{name},{_AN}))",
            f"={guard}SUMIFS([[Analytics.impressions]],"
            f"[[Analytics.platform]],{name},{_AN}))",
            f"={guard}SUMIFS([[Analytics.engagements]],"
            f"[[Analytics.platform]],{name},{_AN}))",
            f"={guard}IFERROR(SUMIFS([[Analytics.engagements]],"
            f"[[Analytics.platform]],{name},{_AN})"
            f"/SUMIFS([[Analytics.reach]],[[Analytics.platform]],{name},"
            f'{_AN}),""))',
            f"={guard}SUMIFS([[Analytics.clicks]],"
            f"[[Analytics.platform]],{name},{_AN}))",
            f"={guard}SUMIFS([[Analytics.leads]],"
            f"[[Analytics.platform]],{name},{_AN}))",
            f"={guard}SUMIFS([[Analytics.conversions]],"
            f"[[Analytics.platform]],{name},{_AN}))",
            f"={guard}SUMIFS([[Analytics.revenue]],"
            f"[[Analytics.platform]],{name},{_AN}))",
            f"={guard}SUMIFS([[Analytics.net_growth]],"
            f"[[Analytics.platform]],{name},{_AN}))",
        ], ["text", "int", "int", "int", "int", "percent", "int", "int",
            "int", "money", "int"])
    ws.conditional_format(first_platform, 2, first_platform + 7, 2,
                          {"type": "data_bar", "bar_color": "#6350A8"})
    row = first_platform + 9

    # ----------------------------------------------------------- pillars
    row = s.section(row, "By content pillar", LAST_COL)
    s.table_head(row, 0, ["Pillar", "Posts published", "Reach",
                          "Engagements", "Eng. rate", "Clicks", "Sales",
                          "Revenue"])
    first_pillar = row + 1
    pillar_rows = min(10, cfg.row_count("Content Pillars"))
    for i in range(pillar_rows):
        r = first_pillar + i
        cr = FIRST_CALC_ROW + i
        name = f"[[Calc.pl_name#{cr}]]"
        guard = f'IF({name}="","",'
        s.calc_row(r, 0, [
            f"={guard}{name})",
            f'={guard}COUNTIFS([[Content Calendar.pillar]],{name},'
            f'{_CAL},[[Content Calendar.status]],"Published"))',
            f"={guard}SUMIFS([[Performance.reach]],"
            f"[[Performance.pillar]],{name},{_PF}))",
            f"={guard}SUMIFS([[Performance.engagements]],"
            f"[[Performance.pillar]],{name},{_PF}))",
            f"={guard}IFERROR(SUMIFS([[Performance.engagements]],"
            f"[[Performance.pillar]],{name},{_PF})"
            f"/SUMIFS([[Performance.reach]],[[Performance.pillar]],"
            f'{name},{_PF}),""))',
            f"={guard}SUMIFS([[Performance.clicks]],"
            f"[[Performance.pillar]],{name},{_PF}))",
            f"={guard}SUMIFS([[Performance.conversions]],"
            f"[[Performance.pillar]],{name},{_PF}))",
            f"={guard}SUMIFS([[Performance.revenue]],"
            f"[[Performance.pillar]],{name},{_PF}))",
        ], ["text", "int", "int", "int", "percent", "int", "int", "money"])
    ws.conditional_format(first_pillar, 4, first_pillar + pillar_rows - 1,
                          4, {"type": "3_color_scale",
                              "min_color": "#FBD9DD",
                              "mid_color": "#FFF1CF",
                              "max_color": "#D9F2E4"})
    row = first_pillar + pillar_rows + 1

    # -------------------------------------------------------- top posts
    row = s.section(row, "Top 10 posts in this range", LAST_COL)
    s.table_head(row, 0, ["#", "Post", "Platform", "Pillar", "Date",
                          "Reach", "Engagements", "Eng. rate", "Clicks",
                          "Sales", "Revenue"])
    first_top = row + 1
    for i in range(10):
        r = first_top + i
        n = i + 1
        pick = (f"MATCH(LARGE([[Calc.pf_key_r]],{n}),"
                f"[[Calc.pf_key_r]],0)")
        s.calc_row(r, 0, [
            n,
            f'=IFERROR(INDEX([[Calc.pf_title]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_platform]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_pillar]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_date]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_reach]],{pick}),"")',
            f'=IFERROR(INDEX([[Performance.engagements]],{pick}),"")',
            f'=IFERROR(INDEX([[Calc.pf_er]],{pick}),"")',
            f'=IFERROR(INDEX([[Performance.clicks]],{pick}),"")',
            f'=IFERROR(INDEX([[Performance.conversions]],{pick}),"")',
            f'=IFERROR(INDEX([[Performance.revenue]],{pick}),"")',
        ], ["int", "text", "text", "text", "date", "int", "int",
            "percent", "int", "int", "money"])
    row = first_top + 11

    # ------------------------------------------------------- campaigns
    row = s.section(row, "Campaigns in this range", LAST_COL)
    s.table_head(row, 0, ["Campaign", "Start", "End", "Posts", "Reach",
                          "Engagements", "Clicks", "Leads", "Sales",
                          "Spend", "Revenue", "Return"])
    first_campaign = row + 1
    campaign_rows = min(10, cfg.row_count("Campaigns"))
    for i in range(campaign_rows):
        r = first_campaign + i
        cr = FIRST_CALC_ROW + i
        guard = f'IF([[Campaigns.name#{cr}]]="","",'
        s.calc_row(r, 0, [
            f"={guard}[[Campaigns.name#{cr}]])",
            f"={guard}[[Campaigns.start#{cr}]])",
            f"={guard}[[Campaigns.end#{cr}]])",
            f"={guard}[[Campaigns.posts#{cr}]])",
            f"={guard}[[Campaigns.reach#{cr}]])",
            f"={guard}[[Campaigns.engagements#{cr}]])",
            f"={guard}[[Campaigns.clicks#{cr}]])",
            f"={guard}[[Campaigns.leads#{cr}]])",
            f"={guard}[[Campaigns.conversions#{cr}]])",
            f"={guard}[[Campaigns.spend#{cr}]])",
            f"={guard}[[Campaigns.revenue#{cr}]])",
            f"={guard}[[Campaigns.roas#{cr}]])",
        ], ["text", "date", "date", "int", "int", "int", "int", "int",
            "int", "money", "money", "number1"])
    row = first_campaign + campaign_rows + 1

    # ----------------------------------------------------- what to say
    row = s.section(row, "What this says, in a sentence", LAST_COL)
    sentence = (
            '=IF(SUMIFS([[Analytics.reach]],' + _AN + ')=0,'
            '"Fill in the Analytics tab to see a summary here.",'
            '"In this period you published "&'
            'TEXT(COUNTIFS(' + _CAL + ',[[Content Calendar.status]],'
            '"Published"),"#,##0")&" posts, reached "&'
            'TEXT(SUMIFS([[Analytics.reach]],' + _AN + '),"#,##0")&'
            '" people, earned "&TEXT(SUMIFS('
            '[[Analytics.engagements]],' + _AN + '),"#,##0")&'
            '" engagements ("&TEXT(IFERROR(SUMIFS('
            '[[Analytics.engagements]],' + _AN + ')/SUMIFS('
            '[[Analytics.reach]],' + _AN + '),0),"0.0%")&'
            '") and gained "&TEXT(SUMIFS([[Analytics.net_growth]],'
            + _AN + '),"#,##0")&" followers.")')
    s.merge(row, 0, LAST_COL, s.reg.render(sentence), fmts.calc("wrap"))
    s.row_height(row, 32)
    row += 2

    row = s.section(row, "How to share this", LAST_COL)
    row = s.bullets(row, [
        "PRINT - File > Print. The page is already set to fit one page "
        "wide; choose landscape for the widest tables.",
        "PDF - File > Save As > PDF, or print to PDF. Perfect for a "
        "client or a monthly review.",
        "COPY - select the area you want, copy, then paste into an email "
        "as a picture so the formatting survives.",
        "EXPORT - File > Save As > CSV exports the active sheet only. "
        "Use it to move data into another tool.",
        "SHARE - keep the file in Drive, OneDrive or Dropbox so you "
        "always open the same copy on every device.",
    ], LAST_COL)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    ws.set_landscape()
    return s
