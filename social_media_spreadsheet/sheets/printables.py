"""The print-ready tabs: a single content brief and a weekly plan.

Both are driven by a drop down at the top of the page, so you pick the
post (or the week) and press print.
"""

from __future__ import annotations

from typing import Optional, Tuple

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, date_check

BRIEF_SHEET = "Print Brief"
WEEK_SHEET = "Print Week"

LAST_COL = 7          # column H - a comfortable portrait page


def _pick(sheet: SheetWriter, fmts, label: str, list_name: str,
          define_as: str, hint: str, default: str = "") -> None:
    """Drop down selector on row 1 (it sits outside the print area)."""
    ws = sheet.ws
    sheet.merge(0, 0, 1, label, fmts.chrome("label_bold"))
    fmt = fmts.input("text", bold=True, font_size=11)
    ws.merge_range(0, 2, 0, 3, default, fmt)
    ws.data_validation(0, 2, 0, 3, {
        "validate": "list", "source": f"={list_name}",
        "input_title": "Choose a record",
        "input_message": hint[:255]})
    sheet.merge(0, 4, LAST_COL, hint, fmts.chrome("note"))
    sheet.row_height(0, 22)
    sheet.book.define_name(define_as,
                           f"={quote_sheet(sheet.name)}!$C$1")


def _line(sheet, fmts, row: int, label: str,
          formula: Optional[str] = None, style: str = "text",
          label_cols: Tuple[int, int] = (0, 1),
          value_cols: Tuple[int, int] = (2, LAST_COL),
          height: Optional[float] = None) -> int:
    sheet.merge(row, label_cols[0], label_cols[1], label,
                fmts.chrome("label_bold"))
    fmt = fmts.calc(style)
    if value_cols[1] > value_cols[0]:
        sheet.ws.merge_range(row, value_cols[0], row, value_cols[1], "",
                             fmt)
    if formula:
        sheet.ws.write_formula(row, value_cols[0],
                               sheet.reg.render(formula), fmt)
    if height:
        sheet.row_height(row, height)
    return row + 1


def _blank_box(sheet, fmts, row: int, label: str, lines: int = 3) -> int:
    sheet.merge(row, 0, 1, label, fmts.chrome("label_bold"))
    sheet.ws.merge_range(row, 2, row + lines - 1, LAST_COL, "",
                         fmts.chrome("form_box"))
    for i in range(lines):
        sheet.row_height(row + i, 16)
    return row + lines


# ---------------------------------------------------------------------------
# Print Brief - one post, everything the maker needs
# ---------------------------------------------------------------------------

def build_brief(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, BRIEF_SHEET,
                    tab_color=Palette.accent)
    ws = s.ws
    s.set_col(0, 0, 18)
    s.set_col(1, 1, 16)
    for c in range(2, LAST_COL + 1):
        s.set_col(c, c, 13)

    default = "C-0001" if cfg.include_sample_data else ""
    _pick(s, fmts, "Content ID", "ContentIDs", "BriefID",
          "Pick a post from the Content Calendar - the whole brief fills "
          "itself in. This row does not print.", default)

    def pull(column: str) -> str:
        return (f'=IF(BriefID="","",IFERROR(INDEX('
                f'[[Content Calendar.{column}]],MATCH(BriefID,'
                f'[[Content Calendar.id]],0)),"Not found"))')

    row = 2
    s.merge(row, 0, LAST_COL, "CONTENT BRIEF", fmts.chrome("print_title"))
    s.row_height(row, 26)
    row += 1
    s.merge(row, 0, LAST_COL, "=BrandName&\"  |  \"&TEXT(TODAY(),\"d mmmm "
                              "yyyy\")", fmts.chrome("print_sub"))
    row += 2

    row = s.section(row, "The post", LAST_COL)
    row = _line(s, fmts, row, "Content ID", '=IF(BriefID="","",BriefID)')
    row = _line(s, fmts, row, "Title", pull("title"))
    row = _line(s, fmts, row, "Publish date", pull("date"), "date")
    row = _line(s, fmts, row, "Time slot", pull("slot"))
    row = _line(s, fmts, row, "Platform", pull("platform"))
    row = _line(s, fmts, row, "Format", pull("format"))
    row = _line(s, fmts, row, "Owner", pull("owner"))
    row = _line(s, fmts, row, "Status", pull("status"), "status")
    row += 1

    row = s.section(row, "Why we are making it", LAST_COL)
    row = _line(s, fmts, row, "Content pillar", pull("pillar"))
    row = _line(s, fmts, row, "Objective", pull("objective"))
    row = _line(s, fmts, row, "Funnel stage", pull("funnel"))
    row = _line(s, fmts, row, "Written for", pull("persona"))
    row = _line(s, fmts, row, "Campaign", pull("campaign"))
    row = _line(s, fmts, row, "What it is about", pull("description"),
                "wrap", height=34)
    row += 1

    row = s.section(row, "What to say", LAST_COL)
    caption = ('=IF(BriefID="","",IFERROR(INDEX([[Caption Bank.%s]],'
               'MATCH(INDEX([[Content Calendar.caption]],MATCH(BriefID,'
               '[[Content Calendar.id]],0)),[[Caption Bank.title]],0)),'
               '"- add a caption on the Caption Bank tab -"))')
    row = _line(s, fmts, row, "Caption name", pull("caption"))
    row = _line(s, fmts, row, "Hook", caption % "hook", "wrap", height=30)
    row = _line(s, fmts, row, "Caption", caption % "full", "wrap",
                height=70)
    row = _line(s, fmts, row, "Call to action", pull("cta"))
    row = _line(s, fmts, row, "Hashtag group", pull("hashtag_group"))
    row = _line(s, fmts, row, "Tags in the group", pull("hashtag_count"),
                "int")
    row += 1

    row = s.section(row, "What to make", LAST_COL)
    asset = ('=IF(BriefID="","",IFERROR(INDEX([[Assets.%s]],MATCH('
             'INDEX([[Content Calendar.asset]],MATCH(BriefID,'
             '[[Content Calendar.id]],0)),[[Assets.id]],0)),'
             '"- no asset linked -"))')
    row = _line(s, fmts, row, "Asset ID", pull("asset"))
    row = _line(s, fmts, row, "Asset name", asset % "title")
    row = _line(s, fmts, row, "File link", asset % "link", "wrap")
    row = _line(s, fmts, row, "Size", asset % "dimensions")
    row = _line(s, fmts, row, "Draft due", pull("draft_due"), "date")
    row = _line(s, fmts, row, "Approval due", pull("approve_due"), "date")
    row += 1

    row = s.section(row, "Notes & shot list (write on the printout)",
                    LAST_COL)
    row = _line(s, fmts, row, "Notes from the calendar", pull("notes"),
                "wrap", height=30)
    row = _blank_box(s, fmts, row, "Shot list", 4)
    row = _blank_box(s, fmts, row, "Props / setup", 3)
    row = _blank_box(s, fmts, row, "Approved by", 2)
    row += 1

    row = s.section(row, "After it goes live", LAST_COL)
    row = _line(s, fmts, row, "Published link", pull("url"), "wrap")
    row = _line(s, fmts, row, "Reach", pull("reach"), "int")
    row = _line(s, fmts, row, "Engagements", pull("engagements"), "int")
    row = _line(s, fmts, row, "Engagement rate", pull("er"), "percent")
    row += 1
    s.note(row,
           "Record the numbers on the Performance tab 48-72 hours after "
           "publishing - they appear here and on the calendar "
           "automatically.",
           LAST_COL, height=26)
    row += 2

    s.finish(freeze_row=None, gridlines=False, fit_width=True)
    ws.print_area(f"A3:{chr(65 + LAST_COL)}{row}")
    ws.set_portrait()
    ws.center_horizontally()
    return s


# ---------------------------------------------------------------------------
# Print Week - seven days of posts on one page
# ---------------------------------------------------------------------------

def build_week(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, WEEK_SHEET,
                    tab_color=Palette.accent)
    ws = s.ws
    s.set_col(0, 0, 13)
    s.set_col(1, 1, 13)
    s.set_col(2, 2, 34)
    s.set_col(3, 3, 16)
    s.set_col(4, 4, 18)
    s.set_col(5, 5, 16)
    s.set_col(6, 6, 14)
    s.set_col(7, 7, 14)

    s.merge(0, 0, 1, "Week starting", fmts.chrome("label_bold"))
    s.write(0, 2, cfg.today, fmts.input("date", bold=True))
    ws.data_validation(0, 2, 0, 2, date_check())
    book.define_name("WeekStartDate",
                     f"={quote_sheet(WEEK_SHEET)}!$C$1")
    s.merge(0, 3, LAST_COL,
            "Pick any date - the sheet prints the seven days that follow. "
            "This row does not print.", fmts.chrome("note"))
    s.row_height(0, 22)

    row = 2
    s.merge(row, 0, LAST_COL, "WEEKLY CONTENT PLAN",
            fmts.chrome("print_title"))
    s.row_height(row, 26)
    row += 1
    s.merge(row, 0, LAST_COL,
            '=BrandName&"  |  "&TEXT(WeekStartDate,"d mmm")&" to "&'
            'TEXT(WeekStartDate+6,"d mmm yyyy")',
            fmts.chrome("print_sub"))
    row += 2

    key = "[[Content Calendar.day_key]]"
    per_day = 4
    for day in range(7):
        day_cell = f"WeekStartDate+{day}"
        row = s.section(row, f'=TEXT({day_cell},"dddd d mmmm")', LAST_COL)
        s.table_head(row, 0, ["Time", "Platform", "Post", "Format",
                              "Pillar", "Owner", "Status", "Done"])
        first = row + 1
        for n in range(1, per_day + 1):
            r = first + n - 1
            pick = (f'SMALL({key},COUNTIF({key},"<"&{day_cell})+{n})')
            base = (f'=IFERROR(IF({pick}>={day_cell}+1,"",INDEX('
                    f'[[Content Calendar.%s]],MATCH({pick},{key},0))),"")')
            s.calc_row(r, 0, [
                base % "slot",
                base % "platform",
                base % "title",
                base % "format",
                base % "pillar",
                base % "owner",
                base % "status",
                "",
            ], ["text", "text", "text", "text", "text", "text", "status",
                "text"])
            ws.write_blank(r, 7, None, fmts.chrome("form_line"))
        row = first + per_day

    row += 1
    row = s.section(row, "This week in numbers", LAST_COL)
    window = ('[[Content Calendar.date]],">="&WeekStartDate,'
              '[[Content Calendar.date]],"<="&WeekStartDate+6')
    s.table_head(row, 0, ["Posts", "Published", "Still to do",
                          "Weekly target", "Against target", "Platforms",
                          "Captions missing", "Assets missing"])
    s.calc_row(row + 1, 0, [
        f"=COUNTIFS({window})",
        f'=COUNTIFS({window},[[Content Calendar.status]],"Published")',
        f'=COUNTIFS({window})-COUNTIFS({window},'
        f'[[Content Calendar.status]],"Published")',
        "=PostsPerWeek",
        f'=IFERROR(COUNTIFS({window})/PostsPerWeek,"")',
        f'=SUMPRODUCT((COUNTIFS([[Content Calendar.platform]],'
        f'[[Calc.p_name]],{window})>0)*1)',
        f'=COUNTIFS({window},[[Content Calendar.caption]],"")',
        f'=COUNTIFS({window},[[Content Calendar.asset]],"")',
    ], ["int", "int", "int", "int", "percent", "int", "int", "int"])
    row += 3

    s.note(row,
           "Print this on a Monday morning, stick it on the wall and tick "
           "the last column as each post goes out.",
           LAST_COL, height=24)
    row += 2

    s.finish(freeze_row=None, gridlines=False, fit_width=True)
    ws.print_area(f"A3:{chr(65 + LAST_COL)}{row}")
    ws.set_landscape()
    ws.center_horizontally()
    return s
