"""`Month View` - the calendar as a calendar: six weeks across seven days.

Excel 2010 has no FILTER, so each day cell finds its posts with a trick:
the Content Calendar carries a hidden ``day_key`` (the date plus a tiny
fraction of the row number).  Because that key sorts by date, the n-th
post on a given day is ``SMALL(day_key, COUNTIF(day_key,"<"&day) + n)``,
which needs no array formulas at all.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, list_from

NAME = "Month View"
DAYS = 7
WEEKS = 6
PER_DAY = 3              # posts shown in each day cell
COL_WIDTH = 19.0
LAST_COL = DAYS - 1

_KEY = "[[Content Calendar.day_key]]"


def _nth(day_cell: str, n: int, column: str) -> str:
    """Formula for the n-th post on `day_cell`, pulling `column`."""
    pick = f'SMALL({_KEY},COUNTIF({_KEY},"<"&{day_cell})+{n})'
    return (f'=IF({day_cell}="","",IFERROR(IF({pick}>={day_cell}+1,"",'
            f'INDEX([[Content Calendar.{column}]],'
            f'MATCH({pick},{_KEY},0))),""))')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.accent)
    ws = s.ws

    s.banner("Month View",
             "The month laid out as a grid, straight from the Content "
             "Calendar. Pick a month, then print it and pin it up.",
             LAST_COL)
    for c in range(DAYS):
        s.set_col(c, c, COL_WIDTH)

    # ------------------------------------------------------ month picker
    row = 4
    s.merge(row, 0, 1, "Month", fmts.chrome("label_bold"))
    s.merge(row, 2, 3, cfg.today.strftime("%B"), fmts.input("text"))
    ws.data_validation(row, 2, row, 2, list_from("MonthList"))
    book.define_name("ViewMonthName", f"={quote_sheet(NAME)}!$C${row + 1}")
    s.merge(row, 4, 4, "First day", fmts.chrome("label_bold"))
    s.merge(row, 5, 6, None, fmts.calc("date"))
    ws.write_formula(
        row, 5,
        '=DATE(PlanYear,MAX(1,IF(ViewMonthName="All year",'
        'MONTH(TODAY()),MATCH(ViewMonthName,MonthList,0)-1)),1)',
        fmts.calc("date"))
    book.define_name("ViewStart", f"={quote_sheet(NAME)}!$F${row + 1}")
    row += 1

    s.note(row,
           "Each box shows up to three posts for that day: platform, "
           "then the title. Days with more than three are marked '+ more' "
           "- the Content Calendar has the full list.",
           LAST_COL, height=24)
    row += 1

    # ----------------------------------------------------- weekday header
    first_cell = ('=ViewStart-MOD(WEEKDAY(ViewStart,IF(WeekStart="Sunday",'
                  '1,2))-1,7)')
    header_row = row
    for d in range(DAYS):
        ws.write_formula(
            header_row, d,
            f'=TEXT({first_cell.lstrip("=")}+{d},"dddd")',
            fmts.chrome("header"))
    s.row_height(header_row, 22)
    row += 1

    # ------------------------------------------------------- the day grid
    date_fmt = fmts.chrome("day_number")
    day_cells = []
    for week in range(WEEKS):
        date_row = row
        s.row_height(date_row, 20)
        for d in range(DAYS):
            offset = week * 7 + d
            ws.write_formula(
                date_row, d,
                f'=IF(MONTH({first_cell.lstrip("=")}+{offset})'
                f'<>MONTH(ViewStart),"",'
                f'{first_cell.lstrip("=")}+{offset})',
                date_fmt)
            day_cells.append((date_row, d))
        for line in range(PER_DAY):
            r = date_row + 1 + line
            s.row_height(r, 15)
            for d in range(DAYS):
                cell = f"{chr(65 + d)}{date_row + 1}"
                if line < PER_DAY - 1:
                    formula = (
                        f'=IF({cell}="","",IFERROR(IF('
                        f'SMALL({_KEY},COUNTIF({_KEY},"<"&{cell})'
                        f'+{line + 1})>={cell}+1,"",'
                        f'INDEX([[Content Calendar.platform]],MATCH(SMALL('
                        f'{_KEY},COUNTIF({_KEY},"<"&{cell})+{line + 1}),'
                        f'{_KEY},0))&" - "&INDEX('
                        f'[[Content Calendar.title]],MATCH(SMALL({_KEY},'
                        f'COUNTIF({_KEY},"<"&{cell})+{line + 1}),'
                        f'{_KEY},0))),""))')
                else:
                    # last line: either the third post or a "+ n more" note
                    formula = (
                        f'=IF({cell}="","",IF(COUNTIF({_KEY},"<"&{cell}+1)'
                        f'-COUNTIF({_KEY},"<"&{cell})>{PER_DAY},'
                        f'"+ "&COUNTIF({_KEY},"<"&{cell}+1)'
                        f'-COUNTIF({_KEY},"<"&{cell})-{PER_DAY - 1}'
                        f'&" more",IFERROR(IF(SMALL({_KEY},'
                        f'COUNTIF({_KEY},"<"&{cell})+{line + 1})'
                        f'>={cell}+1,"",INDEX('
                        f'[[Content Calendar.platform]],MATCH(SMALL({_KEY},'
                        f'COUNTIF({_KEY},"<"&{cell})+{line + 1}),'
                        f'{_KEY},0))&" - "&INDEX('
                        f'[[Content Calendar.title]],MATCH(SMALL({_KEY},'
                        f'COUNTIF({_KEY},"<"&{cell})+{line + 1}),'
                        f'{_KEY},0))),"")))')
                ws.write_formula(r, d, s.reg.render(formula),
                                 fmts.chrome("day_item"))
        ws.conditional_format(
            date_row, 0, date_row + PER_DAY, DAYS - 1,
            {"type": "formula",
             "criteria": f"=A${date_row + 1}=TODAY()",
             "format": fmts.cf("info")})
        row = date_row + PER_DAY + 1


    # ---------------------------------------------------- month summary --
    row += 1
    row = s.section(row, "This month at a glance", LAST_COL)
    s.table_head(row, 0, ["Posts planned", "Published", "Still to go",
                          "Target", "Against target", "Reach",
                          "Engagements"])
    window = ('[[Content Calendar.date]],">="&ViewStart,'
              '[[Content Calendar.date]],"<="&EOMONTH(ViewStart,0)')
    an_window = ('[[Analytics.month]],">="&ViewStart,'
                 '[[Analytics.month]],"<="&EOMONTH(ViewStart,0)')
    s.calc_row(row + 1, 0, [
        f"=COUNTIFS({window})",
        f'=COUNTIFS({window},[[Content Calendar.status]],"Published")',
        f'=COUNTIFS({window})-COUNTIFS({window},'
        f'[[Content Calendar.status]],"Published")',
        "=MonthlyPostTarget",
        f'=IFERROR(COUNTIFS({window},[[Content Calendar.status]],'
        f'"Published")/MonthlyPostTarget,"")',
        f"=SUMIFS([[Analytics.reach]],{an_window})",
        f"=SUMIFS([[Analytics.engagements]],{an_window})",
    ], ["int", "int", "int", "int", "percent", "int", "int"])
    row += 3

    s.note(row,
           "Posts appear here the moment you add them to the Content "
           "Calendar - there is nothing to copy across. Print this page "
           "on one sheet of landscape A4 and it makes a very good "
           "wall planner.",
           LAST_COL, height=26)

    s.finish(freeze_row=header_row + 1, gridlines=False, fit_width=True)
    ws.set_landscape()
    return s
