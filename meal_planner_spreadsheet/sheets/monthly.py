"""`Monthly Planner` - the month as a calendar.

Six weeks of boxes, each showing the dinner (and the lunch, if there is
one) planned for that day with its cost underneath. Pick a month at the
top; everything reads from the Meal Plan, so the grid is never out of
step with the shopping list.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, list_from

NAME = "Monthly Planner"
DAYS = 7
WEEKS = 6
LAST_COL = DAYS - 1
COL_WIDTH = 21.0

_KEY = "[[Meal Plan.key]]"


def _slot(day_cell: str, slot: str) -> str:
    key = f'TEXT({day_cell},"yyyy-mm-dd")&"|"&"{slot}"'
    return (f'=IF({day_cell}="","",IFERROR(INDEX('
            f'[[Meal Plan.recipe_name]],MATCH({key},{_KEY},0)),""))')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.accent)
    ws = s.ws

    s.banner("Monthly Meal Planner",
             "The whole month on one page - dinners, lunches and what "
             "each day costs. Print it and pin it up.", LAST_COL)
    for c in range(DAYS):
        s.set_col(c, c, COL_WIDTH)

    # --------------------------------------------------- month controls --
    row = 4
    s.merge(row, 0, 0, "Month", fmts.chrome("label_bold"))
    s.merge(row, 1, 1, cfg.today.strftime("%B"), fmts.input("text"))
    ws.data_validation(row, 1, row, 1, list_from("MonthList"))
    book.define_name("PlanMonthName", f"={quote_sheet(NAME)}!$B${row + 1}")
    s.merge(row, 2, 2, "First day", fmts.chrome("label_right"))
    ws.write_formula(
        row, 3,
        '=DATE(PlanYear,MAX(1,IFERROR(MATCH(PlanMonthName,MonthList,0),'
        'MONTH(TODAY()))),1)', fmts.calc("date"))
    book.define_name("PlanMonthStart", f"={quote_sheet(NAME)}!$D${row + 1}")
    s.merge(row, 4, 4, "Meals planned", fmts.chrome("label_right"))
    ws.write_formula(
        row, 5,
        s.reg.render('=COUNTIFS([[Meal Plan.date]],">="&PlanMonthStart,'
                     '[[Meal Plan.date]],"<="&EOMONTH(PlanMonthStart,0))'),
        fmts.calc("int"))
    s.merge(row, 6, 6, None, fmts.calc("money"))
    ws.write_formula(
        row, 6,
        s.reg.render('=SUMIFS([[Meal Plan.cost]],[[Meal Plan.date]],'
                     '">="&PlanMonthStart,[[Meal Plan.date]],'
                     '"<="&EOMONTH(PlanMonthStart,0))'),
        fmts.calc("money"))
    row += 1

    row = s.note(row,
                 "Each box shows the dinner in bold type, then the "
                 "lunch, then what the day costs. Empty boxes are days "
                 "with nothing planned - the gaps are the point.",
                 LAST_COL, height=26)

    # -------------------------------------------------- weekday headings --
    first_cell = ('=PlanMonthStart-MOD(WEEKDAY(PlanMonthStart,'
                  'IF(WeekStart="Sunday",1,2))-1,7)')
    head = row
    for i in range(DAYS):
        ws.write_formula(
            head, i,
            f'=TEXT({first_cell.lstrip("=")}+{i},"dddd")',
            fmts.chrome("header"))
    s.row_height(head, 24)
    row = head + 1

    # ------------------------------------------------------- the calendar
    for week in range(WEEKS):
        date_row = row
        meal_row = row + 1
        lunch_row = row + 2
        cost_row = row + 3
        for i in range(DAYS):
            offset = week * 7 + i
            ws.write_formula(
                date_row, i,
                f'={first_cell.lstrip("=")}+{offset}',
                fmts.chrome("day_number", num_format="d mmm"))
            cell = f"{chr(65 + i)}${date_row + 1}"
            ws.write_formula(meal_row, i,
                             s.reg.render(_slot(cell, "Dinner")),
                             fmts.chrome("day_item"))
            ws.write_formula(lunch_row, i,
                             s.reg.render(_slot(cell, "Lunch")),
                             fmts.cell("wrap", "calc"))
            ws.write_formula(
                cost_row, i,
                s.reg.render(f'=IF({cell}="","",IF(SUMIFS('
                             f'[[Meal Plan.cost]],[[Meal Plan.date]],'
                             f'{cell})=0,"",SUMIFS([[Meal Plan.cost]],'
                             f'[[Meal Plan.date]],{cell})))'),
                fmts.calc("money"))
            # grey out the days that belong to the month either side
            ws.conditional_format(
                date_row, i, cost_row, i,
                {"type": "formula",
                 "criteria": f'=MONTH({cell})<>MONTH(PlanMonthStart)',
                 "format": fmts.cf("muted")})
        s.row_height(date_row, 16)
        s.row_height(meal_row, 26)
        s.row_height(lunch_row, 22)
        s.row_height(cost_row, 16)
        row = cost_row + 1
    row += 1

    # ------------------------------------------------------ month summary
    row = s.section(row, "The month in numbers", LAST_COL)
    row = s.table_head(row, 0, ["Meals planned", "Days covered",
                                "Planned food cost", "Monthly budget",
                                "Difference", "Cooked at home",
                                "Eaten out"],
                       [14, 13, 17, 14, 13, 15, 13])
    span = ('[[Meal Plan.date]],">="&PlanMonthStart,'
            '[[Meal Plan.date]],"<="&EOMONTH(PlanMonthStart,0)')
    diff_row = row
    row = s.calc_row(row, 0, [
        f'=COUNTIFS({span})',
        f'=SUMPRODUCT((COUNTIFS({span},[[Meal Plan.date]],'
        f'[[Meal Plan.date]])>0)*1)/MAX(1,1)',
        f'=SUMIFS([[Meal Plan.cost]],{span})',
        "=MonthlyBudget",
        f'=MonthlyBudget-SUMIFS([[Meal Plan.cost]],{span})',
        f'=COUNTIFS({span},[[Meal Plan.home_cooked]],"Yes")',
        f'=COUNTIFS({span},[[Meal Plan.home_cooked]],"No")',
    ], ["int", "int", "money", "money", "money", "int", "int"])
    ws.conditional_format(diff_row, 4, diff_row, 4,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    row += 1

    s.note(row, "Tip: plan the month roughly - the same six or seven "
                "dinners on a rotation is fine - then fine tune one "
                "week at a time on the Weekly Planner.", LAST_COL,
           height=24)

    s.finish(freeze_row=head + 1, print_area=f"A1:G{row + 2}",
             gridlines=False)
    return s
