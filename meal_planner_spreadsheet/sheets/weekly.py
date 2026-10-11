"""`Weekly Planner` - the week on one page.

Seven days across, your meal slots down, with the cost and the calories
worked out underneath and a shopping summary at the bottom. Pick a week
with the date at the top and the whole page re-reads the Meal Plan.

Every cell here is a lookup, not a place to type: meals are added on the
Meal Plan tab (one row each, with drop-downs) and this page follows
along. That keeps one copy of the truth, which is what lets the grocery
list be trustworthy.
"""

from __future__ import annotations

from ..config import Config, MAIN_SLOTS, MEAL_SLOTS, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, date_check

NAME = "Weekly Planner"
DAYS = 7
LAST_COL = 8                 # A = labels, B..H = the week, I = totals
DAY_WIDTH = 19.0

_KEY = "[[Meal Plan.key]]"


def _slot_cell(day_cell: str, slot: str, column: str) -> str:
    """Pull `column` from the meal planned in `slot` on `day_cell`."""
    key = f'TEXT({day_cell},"yyyy-mm-dd")&"|"&"{slot}"'
    return (f'=IF({day_cell}="","",IFERROR(INDEX('
            f'[[Meal Plan.{column}]],MATCH({key},{_KEY},0)),""))')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.accent)
    ws = s.ws

    s.banner("Weekly Meal Planner",
             "One week at a glance - what you are eating, what it costs "
             "and what still needs cooking.", LAST_COL)
    s.set_col(0, 0, 16)
    for c in range(1, DAYS + 1):
        s.set_col(c, c, DAY_WIDTH)
    s.set_col(DAYS + 1, DAYS + 1, 15)

    # ------------------------------------------------------ week picker --
    row = 4
    s.merge(row, 0, 0, "Week starting", fmts.chrome("label_bold"))
    s.write(row, 1, cfg.start_date, fmts.input("date"))
    ws.data_validation(row, 1, row, 1, date_check())
    book.define_name("WeekPick", f"={quote_sheet(NAME)}!$B${row + 1}")
    s.merge(row, 2, 2, "Week number", fmts.chrome("label_right"))
    ws.write_formula(row, 3,
                     '=WEEKNUM(WeekPick,IF(WeekStart="Sunday",1,2))',
                     fmts.calc("int"))
    s.merge(row, 4, 4, "People", fmts.chrome("label_right"))
    ws.write_formula(row, 5, "=DefaultServings", fmts.calc("int"))
    s.merge(row, 6, 6, "Weekly budget", fmts.chrome("label_right"))
    ws.write_formula(row, 7, "=WeeklyBudget", fmts.calc("money"))
    row += 1

    row = s.note(row,
                 "Change the date above to any Monday - every box on "
                 "the page follows it. Meals themselves are added on "
                 "the Meal Plan tab; this page is the view, not the "
                 "entry form.", LAST_COL, height=26)

    # ----------------------------------------------------- day headings --
    head = row
    s.write(head, 0, "Meal", fmts.chrome("header_left"))
    for i in range(DAYS):
        ws.write_formula(head, i + 1, f"=WeekPick+{i}",
                         fmts.chrome("header", num_format='ddd d mmm'))
    s.write(head, DAYS + 1, "Week total", fmts.chrome("header"))
    s.row_height(head, 26)
    day_cells = [f"{chr(66 + i)}${head + 1}" for i in range(DAYS)]
    row = head + 1

    # ----------------------------------------------------- the meal grid --
    slot_rows = {}
    for slot in MEAL_SLOTS:
        s.write(row, 0, slot, fmts.chrome("label_bold"))
        for i in range(DAYS):
            ws.write_formula(row, i + 1,
                             s.reg.render(_slot_cell(day_cells[i], slot,
                                                     "recipe_name")),
                             fmts.cell("wrap", "calc"))
        ws.write_formula(
            row, DAYS + 1,
            s.reg.render(f'=COUNTIFS([[Meal Plan.date]],">="&WeekPick,'
                         f'[[Meal Plan.date]],"<="&WeekPick+6,'
                         f'[[Meal Plan.slot]],"{slot}")'),
            fmts.calc("int"))
        s.row_height(row, 30)
        slot_rows[slot] = row
        row += 1

    # cost and calories strips -------------------------------------------
    s.write(row, 0, "Cost of the day", fmts.chrome("label_bold"))
    for i in range(DAYS):
        ws.write_formula(
            row, i + 1,
            s.reg.render(f'=IF({day_cells[i]}="","",'
                         f'SUMIFS([[Meal Plan.cost]],[[Meal Plan.date]],'
                         f'{day_cells[i]}))'),
            fmts.calc("money"))
    ws.write_formula(row, DAYS + 1,
                     s.reg.render('=SUMIFS([[Meal Plan.cost]],'
                                  '[[Meal Plan.date]],">="&WeekPick,'
                                  '[[Meal Plan.date]],"<="&WeekPick+6)'),
                     fmts.chrome("total_value",
                                 num_format=cfg.money_format))
    cost_row = row
    row += 1

    s.write(row, 0, "Calories a head", fmts.chrome("label_bold"))
    for i in range(DAYS):
        ws.write_formula(
            row, i + 1,
            s.reg.render(f'=IF({day_cells[i]}="","",IFERROR('
                         f'SUMIFS([[Meal Plan.cals_total]],'
                         f'[[Meal Plan.date]],{day_cells[i]})/'
                         f'DefaultServings,""))'),
            fmts.calc("number"))
    ws.write_formula(
        row, DAYS + 1,
        s.reg.render('=IFERROR(SUMIFS([[Meal Plan.cals_total]],'
                     '[[Meal Plan.date]],">="&WeekPick,'
                     '[[Meal Plan.date]],"<="&WeekPick+6)/'
                     '(DefaultServings*7),"")'),
        fmts.calc("number"))
    cals_row = row
    row += 1

    s.write(row, 0, "Who is cooking", fmts.chrome("label_bold"))
    for i in range(DAYS):
        ws.write_formula(
            row, i + 1,
            s.reg.render(_slot_cell(day_cells[i], MAIN_SLOTS[-1], "cook")),
            fmts.calc("text"))
    ws.write_formula(row, DAYS + 1,
                     s.reg.render('=COUNTIFS([[Meal Plan.date]],'
                                  '">="&WeekPick,[[Meal Plan.date]],'
                                  '"<="&WeekPick+6,[[Meal Plan.cook]],'
                                  '"<>")'),
                     fmts.calc("int"))
    row += 2

    # cost colouring: green under budget per day, amber over -------------
    daily = "WeeklyBudget/7"
    ws.conditional_format(cost_row, 1, cost_row, DAYS,
                          {"type": "formula",
                           "criteria": f"=AND(B{cost_row + 1}<>\"\","
                                       f"B{cost_row + 1}>{daily})",
                           "format": fmts.cf("warn")})
    ws.conditional_format(cost_row, 1, cost_row, DAYS,
                          {"type": "formula",
                           "criteria": f"=AND(B{cost_row + 1}<>\"\","
                                       f"B{cost_row + 1}<={daily})",
                           "format": fmts.cf("good")})
    ws.conditional_format(cals_row, 1, cals_row, DAYS,
                          {"type": "formula",
                           "criteria": f"=AND(B{cals_row + 1}<>\"\","
                                       f"B{cals_row + 1}>CalorieTarget)",
                           "format": fmts.cf("warn")})

    # ------------------------------------------------- the week in numbers
    row = s.section(row, "This week in numbers", LAST_COL)
    kpis = [
        ("Meals planned",
         '=COUNTIFS([[Meal Plan.date]],">="&WeekPick,'
         '[[Meal Plan.date]],"<="&WeekPick+6)', "int",
         "every slot with something in it"),
        ("Planned food cost",
         '=SUMIFS([[Meal Plan.cost]],[[Meal Plan.date]],">="&WeekPick,'
         '[[Meal Plan.date]],"<="&WeekPick+6)', "money",
         "recipes only - the shop may differ"),
        ("Against the budget",
         '=WeeklyBudget-SUMIFS([[Meal Plan.cost]],[[Meal Plan.date]],'
         '">="&WeekPick,[[Meal Plan.date]],"<="&WeekPick+6)', "money",
         "green is money left"),
        ("Cooked at home",
         '=COUNTIFS([[Meal Plan.date]],">="&WeekPick,'
         '[[Meal Plan.date]],"<="&WeekPick+6,'
         '[[Meal Plan.home_cooked]],"Yes")', "int",
         f"target {cfg.home_cooked_target} a week"),
    ]
    for i, (label, formula, style, note) in enumerate(kpis):
        s.kpi(row, i * 2, label, s.reg.render(formula), style=style,
              span=2, note=note)
    row += 4

    # ---------------------------------------------------- shopping & prep
    row = s.section(row, "Before the week starts", LAST_COL)
    row = s.table_head(row, 0, ["Job", "Where it stands", "Detail"],
                       [26, 20, 44])
    jobs = [
        ("Shopping list built",
         '=IF(COUNTIF([[Grocery List.ingredient]],"<>")=0,'
         '"Nothing on it","'
         'Ready")',
         '="The list covers "&TEXT(ShopFrom,"d mmm")&" to "&'
         'TEXT(ShopTo,"d mmm")&" - "&'
         'COUNTIF([[Grocery List.ingredient]],"<>")&" items, about "&'
         'TEXT(SUM([[Grocery List.est_cost]]),"0.00")'),
        ("Shopping done",
         '=IF(COUNTIF([[Grocery List.got]],"Yes")=0,"Not started",'
         'IF(COUNTIFS([[Grocery List.short]],">0",'
         '[[Grocery List.got]],"<>Yes")=0,"Done","In progress"))',
         '=COUNTIF([[Grocery List.got]],"Yes")&" of "&'
         'COUNTIFS([[Grocery List.short]],">0")&" items ticked off"'),
        ("Prep jobs for this week",
         '=IF(COUNTIFS([[Meal Prep.date]],">="&WeekPick,'
         '[[Meal Prep.date]],"<="&WeekPick+6)=0,"None planned",'
         'IF(COUNTIFS([[Meal Prep.date]],">="&WeekPick,'
         '[[Meal Prep.date]],"<="&WeekPick+6,'
         '[[Meal Prep.status]],"<>Done")=0,"Done","To do"))',
         '=COUNTIFS([[Meal Prep.date]],">="&WeekPick,'
         '[[Meal Prep.date]],"<="&WeekPick+6)&" job(s) on the prep day"'),
        ("Leftovers to eat first",
         '=IF(COUNTIFS([[Leftovers.status]],"<>Eaten",'
         '[[Leftovers.status]],"<>Binned")=0,"None","Use soon")',
         '=SUMIFS([[Leftovers.portions]],[[Leftovers.status]],'
         '"<>Eaten",[[Leftovers.status]],"<>Binned")&'
         '" portion(s) in the fridge or freezer"'),
        ("Food about to go off",
         '=IF(COUNTIF([[Pantry.status]],"Use soon")=0,"All good",'
         '"Use soon")',
         '=COUNTIF([[Pantry.status]],"Use soon")&" pantry item(s) '
         'within "&ExpiryWarningDays&" days of their date"'),
    ]
    for label, state, detail in jobs:
        s.write(row, 0, label, fmts.chrome("label"))
        ws.write_formula(row, 1, s.reg.render(state), fmts.calc("status"))
        ws.write_formula(row, 2, s.reg.render(detail), fmts.calc("text"))
        s.merge(row, 2, LAST_COL, None, fmts.calc("text"))
        ws.write_formula(row, 2, s.reg.render(detail), fmts.calc("text"))
        row += 1
    for kind, word in (("good", "Done"), ("good", "Ready"),
                       ("warn", "To do"), ("warn", "In progress"),
                       ("warn", "Use soon"), ("muted", "None"),
                       ("muted", "Nothing on it"), ("good", "All good"),
                       ("warn", "Not started")):
        ws.conditional_format(row - len(jobs), 1, row - 1, 1,
                              {"type": "text", "criteria": "containing",
                               "value": word, "format": fmts.cf(kind)})
    row += 1

    s.note(row, "Printing: this page is set to one page wide. Print it "
                "on a Sunday, stick it on the fridge and the week looks "
                "after itself.", LAST_COL, height=24)

    s.finish(freeze_row=head + 1, freeze_col=1,
             print_area=f"A1:I{row + 2}", gridlines=False)
    return s
