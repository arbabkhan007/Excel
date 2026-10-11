"""`Dashboard` - the one page that answers 'are we alright?'

Eight numbers across the top, a list of things that need doing, this
week's dinners, and twelve months of spending, waste and home cooking
with a little column chart beside each row. Everything reads straight
from the registers, so it is never out of date.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, date_check

NAME = "Dashboard"
LAST_COL = 11
MONTHS = 12


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws

    s.banner("Dashboard",
             "This week, this month and the year so far - all from the "
             "numbers you have already entered.", LAST_COL)
    s.set_col(0, 0, 20)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 13)

    # ------------------------------------------------------- the picker --
    row = 4
    s.merge(row, 0, 0, "Week of", fmts.chrome("label_bold"))
    ws.write_formula(row, 1, "=WeekPick", fmts.input("date"))
    ws.data_validation(row, 1, row, 1, date_check())
    book.define_name("DashWeek", f"={quote_sheet(NAME)}!$B${row + 1}")
    s.merge(row, 2, 2, "Month", fmts.chrome("label_right"))
    ws.write_formula(row, 3, '=TEXT(TODAY(),"mmmm yyyy")',
                     fmts.calc("text"))
    book.define_name("DashMonth", f"={quote_sheet(NAME)}!$D${row + 1}")
    s.merge(row, 4, 4, "Month key", fmts.chrome("label_right"))
    ws.write_formula(row, 5, '=TEXT(TODAY(),"yyyy-mm")',
                     fmts.calc("text"))
    book.define_name("DashKey", f"={quote_sheet(NAME)}!$F${row + 1}")
    s.merge(row, 6, 6, "Household", fmts.chrome("label_right"))
    ws.write_formula(row, 7, "=HouseMembers", fmts.calc("int"))
    s.merge(row, 8, 9, "Updated", fmts.chrome("label_right"))
    ws.write_formula(row, 10, "=TODAY()", fmts.calc("date"))
    row += 2

    # ---------------------------------------------------------- the KPIs --
    tiles_1 = [
        ("Meals planned this week",
         '=COUNTIFS([[Meal Plan.date]],">="&DashWeek,'
         '[[Meal Plan.date]],"<="&DashWeek+6)', "int",
         "every slot with something in it"),
        ("Planned cost this week",
         '=SUMIFS([[Meal Plan.cost]],[[Meal Plan.date]],">="&DashWeek,'
         '[[Meal Plan.date]],"<="&DashWeek+6)', "money",
         "at price book prices"),
        ("This week's shopping list",
         '=SUM([[Grocery List.est_cost]])', "money",
         "what the list adds up to"),
        ("Left in the weekly budget",
         '=WeeklyBudget-SUM([[Grocery List.est_cost]])', "money",
         "green is money left"),
    ]
    tiles_2 = [
        ("Spent this month",
         '=SUMIFS([[Grocery Spending.amount]],'
         '[[Grocery Spending.month]],DashKey)', "money",
         "every shop, takeaway and meal out"),
        ("Pantry value",
         '=SUM([[Pantry.value]])', "money",
         "what is sitting in the cupboards"),
        ("Food wasted this month",
         '=SUMIFS([[Food Waste.value]],[[Food Waste.month]],DashKey)',
         "money", f"target under {cfg.currency_symbol}"
                  f"{cfg.waste_cost_target:,.0f}"),
        ("Cooked at home this week",
         '=COUNTIFS([[Meal Plan.date]],">="&DashWeek,'
         '[[Meal Plan.date]],"<="&DashWeek+6,'
         '[[Meal Plan.home_cooked]],"Yes")', "int",
         f"target {cfg.home_cooked_target} a week"),
    ]
    for tiles in (tiles_1, tiles_2):
        for i, (label, formula, style, note) in enumerate(tiles):
            s.kpi(row, i * 3, label, s.reg.render(formula), style=style,
                  span=3, note=note)
        row += 4

    # ----------------------------------------------- needs your attention
    row = s.section(row, "Needs your attention", LAST_COL)
    row = s.table_head(row, 0, ["What", "How many", "What it means",
                                "Where to fix it"],
                       [26, 11, 46, 20])
    alerts = [
        ("Food about to go off",
         '=COUNTIF([[Pantry.status]],"Use soon")',
         f'="Use these within {cfg.expiry_warning_days} days or they '
         f'are money in the bin"', "Pantry"),
        ("Already expired",
         '=COUNTIF([[Pantry.status]],"Expired")',
         '="Past their date - check, then bin and log them"',
         "Pantry"),
        ("Leftovers waiting",
         '=SUMIFS([[Leftovers.portions]],[[Leftovers.status]],'
         '"<>Eaten",[[Leftovers.status]],"<>Binned")',
         '="Portions in the fridge or freezer worth "&'
         'TEXT(SUMIFS([[Leftovers.value]],[[Leftovers.status]],'
         '"<>Eaten",[[Leftovers.status]],"<>Binned"),"0.00")',
         "Leftovers"),
        ("Staples below par",
         '=COUNTIF([[Staples.status]],"Low stock")+'
         'COUNTIF([[Staples.status]],"Out of stock")',
         '="Things you said you never want to run out of"',
         "Staples"),
        ("Ingredients with no price",
         '=COUNTIFS([[Ingredients.name]],"<>",'
         '[[Ingredients.pack_price]],"")',
         '="Until these have a price, recipes using them are '
         'under-costed"', "Ingredients"),
        ("Recipe lines to fix",
         '=COUNTIF([[Recipe Ingredients.alert]],"Missing*")+'
         'COUNTIF([[Recipe Ingredients.alert]],"Check - *")',
         '="Unit mismatches or ingredients not in the price book - '
         'they are left out of the shopping list"',
         "Recipe Ingredients"),
        ("Days with nothing planned",
         "=" + "+".join(
             f'(COUNTIFS([[Meal Plan.date]],DashWeek+{i},'
             f'[[Meal Plan.slot]],"Dinner")=0)*1'
             for i in range(7)),
         '="Dinner slots still empty in the week you are looking at"',
         "Meal Plan"),
        ("Prep jobs outstanding",
         '=COUNTIFS([[Meal Prep.status]],"<>Done",'
         '[[Meal Prep.job]],"<>")',
         '="Batch cooking you planned but have not made yet"',
         "Meal Prep"),
        ("Goals behind target",
         '=COUNTIF([[Goals.on_track]],"No")',
         '="Worth five minutes with a cup of tea"', "Goals"),
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

    # --------------------------------------------------- this week's food
    row = s.section(row, "This week's dinners", LAST_COL)
    s.write(row, 0, "Day", fmts.chrome("header_left"))
    for i in range(7):
        ws.write_formula(row, i + 1, f"=DashWeek+{i}",
                         fmts.chrome("header", num_format="ddd d"))
    s.write(row, 8, "Week", fmts.chrome("header"))
    day_row = row
    row += 1
    s.write(row, 0, "Dinner", fmts.chrome("label_bold"))
    for i in range(7):
        cell = f"{chr(66 + i)}${day_row + 1}"
        ws.write_formula(
            row, i + 1,
            s.reg.render(f'=IFERROR(INDEX([[Meal Plan.recipe_name]],'
                         f'MATCH(TEXT({cell},"yyyy-mm-dd")&"|Dinner",'
                         f'[[Meal Plan.key]],0)),"-")'),
            fmts.cell("wrap", "calc"))
    ws.write_formula(row, 8,
                     s.reg.render('=COUNTIFS([[Meal Plan.date]],'
                                  '">="&DashWeek,[[Meal Plan.date]],'
                                  '"<="&DashWeek+6,[[Meal Plan.slot]],'
                                  '"Dinner")'), fmts.calc("int"))
    s.row_height(row, 30)
    row += 1
    s.write(row, 0, "Cost", fmts.chrome("label_bold"))
    for i in range(7):
        cell = f"{chr(66 + i)}${day_row + 1}"
        ws.write_formula(
            row, i + 1,
            s.reg.render(f'=IF({cell}="","",SUMIFS([[Meal Plan.cost]],'
                         f'[[Meal Plan.date]],{cell}))'),
            fmts.calc("money"))
    ws.write_formula(row, 8,
                     s.reg.render('=SUMIFS([[Meal Plan.cost]],'
                                  '[[Meal Plan.date]],">="&DashWeek,'
                                  '[[Meal Plan.date]],"<="&DashWeek+6)'),
                     fmts.chrome("total_value",
                                 num_format=cfg.money_format))
    row += 2

    # ----------------------------------------------------- the year table
    row = s.section(row, "The year so far", LAST_COL)
    row = s.table_head(row, 0, ["Month", "Grocery spend", "Budget",
                                "Difference", "Eating out", "Food waste",
                                "Meals cooked at home", "Trend"],
                       [12, 14, 12, 12, 12, 12, 20, 16])
    first_month = row
    for m in range(MONTHS):
        key = f'TEXT(DATE(PlanYear,{m + 1},1),"yyyy-mm")'
        ws.write_formula(row, 0,
                         f'=TEXT(DATE(PlanYear,{m + 1},1),"mmm")',
                         fmts.calc("text"))
        s.calc_row(row, 1, [
            f'=SUMIFS([[Grocery Spending.amount]],'
            f'[[Grocery Spending.month]],{key})',
            "=MonthlyBudget",
            f'=MonthlyBudget-SUMIFS([[Grocery Spending.amount]],'
            f'[[Grocery Spending.month]],{key})',
            f'=SUMIFS([[Grocery Spending.amount]],'
            f'[[Grocery Spending.month]],{key},'
            f'[[Grocery Spending.type]],"Eating out")+'
            f'SUMIFS([[Grocery Spending.amount]],'
            f'[[Grocery Spending.month]],{key},'
            f'[[Grocery Spending.type]],"Takeaway")',
            f'=SUMIFS([[Food Waste.value]],[[Food Waste.month]],{key})',
            f'=COUNTIFS([[Meal Plan.month]],{key},'
            f'[[Meal Plan.home_cooked]],"Yes")',
        ], ["money", "money", "money", "money", "money", "int"])
        row += 1
    last_month = row - 1
    ws.conditional_format(first_month, 3, last_month, 3,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(first_month, 1, last_month, 1,
                          {"type": "data_bar", "bar_color": "#4E8A63"})
    ws.conditional_format(first_month, 5, last_month, 5,
                          {"type": "data_bar", "bar_color": "#D9703C"})

    # sparklines beside the table
    for col_idx, (src, colour) in enumerate(
            [(1, "#2F5E43"), (5, "#D9703C"), (6, "#4E8A63")]):
        letter = chr(65 + src)
        try:
            ws.add_sparkline(first_month + col_idx, 7, {
                "range": (f"{quote_sheet(NAME)}!{letter}"
                          f"{first_month + 1}:{letter}{last_month + 1}"),
                "type": "column",
                "series_color": colour,
            })
        except Exception:        # pragma: no cover - library differences
            pass
    row += 1

    s.merge(row, 0, LAST_COL,
            "Spend, waste and home cooking by month. The little charts "
            "on the right are, top to bottom: grocery spend, food "
            "waste, meals cooked at home.", fmts.chrome("note"))
    row += 2

    # --------------------------------------------------------- the totals
    row = s.section(row, "The year in one line", LAST_COL)
    row = s.table_head(row, 0, ["Spent on food", "Food budget",
                                "Difference", "Wasted", "Waste as a "
                                "share", "Meals cooked at home",
                                "Recipes in the book", "Pantry value"],
                       [14, 13, 12, 11, 15, 20, 17, 13])
    s.calc_row(row, 0, [
        '=SUM([[Grocery Spending.amount]])',
        "=AnnualBudget",
        '=AnnualBudget-SUM([[Grocery Spending.amount]])',
        '=SUM([[Food Waste.value]])',
        '=IFERROR(SUM([[Food Waste.value]])/'
        'SUM([[Grocery Spending.amount]]),"")',
        '=COUNTIF([[Meal Plan.home_cooked]],"Yes")',
        '=COUNTIF([[Recipes.name]],"<>")',
        '=SUM([[Pantry.value]])',
    ], ["money", "money", "money", "money", "percent", "int", "int",
        "money"])
    ws.conditional_format(row, 4, row, 4,
                          {"type": "cell", "criteria": ">",
                           "value": cfg.waste_reduction_target,
                           "format": fmts.cf("warn")})
    row += 2

    s.note(row, "Everything on this page is live. If a number looks "
                "wrong, the fix is always on the tab named in the "
                "'Where to fix it' column - never here.", LAST_COL,
           height=26)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
