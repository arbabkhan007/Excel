"""`Sustainability` - the waste report, and what to do about it.

Where the binned food comes from, what it costs a year, which reasons
keep repeating, and a handful of changes that actually work. Reducing
food waste is the one 'green' thing that also puts money straight back
in your pocket, which is why it has a tab of its own.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Sustainability"
LAST_COL = 8

REASONS = ["Expired", "Went mouldy", "Cooked too much",
           "Did not like it", "Forgot about it", "Freezer burn",
           "Left out too long", "Bought too much", "Plate scrapings",
           "Other"]

HABITS = [
    ("Shop your own cupboards first",
     "The Pantry deduction on the grocery list does this for you - "
     "the saving only lands if you keep the pantry roughly honest."),
    ("Freeze half on the day you buy it",
     "Bread, milk, cooked rice, herbs, cheese. The freezer window on "
     "the Settings tab tracks it."),
    ("Plan one 'use it up' meal a week",
     "Soup, fried rice, frittata, traybake. Add it to the Meal Plan "
     "with no recipe and the ingredients come out of the fridge "
     "rather than the shop."),
    ("Buy the smaller bag of salad",
     "The big bag is cheaper per 100 g only if you eat all of it. "
     "Check the Food Waste log before you decide."),
    ("Use the Leftovers tab as a menu",
     "Portions with a value attached get eaten; mystery boxes get "
     "binned."),
    ("Keep a 'eat me first' shelf",
     "One shelf, everything close to its date. Costs nothing, works "
     "better than any app."),
    ("Cook from the freezer once a week",
     "The Fridge & Freezer page sorts it by how long it has been in "
     "there."),
]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_soft)
    ws = s.ws

    s.banner("Food Waste & Sustainability",
             "What goes in the bin, what it costs, and the handful of "
             "habits that change it.", LAST_COL)
    s.set_col(0, 0, 26)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 15)

    row = 4
    kpis = [
        ("Wasted this month",
         '=SUMIFS([[Food Waste.value]],[[Food Waste.month]],'
         'TEXT(TODAY(),"yyyy-mm"))', "money",
         f"target under {cfg.currency_symbol}"
         f"{cfg.waste_cost_target:,.0f}"),
        ("A year at this rate",
         '=SUMIFS([[Food Waste.value]],[[Food Waste.month]],'
         'TEXT(TODAY(),"yyyy-mm"))*12', "money",
         "the number worth staring at"),
        ("Avoidable share",
         '=IFERROR(SUMIFS([[Food Waste.value]],'
         '[[Food Waste.avoidable]],"Yes")/SUM([[Food Waste.value]]),"")',
         "percent0", "peelings and bones are not avoidable"),
        ("Waste as a share of the shop",
         '=IFERROR(SUM([[Food Waste.value]])/'
         'SUM([[Grocery Spending.amount]]),"")', "percent0",
         f"target under {cfg.waste_reduction_target:.0%}"),
    ]
    for i, (label, formula, style, note) in enumerate(kpis):
        s.kpi(row, i * 2, label, s.reg.render(formula), style=style,
              span=2, note=note)
    row += 4

    # ------------------------------------------------------ by reason --
    row = s.section(row, "Why food gets thrown away", LAST_COL)
    row = s.table_head(row, 0, ["Reason", "Times", "Cost",
                                "Share of the waste", "Avoidable?",
                                "Most common item"],
                       [24, 10, 13, 18, 12, 24])
    first = row
    for reason in REASONS:
        s.write(row, 0, reason, fmts.chrome("label"))
        s.calc_row(row, 1, [
            f'=COUNTIFS([[Food Waste.reason]],"{reason}")',
            f'=SUMIFS([[Food Waste.value]],[[Food Waste.reason]],'
            f'"{reason}")',
            f'=IFERROR(SUMIFS([[Food Waste.value]],'
            f'[[Food Waste.reason]],"{reason}")/'
            f'SUM([[Food Waste.value]]),"")',
            f'=IF(COUNTIFS([[Food Waste.reason]],"{reason}")=0,"",'
            f'IF(COUNTIFS([[Food Waste.reason]],"{reason}",'
            f'[[Food Waste.avoidable]],"Yes")>0,"Yes","No"))',
            f'=IFERROR(INDEX([[Food Waste.item]],MATCH("{reason}",'
            f'[[Food Waste.reason]],0)),"")',
        ], ["int", "money", "percent0", "flag", "text"])
        row += 1
    ws.conditional_format(first, 3, row - 1, 3,
                          {"type": "data_bar", "bar_color": "#D9703C"})
    row += 1

    # ------------------------------------------------- worst offenders
    row = s.section(row, "The things you bin most", LAST_COL)
    row = s.table_head(row, 0, ["", "Item", "Times binned",
                                "Cost so far", "Still buying it?",
                                "What would stop it"],
                       [5, 24, 14, 14, 16, 34])
    for n in range(cfg.top_rows):
        pick = f'MATCH({n + 1},[[Ingredients.waste_rank]],0)'
        big = f'INDEX([[Ingredients.waste_count]],{pick})'
        s.write(row, 0, n + 1, fmts.chrome("label"))
        ws.write_formula(
            row, 1,
            s.reg.render(f'=IFERROR(IF({big}=0,"",INDEX('
                         f'[[Ingredients.name]],{pick})),"")'),
            fmts.calc("text"))
        ws.write_formula(
            row, 2, s.reg.render(f'=IFERROR(IF({big}=0,"",{big}),"")'),
            fmts.calc("int"))
        item = f"B{row + 1}"
        s.calc_row(row, 3, [
            f'=IF({item}="","",SUMIFS([[Food Waste.value]],'
            f'[[Food Waste.item]],{item}))',
            f'=IF({item}="","",IF(COUNTIFS([[Grocery List.ingredient]],'
            f'{item})>0,"On this week\'s list","Not this week"))',
            f'=IF({item}="","",IFERROR(INDEX([[Food Waste.prevent]],'
            f'MATCH({item},[[Food Waste.item]],0)),""))',
        ], ["money", "text", "text"])
        row += 1
    row += 1

    # ---------------------------------------------------- disposal ----
    row = s.section(row, "Where it ends up", LAST_COL)
    row = s.table_head(row, 0, ["Disposal route", "Times", "Cost",
                                "Share"], [24, 12, 14, 14])
    for route in ["General bin", "Food waste caddy", "Compost",
                  "Fed to pets", "Given away"]:
        s.write(row, 0, route, fmts.chrome("label"))
        s.calc_row(row, 1, [
            f'=COUNTIFS([[Food Waste.disposal]],"{route}")',
            f'=SUMIFS([[Food Waste.value]],[[Food Waste.disposal]],'
            f'"{route}")',
            f'=IFERROR(COUNTIFS([[Food Waste.disposal]],"{route}")/'
            f'COUNTIF([[Food Waste.date]],">0"),"")',
        ], ["int", "money", "percent0"])
        row += 1
    row += 1

    # -------------------------------------------------- what to change
    row = s.section(row, "Seven things that actually work", LAST_COL)
    for habit, detail in HABITS:
        s.write(row, 0, habit, fmts.chrome("label_bold"))
        s.merge(row, 1, LAST_COL, detail, fmts.chrome("note"))
        s.row_height(row, 28)
        row += 1
    row += 1

    # ---------------------------------------------------- seasonality -
    row = s.section(row, "Eating with the seasons", LAST_COL)
    row = s.table_head(row, 0, ["Season", "Ingredients tagged",
                                "Recipes tagged", "On the list now"],
                       [16, 20, 18, 18])
    for season in ["All year", "Spring", "Summer", "Autumn", "Winter"]:
        s.write(row, 0, season, fmts.chrome("label"))
        s.calc_row(row, 1, [
            f'=COUNTIFS([[Ingredients.season]],"{season}")',
            f'=COUNTIFS([[Recipes.season]],"{season}")',
            f'=SUMPRODUCT(([[Grocery List.ingredient]]<>"")*'
            f'(COUNTIFS([[Ingredients.name]],'
            f'[[Grocery List.ingredient]],[[Ingredients.season]],'
            f'"{season}")>0))',
        ], ["int", "int", "int"])
        row += 1
    row += 1

    s.note(row,
           "The honest version: a spreadsheet cannot measure your "
           "carbon footprint, and anyone who says their meal planner "
           "can is guessing. What this page does measure is money in "
           "the bin, which is the part you control and the part that "
           "changes behaviour. Log a month properly and the pattern "
           "is usually obvious - and usually three or four items.",
           LAST_COL, height=46)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
