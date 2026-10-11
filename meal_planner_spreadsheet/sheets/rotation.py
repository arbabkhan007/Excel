"""`Rotation` - stop eating the same six dinners.

Which recipes you lean on, which you have not cooked for months, how
much variety the last four weeks actually had, and a shortlist of
things to bring back. Nobody needs a hundred recipes - they need the
twenty they like, in an order that does not feel like a loop.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Rotation"
LAST_COL = 8
ROWS = 20


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_soft)
    ws = s.ws

    s.banner("Meal Rotation & Variety",
             "What you cook, how often, and what has been forgotten at "
             "the back of the recipe book.", LAST_COL)
    s.set_col(0, 0, 6)
    s.set_col(1, 1, 30)
    for c in range(2, LAST_COL + 1):
        s.set_col(c, c, 15)

    # ------------------------------------------------------------ KPIs
    row = 4
    kpis = [
        ("Recipes in the book",
         '=COUNTIF([[Recipes.name]],"<>")', "int",
         "everything you could cook"),
        ("Cooked in the last 30 days",
         '=SUMPRODUCT((COUNTIFS([[Meal Plan.recipe]],[[Recipes.id]],'
         '[[Meal Plan.date]],">="&TODAY()-30,[[Meal Plan.date]],'
         '"<="&TODAY())>0)*([[Recipes.name]]<>""))', "int",
         "the ones in active service"),
        ("Never cooked",
         '=COUNTIFS([[Recipes.name]],"<>",[[Recipes.times_planned]],0)',
         "int", "worth a go, or worth deleting"),
        ("Variety score",
         '=IFERROR(SUMPRODUCT((COUNTIFS([[Meal Plan.recipe]],'
         '[[Recipes.id]],[[Meal Plan.date]],">="&TODAY()-30,'
         '[[Meal Plan.date]],"<="&TODAY())>0)*'
         '([[Recipes.name]]<>""))/MAX(1,COUNTIFS([[Meal Plan.date]],'
         '">="&TODAY()-30,[[Meal Plan.date]],"<="&TODAY(),'
         '[[Meal Plan.recipe]],"<>")),"")', "percent0",
         "different recipes / meals cooked"),
    ]
    for i, (label, formula, style, note) in enumerate(kpis):
        s.kpi(row, i * 2, label, s.reg.render(formula), style=style,
              span=2, note=note)
    row += 4
    row = s.note(row, "A variety score of 100% means every meal in the "
                      "last month was a different recipe; 25% means "
                      "you cooked the same four things. Neither is "
                      "right or wrong - but if it is low and you are "
                      "bored, the list underneath is where to look.",
                 LAST_COL, height=30)

    # ------------------------------------------- longest untouched ----
    row = s.section(row, "Not cooked for the longest", LAST_COL)
    row = s.table_head(row, 0, ["#", "Recipe", "Category", "Last cooked",
                                "Days ago", "Rating", "Cost a serving",
                                "Minutes"],
                       [5, 30, 15, 14, 12, 10, 15, 11])
    first = row
    for n in range(ROWS):
        pick = f'MATCH({n + 1},[[Recipes.rotation_rank]],0)'
        s.write(row, 0, n + 1, fmts.chrome("label"))
        for col, key, style in [(1, "name", "text"),
                                (2, "category", "text"),
                                (3, "last_cooked", "date"),
                                (5, "rating", "rating"),
                                (6, "cost_serving", "money"),
                                (7, "total_min", "int")]:
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Recipes.{key}]],{pick}),'
                             f'"")'), fmts.calc(style))
        ws.write_formula(
            row, 4,
            s.reg.render(f'=IFERROR(IF(INDEX([[Recipes.last_cooked]],'
                         f'{pick})="","Never",TODAY()-INDEX('
                         f'[[Recipes.last_cooked]],{pick})),"")'),
            fmts.calc("int"))
        row += 1
    last = row - 1
    ws.conditional_format(first, 4, last, 4,
                          {"type": "text", "criteria": "containing",
                           "value": "Never", "format": fmts.cf("info")})
    ws.conditional_format(first, 5, last, 5,
                          {"type": "cell", "criteria": ">=", "value": 4,
                           "format": fmts.cf("good")})
    row += 1

    # ------------------------------------------------- most cooked ----
    row = s.section(row, "Your greatest hits", LAST_COL)
    row = s.table_head(row, 0, ["", "Recipe", "Times planned",
                                "Times cooked", "Rating",
                                "Cost a serving", "Last cooked"],
                       [5, 30, 15, 14, 10, 15, 14])
    for n in range(cfg.top_rows * 2):
        pick = f'MATCH({n + 1},[[Recipes.times_planned_rank]],0)'
        big = f'INDEX([[Recipes.times_planned]],{pick})'
        s.write(row, 0, n + 1, fmts.chrome("label"))
        for col, key, style in [(1, "name", "text"),
                                (3, "times_cooked", "int"),
                                (4, "rating", "rating"),
                                (5, "cost_serving", "money"),
                                (6, "last_cooked", "date")]:
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(IF({big}=0,"",INDEX('
                             f'[[Recipes.{key}]],{pick})),"")'),
                fmts.calc(style))
        ws.write_formula(row, 2,
                         s.reg.render(f'=IFERROR(IF({big}=0,"",{big}),"")'),
                         fmts.calc("int"))
        row += 1
    row += 1

    # ------------------------------------------------ variety by type -
    row = s.section(row, "Variety by cuisine and method", LAST_COL)
    row = s.table_head(row, 0, ["Cuisine", "Recipes", "Meals planned",
                                "Share of meals", "", "Method",
                                "Recipes", "Meals planned"],
                       [16, 10, 14, 14, 3, 16, 10, 14])
    cuisines = ["British", "Italian", "Indian", "Mexican", "Chinese",
                "Japanese", "Middle Eastern", "Thai", "American",
                "Other"]
    methods = ["Oven", "Hob", "One pot", "Slow cooker", "Air fryer",
               "No cook", "Grill", "Pressure cooker", "Microwave",
               "Barbecue"]
    for cuisine, method in zip(cuisines, methods):
        s.write(row, 0, cuisine, fmts.chrome("label"))
        s.calc_row(row, 1, [
            f'=COUNTIFS([[Recipes.cuisine]],"{cuisine}")',
            f'=SUMPRODUCT(([[Recipes.cuisine]]="{cuisine}")*'
            f'IFERROR([[Recipes.times_planned]]*1,0))',
            f'=IFERROR(SUMPRODUCT(([[Recipes.cuisine]]="{cuisine}")*'
            f'IFERROR([[Recipes.times_planned]]*1,0))/'
            f'MAX(1,SUM(IFERROR([[Recipes.times_planned]]*1,0))),"")',
        ], ["int", "int", "percent0"])
        s.write(row, 5, method, fmts.chrome("label"))
        s.calc_row(row, 6, [
            f'=COUNTIFS([[Recipes.method]],"{method}")',
            f'=SUMPRODUCT(([[Recipes.method]]="{method}")*'
            f'IFERROR([[Recipes.times_planned]]*1,0))',
        ], ["int", "int"])
        row += 1
    row += 1

    # ------------------------------------------------ repeat warning --
    row = s.section(row, "Cooked twice inside the rotation gap", LAST_COL)
    row = s.note(row, f"Your rotation gap is set to "
                      f"{cfg.rotation_gap_days} days on the Settings "
                      f"tab. These recipes have come round faster than "
                      f"that - fine if they are favourites, worth "
                      f"swapping if they are just habit.",
                 LAST_COL, height=28)
    row = s.table_head(row, 0, ["Recipe", "Days since", "Times planned",
                                "Rating", "Suggested swap"],
                       [30, 13, 14, 10, 30])
    for n in range(cfg.top_rows):
        s.write(row, 0, None, fmts.calc("text"))
        rev = f'COUNTIF([[Recipes.name]],"<>")-{n}'
        ws.write_formula(
            row, 0,
            s.reg.render(f'=IFERROR(INDEX([[Recipes.name]],MATCH({rev},'
                         f'[[Recipes.rotation_rank]],0)),"")'),
            fmts.calc("text"))
        for col, key, style in [(1, "days_since", "int"),
                                (2, "times_planned", "int"),
                                (3, "rating", "rating")]:
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Recipes.{key}]],'
                             f'MATCH({rev},[[Recipes.rotation_rank]],0)),'
                             f'"")'), fmts.calc(style))
        ws.write_formula(
            row, 4,
            s.reg.render(f'=IFERROR(INDEX([[Recipes.name]],MATCH({n + 1},'
                         f'[[Recipes.rotation_rank]],0)),"")'),
            fmts.calc("text"))
        row += 1
    row += 1

    s.note(row, "The 'suggested swap' column is deliberately simple: "
                "it offers the recipe you have left alone the longest. "
                "A spreadsheet cannot know that nobody fancies fish on "
                "a Monday.", LAST_COL, height=26)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
