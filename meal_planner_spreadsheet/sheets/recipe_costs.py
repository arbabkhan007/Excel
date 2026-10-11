"""`Recipe Costs` - what every meal actually costs to put on the table.

Recipes ranked cheapest first, with the cost a serving, the cost per
100 calories and what the same meal would cost as a takeaway. It is the
tab that settles arguments: the curry you think is expensive is 90p a
head, and the 'cheap' pasta bake is not.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter, number_check

NAME = "Recipe Costs"
LAST_COL = 9
ROWS = 25


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.accent_soft)
    ws = s.ws

    s.banner("Recipe Costing",
             "Every recipe priced from your own shopping, cheapest "
             "first.", LAST_COL)
    s.set_col(0, 0, 6)
    s.set_col(1, 1, 28)
    for c in range(2, LAST_COL + 1):
        s.set_col(c, c, 14)

    # --------------------------------------------------------- the KPIs
    row = 4
    kpis = [
        ("Cheapest meal a serving",
         '=IFERROR(MIN(IFERROR([[Recipes.cost_serving]]*1,999999)),"")',
         "money", "the one to cook when money is tight"),
        ("Average meal a serving",
         '=IFERROR(AVERAGE([[Recipes.cost_serving]]),"")', "money",
         "across every recipe with a price"),
        ("Dearest meal a serving",
         '=IFERROR(MAX(IFERROR([[Recipes.cost_serving]]*1,0)),"")',
         "money", "worth saving for a Friday"),
        ("Recipes still unpriced",
         '=COUNTIFS([[Recipes.name]],"<>",[[Recipes.cost]],0)+'
         'COUNTIFS([[Recipes.name]],"<>",[[Recipes.cost]],"")',
         "int", "add the ingredient prices and they appear"),
    ]
    for i, (label, formula, style, note) in enumerate(kpis):
        s.kpi(row, i * 2 + 1, label, s.reg.render(formula), style=style,
              span=2, note=note)
    row += 4

    # ----------------------------------------------- the cost benchmark
    row = s.section(row, "What you are comparing against", LAST_COL)
    row = s.field(row, "A takeaway for the household costs",
                  round(cfg.eating_out_budget / 4, 2), style="money",
                  name="TakeawayCost",
                  validate=number_check(0, 1000),
                  help_text="Your own figure for one takeaway meal for "
                            "everybody. Used in the saving column "
                            "below.")
    row = s.field(row, "A bought lunch costs", 6.50, style="money",
                  name="BoughtLunch", validate=number_check(0, 200),
                  help_text="Meal deal, café sandwich, canteen - "
                            "whatever you would otherwise pay.")
    row += 1

    # ------------------------------------------------- the ranked table
    row = s.section(row, "Every recipe, cheapest first", LAST_COL)
    row = s.table_head(row, 0, ["#", "Recipe", "Serves", "Cost to make",
                                "Cost a serving", "Per 100 calories",
                                "Minutes", "Saving against a takeaway",
                                "Rating"],
                       [5, 28, 9, 14, 14, 16, 11, 24, 10])
    first = row
    for n in range(ROWS):
        pick = f'MATCH({n + 1},[[Recipes.cost_rank]],0)'
        s.write(row, 0, n + 1, fmts.chrome("label"))
        for col, key, style in [(1, "name", "text"),
                                (2, "servings", "int"),
                                (3, "cost", "money"),
                                (4, "cost_serving", "money"),
                                (6, "total_min", "int"),
                                (8, "rating", "rating")]:
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Recipes.{key}]],{pick}),'
                             f'"")'), fmts.calc(style))
        r = row + 1
        cals = f'IFERROR(INDEX([[Recipes.calories]],{pick}),0)'
        ws.write_formula(
            row, 5,
            s.reg.render(f'=IF(E{r}="","",IFERROR(E{r}/({cals}/100),""))'),
            fmts.calc("money"))
        ws.write_formula(
            row, 7,
            f'=IF(D{r}="","",TakeawayCost-D{r})', fmts.calc("money"))
        row += 1
    last = row - 1
    ws.conditional_format(first, 4, last, 4,
                          {"type": "3_color_scale",
                           "min_color": "#D9F2E4", "mid_color": "#FFF1CF",
                           "max_color": "#FBD9DD"})
    ws.conditional_format(first, 7, last, 7,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("good")})
    ws.conditional_format(first, 3, last, 3,
                          {"type": "data_bar", "bar_color": "#4E8A63"})
    row += 1

    # ------------------------------------------- where the money goes --
    row = s.section(row, "The expensive ingredients", LAST_COL)
    row = s.note(row, "Sorted by what you spend on them across every "
                      "recipe - the five lines at the top are where a "
                      "swap or a cheaper shop pays off most.",
                 LAST_COL, height=24)
    row = s.table_head(row, 0, ["", "Ingredient", "Cost per 100",
                                "Used in recipes", "Total across "
                                "recipes", "Cheapest store",
                                "Saving available"],
                       [5, 26, 14, 15, 20, 18, 17])
    for n in range(cfg.top_rows * 2):
        pick = f'MATCH({n + 1},[[Ingredients.spend_rank]],0)'
        big = f'INDEX([[Ingredients.spend_key]],{pick})'
        s.write(row, 0, n + 1, fmts.chrome("label"))
        for col, key, style in [(1, "name", "text"),
                                (2, "per_100", "money"),
                                (3, "used_in", "int"),
                                (5, "best_store", "text"),
                                (6, "saving", "money")]:
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(IF({big}=0,"",INDEX('
                             f'[[Ingredients.{key}]],{pick})),"")'),
                fmts.calc(style))
        ws.write_formula(
            row, 4,
            s.reg.render(f'=IFERROR(IF({big}=0,"",{big}),"")'),
            fmts.calc("money"))
        row += 1
    row += 1

    # ------------------------------------------------ cost by category
    row = s.section(row, "Cost by kind of meal", LAST_COL)
    row = s.table_head(row, 0, ["Category", "Recipes", "Average a "
                                "serving", "Cheapest", "Dearest",
                                "Average minutes"],
                       [18, 11, 18, 13, 13, 16])
    for kind in ["Main", "Side", "Soup", "Salad", "Breakfast", "Snack",
                 "Dessert", "Lunchbox", "Batch base"]:
        s.write(row, 0, kind, fmts.chrome("label"))
        s.calc_row(row, 1, [
            f'=COUNTIFS([[Recipes.category]],"{kind}")',
            f'=IFERROR(AVERAGEIFS([[Recipes.cost_serving]],'
            f'[[Recipes.category]],"{kind}"),"")',
            f'=IFERROR(IF(COUNTIFS([[Recipes.category]],"{kind}")=0,"",'
            f'SUMPRODUCT(MIN((([[Recipes.category]]="{kind}")*'
            f'IFERROR([[Recipes.cost_serving]]*1,0)>0)*'
            f'IFERROR([[Recipes.cost_serving]]*1,0)+'
            f'(1-(([[Recipes.category]]="{kind}")*'
            f'IFERROR([[Recipes.cost_serving]]*1,0)>0))*9999))),"")',
            f'=IFERROR(SUMPRODUCT(MAX(([[Recipes.category]]="{kind}")*'
            f'IFERROR([[Recipes.cost_serving]]*1,0))),"")',
            f'=IFERROR(AVERAGEIFS([[Recipes.total_min]],'
            f'[[Recipes.category]],"{kind}"),"")',
        ], ["int", "money", "money", "money", "int"])
        row += 1
    row += 1

    s.note(row,
           "A recipe is only as well priced as its ingredient lines. "
           "Anything flagged on the Recipe Ingredients tab - a missing "
           "price, a unit that does not match the price book - is left "
           "out of these figures rather than guessed at, so a "
           "suspiciously cheap recipe usually means a missing line "
           "rather than a bargain.",
           LAST_COL, height=44)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
