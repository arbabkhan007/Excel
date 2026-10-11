"""`Nutrition` - calories and macros, without the guilt.

Everything here is worked out from the recipes you planned and the
nutrition you copied off the packets: calories, protein, carbs, fat and
fibre a head, day by day, against the targets on the Settings tab. It
is a guide to what you are eating, not medical advice, and it is only
as accurate as the numbers in the price book.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Nutrition"
LAST_COL = 9

MACROS = [
    ("Calories", "cals_total", "CalorieTarget", "number", "kcal"),
    ("Protein", "protein", "ProteinTarget", "number1", "g"),
    ("Carbs", "carbs", "CarbTarget", "number1", "g"),
    ("Fat", "fat", "FatTarget", "number1", "g"),
    ("Fibre", "fibre", "FibreTarget", "number1", "g"),
]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary_mid)
    ws = s.ws

    s.banner("Calories & Macros",
             "What the plan adds up to, a head, a day - measured "
             "against your own targets.", LAST_COL)
    s.set_col(0, 0, 18)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 13)

    # ------------------------------------------------------ the targets --
    row = 4
    row = s.section(row, "Your daily targets", LAST_COL)
    row = s.table_head(row, 0, ["Calories", "Protein (g)", "Carbs (g)",
                                "Fat (g)", "Fibre (g)", "Sugar (g)",
                                "Sodium (mg)", "Water (glasses)"],
                       [12, 12, 12, 11, 12, 12, 13, 15])
    row = s.calc_row(row, 0, [
        "=CalorieTarget", "=ProteinTarget", "=CarbTarget", "=FatTarget",
        "=FibreTarget", "=SugarLimit", "=SodiumLimit", "=WaterTarget",
    ], ["number", "number1", "number1", "number1", "number1",
        "number1", "number", "int"])
    row = s.note(row, "Set these on the Settings tab. They are one "
                      "set of numbers for the household - the "
                      "Household tab holds a calorie figure per person "
                      "if you want to be more precise.", LAST_COL,
                 height=26)

    # --------------------------------------------------- the week ------
    row = s.section(row, "This week, a head, a day", LAST_COL)
    head = row
    s.write(head, 0, "", fmts.chrome("header_left"))
    for i in range(7):
        ws.write_formula(head, i + 1, f"=WeekPick+{i}",
                         fmts.chrome("header", num_format="ddd d"))
    s.write(head, 8, "Average", fmts.chrome("header"))
    s.row_height(head, 24)
    row = head + 1

    per_head = "MAX(1,DefaultServings)"
    for label, key, target, style, unit in MACROS:
        s.write(row, 0, f"{label} ({unit})", fmts.chrome("label_bold"))
        for i in range(7):
            cell = f"{chr(66 + i)}${head + 1}"
            if key == "cals_total":
                total = (f'SUMIFS([[Meal Plan.cals_total]],'
                         f'[[Meal Plan.date]],{cell})')
            else:
                total = (f'SUMPRODUCT(([[Meal Plan.date]]={cell})*'
                         f'IFERROR([[Meal Plan.{key}]]*1,0)*'
                         f'IFERROR([[Meal Plan.eff_servings]]*1,0))')
            ws.write_formula(
                row, i + 1,
                s.reg.render(f'=IF({cell}="","",IFERROR({total}/'
                             f'{per_head},""))'),
                fmts.calc(style))
        letter_first, letter_last = "B", "H"
        ws.write_formula(row, 8,
                         f'=IFERROR(AVERAGE({letter_first}{row + 1}:'
                         f'{letter_last}{row + 1}),"")',
                         fmts.chrome("total_value"))
        ws.conditional_format(row, 1, row, 7,
                              {"type": "formula",
                               "criteria": f'=AND(B{row + 1}<>"",'
                                           f'B{row + 1}>{target}*1.15)',
                               "format": fmts.cf("warn")})
        ws.conditional_format(row, 1, row, 7,
                              {"type": "formula",
                               "criteria": f'=AND(B{row + 1}<>"",'
                                           f'B{row + 1}>={target}*0.85,'
                                           f'B{row + 1}<={target}*1.15)',
                               "format": fmts.cf("good")})
        row += 1

    s.write(row, 0, "Meals logged", fmts.chrome("label_bold"))
    for i in range(7):
        cell = f"{chr(66 + i)}${head + 1}"
        ws.write_formula(
            row, i + 1,
            s.reg.render(f'=IF({cell}="","",COUNTIFS([[Meal Plan.date]],'
                         f'{cell}))'), fmts.calc("int"))
    ws.write_formula(row, 8, f'=SUM(B{row + 1}:H{row + 1})',
                     fmts.chrome("total_value"))
    row += 2

    # -------------------------------------------------- against target --
    row = s.section(row, "How the week compares", LAST_COL)
    row = s.table_head(row, 0, ["", "Average a day", "Target",
                                "Difference", "% of target",
                                "Verdict"],
                       [18, 16, 13, 14, 14, 26])
    first_cmp = row
    for idx, (label, key, target, style, unit) in enumerate(MACROS):
        src = f"I{head + 2 + idx}"
        s.write(row, 0, f"{label} ({unit})", fmts.chrome("label_bold"))
        s.calc_row(row, 1, [
            f"={src}",
            f"={target}",
            f'=IFERROR({src}-{target},"")',
            f'=IFERROR({src}/{target},"")',
            f'=IF({src}="","No data yet",IF({src}<{target}*0.85,'
            f'"Under - add a bit more",IF({src}>{target}*1.15,'
            f'"Over - worth a look","About right")))',
        ], ["number1", "number1", "number1", "percent0", "text"])
        row += 1
    ws.conditional_format(first_cmp, 5, row - 1, 5,
                          {"type": "text", "criteria": "containing",
                           "value": "About right",
                           "format": fmts.cf("good")})
    ws.conditional_format(first_cmp, 5, row - 1, 5,
                          {"type": "text", "criteria": "containing",
                           "value": "Over",
                           "format": fmts.cf("warn")})
    ws.conditional_format(first_cmp, 5, row - 1, 5,
                          {"type": "text", "criteria": "containing",
                           "value": "Under",
                           "format": fmts.cf("info")})
    ws.conditional_format(first_cmp, 4, row - 1, 4,
                          {"type": "data_bar", "bar_color": "#4E8A63"})
    row += 1

    # ------------------------------------------------ by meal and person
    row = s.section(row, "Where the calories come from", LAST_COL)
    row = s.table_head(row, 0, ["Meal", "Calories a week",
                                "Share of the week", "Meals planned",
                                "Average a meal"],
                       [18, 17, 17, 15, 16])
    week = ('[[Meal Plan.date]],">="&WeekPick,[[Meal Plan.date]],'
            '"<="&WeekPick+6')
    for slot in ["Breakfast", "Lunch", "Dinner", "Dessert"]:
        s.write(row, 0, slot, fmts.chrome("label"))
        s.calc_row(row, 1, [
            f'=SUMIFS([[Meal Plan.cals_total]],{week},'
            f'[[Meal Plan.slot]],"{slot}")',
            f'=IFERROR(SUMIFS([[Meal Plan.cals_total]],{week},'
            f'[[Meal Plan.slot]],"{slot}")/SUMIFS('
            f'[[Meal Plan.cals_total]],{week}),"")',
            f'=COUNTIFS({week},[[Meal Plan.slot]],"{slot}")',
            f'=IFERROR(SUMIFS([[Meal Plan.cals_total]],{week},'
            f'[[Meal Plan.slot]],"{slot}")/COUNTIFS({week},'
            f'[[Meal Plan.slot]],"{slot}"),"")',
        ], ["number", "percent0", "int", "number"])
        row += 1
    row += 1

    # ------------------------------------------------- top recipes -----
    row = s.section(row, "The recipes doing the heavy lifting", LAST_COL)
    row = s.table_head(row, 0, ["Highest calorie recipes", "A serving",
                                "Most protein", "A serving",
                                "Most fibre", "A serving"],
                       [26, 12, 22, 12, 22, 12])
    for n in range(cfg.top_rows):
        for col, (value_key, label_col) in enumerate(
                [("calories", 0), ("protein", 2), ("fibre", 4)]):
            pick = (f'MATCH({n + 1},[[Recipes.{value_key}_rank]],0)')
            big = f'INDEX([[Recipes.{value_key}]],{pick})'
            s.write(row, label_col, None, fmts.calc("text"))
            ws.write_formula(
                row, label_col,
                s.reg.render(f'=IFERROR(IF({big}=0,"",INDEX('
                             f'[[Recipes.name]],{pick})),"")'),
                fmts.calc("text"))
            ws.write_formula(
                row, label_col + 1,
                s.reg.render(f'=IFERROR(IF({big}=0,"",{big}),"")'),
                fmts.calc("number1"))
        row += 1
    row += 1

    s.note(row,
           "Honest limits: these numbers come from the 'per 100' "
           "figures you typed into the Ingredients tab, multiplied by "
           "the quantities in your recipes. Cooking losses, oil left "
           "in the pan, who actually ate the bigger portion - none of "
           "that is in here. Treat it as a direction of travel, not a "
           "food diary, and never as medical or dietary advice.",
           LAST_COL, height=50)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
