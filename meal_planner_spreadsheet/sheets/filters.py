"""`Dietary Filters` - 'what can I actually make tonight?'

Six drop-downs at the top: diet, category, method, how long you have,
what it can cost a serving and an allergen to avoid. Underneath, every
recipe that passes, with its time, cost, calories and the reason it
suits. It is the answer to the six o'clock question, and it is also how
you cook for the vegetarian, the coeliac and the fussy one at once.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter, list_from, number_check

NAME = "Dietary Filters"
LAST_COL = 9
RESULTS = 25


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_soft)
    ws = s.ws

    s.banner("Dietary Filters",
             "Pick the rules, get the list. Every recipe that fits, "
             "with the time, the cost and the calories.", LAST_COL)
    s.set_col(0, 0, 18)
    s.set_col(1, 1, 24)
    for c in range(2, LAST_COL + 1):
        s.set_col(c, c, 14)

    # ----------------------------------------------------- the controls --
    row = 4
    row = s.section(row, "Set the rules", LAST_COL)
    controls = [
        ("Diet", "FilterDiet", "Any", list_from("DietList"),
         "Vegetarian, vegan, gluten-free... 'Any' turns the filter off."),
        ("Category", "FilterCategory", "Any",
         list_from("RecipeCategoryList"),
         "Mains, soups, lunchboxes, puddings."),
        ("Method", "FilterMethod", "Any",
         list_from("CookingMethodList"),
         "Only the slow cooker? Only no-cook? Choose it here."),
        ("Ready within (minutes)", "FilterTime", 0,
         number_check(0, 600, decimal=False),
         "Total time, prep plus cook. 0 means no limit."),
        ("Cost a serving up to", "FilterCost", 0,
         number_check(0, 1000),
         "0 means no limit. Try your per-meal budget from the Budget "
         "tab."),
        ("Avoid this allergen", "FilterAllergen", "Any",
         list_from("AllergenList"),
         "Hides every recipe with that allergen anywhere in its "
         "ingredients."),
    ]
    for label, name, default, validate, help_text in controls:
        style = "int" if isinstance(default, int) else "text"
        row = s.field(row, label, default, style=style, name=name,
                      validate=validate, help_text=help_text)
    row += 1

    # --------------------------------------------------- what matched ---
    row = s.section(row, "What you can make", LAST_COL)
    s.merge(row, 0, 2, None, fmts.chrome("kpi_value"))
    ws.write_formula(row, 0,
                     s.reg.render('="Recipes that fit: "&'
                                  'SUM([[Recipes.match]])'),
                     fmts.chrome("kpi_value", align="left"))
    s.merge(row, 3, LAST_COL,
            "Blank rows mean nothing else matched - loosen a rule.",
            fmts.chrome("note"))
    row += 2

    row = s.table_head(row, 0, ["#", "Recipe", "Category", "Suits",
                                "Total time", "Cost a serving",
                                "Calories", "Rating", "Last cooked",
                                "Freezes?"],
                       [5, 26, 14, 15, 12, 14, 11, 10, 12, 10])
    first = row
    for n in range(RESULTS):
        s.write(row, 0, n + 1, fmts.chrome("label"))
        pick = (f'MATCH({n + 1},[[Recipes.match_rank]],0)')
        for col, key, style in [(1, "name", "text"),
                                (2, "category", "text"),
                                (3, "diet", "text"),
                                (4, "total_min", "int"),
                                (5, "cost_serving", "money"),
                                (6, "calories", "number"),
                                (7, "rating", "rating"),
                                (8, "last_cooked", "date"),
                                (9, "freezes", "flag")]:
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Recipes.{key}]],'
                             f'{pick}),"")'),
                fmts.calc(style))
        row += 1
    last = row - 1
    ws.conditional_format(first, 4, last, 4,
                          {"type": "cell", "criteria": "<=", "value": 30,
                           "format": fmts.cf("good")})
    ws.conditional_format(first, 5, last, 5,
                          {"type": "data_bar", "bar_color": "#4E8A63"})
    row += 1

    # ---------------------------------------------- who can eat what ----
    row = s.section(row, "Who in the house this suits", LAST_COL)
    row = s.table_head(row, 0, ["Person", "Diet", "Allergies",
                                "Recipes that suit their diet",
                                "Recipes to avoid"],
                       [18, 16, 20, 26, 20])
    for i in range(6):
        cell = f"A{row + 1}"
        ws.write_formula(row, 0,
                         f'=IFERROR(IF(INDEX(MemberNames,{i + 1})=0,"",'
                         f'INDEX(MemberNames,{i + 1})),"")',
                         fmts.calc("text"))
        diet = (f'IFERROR(INDEX([[Household.diet]],MATCH({cell},'
                f'[[Household.name]],0)),"")')
        allergy = (f'IFERROR(INDEX([[Household.allergies]],MATCH({cell},'
                   f'[[Household.name]],0)),"")')
        s.calc_row(row, 1, [
            f'=IF({cell}="","",{diet})',
            f'=IF({cell}="","",{allergy})',
            f'=IF({cell}="","",IF({diet}="No restriction",'
            f'COUNTIF([[Recipes.name]],"<>"),COUNTIFS([[Recipes.diet]],'
            f'{diet})))',
            f'=IF({cell}="","",IF(OR({allergy}="",{allergy}="None"),0,'
            f'SUMPRODUCT((COUNTIFS([[Recipe Ingredients.recipe]],'
            f'[[Recipes.id]],[[Recipe Ingredients.allergen]],'
            f'{allergy})>0)*([[Recipes.name]]<>""))))',
        ], ["text", "text", "int", "int"])
        row += 1
    row += 1

    s.note(row,
           "Allergen filtering looks at the allergen tag on each "
           "ingredient in the price book, so it is only as complete as "
           "those tags. It is a sorting tool, not a safety check: for "
           "a real allergy, read the label on the actual packet every "
           "single time, and remember that 'may contain' warnings "
           "never appear in a spreadsheet.",
           LAST_COL, height=46)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
