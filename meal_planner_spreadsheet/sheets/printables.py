"""The three printable pages.

* `Print Week` - the week's meals on one sheet of A4 for the fridge
  door, with the prep jobs and the leftovers to eat.
* `Print Shopping` - the grocery list grouped by aisle with tick boxes
  and a running total, sized for a pocket.
* `Print Recipe` - one recipe as a card: ingredients scaled to the
  number of people you choose, method, cost and nutrition.

All three read from the same registers as everything else, so they are
never a stale copy.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, list_from, number_check

WEEK_NAME = "Print Week"
SHOP_NAME = "Print Shopping"
RECIPE_NAME = "Print Recipe"


# ---------------------------------------------------------------------------
# Print Week
# ---------------------------------------------------------------------------

def build_week(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    NAME, LAST_COL = WEEK_NAME, 5
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_soft)
    ws = s.ws

    s.banner("This Week's Meals", "", LAST_COL, tail_text="")
    s.set_col(0, 0, 14)
    s.set_col(1, 1, 34)
    s.set_col(2, 2, 30)
    s.set_col(3, 3, 12)
    s.set_col(4, 5, 14)

    row = 4
    s.merge(row, 0, 1, None, fmts.chrome("print_title"))
    ws.write_formula(row, 0,
                     '=HouseName&" - week of "&TEXT(WeekPick,'
                     '"d mmmm yyyy")', fmts.chrome("print_title"))
    s.merge(row, 2, LAST_COL, None, fmts.chrome("print_sub"))
    ws.write_formula(row, 2,
                     s.reg.render('="Planned cost "&TEXT(SUMIFS('
                                  '[[Meal Plan.cost]],[[Meal Plan.date]],'
                                  '">="&WeekPick,[[Meal Plan.date]],'
                                  '"<="&WeekPick+6),"0.00")&'
                                  '"   |   Shopping day: "&ShoppingDay&'
                                  '"   |   Prep day: "&PrepDay'),
                     fmts.chrome("print_sub"))
    row += 2

    row = s.table_head(row, 0, ["Day", "Dinner", "Lunch", "Cost",
                                "Who is cooking", "Done"],
                       [14, 34, 30, 12, 16, 10])
    first = row
    for i in range(7):
        day = f"A{row + 1}"
        ws.write_formula(row, 0, f"=WeekPick+{i}",
                         fmts.chrome("label_bold", num_format="ddd d mmm"))
        for col, slot in ((1, "Dinner"), (2, "Lunch")):
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Meal Plan.recipe_name]],'
                             f'MATCH(TEXT({day},"yyyy-mm-dd")&"|{slot}",'
                             f'[[Meal Plan.key]],0)),"")'),
                fmts.cell("wrap", "plain"))
        ws.write_formula(
            row, 3,
            s.reg.render(f'=IF(SUMIFS([[Meal Plan.cost]],'
                         f'[[Meal Plan.date]],{day})=0,"",SUMIFS('
                         f'[[Meal Plan.cost]],[[Meal Plan.date]],'
                         f'{day}))'),
            fmts.cell("money", "plain"))
        ws.write_formula(
            row, 4,
            s.reg.render(f'=IFERROR(INDEX([[Meal Plan.cook]],MATCH('
                         f'TEXT({day},"yyyy-mm-dd")&"|Dinner",'
                         f'[[Meal Plan.key]],0)),"")'),
            fmts.cell("text", "plain"))
        s.write(row, 5, None, fmts.chrome("form_box"))
        s.row_height(row, 30)
        row += 1
    s.merge(row, 0, 2, "Week total", fmts.chrome("total_label"))
    ws.write_formula(row, 3,
                     f"=SUM(D{first + 1}:D{row})",
                     fmts.chrome("total_value",
                                 num_format=cfg.money_format))
    s.merge(row, 4, LAST_COL, None, fmts.chrome("total_value"))
    row += 2

    row = s.section(row, "Before the week starts", LAST_COL)
    row = s.table_head(row, 0, ["Prep job", "When", "Portions",
                                "Who", "Done"], [34, 14, 12, 16, 10])
    for n in range(5):
        pick = f'MATCH({n + 1},[[Meal Prep.todo_rank]],0)'
        for col, key, style in [(0, "job", "text"), (1, "date", "date"),
                                (2, "portions", "number1"),
                                (3, "cook", "text")]:
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Meal Prep.{key}]],'
                             f'{pick}),"")'), fmts.cell(style, "plain"))
        s.write(row, 4, None, fmts.chrome("form_box"))
        row += 1
    row += 1

    row = s.section(row, "Eat these first", LAST_COL)
    row = s.table_head(row, 0, ["Leftover or item", "Where", "Eat by",
                                "Portions", "Done"],
                       [34, 14, 14, 12, 10])
    for n in range(5):
        pick = f'MATCH({n + 1},[[Leftovers.queue_rank]],0)'
        for col, key, style in [(0, "item", "text"),
                                (1, "location", "text"),
                                (2, "use_by", "date"),
                                (3, "portions", "number1")]:
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Leftovers.{key}]],'
                             f'{pick}),"")'), fmts.cell(style, "plain"))
        s.write(row, 4, None, fmts.chrome("form_box"))
        row += 1
    row += 1

    s.merge(row, 0, LAST_COL, None, fmts.chrome("print_sub"))
    ws.write_formula(row, 0,
                     f'="{cfg.product_name}  |  "&HouseName&"  |  "&'
                     f'TEXT(TODAY(),"d mmm yyyy")',
                     fmts.chrome("print_sub"))

    s.finish(freeze_row=None, print_area=f"A1:F{row + 1}",
             gridlines=False, fit_width=True)
    return s


# ---------------------------------------------------------------------------
# Print Shopping
# ---------------------------------------------------------------------------

def build_shopping(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    NAME, LAST_COL = SHOP_NAME, 5
    ROWS = 45
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.accent)
    ws = s.ws

    s.banner("Shopping List", "", LAST_COL, tail_text="")
    s.set_col(0, 0, 6)
    s.set_col(1, 1, 30)
    s.set_col(2, 2, 18)
    s.set_col(3, 3, 14)
    s.set_col(4, 4, 12)
    s.set_col(5, 5, 10)

    row = 4
    s.merge(row, 0, 2, None, fmts.chrome("print_title"))
    ws.write_formula(row, 0,
                     '=HouseName&" - "&TEXT(ShopFrom,"d mmm")&" to "&'
                     'TEXT(ShopTo,"d mmm")', fmts.chrome("print_title"))
    s.merge(row, 3, LAST_COL, None, fmts.chrome("print_sub"))
    ws.write_formula(row, 3,
                     s.reg.render('="Estimated "&TEXT(SUM('
                                  '[[Grocery List.est_cost]]),"0.00")&'
                                  '"  |  budget "&TEXT(WeeklyBudget,'
                                  '"0.00")'),
                     fmts.chrome("print_sub"))
    row += 2

    row = s.table_head(row, 0, ["Got", "Item", "Aisle", "How much",
                                "Cost", "Store"],
                       [6, 30, 18, 14, 12, 14])
    first = row
    for n in range(ROWS):
        pick = f'MATCH({n + 1},[[Grocery List.buy_rank]],0)'
        s.write(row, 0, None, fmts.chrome("form_box"))
        for col, key, style in [(1, "ingredient", "text"),
                                (2, "aisle", "text"),
                                (4, "est_cost", "money"),
                                (5, "store", "text")]:
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Grocery List.{key}]],'
                             f'{pick}),"")'), fmts.cell(style, "plain"))
        ws.write_formula(
            row, 3,
            s.reg.render(f'=IFERROR(IF(INDEX([[Grocery List.buy_qty]],'
                         f'{pick})="","",ROUND(INDEX('
                         f'[[Grocery List.buy_qty]],{pick}),0)&" "&'
                         f'INDEX([[Grocery List.unit]],{pick})),"")'),
            fmts.cell("text", "plain"))
        s.row_height(row, 18)
        row += 1
    s.merge(row, 0, 3, "Total", fmts.chrome("total_label"))
    ws.write_formula(row, 4, f"=SUM(E{first + 1}:E{row})",
                     fmts.chrome("total_value",
                                 num_format=cfg.money_format))
    s.merge(row, 5, LAST_COL, None, fmts.chrome("total_value"))
    row += 2

    row = s.section(row, "Extras", LAST_COL)
    for n in range(6):
        s.write(row, 0, None, fmts.chrome("form_box"))
        s.merge(row, 1, 3, None, fmts.chrome("form_line"))
        s.merge(row, 4, LAST_COL, None, fmts.chrome("form_line"))
        s.row_height(row, 20)
        row += 1
    row += 1

    s.merge(row, 0, LAST_COL,
            "Printed from the Grocery List tab - change the shopping "
            "window on Settings and print again.",
            fmts.chrome("print_sub"))

    s.finish(freeze_row=None, print_area=f"A1:F{row + 1}",
             gridlines=False, fit_width=True)
    return s


# ---------------------------------------------------------------------------
# Print Recipe
# ---------------------------------------------------------------------------

def build_recipe(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    NAME, LAST_COL = RECIPE_NAME, 5
    LINES = 16
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.accent_soft)
    ws = s.ws

    s.banner("Recipe Card", "", LAST_COL, tail_text="")
    s.set_col(0, 0, 26)
    s.set_col(1, 1, 14)
    s.set_col(2, 2, 12)
    s.set_col(3, 3, 22)
    s.set_col(4, 5, 16)

    # the two controls
    row = 4
    s.merge(row, 0, 0, "Recipe", fmts.chrome("label_bold"))
    s.write(row, 1, "R-0001" if cfg.include_sample_data else "",
            fmts.input("text"))
    ws.data_validation(row, 1, row, 1, list_from("RecipeIDs"))
    book.define_name("CardRecipe", f"={quote_sheet(NAME)}!$B${row + 1}")
    s.merge(row, 2, 2, "For how many", fmts.chrome("label_right"))
    s.write(row, 3, cfg.default_servings, fmts.input("int"))
    ws.data_validation(row, 3, row, 3, number_check(1, 200,
                                                    decimal=False))
    book.define_name("CardPeople", f"={quote_sheet(NAME)}!$D${row + 1}")
    s.merge(row, 4, LAST_COL, "Pick a recipe ID, set the number of "
                              "people, print.", fmts.chrome("note"))
    row += 1

    base = 'MATCH(CardRecipe,' + s.reg.render('[[Recipes.id]]') + ',0)'
    ws.write_formula(row, 7,
                     s.reg.render(f'=IFERROR(CardPeople/INDEX('
                                  f'[[Recipes.servings]],{base}),1)'),
                     fmts.calc("number1"))
    book.define_name("CardFactor", f"={quote_sheet(NAME)}!$H${row + 1}")
    ws.set_column(7, 7, None, None, {"hidden": 1})
    row += 1

    # the title block
    s.merge(row, 0, LAST_COL, None, fmts.chrome("print_title"))
    ws.write_formula(row, 0,
                     s.reg.render(f'=IFERROR(INDEX([[Recipes.name]],'
                                  f'{base}),"Choose a recipe above")'),
                     fmts.chrome("print_title"))
    s.row_height(row, 30)
    row += 1
    s.merge(row, 0, LAST_COL, None, fmts.chrome("print_sub"))
    ws.write_formula(
        row, 0,
        s.reg.render(f'=IFERROR(INDEX([[Recipes.category]],{base})&'
                     f'"  |  "&INDEX([[Recipes.cuisine]],{base})&'
                     f'"  |  serves "&TEXT(CardPeople,"0")&'
                     f'"  |  "&TEXT(INDEX([[Recipes.total_min]],{base}),'
                     f'"0")&" minutes  |  "&TEXT(INDEX('
                     f'[[Recipes.cost_serving]],{base})*1,"0.00")&'
                     f'" a serving","")'), fmts.chrome("print_sub"))
    row += 2

    # ingredients
    row = s.table_head(row, 0, ["Ingredient", "You need", "Unit",
                                "Prepared", "Optional?"],
                       [26, 14, 12, 22, 12])
    for n in range(LINES):
        pick = (f'MATCH(CardRecipe&"#"&{n + 1},'
                f'[[Recipe Ingredients.scale_key]],0)')
        ws.write_formula(
            row, 0,
            s.reg.render(f'=IFERROR(INDEX('
                         f'[[Recipe Ingredients.ingredient]],{pick}),"")'),
            fmts.cell("text", "plain"))
        ws.write_formula(
            row, 1,
            s.reg.render(f'=IFERROR(ROUND(INDEX('
                         f'[[Recipe Ingredients.qty]],{pick})*'
                         f'CardFactor,2),"")'),
            fmts.cell("number1", "plain"))
        for col, key in ((2, "unit"), (3, "prep"), (4, "optional")):
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX('
                             f'[[Recipe Ingredients.{key}]],{pick}),"")'),
                fmts.cell("text", "plain"))
        s.row_height(row, 18)
        row += 1
    row += 1

    # method
    row = s.section(row, "Method", LAST_COL)
    s.merge(row, 0, LAST_COL, None, fmts.cell("wrap", "plain"))
    ws.write_formula(row, 0,
                     s.reg.render(f'=IFERROR(INDEX('
                                  f'[[Recipes.method_notes]],{base}),"")'),
                     fmts.cell("wrap", "plain"))
    s.row_height(row, 90)
    row += 2

    # the numbers
    row = s.section(row, "Per serving", LAST_COL)
    row = s.table_head(row, 0, ["Cost", "Calories", "Protein (g)",
                                "Carbs (g)", "Fat (g)", "Fibre (g)"])
    s.calc_row(row, 0, [
        s.reg.render(f'=IFERROR(INDEX([[Recipes.cost_serving]],{base}),'
                     f'"")'),
        s.reg.render(f'=IFERROR(INDEX([[Recipes.calories]],{base}),"")'),
        s.reg.render(f'=IFERROR(INDEX([[Recipes.protein]],{base}),"")'),
        s.reg.render(f'=IFERROR(INDEX([[Recipes.carbs]],{base}),"")'),
        s.reg.render(f'=IFERROR(INDEX([[Recipes.fat]],{base}),"")'),
        s.reg.render(f'=IFERROR(INDEX([[Recipes.fibre]],{base}),"")'),
    ], ["money", "number", "number1", "number1", "number1", "number1"],
        mode="plain")
    row += 2

    s.merge(row, 0, LAST_COL, None, fmts.chrome("note_box"))
    ws.write_formula(
        row, 0,
        s.reg.render(f'=IFERROR("Allergens: "&INDEX('
                     f'[[Recipes.allergens]],{base})&"   |   Notes: "&'
                     f'INDEX([[Recipes.notes]],{base}),"")'),
        fmts.chrome("note_box"))
    s.row_height(row, 32)
    row += 2

    s.merge(row, 0, LAST_COL, None, fmts.chrome("print_sub"))
    ws.write_formula(row, 0,
                     f'=HouseName&"  |  {cfg.product_name}  |  "&'
                     f'TEXT(TODAY(),"d mmm yyyy")',
                     fmts.chrome("print_sub"))

    s.finish(freeze_row=None, print_area=f"A1:F{row + 1}",
             gridlines=False, fit_width=True)
    return s
