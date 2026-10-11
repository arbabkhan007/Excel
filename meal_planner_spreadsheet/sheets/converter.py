"""`Converter` - unit conversion and recipe scaling, on one page.

Two tools. The top half converts any quantity between the units in the
table on the Lists tab - grams to ounces, millilitres to cups, a
tablespoon to a teaspoon. The bottom half takes one of your recipes and
rewrites every ingredient line for a different number of people, with
the new cost.

Weight and volume are deliberately not interchangeable: 100 ml of oil
is not 100 g of oil, and a workbook that pretends otherwise quietly
ruins your baking.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, list_from, number_check

NAME = "Converter"
LAST_COL = 9
QUICK_ROWS = 10
SCALE_ROWS = 18


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_mid)
    ws = s.ws

    s.banner("Converter & Recipe Scaler",
             "Change the units, or change how many people you are "
             "feeding - the quantities follow.", LAST_COL)
    s.set_col(0, 0, 22)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 14)

    # ------------------------------------------------- the quick convert
    row = 4
    row = s.section(row, "Convert a quantity", LAST_COL)
    row = s.note(row, "Type an amount and a unit on the left, choose "
                      "the unit you want on the right. Anything that "
                      "cannot be converted - grams into millilitres, "
                      "say - says so instead of guessing.",
                 LAST_COL, height=26)
    row = s.table_head(row, 0, ["Amount", "From", "To", "Answer",
                                "In base units", "Measures"],
                       [12, 14, 14, 16, 16, 14])
    first_q = row
    for i in range(QUICK_ROWS):
        s.write(row, 0, 100 if i == 0 else None, fmts.input("number1"))
        ws.data_validation(row, 1, row, 1, list_from("UnitList"))
        s.write(row, 1, "g" if i == 0 else None, fmts.input("text"))
        ws.data_validation(row, 2, row, 2, list_from("UnitList"))
        s.write(row, 2, "oz" if i == 0 else None, fmts.input("text"))
        r = row + 1
        base = (f'A{r}*INDEX(UnitFactors,MATCH(B{r},UnitNames,0))')
        to_factor = f'INDEX(UnitFactors,MATCH(C{r},UnitNames,0))'
        same = (f'INDEX(UnitBases,MATCH(B{r},UnitNames,0))='
                f'INDEX(UnitBases,MATCH(C{r},UnitNames,0))')
        ws.write_formula(
            row, 3,
            f'=IF(OR(A{r}="",B{r}="",C{r}=""),"",IFERROR(IF({same},'
            f'{base}/{to_factor},"Cannot convert "&'
            f'INDEX(UnitKinds,MATCH(B{r},UnitNames,0))&" to "&'
            f'INDEX(UnitKinds,MATCH(C{r},UnitNames,0))),"Check the '
            f'units"))', fmts.calc("number1"))
        ws.write_formula(
            row, 4,
            f'=IF(OR(A{r}="",B{r}=""),"",IFERROR({base}&" "&'
            f'INDEX(UnitBases,MATCH(B{r},UnitNames,0)),""))',
            fmts.calc("text"))
        ws.write_formula(
            row, 5,
            f'=IF(B{r}="","",IFERROR(INDEX(UnitKinds,MATCH(B{r},'
            f'UnitNames,0)),""))', fmts.calc("text"))
        row += 1
    ws.conditional_format(first_q, 3, row - 1, 3,
                          {"type": "text", "criteria": "containing",
                           "value": "Cannot", "format": fmts.cf("warn")})
    row += 1

    # ------------------------------------------------- the common ones --
    row = s.section(row, "The ones you look up every time", LAST_COL)
    row = s.table_head(row, 0, ["Metric", "=", "Imperial / cups",
                                "Metric", "=", "Imperial / cups"],
                       [18, 4, 22, 18, 4, 22])
    pairs = [
        ("1 teaspoon", "5 ml", "100 g flour", "about 3/4 cup"),
        ("1 tablespoon", "15 ml (3 tsp)", "100 g sugar", "1/2 cup"),
        ("1 cup", "240 ml", "100 g butter", "7 tablespoons"),
        ("250 ml", "about 1 cup", "100 g rice", "1/2 cup uncooked"),
        ("500 ml", "about 2 cups", "1 egg (medium)", "about 50 g"),
        ("1 litre", "about 4 cups", "1 clove garlic", "1 teaspoon minced"),
        ("100 g", "3.5 oz", "Oven 180C fan", "350F / gas 4"),
        ("500 g", "1 lb 2 oz", "Oven 200C fan", "400F / gas 6"),
        ("1 kg", "2 lb 3 oz", "Oven 220C fan", "425F / gas 7"),
    ]
    for left, left_v, right, right_v in pairs:
        s.write(row, 0, left, fmts.chrome("label"))
        s.write(row, 1, "=", fmts.chrome("label"))
        s.write(row, 2, left_v, fmts.chrome("label_bold"))
        s.write(row, 3, right, fmts.chrome("label"))
        s.write(row, 4, "=", fmts.chrome("label"))
        s.write(row, 5, right_v, fmts.chrome("label_bold"))
        row += 1
    row = s.note(row, "Cup measures are the US cup (240 ml). Flour and "
                      "sugar weights are approximate because a cup of "
                      "flour depends on how hard you pack it - weigh "
                      "it if the baking matters.", LAST_COL, height=26)

    # ----------------------------------------------------- the scaler ---
    row = s.section(row, "Scale a recipe", LAST_COL)
    row = s.field(row, "Recipe",
                  "R-0001" if cfg.include_sample_data else "",
                  name="ScaleRecipe",
                  validate=list_from("RecipeIDs"),
                  help_text="Pick the recipe ID. The Recipes tab lists "
                            "them.")
    row = s.field(row, "People to feed", cfg.default_servings,
                  style="int", name="ScalePeople",
                  validate=number_check(1, 200, decimal=False),
                  help_text="Every quantity below is multiplied to "
                            "suit.")
    s.merge(row, 1, 2, "Recipe name", fmts.chrome("label_right"))
    ws.write_formula(row, 3,
                     s.reg.render('=IFERROR(INDEX([[Recipes.name]],'
                                  'MATCH(ScaleRecipe,[[Recipes.id]],0)),'
                                  '"Pick a recipe above")'),
                     fmts.calc("text"))
    s.merge(row, 4, 4, "Normally serves", fmts.chrome("label_right"))
    ws.write_formula(row, 5,
                     s.reg.render('=IFERROR(INDEX([[Recipes.servings]],'
                                  'MATCH(ScaleRecipe,[[Recipes.id]],0)),'
                                  '"")'), fmts.calc("int"))
    s.merge(row, 6, 6, "Scale factor", fmts.chrome("label_right"))
    ws.write_formula(row, 7,
                     s.reg.render('=IFERROR(ScalePeople/INDEX('
                                  '[[Recipes.servings]],MATCH('
                                  'ScaleRecipe,[[Recipes.id]],0)),"")'),
                     fmts.calc("number1"))
    book.define_name("ScaleFactor", f"={quote_sheet(NAME)}!$H${row + 1}")
    row += 2

    row = s.table_head(row, 0, ["Ingredient", "Original", "Unit",
                                "You need", "In base units", "Prepared",
                                "Cost", "Optional?"],
                       [26, 12, 10, 14, 16, 18, 12, 10])
    first_s = row
    for n in range(SCALE_ROWS):
        pick = (f'MATCH(ScaleRecipe&"#"&{n + 1},'
                f'[[Recipe Ingredients.scale_key]],0)')
        for col, key, style in [(0, "ingredient", "text"),
                                (1, "qty", "number1"),
                                (2, "unit", "text"),
                                (5, "prep", "text"),
                                (7, "optional", "flag")]:
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX('
                             f'[[Recipe Ingredients.{key}]],{pick}),"")'),
                fmts.calc(style))
        r = row + 1
        ws.write_formula(row, 3,
                         f'=IF(B{r}="","",IFERROR(ROUND(B{r}*'
                         f'ScaleFactor,2),""))', fmts.calc("number1"))
        ws.write_formula(
            row, 4,
            s.reg.render(f'=IF(A{r}="","",IFERROR(ROUND(INDEX('
                         f'[[Recipe Ingredients.base_qty]],{pick})*'
                         f'ScaleFactor,0)&" "&INDEX('
                         f'[[Recipe Ingredients.base_unit]],{pick}),""))'),
            fmts.calc("text"))
        ws.write_formula(
            row, 6,
            s.reg.render(f'=IF(A{r}="","",IFERROR(INDEX('
                         f'[[Recipe Ingredients.line_cost]],{pick})*'
                         f'ScaleFactor,""))'), fmts.calc("money"))
        row += 1
    s.merge(row, 0, 2, "Scaled totals", fmts.chrome("total_label"))
    ws.write_formula(row, 6, f'=SUM(G{first_s + 1}:G{row})',
                     fmts.chrome("total_value",
                                 num_format=cfg.money_format))
    s.merge(row, 7, LAST_COL, None, fmts.chrome("total_value",
                                                align="left"))
    ws.write_formula(
        row, 7,
        s.reg.render('=IFERROR("That is "&TEXT(ScalePeople,"0")&'
                     '" servings at "&TEXT(SUM(G'
                     + str(first_s + 1) + ':G' + str(row) + ')/'
                     'MAX(1,ScalePeople),"0.00")&" each","")'),
        fmts.chrome("total_value", align="left"))
    row += 2

    s.note(row,
           "Scaling multiplies every quantity by the same number, "
           "which is right for ingredients and wrong for cooking "
           "times. A double batch of stew takes a little longer to "
           "come to the boil, not twice as long; a doubled cake needs "
           "a bigger tin, not two of the same. Spices and salt "
           "rarely want doubling either - taste as you go.",
           LAST_COL, height=46)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
