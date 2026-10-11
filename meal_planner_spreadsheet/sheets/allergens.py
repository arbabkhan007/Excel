"""`Food Safety` - allergens, storage and the things worth checking.

Three parts: a grid showing which allergens each person reacts to and
how many recipes contain them, a storage and shelf-life reference, and
the cooking temperatures worth keeping somewhere you can see them.

Everything here is a prompt, not a guarantee. The workbook only knows
the allergen tags you typed into the price book, and a tag is not a
label.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Food Safety"
LAST_COL = 9

ALLERGENS = ["Celery", "Cereals with gluten", "Crustaceans", "Eggs",
             "Fish", "Lupin", "Milk", "Molluscs", "Mustard", "Nuts",
             "Peanuts", "Sesame", "Soya", "Sulphites"]

STORAGE = [
    ("Fridge", "0-5C", "Milk, cooked leftovers, opened jars, raw meat "
                       "on the bottom shelf"),
    ("Freezer", "-18C", "Batch cooking, bread, herbs, anything near "
                        "its date"),
    ("Cupboard", "Cool and dark", "Tins, pasta, rice, flour, oil away "
                                  "from the hob"),
    ("Fruit bowl", "Room temperature", "Bananas, tomatoes, citrus, "
                                       "avocados"),
    ("Not the fridge", "Room temperature", "Potatoes, onions, garlic, "
                                           "bread, honey"),
]

TEMPS = [
    ("Poultry", "75C / 165F", "No pink, juices run clear"),
    ("Minced meat & burgers", "70C / 160F", "All the way through"),
    ("Pork", "70C / 160F", "Rest it for five minutes"),
    ("Whole cuts of beef & lamb", "55-70C", "Rare to well done - your "
                                            "call"),
    ("Fish", "63C / 145F", "Flakes easily"),
    ("Reheating anything", "75C / 165F", "Piping hot right through, "
                                         "once only"),
    ("Fridge", "0-5C", "Check it with a thermometer once a year"),
    ("Freezer", "-18C or colder", "Food stays safe indefinitely; "
                                  "quality drops"),
]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.accent)
    ws = s.ws

    s.banner("Food Safety & Allergens",
             "Who reacts to what, where things live and how hot it "
             "needs to be.", LAST_COL)
    s.set_col(0, 0, 24)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 15)

    # ---------------------------------------------------- the headline --
    row = 4
    kpis = [
        ("People with an allergy",
         '=COUNTIFS([[Household.allergies]],"<>",'
         '[[Household.allergies]],"<>None")', "int",
         "from the Household tab"),
        ("Ingredients carrying an allergen",
         '=COUNTIFS([[Ingredients.allergen]],"<>",'
         '[[Ingredients.allergen]],"<>None")', "int",
         "tagged in the price book"),
        ("Recipes with a flagged ingredient",
         '=COUNTIF([[Recipes.allergens]],"*flagged*")', "int",
         "check these against the real labels"),
        ("Pantry items past their date",
         '=COUNTIF([[Pantry.status]],"Expired")', "int",
         "deal with these first"),
    ]
    for i, (label, formula, style, note) in enumerate(kpis):
        s.kpi(row, i * 2, label, s.reg.render(formula), style=style,
              span=2, note=note)
    row += 4

    # ----------------------------------------------- the allergen grid --
    row = s.section(row, "Allergen grid", LAST_COL)
    row = s.note(row, "Each row is one of the fourteen allergens that "
                      "must be declared on food labels in the UK and "
                      "EU. 'In the house' counts the ingredients in "
                      "your price book carrying that tag.",
                 LAST_COL, height=26)
    head = row
    s.write(head, 0, "Allergen", fmts.chrome("header_left"))
    s.write(head, 1, "Ingredients", fmts.chrome("header"))
    s.write(head, 2, "Recipes", fmts.chrome("header"))
    s.write(head, 3, "In the pantry", fmts.chrome("header"))
    for i in range(5):
        ws.write_formula(head, 4 + i,
                         f'=IFERROR(IF(INDEX(MemberNames,{i + 1})=0,"",'
                         f'INDEX(MemberNames,{i + 1})),"")',
                         fmts.chrome("header"))
    s.row_height(head, 26)
    row = head + 1
    first_grid = row
    for allergen in ALLERGENS:
        s.write(row, 0, allergen, fmts.chrome("label_bold"))
        s.calc_row(row, 1, [
            f'=COUNTIFS([[Ingredients.allergen]],"{allergen}")',
            f'=SUMPRODUCT((COUNTIFS([[Recipe Ingredients.recipe]],'
            f'[[Recipes.id]],[[Recipe Ingredients.allergen]],'
            f'"{allergen}")>0)*([[Recipes.name]]<>""))',
            f'=SUMPRODUCT((COUNTIFS([[Ingredients.name]],'
            f'[[Pantry.ingredient]])>0)*0)+COUNTIFS('
            f'[[Pantry.ingredient]],"<>")*0+SUMPRODUCT('
            f'([[Pantry.ingredient]]<>"")*(COUNTIFS('
            f'[[Ingredients.name]],[[Pantry.ingredient]],'
            f'[[Ingredients.allergen]],"{allergen}")>0))',
        ], ["int", "int", "int"])
        for i in range(5):
            person = f"{chr(69 + i)}${head + 1}"
            ws.write_formula(
                row, 4 + i,
                s.reg.render(f'=IF({person}="","",IF(IFERROR(INDEX('
                             f'[[Household.allergies]],MATCH({person},'
                             f'[[Household.name]],0)),"")="{allergen}",'
                             f'"AVOID","ok"))'),
                fmts.calc("flag"))
        row += 1
    ws.conditional_format(first_grid, 4, row - 1, 8,
                          {"type": "text", "criteria": "containing",
                           "value": "AVOID", "format": fmts.cf("bad")})
    ws.conditional_format(first_grid, 4, row - 1, 8,
                          {"type": "text", "criteria": "containing",
                           "value": "ok", "format": fmts.cf("muted")})
    ws.conditional_format(first_grid, 1, row - 1, 3,
                          {"type": "data_bar", "bar_color": "#D9703C"})
    row += 1

    # --------------------------------------------- recipes to watch ----
    row = s.section(row, "Recipes to check before you serve them",
                    LAST_COL)
    row = s.table_head(row, 0, ["Recipe", "Allergen note", "Serves",
                                "Category", "Last cooked"],
                       [30, 26, 10, 15, 14])
    for n in range(cfg.top_rows * 2):
        pick = f'MATCH({n + 1},[[Recipes.cost_rank]],0)'
        for col, key, style in [(0, "name", "text"),
                                (1, "allergens", "text"),
                                (2, "servings", "int"),
                                (3, "category", "text"),
                                (4, "last_cooked", "date")]:
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(IF(INDEX([[Recipes.allergens]],'
                             f'{pick})="None listed","",INDEX('
                             f'[[Recipes.{key}]],{pick})),"")'),
                fmts.calc(style))
        row += 1
    row += 1

    # --------------------------------------------------- storage rules -
    row = s.section(row, "Where things live", LAST_COL)
    row = s.table_head(row, 0, ["Place", "Temperature", "What goes there",
                                "In your pantry now"],
                       [18, 18, 44, 18])
    for place, temp, what in STORAGE:
        s.write(row, 0, place, fmts.chrome("label_bold"))
        s.write(row, 1, temp, fmts.chrome("label"))
        s.merge(row, 2, LAST_COL - 1, what, fmts.chrome("label"))
        ws.write_formula(
            row, LAST_COL,
            s.reg.render(f'=COUNTIFS([[Pantry.location]],"{place}")'),
            fmts.calc("int"))
        row += 1
    row += 1

    # ------------------------------------------------------ temperatures
    row = s.section(row, "Cooked through means", LAST_COL)
    row = s.table_head(row, 0, ["What", "Temperature", "How you know"],
                       [28, 18, 50])
    for what, temp, how in TEMPS:
        s.write(row, 0, what, fmts.chrome("label_bold"))
        s.write(row, 1, temp, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, how, fmts.chrome("label"))
        row += 1
    row += 1

    # ------------------------------------------------- leftovers rules -
    row = s.section(row, "Leftovers, reheating and the freezer", LAST_COL)
    row = s.bullets(row, [
        "COOL IT FAST - get leftovers into the fridge within two "
        "hours, in a shallow container so they cool quickly.",
        f"THE FRIDGE WINDOW - this workbook assumes "
        f"{cfg.leftover_safe_days} days for cooked leftovers; change "
        f"it on the Settings tab if your household rule differs.",
        f"THE FREEZER WINDOW - {cfg.freezer_review_days} days is the "
        f"quality reminder, not a safety limit. Frozen food stays "
        f"safe; it just stops tasting of much.",
        "REHEAT ONCE - reheat a portion once, until it is piping hot "
        "all the way through. Not twice.",
        "DEFROST IN THE FRIDGE - overnight, not on the worktop. Plan "
        "it on the Meal Prep tab the day before.",
        "RICE - cool it and refrigerate it quickly, eat it within a "
        "day, reheat it once, properly.",
        "WHEN IN DOUBT - the use-by date on the packet beats anything "
        "this spreadsheet works out. Throw it out and log it on the "
        "Food Waste tab.",
    ], LAST_COL, height=26)
    row += 1

    s.note(row,
           "Plain English: this tab organises what you told it. It "
           "cannot see 'may contain' warnings, recipe changes by the "
           "manufacturer, or cross-contamination in your own kitchen. "
           "For a serious allergy, the label on the packet in your "
           "hand is the only source that counts - every time, even "
           "for a product you have bought for years.",
           LAST_COL, height=46)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
