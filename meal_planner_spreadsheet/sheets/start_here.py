"""`Start Here` - the welcome page and the six step set-up."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Start Here"
LAST_COL = 11

STEPS = [
    ("1", "Settings",
     "Your household, how many you feed, the weekly budget, your "
     "shopping day and your nutrition targets. Ten minutes, and it "
     "feeds every other tab."),
    ("2", "Household",
     "One row per person: portion size, diet, allergies, what they "
     "will not touch and which meals they are in for."),
    ("3", "Ingredients",
     "The price book. Pack size and pack price for the thirty things "
     "you buy most. This is what prices every recipe and every "
     "shopping list, so it is the tab worth the effort."),
    ("4", "Recipes + Recipe Ingredients",
     "Add a recipe, then list its ingredients one line each with a "
     "quantity and a unit. The cost, the calories and the allergens "
     "are worked out for you."),
    ("5", "Pantry",
     "What is already in the cupboard, the fridge and the freezer. "
     "This is what stops the grocery list buying things you own."),
    ("6", "Meal Plan",
     "Plan the week: date, meal, recipe, how many you are feeding. "
     "The Weekly Planner shows it as a week and the Grocery List "
     "builds itself."),
]

TABS = [
    ("Dashboard", "This week at a glance - cost, meals, alerts"),
    ("Weekly Planner", "Seven days on one page, ready to print"),
    ("Monthly Planner", "The month as a calendar grid"),
    ("Settings", "Household, budget, targets, shopping window"),
    ("Lists", "Every drop-down, plus the unit conversion table"),
    ("Household", "Who eats here, portions, diets and allergies"),
    ("Recipes", "THE RECIPE BOOK - cost and nutrition worked out"),
    ("Recipe Ingredients", "What goes in each recipe. The engine room"),
    ("Ingredients", "The price book: pack sizes, prices, nutrition"),
    ("Meal Plan", "THE MAIN TAB - one row per planned meal"),
    ("Grocery List", "Builds itself from the plan, minus the pantry"),
    ("Pantry", "What you own, what it is worth, what expires when"),
    ("Staples", "Par levels for the things you never want to run out of"),
    ("Leftovers", "What is in the fridge and how long you have"),
    ("Food Waste", "What got binned and what it cost"),
    ("Meal Prep", "Batch cooking: portions made, minutes per portion"),
    ("Grocery Spending", "Every shop and takeaway against the budget"),
    ("Store Prices", "The same item priced in three shops"),
    ("Stores", "Where you shop, hours, delivery, loyalty"),
    ("Events", "Birthdays and occasions with a head count and budget"),
    ("Goals", "Three or four food habits, measured automatically"),
    ("Help", "Every question answered, including the honest limits"),
]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)

    s.banner(cfg.product_name,
             "Plan the week, price it before you shop, and stop buying "
             "things you already own.", LAST_COL)
    s.set_col(0, 0, 4)
    s.set_col(1, 1, 22)
    s.set_col(2, 2, 30)
    for c in range(3, LAST_COL + 1):
        s.set_col(c, c, 13)

    row = 4
    row = s.note(
        row,
        "Everything in this workbook is linked. Type something once - a "
        "recipe, a price, a meal - and it turns up everywhere it is "
        "needed: the planner, the shopping list, the budget, the "
        "nutrition tracker and the alerts. Cream cells are yours to "
        "fill in; grey cells are formulas and are locked so they "
        "cannot be broken by accident.",
        LAST_COL, first_col=1, height=46)
    row += 1

    row = s.section(row, "Six steps to set it up", LAST_COL)
    s.table_head(row, 1, ["Step", "Where", "What to do"], [8, 26, 68])
    first = row + 1
    for i, (number, where, what) in enumerate(STEPS):
        r = first + i
        s.row_height(r, 32)
        s.write(r, 1, number, fmts.chrome("kpi_label"))
        s.write(r, 2, where, fmts.chrome("label_bold"))
        s.merge(r, 3, LAST_COL, what, fmts.chrome("note"))
    row = first + len(STEPS) + 1

    row = s.section(row, "How the automatic grocery list works", LAST_COL)
    row = s.bullets(row, [
        "1. The Meal Plan says you are cooking recipe R-0002 on Tuesday "
        "for 4 people.",
        "2. The recipe serves 6, so everything in it is multiplied by "
        "0.67 - that is the scaling, done for you.",
        "3. Each ingredient line is converted to a base unit: 2 tbsp "
        "of oil becomes 30 ml, 1 kg of mince becomes 1000 g.",
        "4. Every meal in the shopping window is added up by "
        "ingredient, so three recipes using onions make one line.",
        "5. Whatever the Pantry says you already own is taken off.",
        "6. What is left is rounded up to whole packs and priced from "
        "the Ingredients tab - that is your list, with a total, before "
        "you leave the house.",
    ], LAST_COL, first_col=1, height=26)
    row += 1

    row = s.section(row, "The weekly rhythm", LAST_COL)
    row = s.bullets(row, [
        "FRIDAY (10 minutes) - look at the Pantry 'Use soon' list and "
        "the Leftovers tab, then plan next week on the Meal Plan tab.",
        "SATURDAY MORNING - open the Grocery List, tick off anything "
        "you already have, print it or take the laptop, and shop.",
        "AFTER THE SHOP (2 minutes) - log the total on the Grocery "
        "Spending tab and update the Pantry for the big items.",
        "PREP DAY - work through the Meal Prep tab: two or three jobs, "
        "not ten.",
        "DURING THE WEEK - mark meals 'Eaten' as you go and log "
        "leftovers while you are putting them in the fridge.",
        "END OF THE MONTH - read the Dashboard and the Food Waste tab, "
        "and change one thing.",
    ], LAST_COL, first_col=1, height=26)
    row += 1

    row = s.section(row, "What is in the workbook", LAST_COL)
    s.table_head(row, 1, ["Tab", "What it is for"], [26, 70])
    first = row + 1
    for i, (tab, purpose) in enumerate(TABS):
        r = first + i
        s.write(r, 1, tab, fmts.chrome("label_bold"))
        s.merge(r, 2, LAST_COL, purpose, fmts.chrome("label"))
    row = first + len(TABS) + 1

    row = s.section(row, "Good to know", LAST_COL)
    row = s.bullets(row, [
        f"PROTECTION - every sheet is protected so the formulas cannot "
        f"be edited by accident. The password is '{cfg.password}'. "
        f"Review > Unprotect Sheet if you want to change something.",
        "ADDING ROWS - every register has spare rows ready to use. "
        "Right click a row inside the list and choose Insert to add "
        "more; the formulas copy down.",
        "SORTING - use the filter arrows in the header row rather than "
        "re-sorting the sheet, so the recipe IDs keep matching the "
        "other tabs.",
        "DROP DOWNS - every list lives on the Lists tab, including the "
        "unit table. Add a unit there and it works everywhere.",
        "UNITS - weights convert to grams, volumes to millilitres, and "
        "counted things stay as 'each'. Weight and volume are "
        "deliberately kept apart because 100 ml of oil is not 100 g.",
        "GOOGLE SHEETS - File > Import works, and everything except "
        "the sheet protection carries over.",
        "YOUR CURRENCY - set the symbol on the Settings tab, or "
        "rebuild the workbook with a different currency.",
    ], LAST_COL, first_col=1, height=28)
    row += 1

    row = s.section(row, "What a spreadsheet can and cannot do", LAST_COL)
    row = s.note(
        row,
        "Being straight with you. This workbook does the arithmetic on "
        "the information you give it: it cannot read the barcodes in "
        "your cupboard, it cannot check tonight's shelf prices, and it "
        "will not order the shopping for you. The automatic grocery "
        "list is only as good as your ingredient lines and units - a "
        "recipe that says 'a splash of oil' cannot be added up. "
        "Nutrition comes from the numbers you copy off the packet and "
        "is a guide, not medical advice. Allergen flags and use-by "
        "dates are a second pair of eyes, not a guarantee: always read "
        "the label and trust your nose. Reminders appear on the tabs - "
        "a spreadsheet cannot push a notification to your phone. What "
        "it does do is turn a vague plan into a priced shopping list "
        "in ten minutes a week, and show you exactly where the food "
        "money goes.",
        LAST_COL, first_col=1, height=92)
    row += 1

    s.merge(row, 1, LAST_COL,
            f"{cfg.product_name}  |  {cfg.author}  |  version "
            f"{cfg.version}", fmts.chrome("print_sub"))

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
