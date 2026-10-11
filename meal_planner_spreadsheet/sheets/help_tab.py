"""`Help` - how the thing works, in the order you will need it.

Setup in ten minutes, then a tab-by-tab reference, the automation
chain written out in plain English, what to do when a number looks
wrong, the honest limits, and the keyboard shortcuts worth knowing.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Help"
LAST_COL = 7

SETUP = [
    ("1. Settings",
     "Household name, who cooks, the week start, your shopping window, "
     "budgets and targets. Ten minutes here and everything else has "
     "something to work with."),
    ("2. Household",
     "One row per person. The portion factor is the important one - a "
     "teenager eating 1.4 portions and a toddler eating 0.6 is why "
     "'serves 4' never quite works."),
    ("3. Lists",
     "Only if you want different categories, cuisines, stores or "
     "units. Everything else reads its dropdowns from here, so edit "
     "the list rather than typing a one-off value."),
    ("4. Ingredients",
     "Name, pack size, pack price, base unit. This is the tab that "
     "makes costing work; without prices the grocery list still works "
     "but the money columns stay blank."),
    ("5. Recipes",
     "Name, servings, timings, category. Then the lines go on Recipe "
     "Ingredients - one row per ingredient per recipe."),
    ("6. Meal Plan",
     "Pick a date, a slot and a recipe. The Weekly Planner and Monthly "
     "Planner are views of this tab, so you can plan from whichever "
     "one suits you."),
    ("7. Grocery List",
     "Nothing to do. Set the shopping window on Settings and the list "
     "builds itself from the plan, minus what the Pantry says you "
     "already have."),
]

TABS = [
    ("Start Here", "The ten minute setup and where everything lives."),
    ("Dashboard", "The week at a glance: cost, calories, what needs "
                  "attention, what is being eaten."),
    ("Weekly Planner", "Seven days by seven meal slots. Pick a week "
                       "with the date at the top."),
    ("Monthly Planner", "The same plan as a calendar month, with "
                        "dinners in the day cells."),
    ("Settings", "Every number the workbook uses to make a decision."),
    ("Lists", "The dropdown lists, and the unit conversion table."),
    ("Household", "Who eats, what they eat, what they cannot eat."),
    ("Recipes", "The recipe book. Cost, calories and macros come back "
                "from the ingredient lines."),
    ("Recipe Ingredients", "One row per ingredient per recipe. The "
                           "engine room."),
    ("Ingredients", "Prices, pack sizes, nutrition per 100, allergens, "
                    "storage."),
    ("Meal Plan", "Every planned meal as a row - the single source of "
                  "truth behind both planners."),
    ("Grocery List", "Built from the plan minus the pantry. Sorted by "
                     "aisle for the actual shop."),
    ("Pantry", "What is in the cupboards, with value and expiry."),
    ("Staples", "The things you always want in: tick when they run "
                "low."),
    ("Leftovers", "Portions in the fridge, with a use-by and a value."),
    ("Food Waste", "What got binned, why, and what it cost."),
    ("Meal Prep", "Batch cooking sessions and what they cover."),
    ("Dietary Filters", "Six dropdowns; the recipes that survive all "
                        "six."),
    ("Nutrition", "Daily calories and macros against the targets."),
    ("Budget", "Twelve months of planned against actual, by type and "
               "by store."),
    ("Recipe Costs", "Cheapest meals first, and where the money goes."),
    ("Converter", "Unit conversion and the recipe scaler."),
    ("Rotation", "Variety: what you cook too often and what you have "
                 "forgotten."),
    ("Food Safety", "Allergens by household member, storage times and "
                    "cooking temperatures."),
    ("Fridge & Freezer", "Cold storage sorted by what needs eating "
                         "first."),
    ("Grocery Spending", "Every receipt. This is what the budget "
                         "compares against."),
    ("Store Prices", "The same item at different shops, and the "
                     "cheapest one."),
    ("Stores", "Where you shop, travel time, delivery, price level."),
    ("Events", "Dinners, parties and holidays with their own budget."),
    ("Goals", "A handful of food habits with a measure attached."),
    ("Sustainability", "Why food gets binned and what stops it."),
    ("Reports", "A month on one page, and the year in one table."),
    ("Print Week", "The week's meals on one sheet for the fridge "
                   "door."),
    ("Print Shopping", "The shopping list with tick boxes."),
    ("Print Recipe", "One recipe scaled to any number of people."),
]

CHAIN = [
    "You set a shopping window on Settings (Shop from / Shop to).",
    "Meal Plan rows inside that window are counted. Each row knows its "
    "recipe and how many people are eating, so it knows its scale.",
    "Recipes totals up how many batches of each recipe the window "
    "needs.",
    "Recipe Ingredients multiplies every line by those batches and "
    "converts it to a base unit - grams, millilitres or each.",
    "Ingredients adds up the total needed for the window, takes off "
    "what the Pantry says you already have, and ranks what is short.",
    "Grocery List reads that ranking, rounds up to whole packs, prices "
    "it from the cheapest store, and groups it by aisle.",
    "Print Shopping prints that same list with tick boxes.",
]

TROUBLE = [
    ("The grocery list is empty",
     "Check Shop from and Shop to on Settings cover days that actually "
     "have meals planned, and that those meals point at a recipe "
     "rather than free text."),
    ("An ingredient is on the list but I have loads",
     "The Pantry quantity is what gets deducted. If the cupboard is "
     "full but the Pantry says zero, the list will buy it again."),
    ("Costs are blank or far too low",
     "An ingredient with no pack price costs nothing. The Recipes tab "
     "has a 'missing a price' flag that tells you which."),
    ("A conversion looks wrong",
     "Weight and volume are never converted into each other - a "
     "spreadsheet cannot know the density of your flour. Enter the "
     "line in the unit you will measure it in."),
    ("#N/A or #VALUE! in a cell",
     "Nearly always a name typed slightly differently from the list - "
     "'Tomatoes' and 'tomatoes ' are two ingredients. Use the "
     "dropdowns."),
    ("A sheet will not let me type",
     "Formula cells are locked on purpose. The password is on the "
     "Start Here tab; Review then Unprotect Sheet if you really want "
     "to change a calculation."),
    ("The numbers did not update",
     "Press F9. If that does not do it, Formulas then Calculation "
     "Options then Automatic."),
]

LIMITS = [
    "Automatic shopping lists need structured ingredients and "
    "consistent units. 'A splash of oil' cannot be added up; '15 ml "
    "oil' can.",
    "Nutrition is only as good as the per-100 figures you type in, and "
    "this is not medical or dietary advice. If something matters "
    "clinically, check it with a professional.",
    "Allergen and storage columns are memory aids, not guarantees. "
    "Always read the actual label - recipes change, so do products.",
    "Nothing here can send you a push notification, email you, or "
    "order the shopping. It is a spreadsheet; it works when you open "
    "it.",
    "Prices go stale. The store price tab is a snapshot of what you "
    "last saw, not a live feed.",
]

SHORTCUTS = [
    ("Ctrl + ;", "Today's date into the cell"),
    ("Alt + Down", "Open the dropdown in a validated cell"),
    ("F9", "Recalculate everything"),
    ("Ctrl + Page Down", "Next tab"),
    ("Ctrl + Shift + Down", "Select to the bottom of the column"),
    ("Ctrl + D", "Copy the cell above - the quickest way to extend a "
                 "formula column into new rows"),
    ("Ctrl + F", "Find - faster than scrolling the Recipes tab"),
]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.ink)

    s.banner("Help & How It Works",
             "Setup, the tab list, the automation chain, and what to "
             "do when a number looks wrong.", LAST_COL)
    s.set_col(0, 0, 24)
    s.set_col(1, LAST_COL, 16)

    row = 4
    row = s.note(row,
                 "Short version: fill in Settings, Household, "
                 "Ingredients and Recipes, then plan some meals. "
                 "Everything else is worked out for you. Every white "
                 "cell is yours to type in; every shaded cell is a "
                 "formula and is locked.",
                 LAST_COL, height=44)
    row += 1

    row = s.section(row, "Setting up, in order", LAST_COL)
    for label, detail in SETUP:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.merge(row, 1, LAST_COL, detail, fmts.chrome("note"))
        s.row_height(row, 32)
        row += 1
    row += 1

    row = s.section(row, "What each tab is for", LAST_COL)
    row = s.table_head(row, 0, ["Tab", "What it does"], [24, 16])
    for tab, what in TABS:
        s.write(row, 0, tab, fmts.chrome("label_bold"))
        s.merge(row, 1, LAST_COL, what, fmts.chrome("note"))
        s.row_height(row, 24)
        row += 1
    row += 1

    row = s.section(row, "How the shopping list builds itself",
                    LAST_COL)
    for i, step in enumerate(CHAIN, start=1):
        s.write(row, 0, f"Step {i}", fmts.chrome("label_bold"))
        s.merge(row, 1, LAST_COL, step, fmts.chrome("note"))
        s.row_height(row, 30)
        row += 1
    row = s.note(row,
                 "Every step is an ordinary formula you can click on "
                 "and read. There are no macros in this file, which is "
                 "also why it opens without a security warning.",
                 LAST_COL, height=30)
    row += 1

    row = s.section(row, "When something looks wrong", LAST_COL)
    row = s.table_head(row, 0, ["Symptom", "What it usually is"],
                       [24, 16])
    for symptom, fix in TROUBLE:
        s.write(row, 0, symptom, fmts.chrome("label_bold"))
        s.merge(row, 1, LAST_COL, fix, fmts.chrome("note"))
        s.row_height(row, 32)
        row += 1
    row += 1

    row = s.section(row, "What this workbook cannot do", LAST_COL)
    for limit in LIMITS:
        s.merge(row, 0, LAST_COL, limit, fmts.chrome("note"))
        s.row_height(row, 30)
        row += 1
    row += 1

    row = s.section(row, "Shortcuts worth knowing", LAST_COL)
    row = s.table_head(row, 0, ["Keys", "What it does"], [24, 16])
    for keys, what in SHORTCUTS:
        s.write(row, 0, keys, fmts.chrome("label_bold"))
        s.merge(row, 1, LAST_COL, what, fmts.chrome("note"))
        s.row_height(row, 22)
        row += 1
    row += 1

    row = s.section(row, "Adding more rows", LAST_COL)
    row = s.bullets(row, [
        "Every register has spare rows below the sample data - just "
        "keep typing.",
        "If you run out, select the last row, copy it, and insert "
        "copied cells. That carries the formulas and the dropdowns "
        "with it.",
        "Formula columns are locked; use the password on Start Here "
        "if you need to unlock a sheet first.",
        "Deleting rows is safe. Deleting whole columns is not - other "
        "tabs point at them by name.",
    ], LAST_COL, height=20)
    row += 1

    row = s.section(row, "Printing", LAST_COL)
    row = s.bullets(row, [
        "Print Week, Print Shopping and Print Recipe are already set "
        "up to fit a page - just press print.",
        "Everything else is set to fit the page width; use landscape "
        "for the wide registers.",
        "The printables read live data, so there is never a stale "
        "copy to tidy up.",
    ], LAST_COL, height=20)
    row += 1

    s.note(row,
           f"{cfg.product_name} - built by {cfg.author}. No macros, no "
           f"add-ins, no internet connection required. Works in Excel "
           f"2010 and later, and in Google Sheets and LibreOffice with "
           f"minor cosmetic differences.",
           LAST_COL, height=36)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
