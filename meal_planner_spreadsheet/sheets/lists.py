"""`Lists` - every drop down in the workbook, plus the unit table.

Each list becomes a workbook level defined name, so a column can point at
it with ``validate=list_from("MealSlotList")``. Edit a list here and every
drop down that uses it follows.

The unit table is the engine behind the automatic grocery list: each unit
carries the factor that converts it to a base unit (grams, millilitres or
"each"), which is what lets the workbook add 2 tbsp of oil to 100 ml of
oil and come out with a sensible number.
"""

from __future__ import annotations

from typing import List, Tuple

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter

NAME = "Lists"
LAST_COL = 9

#: name -> (header, [values]) - written one list per column block
LISTS: List[Tuple[str, List[str]]] = [
    ("YesNoList", ["Yes", "No"]),
    ("YesNoNAList", ["Yes", "No", "N/A"]),
    ("MealSlotList", ["Breakfast", "Morning snack", "Lunch",
                      "Afternoon snack", "Dinner", "Evening snack",
                      "Dessert"]),
    ("MealStatusList", ["Planned", "Prepared", "Eaten", "Skipped",
                        "Ate out", "Takeaway", "Leftovers", "Cancelled"]),
    ("RecipeCategoryList", ["Main", "Side", "Soup", "Salad", "Breakfast",
                            "Snack", "Dessert", "Baking", "Sauce & dip",
                            "Drink", "Batch base", "Lunchbox"]),
    ("CuisineList", ["American", "British", "Chinese", "French", "Greek",
                     "Indian", "Italian", "Japanese", "Korean", "Mexican",
                     "Middle Eastern", "Spanish", "Thai", "Vietnamese",
                     "Caribbean", "African", "Fusion", "Other"]),
    ("CookingMethodList", ["Oven", "Hob", "One pot", "Slow cooker",
                           "Pressure cooker", "Air fryer", "Grill",
                           "Barbecue", "Microwave", "No cook",
                           "Blender", "Bread machine"]),
    ("DifficultyList", ["Easy", "Medium", "Hard"]),
    ("SpiceLevelList", ["None", "Mild", "Medium", "Hot", "Very hot"]),
    ("PortionSizeList", ["Small", "Regular", "Large", "Extra large"]),
    ("DietList", ["Vegetarian", "Vegan", "Pescatarian", "Gluten-free",
                  "Dairy-free", "Egg-free", "Nut-free", "Soy-free",
                  "Low sodium", "Low carb", "High protein", "High fibre",
                  "Halal", "Kosher", "No restriction"]),
    ("AllergenList", ["Celery", "Cereals with gluten", "Crustaceans",
                      "Eggs", "Fish", "Lupin", "Milk", "Molluscs",
                      "Mustard", "Nuts", "Peanuts", "Sesame", "Soya",
                      "Sulphites", "None"]),
    ("GroceryCategoryList", ["Produce", "Meat & seafood", "Dairy & eggs",
                             "Bakery", "Pantry staples", "Frozen",
                             "Beverages", "Spices & seasonings",
                             "Snacks", "Baby & kids", "Household",
                             "Health & personal care", "Pet", "Other"]),
    ("AisleList", ["1 - Produce", "2 - Bakery", "3 - Deli & cheese",
                   "4 - Meat & fish", "5 - Dairy & eggs", "6 - Frozen",
                   "7 - Tins & jars", "8 - Pasta, rice & grains",
                   "9 - Baking", "10 - Spices & sauces", "11 - Snacks",
                   "12 - Drinks", "13 - Household", "14 - Health",
                   "Online only", "Market stall"]),
    ("StorageList", ["Pantry", "Cupboard", "Fridge", "Freezer",
                     "Fruit bowl", "Spice rack", "Garage shelf",
                     "Cellar", "Other"]),
    ("ContainerList", ["Glass box", "Plastic box", "Freezer bag",
                       "Jar", "Foil tray", "Vacuum bag", "Original pack",
                       "Wrapped plate", "Bowl with lid"]),
    ("ShoppingStatusList", ["To buy", "On the list", "Bought",
                            "Substituted", "Out of stock", "Not needed",
                            "Ordered"]),
    ("PriorityList", ["Must have", "Nice to have", "Only if cheap",
                      "Stock up"]),
    ("TaskStatusList", ["To do", "In progress", "Done", "Blocked",
                        "Not needed"]),
    ("PrepTaskList", ["Wash & chop", "Marinate", "Batch cook", "Bake",
                      "Portion & label", "Freeze", "Defrost", "Soak",
                      "Make sauce", "Pack lunchboxes", "Shop", "Tidy"]),
    ("MemberRoleList", ["Cook", "Shopper", "Helper", "Dishes",
                        "Lunchbox duty", "Guest", "Child"]),
    ("EventTypeList", ["Birthday", "Christmas", "Thanksgiving", "Easter",
                       "Dinner party", "Barbecue", "Picnic", "Potluck",
                       "Sunday lunch", "Anniversary", "School event",
                       "Other"]),
    ("LeftoverStatusList", ["In the fridge", "In the freezer", "Planned",
                            "Eaten", "Frozen for later", "Binned",
                            "Given away"]),
    ("WasteReasonList", ["Expired", "Went mouldy", "Cooked too much",
                         "Did not like it", "Forgot about it",
                         "Freezer burn", "Left out too long",
                         "Bought too much", "Plate scrapings", "Other"]),
    ("WasteActionList", ["Binned", "Composted", "Fed to pets",
                         "Given away", "Donated"]),
    ("StoreTypeList", ["Supermarket", "Discounter", "Local grocer",
                       "Butcher", "Fishmonger", "Bakery", "Farm shop",
                       "Market", "Online delivery", "Wholesale",
                       "Convenience"]),
    ("SpendCategoryList", ["Groceries", "Fresh top-up", "Bulk shop",
                           "Takeaway", "Restaurant", "Coffee & drinks",
                           "Work lunch", "School meals", "Snacks",
                           "Subscription box", "Delivery fee",
                           "Household"]),
    ("GoalCategoryList", ["Budget", "Health", "Waste", "Habit",
                          "Variety", "Skill", "Time"]),
    ("GoalStatusList", ["Not started", "In progress", "On track",
                        "Behind", "Complete", "Paused"]),
    ("FrequencyList", ["Daily", "Weekly", "Fortnightly", "Monthly",
                       "Occasionally", "One-off"]),
    ("SeasonList", ["All year", "Spring", "Summer", "Autumn", "Winter"]),
    ("RatingList", ["5 - Make it again", "4 - Good", "3 - Fine",
                    "2 - Needs work", "1 - Never again"]),
    ("TextureList", ["Any", "Soft", "Crunchy", "Creamy", "Crispy",
                     "Smooth", "No lumps", "Not slimy"]),
    ("MonthList", ["January", "February", "March", "April", "May",
                   "June", "July", "August", "September", "October",
                   "November", "December"]),
    ("WeekdayList", ["Monday", "Tuesday", "Wednesday", "Thursday",
                     "Friday", "Saturday", "Sunday"]),
    ("WeekStartList", ["Monday", "Sunday"]),
    ("SystemList", ["Metric", "Imperial"]),
    ("RoundingList", ["Whole packs", "Exact quantity"]),
    ("ActiveList", ["Active", "Occasional", "Away", "Inactive"]),
    ("AgeGroupList", ["Adult", "Teenager", "Child", "Toddler", "Baby",
                      "Senior"]),
    ("BudgetLineList", ["Weekly shop", "Top-up shop", "Eating out",
                        "Takeaway", "Coffee & treats", "Lunches out",
                        "Special occasion", "Household", "Pet food"]),
    ("SpendTypeList", ["Weekly shop", "Top-up shop", "Bulk shop",
                       "Takeaway", "Eating out", "Work lunch",
                       "Coffee & treats", "Delivery fee",
                       "Special occasion", "Household"]),
    ("PaymentList", ["Card", "Cash", "Online", "Loyalty points",
                     "Voucher", "Shared account"]),
    ("PriceLevelList", ["Budget", "Mid-range", "Premium", "Varies"]),
    ("PrepStatusList", ["To do", "Prepping", "Done", "Skipped"]),
    ("EventStatusList", ["Idea", "Planned", "Menu set", "Shopping done",
                         "Done", "Cancelled"]),
    ("DirectionList", ["Hit or beat", "Stay under"]),
    ("DisposalList", ["General bin", "Food waste caddy", "Compost",
                      "Fed to pets", "Given away"]),
    ("MeasureList", ["Home cooked meals a week", "Takeaways a month",
                     "Grocery spend this month", "New recipes this month",
                     "Food waste this month", "Packed lunches a week",
                     "Leftovers eaten", "Meals prepped ahead",
                     "Your own count"]),
]

#: unit -> (base unit, how many base units one of it is, what it measures)
#: The factors are the ordinary cooking conversions; the "each" units are
#: counted rather than measured, so they convert to themselves.
UNITS: List[Tuple[str, str, float, str]] = [
    ("g", "g", 1, "Weight"),
    ("kg", "g", 1000, "Weight"),
    ("mg", "g", 0.001, "Weight"),
    ("oz", "g", 28.3495, "Weight"),
    ("lb", "g", 453.592, "Weight"),
    ("ml", "ml", 1, "Volume"),
    ("l", "ml", 1000, "Volume"),
    ("tsp", "ml", 5, "Volume"),
    ("tbsp", "ml", 15, "Volume"),
    ("cup", "ml", 240, "Volume"),
    ("fl oz", "ml", 29.5735, "Volume"),
    ("pint (US)", "ml", 473.176, "Volume"),
    ("pint (UK)", "ml", 568.261, "Volume"),
    ("quart", "ml", 946.353, "Volume"),
    ("gallon (US)", "ml", 3785.41, "Volume"),
    ("each", "each", 1, "Count"),
    ("clove", "each", 1, "Count"),
    ("slice", "each", 1, "Count"),
    ("pinch", "each", 1, "Count"),
    ("handful", "each", 1, "Count"),
    ("bunch", "each", 1, "Count"),
    ("sprig", "each", 1, "Count"),
    ("can", "each", 1, "Count"),
    ("jar", "each", 1, "Count"),
    ("packet", "each", 1, "Count"),
    ("bag", "each", 1, "Count"),
    ("bottle", "each", 1, "Count"),
    ("loaf", "each", 1, "Count"),
    ("sheet", "each", 1, "Count"),
    ("to taste", "each", 0, "Count"),
]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.grey_text)

    s.banner("Drop Down Lists",
             "Every list the workbook offers you. Edit, add or delete an "
             "entry and the drop down that uses it changes everywhere.",
             LAST_COL)

    row = 4
    row = s.note(
        row,
        "Add new entries directly under the last one in a column - the "
        "lists are sized with room to grow, and blanks are ignored. Do "
        "not delete a whole column: a drop down pointing at a missing "
        "list stops working.",
        LAST_COL, height=30)
    row += 1

    # ------------------------------------------------------------ the lists
    row = s.section(row, "The drop downs", LAST_COL)
    per_row = 5
    spare = 8                       # blank rows kept inside every name
    top = row
    col = 0
    tallest = 0
    for index, (name, values) in enumerate(LISTS):
        if index and index % per_row == 0:
            top += tallest + 2
            col = 0
            tallest = 0
        s.write(top, col, name.replace("List", ""), fmts.chrome("header"))
        s.set_col(col, col, 22)
        for offset, value in enumerate(values):
            s.write(top + 1 + offset, col, value, fmts.input("text"))
        for offset in range(len(values), len(values) + spare):
            s.write(top + 1 + offset, col, None, fmts.input("text"))
        first = top + 2             # 1 based first value row
        last = top + 1 + len(values) + spare
        book.define_name(
            name,
            f"={quote_sheet(NAME)}!${chr(65 + col)}${first}:"
            f"${chr(65 + col)}${last}")
        tallest = max(tallest, len(values) + spare)
        col += 2
    row = top + tallest + 3

    # -------------------------------------------------------- the unit table
    row = s.section(row, "Measurement units and how they convert", LAST_COL)
    row = s.note(
        row,
        "This little table is what makes the grocery list add up. Every "
        "quantity you type on a recipe is multiplied by the factor below "
        "to get grams, millilitres or a simple count, so two tablespoons "
        "of oil and 100 ml of oil end up on one line of the shopping "
        "list. Add a unit here and it works everywhere.",
        LAST_COL, height=42)
    row += 1

    s.table_head(row, 0, ["Unit", "Measures", "Base unit",
                          "1 unit = this many base units"],
                 [14, 12, 12, 30])
    first_unit = row + 1
    for index, (unit, base, factor, kind) in enumerate(UNITS):
        r = first_unit + index
        s.write(r, 0, unit, fmts.input("text"))
        s.write(r, 1, kind, fmts.input("text"))
        s.write(r, 2, base, fmts.input("text"))
        s.write(r, 3, factor, fmts.input("number1"))
    spare_units = 10
    for index in range(len(UNITS), len(UNITS) + spare_units):
        r = first_unit + index
        for c, style in enumerate(["text", "text", "text", "number1"]):
            s.write(r, c, None, fmts.input(style))
    last_unit = first_unit + len(UNITS) + spare_units

    sheet_ref = quote_sheet(NAME)
    book.define_name("UnitNames",
                     f"={sheet_ref}!$A${first_unit + 1}:$A${last_unit}")
    book.define_name("UnitKinds",
                     f"={sheet_ref}!$B${first_unit + 1}:$B${last_unit}")
    book.define_name("UnitBases",
                     f"={sheet_ref}!$C${first_unit + 1}:$C${last_unit}")
    book.define_name("UnitFactors",
                     f"={sheet_ref}!$D${first_unit + 1}:$D${last_unit}")
    book.define_name("UnitList",
                     f"={sheet_ref}!$A${first_unit + 1}:$A${last_unit}")
    row = last_unit + 2

    row = s.note(
        row,
        "A note on cups and grams: converting a volume into a weight "
        "depends on what is in the cup (a cup of flour and a cup of "
        "honey weigh very different amounts). The workbook keeps weights "
        "and volumes apart rather than guessing - if you buy an "
        "ingredient by weight, write its recipe quantities by weight too.",
        LAST_COL, height=44)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
