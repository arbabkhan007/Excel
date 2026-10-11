"""
The `Settings` tab - household setup, planning window, budget and targets.

Everything the rest of the workbook needs to know about your household is
entered once here and published as a defined name (HouseName, PlanStart,
WeeklyBudget, CalorieTarget, ExpiryWarningDays ...) so the formulas stay
readable.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter, date_check, list_from, number_check

SHEET = "Settings"
LAST_COL = 8


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, SHEET, tab_color=Palette.accent)
    h = cfg.household

    s.banner("Household Setup & Settings",
             "Fill in the cream cells once. Grey cells are locked "
             "formulas. Everything else in the workbook reads these "
             "values.", LAST_COL)

    s.set_col(0, 0, 2.5)
    s.set_col(1, 1, 30, fmts.chrome("label"))
    s.set_col(2, 2, 16, fmts.chrome("label"))
    s.set_col(3, 3, 26)
    s.set_col(4, 4, 20)
    s.set_col(5, 8, 17, fmts.chrome("note"))

    r = 4  # 0 based row index -> Excel row 5

    # -------------------------------------------------------- 1. household
    r = s.section(r, "1. Your household", LAST_COL)
    r = s.field(r, "Household name", h.name, name="HouseName", value_end=4,
                help_text="Shown on every printable page and report.")
    r = s.field(r, "Main cook / planner", h.cook, name="HouseCook",
                value_end=4)
    r = s.field(r, "Email", h.email, name="HouseEmail", value_end=4)
    r = s.field(r, "Town / city", h.town, name="HouseTown", value_end=4)
    r = s.field(r, "People to cook for", cfg.members, style="int",
                name="HouseMembers",
                validate=number_check(1, 30, decimal=False),
                help_text="Used as the default serving count and to work "
                          "out the cost per person.")
    r = s.field(r, "Usual supermarket", h.store, name="MainStore",
                value_end=4,
                help_text="The store the grocery list defaults to.")
    r = s.field(r, "Notes about the household", h.notes, name="HouseNotes",
                value_end=4)

    # --------------------------------------------------------- 2. planning
    r += 1
    r = s.section(r, "2. Planning window", LAST_COL)
    r = s.field(r, "Planning year", cfg.year, style="int", name="PlanYear",
                validate=number_check(2000, 2100, decimal=False),
                help_text="The year the monthly planner and the reports "
                          "open on.")
    r = s.field(r, "Meal planning starts on", cfg.start_date, style="date",
                name="PlanStart", validate=date_check(),
                help_text="The first day of week 1 on the weekly planner. "
                          "Usually a Monday.")
    r = s.field(r, "Week starts on", cfg.week_start, name="WeekStart",
                validate=list_from("WeekStartList"),
                help_text="Used by the planner grids and the week numbers.")
    r = s.field(r, "Shopping day", cfg.shopping_day, name="ShoppingDay",
                validate=list_from("WeekdayList"),
                help_text="The grocery list counts down to this day.")
    r = s.field(r, "Meal prep day", cfg.prep_day, name="PrepDay",
                validate=list_from("WeekdayList"))
    r = s.field(r, "Measurement system", cfg.measurement_system,
                name="MeasureSystem", validate=list_from("SystemList"),
                help_text="Metric shows grams and millilitres, Imperial "
                          "shows ounces and cups. Both always work - this "
                          "only changes what the converter leads with.")
    r = s.field(r, "Default servings per meal", cfg.default_servings,
                style="int", name="DefaultServings",
                validate=number_check(1, 50, decimal=False))
    r = s.field(r, "Shopping list covers from", "=PlanStart", style="date",
                name="ShopFrom", validate=date_check(),
                help_text="The grocery list adds up every meal planned "
                          "between these two dates. Type over the dates "
                          "whenever you want a different window - the "
                          "whole list rebuilds itself.")
    r = s.field(r, "Shopping list covers to", "=ShopFrom+6", style="date",
                name="ShopTo", validate=date_check(),
                help_text="Six days after the start gives you one week. "
                          "Make it +13 for a fortnight's shop.")
    r = s.field(r, "Rounding on the grocery list", cfg.rounding,
                name="RoundingRule", value_end=4,
                validate=list_from("RoundingList"),
                help_text="'Whole packs' turns 1.3 tins into 2 tins - "
                          "what the shop actually sells you. 'Exact "
                          "quantity' leaves the raw number.")

    # ----------------------------------------------------------- 3. budget
    r += 1
    r = s.section(r, "3. Food budget", LAST_COL)
    r = s.field(r, "Weekly food budget", cfg.weekly_food_budget,
                style="money", name="WeeklyBudget",
                validate=number_check(0, 100000),
                help_text="Groceries only - eating out has its own line.")
    r = s.field(r, "Monthly food budget", cfg.monthly_food_budget,
                style="money", name="MonthlyBudget",
                validate=number_check(0, 1000000))
    r = s.field(r, "Yearly food budget", cfg.annual_food_budget,
                style="money", name="AnnualBudget",
                validate=number_check(0, 10000000))
    r = s.field(r, "Eating out / takeaway budget (month)",
                cfg.eating_out_budget, style="money", name="EatingOutBudget",
                validate=number_check(0, 100000))
    r = s.field(r, "Food waste cost you will accept (month)",
                cfg.waste_cost_target, style="money", name="WasteCostTarget",
                validate=number_check(0, 100000),
                help_text="The waste dashboard turns red above this.")

    # -------------------------------------------------------- 4. nutrition
    r += 1
    r = s.section(r, "4. Daily nutrition targets (per person)", LAST_COL)
    r = s.note(r, "Optional. Leave them as they are if you only want the "
                  "planning and the shopping - nothing else depends on "
                  "them. These are general placeholders, not medical "
                  "advice; set your own from your own guidance.",
               LAST_COL, first_col=1, height=30)
    r = s.field(r, "Calories a day", cfg.calorie_target, style="int",
                name="CalorieTarget", validate=number_check(0, 20000,
                                                            decimal=False))
    r = s.field(r, "Protein (g)", cfg.protein_target, style="int",
                name="ProteinTarget",
                validate=number_check(0, 1000, decimal=False))
    r = s.field(r, "Carbohydrates (g)", cfg.carb_target, style="int",
                name="CarbTarget",
                validate=number_check(0, 2000, decimal=False))
    r = s.field(r, "Fat (g)", cfg.fat_target, style="int", name="FatTarget",
                validate=number_check(0, 1000, decimal=False))
    r = s.field(r, "Fibre (g)", cfg.fibre_target, style="int",
                name="FibreTarget",
                validate=number_check(0, 500, decimal=False))
    r = s.field(r, "Sugar limit (g)", cfg.sugar_limit, style="int",
                name="SugarLimit",
                validate=number_check(0, 1000, decimal=False))
    r = s.field(r, "Sodium limit (mg)", cfg.sodium_limit, style="int",
                name="SodiumLimit",
                validate=number_check(0, 20000, decimal=False))
    r = s.field(r, "Water target (glasses a day)", cfg.water_target,
                style="int", name="WaterTarget",
                validate=number_check(0, 40, decimal=False))

    # --------------------------------------------------------- 5. reminders
    r += 1
    r = s.section(r, "5. Reminder windows", LAST_COL)
    r = s.field(r, "Warn me this many days before food expires",
                cfg.expiry_warning_days, style="int",
                name="ExpiryWarningDays",
                validate=number_check(0, 120, decimal=False),
                help_text="Drives the 'use it soon' flags on the fridge, "
                          "freezer and pantry tabs.")
    r = s.field(r, "Days a cooked leftover keeps in the fridge",
                cfg.leftover_safe_days, style="int", name="LeftoverDays",
                validate=number_check(0, 30, decimal=False),
                help_text="A guide for the leftovers tab, not a food "
                          "safety ruling - go by how it looks and smells, "
                          "and by local guidance.")
    r = s.field(r, "Days before a freezer item is 'old'",
                cfg.freezer_review_days, style="int", name="FreezerDays",
                validate=number_check(0, 1000, decimal=False))
    r = s.field(r, "Days of stock that counts as low", cfg.low_stock_days,
                style="int", name="LowStockDays",
                validate=number_check(0, 365, decimal=False))
    r = s.field(r, "Flag prep tasks this many days early",
                cfg.prep_reminder_days, style="int", name="PrepLeadDays",
                validate=number_check(0, 30, decimal=False))
    r = s.field(r, "Days before a meal can repeat", cfg.rotation_gap_days,
                style="int", name="RotationGapDays",
                validate=number_check(0, 365, decimal=False),
                help_text="The rotation planner flags anything cooked "
                          "again sooner than this.")

    # ----------------------------------------------------------- 6. habits
    r += 1
    r = s.section(r, "6. Habit targets", LAST_COL)
    r = s.field(r, "Home cooked meals a week", cfg.home_cooked_target,
                style="int", name="HomeCookedTarget",
                validate=number_check(0, 100, decimal=False))
    r = s.field(r, "New recipes a month", cfg.new_recipe_target,
                style="int", name="NewRecipeTarget",
                validate=number_check(0, 100, decimal=False))
    r = s.field(r, "Takeaways a month (limit)", cfg.takeaway_limit,
                style="int", name="TakeawayLimit",
                validate=number_check(0, 100, decimal=False))
    r = s.field(r, "Packed lunches a week", cfg.packed_lunch_target,
                style="int", name="PackedLunchTarget",
                validate=number_check(0, 100, decimal=False))
    r = s.field(r, "Food waste target (share of spend)",
                cfg.waste_reduction_target, style="percent",
                name="WasteTarget", validate=number_check(0, 1))

    # --------------------------------------------------------- 7. prefixes
    r += 1
    r = s.section(r, "7. ID prefixes", LAST_COL)
    r = s.note(r, "The letters in front of the automatic numbers. Change "
                  "them before you start entering data, not after.",
               LAST_COL, first_col=1, height=24)
    for label, value, name in [
            ("Recipe IDs", "R-", "RecipePrefix"),
            ("Ingredient IDs", "ING-", "IngredientPrefix"),
            ("Planned meal IDs", "M-", "MealPrefix"),
            ("Pantry item IDs", "P-", "PantryPrefix"),
            ("Shopping line IDs", "G-", "GroceryPrefix"),
            ("Leftover IDs", "L-", "LeftoverPrefix"),
            ("Event IDs", "E-", "EventPrefix"),
            ("Purchase IDs", "S-", "SpendPrefix")]:
        r = s.field(r, label, value, name=name)

    # ----------------------------------------------------------- 8. product
    r += 1
    r = s.section(r, "8. About this workbook", LAST_COL)
    r = s.field(r, "Product", cfg.product_name, mode="calc", value_end=5)
    r = s.field(r, "Version", cfg.version, mode="calc")
    r = s.field(r, "Created by", cfg.author, mode="calc")
    r = s.field(r, "Sheet protection password", cfg.password, mode="calc",
                help_text="Review > Unprotect Sheet if you need to change "
                          "a formula.")
    r = s.field(r, "Currency", f"{cfg.currency_symbol} "
                               f"({cfg.currency_code})", mode="calc")
    r += 1
    s.note(r,
           "Changed your mind about something here? Everything updates "
           "straight away - the planner, the grocery list, the budget and "
           "the dashboard all read these cells rather than keeping their "
           "own copy.", LAST_COL, first_col=1, height=30)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
