"""Demo data: four weeks in the life of the Harper household.

Two adults, a teenager and a six year old, a £180 weekly budget, a
freezer with half a lasagne in it and a standing argument about
mushrooms. Every tab opens with believable numbers so the dashboard,
the grocery list and the reports mean something the moment you open the
file - and `--blank` strips the lot out for your own data.

Everything is dated relative to the build date, so the workbook never
looks stale.
"""

from __future__ import annotations

import datetime as _dt
import random
from dataclasses import dataclass, field
from typing import Any, Dict, List

from .config import Config


def d(base: _dt.date, offset: int) -> _dt.date:
    return base + _dt.timedelta(days=offset)


def week_monday(day: _dt.date) -> _dt.date:
    return day - _dt.timedelta(days=day.weekday())


@dataclass
class Sample:
    members: List[Dict[str, Any]] = field(default_factory=list)
    stores: List[Dict[str, Any]] = field(default_factory=list)
    ingredients: List[Dict[str, Any]] = field(default_factory=list)
    recipes: List[Dict[str, Any]] = field(default_factory=list)
    recipe_lines: List[Dict[str, Any]] = field(default_factory=list)
    meal_plan: List[Dict[str, Any]] = field(default_factory=list)
    pantry: List[Dict[str, Any]] = field(default_factory=list)
    leftovers: List[Dict[str, Any]] = field(default_factory=list)
    waste: List[Dict[str, Any]] = field(default_factory=list)
    spending: List[Dict[str, Any]] = field(default_factory=list)
    store_prices: List[Dict[str, Any]] = field(default_factory=list)
    prep: List[Dict[str, Any]] = field(default_factory=list)
    staples: List[Dict[str, Any]] = field(default_factory=list)
    events: List[Dict[str, Any]] = field(default_factory=list)
    goals: List[Dict[str, Any]] = field(default_factory=list)


# ---------------------------------------------------------------------------
# The price book
# ---------------------------------------------------------------------------
# name, category, aisle, store, unit, pack, price, kcal, pro, carb, fat,
# fibre, allergen, diet, shelf days, storage
INGREDIENTS = [
    ("Chicken breast", "Meat & seafood", "4 - Meat & fish",
     "Greenfield Market", "g", 600, 6.50, 165, 31, 0, 3.6, 0, "None",
     "No restriction", 3, "Fridge"),
    ("Chicken thighs", "Meat & seafood", "4 - Meat & fish",
     "Discount Larder", "g", 900, 4.80, 209, 26, 0, 11, 0, "None",
     "No restriction", 3, "Fridge"),
    ("Beef mince 5%", "Meat & seafood", "4 - Meat & fish",
     "Greenfield Market", "g", 500, 4.75, 137, 21, 0, 5, 0, "None",
     "No restriction", 3, "Fridge"),
    ("Salmon fillet", "Meat & seafood", "4 - Meat & fish",
     "Hill Farm Shop", "g", 240, 5.20, 208, 20, 0, 13, 0, "Fish",
     "Pescatarian", 2, "Fridge"),
    ("Eggs", "Dairy & eggs", "5 - Dairy & eggs", "Hill Farm Shop",
     "each", 12, 3.60, 72, 6.3, 0.4, 4.8, 0, "Eggs", "Vegetarian", 21,
     "Fridge"),
    ("Milk", "Dairy & eggs", "5 - Dairy & eggs", "Corner Express",
     "ml", 2272, 1.65, 42, 3.4, 4.8, 1, 0, "Milk", "Vegetarian", 7,
     "Fridge"),
    ("Cheddar", "Dairy & eggs", "5 - Dairy & eggs", "Greenfield Market",
     "g", 400, 3.40, 402, 25, 1.3, 33, 0, "Milk", "Vegetarian", 21,
     "Fridge"),
    ("Greek yoghurt", "Dairy & eggs", "5 - Dairy & eggs",
     "Greenfield Market", "g", 500, 2.10, 97, 9, 4, 5, 0, "Milk",
     "Vegetarian", 10, "Fridge"),
    ("Butter", "Dairy & eggs", "5 - Dairy & eggs", "Discount Larder",
     "g", 250, 2.25, 717, 0.9, 0.1, 81, 0, "Milk", "Vegetarian", 40,
     "Fridge"),
    ("Parmesan", "Dairy & eggs", "3 - Deli & cheese",
     "Greenfield Market", "g", 150, 3.20, 431, 38, 4.1, 29, 0, "Milk",
     "Vegetarian", 60, "Fridge"),
    ("Onions", "Produce", "1 - Produce", "Discount Larder", "g", 1000,
     0.95, 40, 1.1, 9.3, 0.1, 1.7, "None", "Vegan", 30, "Cupboard"),
    ("Garlic", "Produce", "1 - Produce", "Discount Larder", "each", 3,
     0.70, 4, 0.2, 1, 0, 0.1, "None", "Vegan", 60, "Cupboard"),
    ("Carrots", "Produce", "1 - Produce", "Hill Farm Shop", "g", 1000,
     0.80, 41, 0.9, 9.6, 0.2, 2.8, "None", "Vegan", 21, "Fridge"),
    ("Potatoes", "Produce", "1 - Produce", "Hill Farm Shop", "g", 2500,
     2.20, 77, 2, 17, 0.1, 2.2, "None", "Vegan", 21, "Cupboard"),
    ("Sweet potato", "Produce", "1 - Produce", "Greenfield Market",
     "g", 1000, 1.90, 86, 1.6, 20, 0.1, 3, "None", "Vegan", 14,
     "Cupboard"),
    ("Broccoli", "Produce", "1 - Produce", "Greenfield Market", "g",
     400, 1.10, 34, 2.8, 7, 0.4, 2.6, "None", "Vegan", 6, "Fridge"),
    ("Peppers", "Produce", "1 - Produce", "Greenfield Market", "each",
     3, 1.75, 31, 1, 6, 0.3, 2.1, "None", "Vegan", 8, "Fridge"),
    ("Courgette", "Produce", "1 - Produce", "Hill Farm Shop", "each",
     2, 1.10, 17, 1.2, 3.1, 0.3, 1, "None", "Vegan", 7, "Fridge"),
    ("Spinach", "Produce", "1 - Produce", "Greenfield Market", "g",
     250, 1.60, 23, 2.9, 3.6, 0.4, 2.2, "None", "Vegan", 4, "Fridge"),
    ("Mushrooms", "Produce", "1 - Produce", "Greenfield Market", "g",
     300, 1.20, 22, 3.1, 3.3, 0.3, 1, "None", "Vegan", 5, "Fridge"),
    ("Tomatoes", "Produce", "1 - Produce", "Hill Farm Shop", "g", 500,
     1.45, 18, 0.9, 3.9, 0.2, 1.2, "None", "Vegan", 7, "Fruit bowl"),
    ("Chopped tomatoes", "Pantry staples", "7 - Tins & jars",
     "Discount Larder", "g", 400, 0.55, 32, 1.6, 7, 0.3, 1.5, "None",
     "Vegan", 540, "Cupboard"),
    ("Lemon", "Produce", "1 - Produce", "Greenfield Market", "each", 4,
     1.00, 29, 1.1, 9.3, 0.3, 2.8, "None", "Vegan", 14, "Fruit bowl"),
    ("Bananas", "Produce", "1 - Produce", "Discount Larder", "each", 6,
     1.05, 89, 1.1, 23, 0.3, 2.6, "None", "Vegan", 6, "Fruit bowl"),
    ("Apples", "Produce", "1 - Produce", "Hill Farm Shop", "each", 6,
     1.80, 52, 0.3, 14, 0.2, 2.4, "None", "Vegan", 14, "Fruit bowl"),
    ("Frozen peas", "Frozen", "6 - Frozen", "Discount Larder", "g",
     900, 1.40, 81, 5.4, 14, 0.4, 5.1, "None", "Vegan", 365, "Freezer"),
    ("Frozen berries", "Frozen", "6 - Frozen", "Greenfield Market",
     "g", 500, 2.75, 58, 0.8, 12, 0.4, 3.5, "None", "Vegan", 365,
     "Freezer"),
    ("Pasta", "Pantry staples", "8 - Pasta, rice & grains",
     "Discount Larder", "g", 1000, 1.15, 371, 13, 75, 1.5, 3.2,
     "Cereals with gluten", "Vegan", 540, "Cupboard"),
    ("Rice", "Pantry staples", "8 - Pasta, rice & grains",
     "Discount Larder", "g", 2000, 2.60, 356, 7.5, 78, 0.9, 1.4,
     "None", "Vegan", 720, "Cupboard"),
    ("Couscous", "Pantry staples", "8 - Pasta, rice & grains",
     "Greenfield Market", "g", 500, 1.40, 376, 13, 77, 0.6, 5,
     "Cereals with gluten", "Vegan", 540, "Cupboard"),
    ("Red lentils", "Pantry staples", "7 - Tins & jars",
     "Discount Larder", "g", 500, 1.20, 352, 24, 60, 1.1, 11, "None",
     "Vegan", 720, "Cupboard"),
    ("Chickpeas", "Pantry staples", "7 - Tins & jars",
     "Discount Larder", "g", 400, 0.60, 139, 7.5, 21, 2.4, 6.4, "None",
     "Vegan", 720, "Cupboard"),
    ("Black beans", "Pantry staples", "7 - Tins & jars",
     "Discount Larder", "g", 400, 0.75, 132, 8.9, 24, 0.5, 8.7, "None",
     "Vegan", 720, "Cupboard"),
    ("Coconut milk", "Pantry staples", "7 - Tins & jars",
     "Online Grocer", "ml", 400, 1.30, 197, 2, 3, 20, 0, "None",
     "Vegan", 540, "Cupboard"),
    ("Olive oil", "Pantry staples", "10 - Spices & sauces",
     "Greenfield Market", "ml", 500, 4.50, 884, 0, 0, 100, 0, "None",
     "Vegan", 540, "Cupboard"),
    ("Soy sauce", "Pantry staples", "10 - Spices & sauces",
     "Online Grocer", "ml", 250, 1.85, 53, 8, 4.9, 0.1, 0.8, "Soya",
     "Vegan", 720, "Cupboard"),
    ("Curry paste", "Pantry staples", "10 - Spices & sauces",
     "Greenfield Market", "g", 180, 2.30, 180, 3, 12, 13, 3, "None",
     "Vegetarian", 365, "Cupboard"),
    ("Stock cubes", "Pantry staples", "10 - Spices & sauces",
     "Discount Larder", "each", 12, 1.10, 6, 0.3, 0.9, 0.1, 0, "None",
     "Vegan", 540, "Cupboard"),
    ("Paprika", "Spices & seasonings", "10 - Spices & sauces",
     "Greenfield Market", "g", 50, 1.25, 282, 14, 54, 13, 35, "None",
     "Vegan", 720, "Spice rack"),
    ("Cumin", "Spices & seasonings", "10 - Spices & sauces",
     "Greenfield Market", "g", 45, 1.20, 375, 18, 44, 22, 11, "None",
     "Vegan", 720, "Spice rack"),
    ("Bread", "Bakery", "2 - Bakery", "Corner Express", "g", 800, 1.35,
     265, 9, 49, 3.2, 2.7, "Cereals with gluten", "Vegan", 5,
     "Cupboard"),
    ("Tortilla wraps", "Bakery", "2 - Bakery", "Greenfield Market",
     "each", 8, 1.50, 218, 7, 36, 5, 2, "Cereals with gluten", "Vegan",
     14, "Cupboard"),
    ("Porridge oats", "Pantry staples", "8 - Pasta, rice & grains",
     "Discount Larder", "g", 1000, 1.30, 379, 13, 67, 6.5, 10, "None",
     "Vegan", 365, "Cupboard"),
    ("Peanut butter", "Pantry staples", "9 - Baking", "Online Grocer",
     "g", 340, 2.95, 588, 25, 20, 50, 6, "Peanuts", "Vegan", 365,
     "Cupboard"),
    ("Honey", "Pantry staples", "9 - Baking", "Hill Farm Shop", "g",
     340, 3.80, 304, 0.3, 82, 0, 0.2, "None", "Vegetarian", 720,
     "Cupboard"),
    ("Flour", "Pantry staples", "9 - Baking", "Discount Larder", "g",
     1500, 0.90, 364, 10, 76, 1, 2.7, "Cereals with gluten", "Vegan",
     365, "Cupboard"),
    ("Tofu", "Pantry staples", "5 - Dairy & eggs", "Online Grocer",
     "g", 396, 2.20, 76, 8, 1.9, 4.8, 0.3, "Soya", "Vegan", 14,
     "Fridge"),
]

# ---------------------------------------------------------------------------
# The recipe book
# ---------------------------------------------------------------------------
# name, category, slot, cuisine, serves, prep, cook, method, difficulty,
# diet, freezes, kid friendly, batch friendly, rating, tags
RECIPES = [
    ("Weeknight chicken traybake", "Main", "Dinner", "British", 4, 15,
     45, "Oven", "Easy", "No restriction", "No", "Yes", "Yes", 5,
     "one tray, no washing up"),
    ("Beef bolognese", "Main", "Dinner", "Italian", 6, 20, 60, "Hob",
     "Easy", "No restriction", "Yes", "Yes", "Yes", 5,
     "batch base, freezer friendly"),
    ("Red lentil dhal", "Main", "Dinner", "Indian", 4, 10, 30,
     "One pot", "Easy", "Vegan", "Yes", "No", "Yes", 4,
     "cheap, store cupboard"),
    ("Salmon, new potatoes & greens", "Main", "Dinner", "British", 2,
     10, 25, "Oven", "Easy", "Pescatarian", "No", "No", "No", 4,
     "payday treat"),
    ("Chickpea & spinach curry", "Main", "Dinner", "Indian", 4, 10, 25,
     "One pot", "Easy", "Vegan", "Yes", "No", "Yes", 4,
     "store cupboard"),
    ("Tomato & mascarpone pasta", "Main", "Dinner", "Italian", 4, 5,
     20, "Hob", "Easy", "Vegetarian", "No", "Yes", "No", 4,
     "twenty minutes flat"),
    ("Chicken fajitas", "Main", "Dinner", "Mexican", 4, 15, 20, "Hob",
     "Easy", "No restriction", "No", "Yes", "No", 5, "friday night"),
    ("Black bean chilli", "Main", "Dinner", "Mexican", 6, 15, 40,
     "Slow cooker", "Easy", "Vegan", "Yes", "Yes", "Yes", 4,
     "batch base"),
    ("Veg-packed shepherd's pie", "Main", "Dinner", "British", 6, 25,
     50, "Oven", "Medium", "No restriction", "Yes", "Yes", "Yes", 5,
     "hides the carrots"),
    ("Teriyaki salmon & rice", "Main", "Dinner", "Japanese", 2, 10, 20,
     "Hob", "Easy", "Pescatarian", "No", "No", "No", 4, ""),
    ("Overnight oats", "Breakfast", "Breakfast", "British", 2, 5, 0,
     "No cook", "Easy", "Vegetarian", "No", "Yes", "Yes", 4,
     "make it the night before"),
    ("Scrambled eggs on toast", "Breakfast", "Breakfast", "British", 2,
     5, 8, "Hob", "Easy", "Vegetarian", "No", "Yes", "No", 4, ""),
    ("Banana pancakes", "Breakfast", "Breakfast", "American", 4, 10,
     15, "Hob", "Easy", "Vegetarian", "Yes", "Yes", "No", 5,
     "saturday morning"),
    ("Couscous lunch jars", "Lunchbox", "Lunch", "Middle Eastern", 4,
     15, 10, "No cook", "Easy", "Vegan", "No", "No", "Yes", 4,
     "lunchbox, prep ahead"),
    ("Hummus & veg wraps", "Lunchbox", "Lunch", "Middle Eastern", 2, 10,
     0, "No cook", "Easy", "Vegan", "No", "Yes", "No", 3, "lunchbox"),
    ("Leftover soup", "Soup", "Lunch", "British", 4, 10, 25, "One pot",
     "Easy", "Vegan", "Yes", "No", "Yes", 4, "uses up leftovers"),
    ("Tofu stir fry", "Main", "Dinner", "Chinese", 4, 15, 15, "Hob",
     "Easy", "Vegan", "No", "No", "No", 3, ""),
    ("Apple crumble", "Dessert", "Dessert", "British", 6, 20, 35,
     "Oven", "Easy", "Vegetarian", "Yes", "Yes", "No", 5,
     "sunday pudding"),
]

#: recipe index -> [(ingredient, qty, unit, prep, optional)]
RECIPE_LINES: Dict[int, List[Any]] = {
    0: [("Chicken thighs", 800, "g", "skin on", "No"),
        ("Potatoes", 700, "g", "cut into chunks", "No"),
        ("Peppers", 2, "each", "sliced", "No"),
        ("Onions", 200, "g", "wedges", "No"),
        ("Olive oil", 2, "tbsp", "", "No"),
        ("Paprika", 2, "tsp", "", "No")],
    1: [("Beef mince 5%", 500, "g", "", "No"),
        ("Onions", 150, "g", "finely chopped", "No"),
        ("Carrots", 150, "g", "grated", "No"),
        ("Garlic", 2, "clove", "crushed", "No"),
        ("Chopped tomatoes", 800, "g", "", "No"),
        ("Pasta", 500, "g", "", "No"),
        ("Parmesan", 40, "g", "grated", "Yes")],
    2: [("Red lentils", 300, "g", "rinsed", "No"),
        ("Onions", 150, "g", "chopped", "No"),
        ("Garlic", 3, "clove", "crushed", "No"),
        ("Curry paste", 60, "g", "", "No"),
        ("Coconut milk", 400, "ml", "", "No"),
        ("Spinach", 100, "g", "", "Yes")],
    3: [("Salmon fillet", 240, "g", "", "No"),
        ("Potatoes", 500, "g", "halved", "No"),
        ("Broccoli", 200, "g", "", "No"),
        ("Lemon", 1, "each", "halved", "No"),
        ("Olive oil", 1, "tbsp", "", "No")],
    4: [("Chickpeas", 800, "g", "drained", "No"),
        ("Spinach", 200, "g", "", "No"),
        ("Onions", 150, "g", "chopped", "No"),
        ("Curry paste", 50, "g", "", "No"),
        ("Chopped tomatoes", 400, "g", "", "No"),
        ("Rice", 300, "g", "", "No")],
    5: [("Pasta", 400, "g", "", "No"),
        ("Chopped tomatoes", 400, "g", "", "No"),
        ("Garlic", 2, "clove", "sliced", "No"),
        ("Olive oil", 2, "tbsp", "", "No"),
        ("Parmesan", 50, "g", "grated", "No")],
    6: [("Chicken breast", 500, "g", "sliced", "No"),
        ("Peppers", 3, "each", "sliced", "No"),
        ("Onions", 150, "g", "sliced", "No"),
        ("Tortilla wraps", 8, "each", "", "No"),
        ("Cheddar", 100, "g", "grated", "Yes"),
        ("Paprika", 2, "tsp", "", "No")],
    7: [("Black beans", 800, "g", "drained", "No"),
        ("Chopped tomatoes", 800, "g", "", "No"),
        ("Onions", 200, "g", "chopped", "No"),
        ("Peppers", 2, "each", "diced", "No"),
        ("Cumin", 2, "tsp", "", "No"),
        ("Rice", 400, "g", "", "No")],
    8: [("Beef mince 5%", 500, "g", "", "No"),
        ("Potatoes", 1000, "g", "mashed", "No"),
        ("Carrots", 300, "g", "diced", "No"),
        ("Frozen peas", 200, "g", "", "No"),
        ("Onions", 150, "g", "chopped", "No"),
        ("Stock cubes", 1, "each", "", "No"),
        ("Butter", 40, "g", "for the mash", "No")],
    9: [("Salmon fillet", 240, "g", "", "No"),
        ("Rice", 150, "g", "", "No"),
        ("Soy sauce", 3, "tbsp", "", "No"),
        ("Honey", 1, "tbsp", "", "No"),
        ("Broccoli", 150, "g", "", "No")],
    10: [("Porridge oats", 100, "g", "", "No"),
         ("Milk", 300, "ml", "", "No"),
         ("Greek yoghurt", 100, "g", "", "No"),
         ("Frozen berries", 100, "g", "", "No"),
         ("Honey", 1, "tbsp", "", "Yes")],
    11: [("Eggs", 4, "each", "beaten", "No"),
         ("Bread", 120, "g", "toasted", "No"),
         ("Butter", 20, "g", "", "No"),
         ("Milk", 30, "ml", "", "Yes")],
    12: [("Flour", 200, "g", "", "No"),
         ("Bananas", 2, "each", "mashed", "No"),
         ("Eggs", 2, "each", "", "No"),
         ("Milk", 250, "ml", "", "No"),
         ("Butter", 20, "g", "for the pan", "No")],
    13: [("Couscous", 200, "g", "", "No"),
         ("Chickpeas", 400, "g", "drained", "No"),
         ("Tomatoes", 200, "g", "diced", "No"),
         ("Courgette", 1, "each", "roasted", "No"),
         ("Lemon", 1, "each", "juiced", "No"),
         ("Olive oil", 2, "tbsp", "", "No")],
    14: [("Tortilla wraps", 2, "each", "", "No"),
         ("Chickpeas", 200, "g", "mashed", "No"),
         ("Peppers", 1, "each", "sliced", "No"),
         ("Spinach", 50, "g", "", "No")],
    15: [("Carrots", 300, "g", "chopped", "No"),
         ("Potatoes", 300, "g", "chopped", "No"),
         ("Onions", 150, "g", "chopped", "No"),
         ("Stock cubes", 2, "each", "", "No"),
         ("Olive oil", 1, "tbsp", "", "No")],
    16: [("Tofu", 396, "g", "cubed", "No"),
         ("Broccoli", 200, "g", "florets", "No"),
         ("Peppers", 2, "each", "sliced", "No"),
         ("Soy sauce", 4, "tbsp", "", "No"),
         ("Rice", 300, "g", "", "No"),
         ("Garlic", 2, "clove", "sliced", "No")],
    17: [("Apples", 6, "each", "peeled and sliced", "No"),
         ("Flour", 200, "g", "", "No"),
         ("Butter", 100, "g", "cold, cubed", "No"),
         ("Honey", 60, "g", "", "Yes")],
}

_CACHE: Dict[int, Sample] = {}


def get(cfg: Config) -> Sample:
    """Build (and cache) the demo data for this configuration."""
    key = id(cfg)
    if key not in _CACHE:
        _CACHE[key] = _build(cfg)
    return _CACHE[key]


def _build(cfg: Config) -> Sample:
    rnd = random.Random(20261011)
    today = cfg.today
    monday = week_monday(today)
    start = d(monday, -14)          # two weeks of history, two weeks ahead
    s = Sample()

    # ---------------------------------------------------------- household --
    s.members = [
        {"name": "Maya Harper", "role": "Cook", "age_group": "Adult",
         "portion": 1, "diet": "No restriction", "allergies": "None",
         "intolerances": "", "dislikes": "Olives",
         "loves": "Anything with chilli", "breakfast": "Yes",
         "lunch": "No", "dinner": "Yes", "packed": "No",
         "cal_target": 2000, "cooks": "Yes",
         "cook_days": "Most nights", "shops": "Yes",
         "jobs": "Meal plan, big shop", "birthday": d(start, 120),
         "status": "Active",
         "notes": "Does the planning on a Friday with a coffee."},
        {"name": "Tom Harper", "role": "Shopper", "age_group": "Adult",
         "portion": 1.2, "diet": "No restriction", "allergies": "None",
         "intolerances": "", "dislikes": "Mushrooms",
         "loves": "Curry night", "breakfast": "No", "lunch": "No",
         "dinner": "Yes", "packed": "Yes", "cal_target": 2400,
         "cooks": "Yes", "cook_days": "Tuesday and Thursday",
         "shops": "Yes", "jobs": "Top-up shop, washing up",
         "birthday": d(start, 35), "status": "Active",
         "notes": "Takes a packed lunch four days out of five."},
        {"name": "Rory Harper", "role": "Helper", "age_group": "Teenager",
         "portion": 1.4, "diet": "Vegetarian", "allergies": "None",
         "intolerances": "", "dislikes": "Anything 'mixed up'",
         "loves": "Pasta, pancakes", "breakfast": "Yes", "lunch": "Yes",
         "dinner": "Yes", "packed": "Yes", "cal_target": 2600,
         "cooks": "Yes", "cook_days": "Sunday pudding",
         "shops": "No", "jobs": "Lunchboxes", "birthday": d(start, 200),
         "status": "Active",
         "notes": "Went vegetarian in March - needs a protein swap on "
                  "meat nights."},
        {"name": "Esme Harper", "role": "Child", "age_group": "Child",
         "portion": 0.6, "diet": "No restriction",
         "allergies": "Peanuts", "intolerances": "",
         "dislikes": "Mushrooms, anything spicy", "loves": "Pancakes",
         "breakfast": "Yes", "lunch": "Yes", "dinner": "Yes",
         "packed": "Yes", "cal_target": 1400, "cooks": "No",
         "cook_days": "", "shops": "No", "jobs": "Laying the table",
         "birthday": d(start, 70), "status": "Active",
         "notes": "Peanut allergy - no peanut butter in the house, "
                  "check every label."},
    ]

    # ------------------------------------------------------------ stores --
    s.stores = [
        {"name": "Greenfield Market", "type": "Supermarket",
         "location": "High Street", "travel_min": 12,
         "best_for": "The weekly shop, decent veg", "delivery": "Yes",
         "delivery_fee": 3.50, "min_order": 40,
         "slot": "Saturday 9-10am", "loyalty": "Greenfield points",
         "open_hours": "7am - 10pm", "price_level": "Mid-range",
         "rating": 4, "status": "Active",
         "notes": "Default store for most of the price book."},
        {"name": "Discount Larder", "type": "Discounter",
         "location": "Retail park", "travel_min": 18,
         "best_for": "Store cupboard, tins, frozen", "delivery": "No",
         "delivery_fee": None, "min_order": None, "slot": "",
         "loyalty": "", "open_hours": "8am - 8pm",
         "price_level": "Budget", "rating": 4, "status": "Active",
         "notes": "Worth a monthly stock-up run, not a weekly trip."},
        {"name": "Hill Farm Shop", "type": "Farm shop",
         "location": "Hill Lane", "travel_min": 20,
         "best_for": "Veg box, eggs, meat for special occasions",
         "delivery": "Yes", "delivery_fee": 2.00, "min_order": 25,
         "slot": "Thursday afternoon", "loyalty": "",
         "open_hours": "9am - 5pm", "price_level": "Premium",
         "rating": 5, "status": "Active",
         "notes": "Dearer per 100g but the veg lasts twice as long."},
        {"name": "Corner Express", "type": "Convenience",
         "location": "End of the road", "travel_min": 3,
         "best_for": "Milk and bread emergencies", "delivery": "No",
         "delivery_fee": None, "min_order": None, "slot": "",
         "loyalty": "", "open_hours": "6am - 11pm",
         "price_level": "Premium", "rating": 3, "status": "Active",
         "notes": "Convenient and expensive - the top-up habit to "
                  "break."},
        {"name": "Online Grocer", "type": "Online delivery",
         "location": "Delivered", "travel_min": 0,
         "best_for": "Heavy and bulky, world food aisle",
         "delivery": "Yes", "delivery_fee": 4.95, "min_order": 50,
         "slot": "Wednesday 7-8pm", "loyalty": "Free delivery pass",
         "open_hours": "Any time", "price_level": "Mid-range",
         "rating": 4, "status": "Active", "notes": ""},
    ]

    # ------------------------------------------------------- ingredients --
    for row in INGREDIENTS:
        (name, cat, aisle, store, unit, pack, price, kcal, pro, carb,
         fat, fibre, allergen, diet, shelf, storage) = row
        s.ingredients.append({
            "name": name, "category": cat, "aisle": aisle,
            "store": store, "base_unit": unit, "pack_size": pack,
            "pack_price": price, "calories": kcal, "protein": pro,
            "carbs": carb, "fat": fat, "fibre": fibre,
            "sugar": round(carb * 0.25, 1), "sodium": int(kcal * 0.8),
            "allergen": allergen, "diet": diet, "season": "All year",
            "shelf_days": shelf, "storage": storage,
            "last_bought": d(monday, -rnd.randint(1, 10)),
            "brand": "" if rnd.random() < 0.6 else "Own brand",
            "substitute": "", "notes": "",
        })

    # ----------------------------------------------------------- recipes --
    for idx, row in enumerate(RECIPES):
        (name, cat, slot, cuisine, serves, prep, cook, method, diff,
         diet, freezes, kid, batch, rating, tags) = row
        lines = RECIPE_LINES.get(idx, [])
        s.recipes.append({
            "name": name, "category": cat, "meal_slot": slot,
            "cuisine": cuisine, "servings": serves, "prep_min": prep,
            "cook_min": cook, "method": method, "difficulty": diff,
            "diet": diet, "freezes": freezes, "kid_friendly": kid,
            "prep_friendly": batch, "rating": rating,
            "family_score": max(1, min(5, rating - (idx % 2))),
            "favourite": "Yes" if rating >= 5 else "No",
            "spice": "Mild" if cuisine in ("Indian", "Mexican") else "None",
            "season": "All year", "tags": tags,
            "one_pot": "Yes" if method in ("One pot", "Slow cooker")
                       else "No",
            "no_cook": "Yes" if method == "No cook" else "No",
            "description": f"{name} - {len(lines)} ingredients, serves "
                           f"{serves}.",
            "equipment": {"Oven": "Roasting tin", "Slow cooker":
                          "Slow cooker", "Hob": "Big pan"}.get(method, ""),
            "source": "Family favourite" if rating >= 5 else "",
            "link": "", "photo": "",
            "method_notes": "1. Prep everything first. 2. Cook as usual. "
                            "3. Taste before you serve it.",
            "added": d(start, -rnd.randint(30, 400)),
            "notes": "",
        })
        for ing, qty, unit, prep_note, optional in lines:
            s.recipe_lines.append({
                "recipe": f"R-{idx + 1:04d}", "ingredient": ing,
                "qty": qty, "unit": unit, "prep": prep_note,
                "optional": optional, "section": "",
            })

    # --------------------------------------------------------- meal plan --
    dinners = [0, 1, 2, 5, 6, 3, 8, 7, 4, 9, 16, 1, 0, 2]
    breakfasts = [11, 10, 11, 12, 10, 11, 12]
    lunches = [13, 14, 15, 14, 13, 15, 14]
    statuses = {True: "Eaten", False: "Planned"}
    for day_offset in range(28):
        day = d(start, day_offset)
        past = day < today
        # breakfast on most days
        if day_offset % 7 not in (2, 4):
            r = breakfasts[day_offset % len(breakfasts)]
            s.meal_plan.append({
                "date": day, "slot": "Breakfast",
                "recipe": f"R-{r + 1:04d}", "servings": 3,
                "shop": "Yes" if not past else "No",
                "status": statuses[past], "cook": "Maya Harper",
                "prep_ahead": "No",
                "notes": "",
            })
        # lunch - packed lunches in the week, leftovers at the weekend
        if day.weekday() < 5:
            r = lunches[day_offset % len(lunches)]
            s.meal_plan.append({
                "date": day, "slot": "Lunch", "recipe": f"R-{r + 1:04d}",
                "servings": 2, "shop": "Yes" if not past else "No",
                "status": statuses[past], "cook": "Rory Harper",
                "prep_ahead": "Yes" if day.weekday() == 0 else "No",
                "notes": "Lunchboxes" if day.weekday() < 5 else "",
            })
        else:
            s.meal_plan.append({
                "date": day, "slot": "Lunch", "recipe": None,
                "freetext": "Leftovers from the fridge", "servings": 4,
                "shop": "No", "status": statuses[past],
                "cook": "", "prep_ahead": "No", "notes": "",
            })
        # dinner every day
        if day_offset == 12:
            s.meal_plan.append({
                "date": day, "slot": "Dinner", "recipe": None,
                "freetext": "Takeaway - curry house", "servings": 4,
                "shop": "No", "status": "Takeaway", "cook": "",
                "actual_cost": 38.50, "prep_ahead": "No",
                "notes": "Friday treat. Over budget but worth it."})
        else:
            r = dinners[day_offset % len(dinners)]
            s.meal_plan.append({
                "date": day, "slot": "Dinner", "recipe": f"R-{r + 1:04d}",
                "servings": 4, "shop": "Yes" if not past else "No",
                "status": statuses[past],
                "cook": "Tom Harper" if day.weekday() in (1, 3)
                        else "Maya Harper",
                "leftovers": 2 if r in (1, 7, 8) else None,
                "prep_ahead": "Yes" if r in (1, 7, 8) else "No",
                "event": "Sunday lunch with Nan"
                         if day.weekday() == 6 and day_offset > 20 else "",
                "notes": "",
            })
        # the occasional pudding
        if day.weekday() == 6:
            s.meal_plan.append({
                "date": day, "slot": "Dessert", "recipe": "R-0018",
                "servings": 6, "shop": "Yes" if not past else "No",
                "status": statuses[past], "cook": "Rory Harper",
                "prep_ahead": "No", "notes": ""})

    # ------------------------------------------------------------ pantry --
    pantry_rows = [
        ("Rice", "Cupboard", 1200, "g", "No", 18, 300),
        ("Pasta", "Cupboard", 700, "g", "Yes", 25, 400),
        ("Chopped tomatoes", "Cupboard", 1600, "g", "No", 300, 800),
        ("Chickpeas", "Cupboard", 1200, "g", "No", 400, 400),
        ("Red lentils", "Cupboard", 400, "g", "Yes", 200, 250),
        ("Olive oil", "Cupboard", 300, "ml", "Yes", 200, 250),
        ("Stock cubes", "Cupboard", 7, "each", "Yes", 200, 6),
        ("Paprika", "Spice rack", 30, "g", "Yes", 300, 20),
        ("Cumin", "Spice rack", 12, "g", "Yes", 300, 20),
        ("Flour", "Cupboard", 900, "g", "Yes", 120, 500),
        ("Porridge oats", "Cupboard", 550, "g", "Yes", 90, 400),
        ("Honey", "Cupboard", 210, "g", "Yes", 400, 100),
        ("Soy sauce", "Cupboard", 180, "ml", "Yes", 300, 100),
        ("Curry paste", "Fridge", 90, "g", "Yes", 30, 60),
        ("Milk", "Fridge", 1100, "ml", "Yes", 4, 1000),
        ("Butter", "Fridge", 160, "g", "Yes", 20, 125),
        ("Cheddar", "Fridge", 220, "g", "Yes", 12, 150),
        ("Eggs", "Fridge", 7, "each", "No", 14, 6),
        ("Greek yoghurt", "Fridge", 300, "g", "Yes", 5, 250),
        ("Carrots", "Fridge", 600, "g", "No", 9, 300),
        ("Potatoes", "Cupboard", 1700, "g", "No", 12, 1000),
        ("Onions", "Cupboard", 700, "g", "No", 20, 500),
        ("Spinach", "Fridge", 120, "g", "Yes", 2, 0),
        ("Mushrooms", "Fridge", 150, "g", "Yes", 1, 0),
        ("Frozen peas", "Freezer", 600, "g", "Yes", 200, 400),
        ("Frozen berries", "Freezer", 350, "g", "Yes", 250, 200),
        ("Beef mince 5%", "Freezer", 500, "g", "No", 120, 0),
        ("Chicken thighs", "Freezer", 900, "g", "No", 150, 0),
        ("Bread", "Cupboard", 400, "g", "Yes", 2, 0),
        ("Tortilla wraps", "Cupboard", 4, "each", "Yes", 9, 0),
    ]
    for name, loc, qty, unit, opened, days, minimum in pantry_rows:
        s.pantry.append({
            "ingredient": name, "location": loc, "qty": qty,
            "unit": unit, "opened": opened,
            "bought": d(today, -rnd.randint(2, 20)),
            "expiry": d(today, days) if days <= 30 else None,
            "min_qty": minimum or None,
            "staple": "Yes" if minimum else "No",
            "notes": "Needs using" if days <= 2 else "",
        })

    # --------------------------------------------------------- leftovers --
    leftover_rows = [
        ("Half the bolognese", "R-0002", -2, 3, "Fridge", "Planned"),
        ("Shepherd's pie - two portions", "R-0009", -1, 2, "Fridge",
         "In the fridge"),
        ("Chilli for the freezer", "R-0008", -6, 4, "Freezer", "Frozen"),
        ("Dhal", "R-0003", -3, 2, "Fridge", "Eaten"),
        ("Roast potatoes", None, -4, 1, "Fridge", "Binned"),
        ("Apple crumble", "R-0018", -2, 3, "Fridge", "Eaten"),
        ("Traybake chicken", "R-0001", -1, 2, "Fridge", "Planned"),
        ("Soup", "R-0016", -5, 2, "Freezer", "Frozen"),
    ]
    for item, recipe, offset, portions, where, status in leftover_rows:
        s.leftovers.append({
            "item": item, "recipe": recipe, "cooked": d(today, offset),
            "portions": portions, "location": where,
            "container": "Glass box" if where == "Fridge"
                         else "Freezer bag",
            "planned_for": d(today, 1) if status == "Planned" else None,
            "eaten_by": "Tom Harper" if status == "Planned" else "",
            "status": status,
            "notes": "Label it or it will be a mystery in a month"
                     if where == "Freezer" else "",
        })

    # -------------------------------------------------------- food waste --
    waste_rows = [
        ("Spinach", 120, "g", "Went mouldy", "Yes", "Fridge",
         "Food waste caddy", "Buy the small bag, or freeze half"),
        ("Bread", 300, "g", "Expired", "Yes", "Cupboard", "Compost",
         "Freeze half the loaf on the day we buy it"),
        ("Milk", 500, "ml", "Expired", "Yes", "Fridge",
         "General bin", "Check the date before the top-up shop"),
        ("Mushrooms", 150, "g", "Forgot about it", "Yes", "Fridge",
         "Compost", "Keep them at eye level, not in the drawer"),
        ("Potatoes", 400, "g", "Went mouldy", "Yes", "Cupboard",
         "Compost", "Smaller bag from the farm shop"),
        ("Greek yoghurt", 200, "g", "Expired", "Yes", "Fridge",
         "General bin", "One pot a week is enough"),
        ("Carrots", 200, "g", "Plate scrapings", "No", "Fridge",
         "Food waste caddy", "Esme's portion is too big"),
        ("Tomatoes", 150, "g", "Bought too much", "Yes", "Fruit bowl",
         "Compost", "Only buy them when a recipe needs them"),
    ]
    for i, (item, qty, unit, reason, avoidable, where, disposal,
            prevent) in enumerate(waste_rows):
        s.waste.append({
            "date": d(today, -(i * 3 + 2)), "item": item, "qty": qty,
            "unit": unit, "reason": reason, "avoidable": avoidable,
            "where": where, "disposal": disposal, "prevent": prevent,
        })

    # ---------------------------------------------------------- spending --
    spend_rows = [
        (-21, "Greenfield Market", "Weekly shop", "Weekly shop", 142.80,
         "Yes", "Weekly shop"),
        (-19, "Corner Express", "Top-up shop", "Milk and bread", 6.40,
         "No", "Top-up shop"),
        (-17, "Hill Farm Shop", "Top-up shop", "Veg box", 18.50, "Yes",
         "Weekly shop"),
        (-14, "Greenfield Market", "Weekly shop", "Weekly shop", 151.25,
         "Yes", "Weekly shop"),
        (-12, "Online Grocer", "Bulk shop", "Store cupboard stock-up",
         62.40, "Yes", "Weekly shop"),
        (-11, "Corner Express", "Top-up shop", "Milk", 2.10, "No",
         "Top-up shop"),
        (-9, "Curry House", "Takeaway", "Friday takeaway", 38.50, "No",
         "Takeaway"),
        (-7, "Greenfield Market", "Weekly shop", "Weekly shop", 138.60,
         "Yes", "Weekly shop"),
        (-6, "Discount Larder", "Bulk shop", "Tins and frozen", 31.20,
         "Yes", "Weekly shop"),
        (-5, "Corner Express", "Top-up shop", "Bread", 1.35, "No",
         "Top-up shop"),
        (-4, "The Lantern", "Eating out", "Birthday lunch", 64.00,
         "Yes", "Special occasion"),
        (-2, "Greenfield Market", "Weekly shop", "Weekly shop", 147.95,
         "Yes", "Weekly shop"),
        (-1, "Hill Farm Shop", "Top-up shop", "Eggs and veg", 12.75,
         "Yes", "Weekly shop"),
    ]
    for offset, store, kind, item, amount, planned, line in spend_rows:
        s.spending.append({
            "date": d(today, offset), "store": store, "type": kind,
            "item": item, "amount": amount, "payer": "Maya Harper"
            if amount > 20 else "Tom Harper",
            "method": "Card", "planned": planned, "budget_line": line,
            "waste": "No", "notes": "",
        })

    # ------------------------------------------------------ store prices --
    compare = ["Chicken breast", "Beef mince 5%", "Pasta", "Rice",
               "Milk", "Cheddar", "Olive oil", "Eggs", "Chopped tomatoes",
               "Porridge oats", "Butter", "Bread"]
    lookup = {row[0]: row for row in INGREDIENTS}
    for name in compare:
        base = lookup[name]
        pack, price, unit = base[5], base[6], base[4]
        for store, factor in (("Greenfield Market", 1.0),
                              ("Discount Larder", 0.82),
                              ("Corner Express", 1.28)):
            s.store_prices.append({
                "ingredient": name, "store": store,
                "brand": "Own brand" if factor < 1 else "Branded",
                "pack_size": pack, "price": round(price * factor, 2),
                "offer": "No",
                "checked": d(today, -rnd.randint(3, 60)),
                "notes": "",
            })

    # --------------------------------------------------------- meal prep --
    prep_rows = [
        ("Big pot of bolognese", "R-0002", 2, "Freezer", 6,
         "Maya Harper", "Done"),
        ("Chilli for the freezer", "R-0008", 2, "Freezer", 6,
         "Tom Harper", "Done"),
        ("Couscous lunch jars", "R-0014", 1, "Fridge", 4,
         "Rory Harper", "Done"),
        ("Chop all the veg for the week", None, 1, "Fridge", 0,
         "Maya Harper", "Done"),
        ("Overnight oats x4", "R-0011", 2, "Fridge", 4, "Rory Harper",
         "To do"),
        ("Shepherd's pie - double batch", "R-0009", 2, "Freezer", 4,
         "Maya Harper", "To do"),
        ("Pancake batter for Saturday", "R-0013", 1, "Fridge", 1,
         "Rory Harper", "To do"),
    ]
    for i, (job, recipe, batches, where, containers, who,
            status) in enumerate(prep_rows):
        day = d(monday, -7 + (i * 3))
        s.prep.append({
            "job": job, "date": day, "recipe": recipe,
            "batches": batches, "storage": where,
            "containers": containers or None,
            "covers_from": day, "covers_to": d(day, 6),
            "cook": who, "status": status,
            "minutes": 45 if batches > 1 else 25,
            "notes": "",
        })

    # ---------------------------------------------------------- staples --
    staple_rows = [
        ("Rice", 1000, 250), ("Pasta", 1000, 400),
        ("Chopped tomatoes", 1600, 400), ("Olive oil", 500, 120),
        ("Stock cubes", 10, 3), ("Flour", 1000, 150),
        ("Porridge oats", 1000, 350), ("Milk", 2000, 1800),
        ("Butter", 250, 80), ("Eggs", 12, 8),
        ("Chickpeas", 1200, 300), ("Red lentils", 500, 150),
        ("Frozen peas", 900, 200), ("Soy sauce", 250, 40),
        ("Honey", 340, 40),
    ]
    for name, par, usage in staple_rows:
        s.staples.append({
            "ingredient": name, "par": par, "usage": usage,
            "bulk": "Yes" if par >= 1000 else "No", "notes": "",
        })

    # ----------------------------------------------------------- events --
    s.events = [
        {"name": "Sunday lunch with Nan", "type": "Sunday lunch",
         "date": d(monday, 6), "time": "13:00", "guests": 6,
         "adults": 5, "children": 1, "dietary": "Nan does not eat pork",
         "budget": 35.00, "menu": "Roast chicken, all the trimmings, "
         "apple crumble", "prep_start": d(monday, 5),
         "shopping_done": "No", "host": "Maya Harper",
         "status": "Planned",
         "notes": "Rory is doing the pudding."},
        {"name": "Esme's birthday tea", "type": "Birthday",
         "date": d(today, 23), "time": "16:00", "guests": 12,
         "adults": 4, "children": 8,
         "dietary": "Two peanut allergies, one dairy free",
         "budget": 60.00, "menu": "Pancakes, fruit, chocolate cake",
         "prep_start": d(today, 21), "shopping_done": "No",
         "host": "Maya Harper", "status": "Planned",
         "notes": "Nothing with peanuts in the house that day."},
        {"name": "Book club", "type": "Dinner party",
         "date": d(today, 9), "time": "19:30", "guests": 7,
         "adults": 7, "children": 0, "dietary": "One vegan",
         "budget": 45.00, "menu": "Dhal, flatbreads, salad",
         "prep_start": d(today, 9), "shopping_done": "No",
         "host": "Maya Harper", "status": "Menu set", "notes": ""},
        {"name": "Tom's work barbecue", "type": "Barbecue",
         "date": d(today, -10), "time": "12:00", "guests": 15,
         "adults": 12, "children": 3, "dietary": "",
         "budget": 70.00, "actual_cost": 82.40,
         "menu": "Burgers, salads, the usual", "prep_start": d(today, -11),
         "shopping_done": "Yes", "host": "Tom Harper", "status": "Done",
         "notes": "Bought far too much bread again."},
    ]

    # ------------------------------------------------------------ goals --
    s.goals = [
        {"goal": "Cook at home six nights a week", "area": "Habit",
         "measure": "Home cooked meals a week", "target": 18,
         "direction": "Hit or beat", "why": "Takeaways were creeping "
         "back up to twice a week", "start": d(today, -60),
         "review": d(today, 7), "owner": "Maya Harper",
         "status": "In progress", "notes": ""},
        {"goal": "Keep the grocery shop under £780 a month",
         "area": "Budget", "measure": "Grocery spend this month",
         "target": 780, "direction": "Stay under",
         "why": "It drifted to £900 last winter", "start": d(today, -90),
         "review": d(today, 14), "owner": "Maya Harper",
         "status": "On track", "notes": ""},
        {"goal": "Two new recipes a month", "area": "Variety",
         "measure": "New recipes this month", "target": 2,
         "direction": "Hit or beat",
         "why": "We were eating the same six dinners on a loop",
         "start": d(today, -120), "review": d(today, 21),
         "owner": "Rory Harper", "status": "In progress", "notes": ""},
        {"goal": "Waste less than £15 of food a month", "area": "Waste",
         "measure": "Food waste this month", "target": 15,
         "direction": "Stay under",
         "why": "The salad drawer was a graveyard",
         "start": d(today, -45), "review": d(today, -2),
         "owner": "Tom Harper", "status": "Behind",
         "notes": "Mostly bagged salad and bread."},
        {"goal": "Four takeaways a month, no more", "area": "Budget",
         "measure": "Takeaways a month", "target": 4,
         "direction": "Stay under", "why": "£40 a go adds up fast",
         "start": d(today, -30), "review": d(today, 10),
         "owner": "Tom Harper", "status": "On track", "notes": ""},
        {"goal": "Five packed lunches a week", "area": "Health",
         "measure": "Packed lunches a week", "target": 5,
         "direction": "Hit or beat",
         "why": "Meal deals were £7 a day", "start": d(today, -20),
         "review": d(today, 12), "owner": "Rory Harper",
         "status": "In progress", "notes": ""},
    ]

    return s
