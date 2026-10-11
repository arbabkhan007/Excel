"""Workbook builder: puts every tab together in the right order.

Typical use::

    from meal_planner_spreadsheet import build_workbook
    build_workbook("Meal Planner & Grocery List.xlsx")
"""

from __future__ import annotations

import datetime as _dt
import inspect
import os
from typing import Callable, Dict, List, Optional

from xlsxwriter.workbook import Workbook

from .config import Config
from .formats import FormatBook
from .layout import Registry, SheetSpec
from .names import define_cross_sheet_names
from .sheets import (events, goals, grocery, ingredients, leftovers, lists,
                     meal_plan, members, pantry, prep, recipe_lines,
                     recipes, settings, spending, staples, store_prices,
                     stores, waste, weekly)
from .writer import SheetWriter

#: register modules, in the order their specs are registered
REGISTER_MODULES = [
    members, recipes, recipe_lines, ingredients, meal_plan, grocery,
    pantry, leftovers, waste, spending, store_prices, stores, prep,
    staples, events, goals,
]

#: tab order of the finished workbook
SHEET_ORDER: List[str] = [
    "Weekly Planner", "Settings", "Lists",
    "Household", "Recipes", "Recipe Ingredients", "Ingredients",
    "Meal Plan", "Grocery List", "Pantry", "Staples", "Leftovers",
    "Food Waste", "Meal Prep", "Grocery Spending", "Store Prices",
    "Stores", "Events", "Goals",
]

#: sheets that are not registers - name -> builder(book, fmts, cfg, reg)
PAGE_BUILDERS: Dict[str, Callable] = {
    "Weekly Planner": weekly.build,
    "Settings": settings.build,
    "Lists": lists.build,
}

#: register modules by sheet name
REGISTER_BY_NAME = {module.NAME: module for module in REGISTER_MODULES}


# ---------------------------------------------------------------------------

def build_register(book, fmts: FormatBook, cfg: Config, reg: Registry,
                   module) -> SheetWriter:
    """Create one register sheet from its :class:`SheetSpec`."""
    spec: SheetSpec = reg[module.NAME]
    sheet = SheetWriter(book, fmts, cfg, reg, spec.name,
                        tab_color=spec.tab_color)
    sheet.grid(spec)
    decorate = getattr(module, "decorate", None)
    if decorate is not None:
        wants_registry = len(inspect.signature(decorate).parameters) >= 4
        if wants_registry:
            decorate(sheet, fmts, cfg, reg)
        else:
            decorate(sheet, fmts, cfg)
    sheet.protect()
    return sheet


def build_workbook(path: str = "Meal Planner & Grocery List.xlsx",
                   cfg: Optional[Config] = None) -> str:
    """Write the whole product to *path* and return the path."""
    cfg = cfg or Config()
    folder = os.path.dirname(os.path.abspath(path))
    if folder and not os.path.isdir(folder):
        os.makedirs(folder)

    book = Workbook(path)
    _set_properties(book, cfg)

    fmts = FormatBook(book, cfg)
    reg = Registry()

    # 1. Register every spec first so cross sheet formulas resolve no
    #    matter which tab happens to be written first.
    for module in REGISTER_MODULES:
        reg.add(module.spec(cfg))

    # 2. Write the tabs in their final order.
    sheets: Dict[str, SheetWriter] = {}
    for name in SHEET_ORDER:
        if name in REGISTER_BY_NAME:
            sheets[name] = build_register(book, fmts, cfg, reg,
                                          REGISTER_BY_NAME[name])
        elif name in PAGE_BUILDERS:
            sheets[name] = PAGE_BUILDERS[name](book, fmts, cfg, reg)
        else:                                   # pragma: no cover
            raise KeyError(f"No builder for sheet '{name}'")

    # 3. Live lookup names (recipe IDs, ingredient names, who lives here).
    define_cross_sheet_names(book, reg)

    # 4. Final touches: open on the first page.
    first = sheets[SHEET_ORDER[0]].ws
    first.activate()
    first.set_first_sheet()

    book.close()
    return path


def _set_properties(book, cfg: Config) -> None:
    book.set_properties({
        "title": cfg.product_name,
        "subject": "Meal planning, grocery lists and food budgeting",
        "author": cfg.author,
        "manager": cfg.author,
        "company": cfg.author,
        "category": "Household spreadsheet template",
        "keywords": ("meal planner, weekly meal planner, monthly meal "
                     "planner, recipe organizer, grocery list, shopping "
                     "list, pantry inventory, food budget, meal prep, "
                     "nutrition tracker, spreadsheet template"),
        "comments": (f"{cfg.product_name}. Created by {cfg.author}. "
                     f"Version {cfg.version}. Formula cells are locked "
                     f"and every sheet is protected with the password "
                     f"'{cfg.password}'."),
        "status": "Final",
        "created": _dt.datetime(cfg.today.year, cfg.today.month,
                                cfg.today.day),
    })
