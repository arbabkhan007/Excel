"""Workbook builder: puts the twenty tabs together in the right order.

Typical use::

    from salon_business_spreadsheet import build_workbook
    build_workbook("Salon Business Management Spreadsheet.xlsx")
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
from .sheets import (analytics, appointments, clients, commissions,
                     dashboard, expenses, goals, inventory, loyalty,
                     marketing, payments, payroll, profit_loss,
                     purchases, reports, sales, services, settings,
                     staff, start_here)
from .writer import SheetWriter

#: register modules, in the order their specs are registered
REGISTER_MODULES = [
    services, clients, staff, appointments, sales, payments,
    commissions, payroll, expenses, inventory, purchases, marketing,
    loyalty,
]

#: tab order of the finished workbook
SHEET_ORDER: List[str] = [
    "Start Here",
    "Dashboard",
    "Appointments",
    "Clients",
    "Services",
    "Sales",
    "Payments",
    "Staff",
    "Commissions",
    "Payroll",
    "Expenses",
    "Profit & Loss",
    "Inventory",
    "Purchases",
    "Marketing",
    "Loyalty",
    "Goals & Budget",
    "Analytics",
    "Reports",
    "Settings",
]

#: sheets that are not registers - name -> builder(book, fmts, cfg, reg)
PAGE_BUILDERS: Dict[str, Callable] = {
    "Start Here": start_here.build,
    "Dashboard": dashboard.build,
    "Profit & Loss": profit_loss.build,
    "Goals & Budget": goals.build,
    "Analytics": analytics.build,
    "Reports": reports.build,
    "Settings": settings.build,
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


def build_workbook(path: str = "Salon Business Spreadsheet.xlsx",
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

    # 3. Live lookup names (client IDs, service IDs, the team).
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
        "subject": "Salon appointments, clients, pricing, sales, staff "
                   "commission, stock and profit",
        "author": cfg.author,
        "manager": cfg.author,
        "company": cfg.author,
        "category": "Small business spreadsheet template",
        "keywords": ("salon spreadsheet, hair salon business planner, "
                     "appointment tracker, client CRM, salon pricing "
                     "calculator, stylist commission calculator, "
                     "salon profit and loss, beauty business "
                     "spreadsheet, nail tech planner, lash artist "
                     "booking tracker, spreadsheet template"),
        "comments": (f"{cfg.product_name}. Created by {cfg.author}. "
                     f"Version {cfg.version}. Formula cells are locked "
                     f"and every sheet is protected with the password "
                     f"'{cfg.password}'."),
        "status": "Final",
        "created": _dt.datetime(cfg.today.year, cfg.today.month,
                                cfg.today.day),
    })
