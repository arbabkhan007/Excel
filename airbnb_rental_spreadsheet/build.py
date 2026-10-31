"""Workbook builder: puts the twenty tabs together in order.

Typical use::

    from airbnb_rental_spreadsheet import build_workbook
    build_workbook("Airbnb Rental Management Spreadsheet.xlsx")
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
from .sheets import (annual, arrivals, bookings, calendar, cleaning,
                     dashboard, expenses, getting_started, guests,
                     maintenance, monthly, payouts, performance,
                     properties, reviews, settings, supplies, tax,
                     turnover, vendors)
from .writer import SheetWriter

#: register modules, in the order their specs are registered
REGISTER_MODULES = [
    properties, bookings, guests, arrivals, cleaning, payouts,
    expenses, supplies, maintenance, reviews, vendors, performance,
]

#: tab order of the finished workbook
SHEET_ORDER: List[str] = [
    "Dashboard",
    "Getting Started",
    "Settings & Lists",
    "Property Database",
    "Booking Tracker",
    "Booking Calendar",
    "Guest Database",
    "Check-in & Check-out",
    "Cleaning Schedule",
    "Turnover Checklist",
    "Platform Revenue & Fees",
    "Expense Tracker",
    "Supplies Inventory",
    "Maintenance Log",
    "Reviews Tracker",
    "Staff & Vendors",
    "Tax Summary",
    "Occupancy & Performance",
    "Monthly Reports",
    "Annual Reports",
]

#: sheets that are not registers - name -> builder(book, fmts, cfg, reg)
PAGE_BUILDERS: Dict[str, Callable] = {
    "Dashboard": dashboard.build,
    "Getting Started": getting_started.build,
    "Settings & Lists": settings.build,
    "Booking Calendar": calendar.build,
    "Turnover Checklist": turnover.build,
    "Tax Summary": tax.build,
    "Monthly Reports": monthly.build,
    "Annual Reports": annual.build,
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


def build_workbook(path: str = "Airbnb Rental Spreadsheet.xlsx",
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

    # 3. Live lookup names (property IDs, booking IDs, guests, vendors).
    define_cross_sheet_names(book, reg)

    # 4. Final touches: open on the dashboard.
    first = sheets[SHEET_ORDER[0]].ws
    first.activate()
    first.set_first_sheet()

    book.close()
    return path


def _set_properties(book, cfg: Config) -> None:
    book.set_properties({
        "title": cfg.product_name,
        "subject": "Short-term rental bookings, cleaning turnovers, "
                   "platform fees, expenses, occupancy and profit",
        "author": cfg.author,
        "manager": cfg.author,
        "company": cfg.author,
        "category": "Small business spreadsheet template",
        "keywords": ("airbnb spreadsheet, short term rental tracker, "
                     "vacation rental management, booking calendar, "
                     "cleaning schedule, guest tracker, occupancy "
                     "calculator, ADR RevPAR, str investor "
                     "spreadsheet, vrbo tracker, property management "
                     "template, rental income expense tracker"),
        "comments": (f"{cfg.product_name}. Created by {cfg.author}. "
                     f"Version {cfg.version}. Formula cells are locked "
                     f"and every sheet is protected with the password "
                     f"'{cfg.password}'. Not affiliated with or "
                     f"connected to Airbnb, Vrbo or Booking.com."),
        "status": "Final",
        "created": _dt.datetime(cfg.today.year, cfg.today.month,
                                cfg.today.day),
    })
