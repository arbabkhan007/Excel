"""Workbook builder: puts the twenty-three tabs together in order.

Typical use::

    from daycare_business_spreadsheet import build_workbook
    build_workbook("Daycare Business Management Spreadsheet.xlsx")
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
from .sheets import (annual, attendance, children, classrooms,
                     communication, daily_reports, dashboard,
                     documents, expenses, getting_started, incidents,
                     invoices, meals, medical, monthly, parents,
                     payroll, schedule, settings, staff, supplies,
                     tuition, waitlist)
from .writer import SheetWriter

#: register modules, in the order their specs are registered
REGISTER_MODULES = [
    classrooms, parents, children, waitlist, attendance, tuition,
    invoices, daily_reports, meals, medical, incidents, staff,
    schedule, payroll, expenses, supplies, communication, documents,
]

#: tab order of the finished workbook
SHEET_ORDER: List[str] = [
    "Dashboard",
    "Getting Started",
    "Settings & Lists",
    "Child Enrollment",
    "Parent Database",
    "Waitlist",
    "Attendance",
    "Tuition & Fees",
    "Invoices & Receipts",
    "Daily Reports",
    "Meal Tracker",
    "Allergy & Medical",
    "Incident Reports",
    "Staff Database",
    "Staff Schedule",
    "Payroll Tracker",
    "Classroom Management",
    "Expense Tracker",
    "Supply Inventory",
    "Parent Communication",
    "Document Tracker",
    "Monthly Reports",
    "Annual Reports",
]

#: sheets that are not registers - name -> builder(book, fmts, cfg, reg)
PAGE_BUILDERS: Dict[str, Callable] = {
    "Dashboard": dashboard.build,
    "Getting Started": getting_started.build,
    "Settings & Lists": settings.build,
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


def build_workbook(path: str = "Daycare Business Spreadsheet.xlsx",
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

    # 3. Live lookup names (child IDs, family IDs, rooms, the team).
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
        "subject": "Childcare enrollment, attendance, tuition, "
                   "invoicing, staff, classrooms and parent records",
        "author": cfg.author,
        "manager": cfg.author,
        "company": cfg.author,
        "category": "Small business spreadsheet template",
        "keywords": ("daycare spreadsheet, childcare management "
                     "system, child care business planner, enrollment "
                     "tracker, attendance tracker, tuition invoice "
                     "template, preschool spreadsheet, nursery "
                     "management, childminder tracker, staff "
                     "scheduling, classroom ratio, spreadsheet "
                     "template"),
        "comments": (f"{cfg.product_name}. Created by {cfg.author}. "
                     f"Version {cfg.version}. Formula cells are locked "
                     f"and every sheet is protected with the password "
                     f"'{cfg.password}'."),
        "status": "Final",
        "created": _dt.datetime(cfg.today.year, cfg.today.month,
                                cfg.today.day),
    })
