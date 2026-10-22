"""Workbook level defined names that point at the register sheets.

The drop-down lists on the `Settings` tab define their own names.
These are the ones that have to follow live data - client IDs,
service IDs, the team, the product list - so they are built from the
registry once every register exists. Each one is an ``OFFSET`` range
that grows and shrinks with the number of filled rows, which is why a
new client appears in the appointment drop-down the moment you type
their name.
"""

from __future__ import annotations

from typing import Dict, Tuple

from .layout import Registry, dynamic_name

#: defined name -> (sheet, column key)
CROSS_SHEET_NAMES: Dict[str, Tuple[str, str]] = {
    # the spine of the workbook - every drop-down that must follow data
    "ClientList": ("Clients", "id"),
    "ClientNames": ("Clients", "name"),
    "ServiceList": ("Services", "id"),
    "ServiceNames": ("Services", "name"),
    "StaffList": ("Staff", "id"),
    "StaffNames": ("Staff", "name"),
    "ProductList": ("Inventory", "id"),
    "ProductNames": ("Inventory", "name"),
    "AppointmentIDs": ("Appointments", "id"),
    "SaleIDs": ("Sales", "id"),
    "PromoCodes": ("Marketing", "code"),
    # columns the dashboard, the analytics and the reports scan
    "ServiceCategories": ("Services", "category"),
    "ServiceRevenueCol": ("Services", "revenue"),
    "SaleDates": ("Sales", "date"),
    "SaleMonths": ("Sales", "month"),
    "SaleNet": ("Sales", "net"),
    "SaleStaff": ("Sales", "staff"),
    "SaleKinds": ("Sales", "kind"),
    "AppointmentDates": ("Appointments", "date"),
    "AppointmentStatus": ("Appointments", "status"),
    "ClientSegments": ("Clients", "segment"),
    "ExpenseMonths": ("Expenses", "month"),
    "ExpenseCategoriesCol": ("Expenses", "category"),
    "ExpenseAmounts": ("Expenses", "amount"),
}


def define_cross_sheet_names(book, reg: Registry) -> None:
    """Create the live lookup names. Call after every register is built."""
    for name, (sheet, key) in CROSS_SHEET_NAMES.items():
        if sheet not in reg:
            continue
        spec = reg[sheet]
        book.define_name(
            name,
            dynamic_name(sheet, spec.letter(key), spec.first_data_row,
                         spec.last_data_row))
