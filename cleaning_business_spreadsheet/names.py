"""Workbook level defined names that point at the register sheets.

The drop down lists on the `Lists` tab define their own names. These are
the ones that have to follow live data - client IDs, property references,
the people on the payroll - so they are built from the registry once
every register exists. Each one is an ``OFFSET`` range that grows and
shrinks with the number of filled rows, which is why a new client appears
in the job scheduler's drop down the moment you type their name.
"""

from __future__ import annotations

from typing import Dict, Tuple

from .layout import Registry, dynamic_name

#: defined name -> (sheet, column key)
CROSS_SHEET_NAMES: Dict[str, Tuple[str, str]] = {
    # the spine of the whole workbook
    "ClientIDs": ("Clients", "id"),
    "ClientNames": ("Clients", "name"),
    "PropertyIDs": ("Properties", "id"),
    "PropertyNames": ("Properties", "name"),
    "ServiceNames": ("Services", "name"),
    "ServiceIDs": ("Services", "id"),
    "JobIDs": ("Jobs", "id"),
    "QuoteIDs": ("Quotes", "id"),
    "InvoiceIDs": ("Invoices", "id"),
    "EmployeeNames": ("Employees", "name"),
    "EmployeeIDs": ("Employees", "id"),
    "SupplyNames": ("Supplies", "name"),
    "EquipmentNames": ("Equipment", "name"),
    "LeadIDs": ("Leads", "id"),
    # columns the dashboard, the calendar and the reports scan
    "JobDates": ("Jobs", "date"),
    "JobClients": ("Jobs", "client"),
    "JobStatuses": ("Jobs", "status"),
    "JobMonths": ("Jobs", "month"),
    "InvoiceMonths": ("Invoices", "month"),
    "ExpenseMonths": ("Expenses", "month"),
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
