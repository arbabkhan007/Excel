"""Workbook level defined names that point at the register sheets.

The drop-down lists on the `Settings & Lists` tab define their own
names.  These are the ones that have to follow live data - property IDs,
booking references, guests, cleaners and contractors, the linen cupboard
- so they are built from the registry once every register exists.  Each
one is an ``OFFSET`` range that grows and shrinks with the number of
filled rows, which is why a property you add this morning turns up in
the booking drop-down this afternoon.
"""

from __future__ import annotations

from typing import Dict, Tuple

from .layout import Registry, dynamic_name

#: defined name -> (sheet, column key)
CROSS_SHEET_NAMES: Dict[str, Tuple[str, str]] = {
    # the spine of the workbook - every drop-down that must follow data
    "PropertyList": ("Property Database", "id"),
    "PropertyNameList": ("Property Database", "name"),
    "BookingList": ("Booking Tracker", "id"),
    "GuestList": ("Guest Database", "id"),
    "GuestNameList": ("Guest Database", "name"),
    "VendorList": ("Staff & Vendors", "id"),
    "VendorNameList": ("Staff & Vendors", "name"),
    "SupplyList": ("Supplies Inventory", "id"),
    "SupplyNameList": ("Supplies Inventory", "name"),
    "CleanTaskList": ("Cleaning Schedule", "id"),
    "MaintenanceList": ("Maintenance Log", "id"),
    # property columns the dashboard and the comparison tables scan
    "PropertyStatusCol": ("Property Database", "status"),
    "PropertyNightlyCol": ("Property Database", "nightly_rate"),
    "PropertyRankCol": ("Property Database", "profit_rank"),
    # booking columns - the busiest ranges in the workbook
    "BookingProperty": ("Booking Tracker", "property"),
    "BookingPlatform": ("Booking Tracker", "platform"),
    "BookingStatusCol": ("Booking Tracker", "status"),
    "BookingCheckIn": ("Booking Tracker", "check_in"),
    "BookingCheckOut": ("Booking Tracker", "check_out"),
    "BookingCheckInNum": ("Booking Tracker", "ci_num"),
    "BookingCheckOutNum": ("Booking Tracker", "co_num"),
    "BookingNights": ("Booking Tracker", "nights"),
    "BookingMonths": ("Booking Tracker", "month"),
    "BookingAccommodation": ("Booking Tracker", "accommodation"),
    "BookingGross": ("Booking Tracker", "gross"),
    "BookingNet": ("Booking Tracker", "net_revenue"),
    "BookingHolds": ("Booking Tracker", "holds"),
    "BookingEarns": ("Booking Tracker", "earns"),
    "BookingRowNum": ("Booking Tracker", "seq_num"),
    # the money columns the reports and the tax summary add up
    "ExpenseMonths": ("Expense Tracker", "month"),
    "ExpenseCategoryCol": ("Expense Tracker", "category"),
    "ExpenseAmounts": ("Expense Tracker", "amount"),
    "ExpensePL": ("Expense Tracker", "pl_amount"),
    "ExpenseProperty": ("Expense Tracker", "property"),
    "CleanMonths": ("Cleaning Schedule", "month"),
    "CleanCost": ("Cleaning Schedule", "total_cost"),
    "CleanProperty": ("Cleaning Schedule", "property"),
    "MaintMonths": ("Maintenance Log", "month"),
    "MaintCost": ("Maintenance Log", "cost_used"),
    "MaintProperty": ("Maintenance Log", "property"),
    "PayoutMonths": ("Platform Revenue & Fees", "month"),
    "PayoutReceived": ("Platform Revenue & Fees", "received"),
    "ReviewMonths": ("Reviews Tracker", "month"),
    "ReviewOverall": ("Reviews Tracker", "rate_overall"),
    "ReviewProperty": ("Reviews Tracker", "property"),
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
