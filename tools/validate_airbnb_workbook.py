"""Static checks on a generated Airbnb & Short-Term Rental workbook.

Run it with any Python that has openpyxl installed::

    python tools/validate_airbnb_workbook.py output/airbnb.xlsx

It never opens Excel - it parses the XML and reports anything that would
break for a buyer: unresolved template tokens, formulas pointing at the
wrong sheet, defined names that do not exist, missing sheet protection,
formula cells that are not locked, and input cells that are.
"""

from __future__ import annotations

import os
import re
import sys
import zipfile
from collections import Counter, defaultdict

try:
    import openpyxl
except ImportError:                                   # pragma: no cover
    print("openpyxl is required: pip install openpyxl")
    raise SystemExit(2)

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from airbnb_rental_spreadsheet.build import SHEET_ORDER as EXPECTED_SHEETS  # noqa: E402,E501

FUNCTION_RE = re.compile(r"\b([A-Z][A-Z0-9\.]{1,})\s*\(")
NAME_RE = re.compile(r"(?<![A-Za-z0-9_.!$\"])([A-Za-z_][A-Za-z0-9_.]{2,})"
                     r"(?![A-Za-z0-9_.(!])")
#: a column letter with no row number - almost always a missing {r}
BARE_COL_RE = re.compile(r"(?<![A-Za-z0-9_.!$\"])([A-Z]{1,3})"
                         r"(?![A-Za-z0-9_.(!$:])")

KNOWN_FUNCTIONS = {
    "IF", "IFERROR", "SUM", "SUMIF", "SUMIFS", "COUNT", "COUNTA", "COUNTIF",
    "COUNTIFS", "AVERAGE", "AVERAGEIF", "AVERAGEIFS", "MAX", "MIN", "ROUND",
    "ROUNDUP", "ROUNDDOWN", "INDEX", "MATCH", "OFFSET", "TEXT", "VALUE",
    "LEFT", "RIGHT", "MID", "LEN", "TRIM", "UPPER", "LOWER", "PROPER",
    "CONCATENATE", "TODAY", "NOW", "DATE", "DAY", "MONTH", "YEAR", "WEEKDAY",
    "WEEKNUM", "EOMONTH", "EDATE", "DATEDIF", "DAYS", "NETWORKDAYS", "HOUR",
    "MINUTE", "TIME", "SUMPRODUCT", "LARGE", "SMALL", "RANK", "ABS", "AND",
    "OR", "NOT", "ISBLANK", "ISNUMBER", "ISTEXT", "ISERROR", "ISNA", "NA",
    "CHOOSE", "VLOOKUP", "HLOOKUP", "LOOKUP", "SUBSTITUTE", "REPT",
    "MEDIAN", "STDEV", "FLOOR", "CEILING", "INT", "MOD", "SIGN", "RAND",
    "SEARCH", "FIND", "EXACT", "N", "T", "TRUE", "FALSE", "ROW", "ROWS",
    "COLUMN", "COLUMNS", "YEARFRAC", "INDIRECT", "ADDRESS", "CELL",
    "SUBTOTAL", "TRANSPOSE", "PRODUCT", "POWER", "SQRT", "COUNTBLANK",
    "ISERR", "ISLOGICAL", "ISREF", "TYPE", "NUMBERVALUE", "DATEVALUE",
    "TIMEVALUE", "WORKDAY", "FREQUENCY", "PERCENTILE",
}

#: functions that do not exist in Excel 2010 / Google Sheets
MODERN_ONLY = {
    "XLOOKUP", "IFS", "MAXIFS", "MINIFS", "TEXTJOIN", "CONCAT", "SWITCH",
    "LET", "LAMBDA", "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE",
    "TEXTSPLIT", "AGGREGATE", "IFNA",
}


def _colnum(ref: str) -> tuple:
    match = re.match(r"([A-Z]+)(\d+)", ref)
    letters, row = match.group(1), int(match.group(2))
    num = 0
    for char in letters:
        num = num * 26 + (ord(char) - 64)
    return num, row


def check_raw_xml(path: str) -> list:
    """Overlapping merges and stray hyperlinks - Excel calls these corrupt."""
    found = []
    with zipfile.ZipFile(path) as zf:
        book = zf.read("xl/workbook.xml").decode()
        sheets = re.findall(r'<sheet name="([^"]+)"[^>]*r:id="rId(\d+)"',
                            book)
        sheet_names = {name for name, _ in sheets}
        for name, rid in sheets:
            try:
                xml = zf.read(f"xl/worksheets/sheet{rid}.xml").decode()
            except KeyError:
                continue
            merges = re.findall(r'<mergeCell ref="([^"]+)"/>', xml)
            boxes = []
            for ref in merges:
                start, _, end = ref.partition(":")
                end = end or start
                c1, r1 = _colnum(start)
                c2, r2 = _colnum(end)
                boxes.append((c1, r1, c2, r2, ref))
            for i, a in enumerate(boxes):
                for b in boxes[i + 1:]:
                    if (a[0] <= b[2] and b[0] <= a[2]
                            and a[1] <= b[3] and b[1] <= a[3]):
                        found.append(f"{name}: merges overlap "
                                     f"{a[4]} / {b[4]}")
            for location in re.findall(
                    r'<hyperlink ref="[^"]+" location="([^"]+)"', xml):
                target = location.split("!")[0].strip("'").replace("''", "'")
                if target not in sheet_names:
                    found.append(f"{name}: hyperlink points at missing "
                                 f"sheet '{target}'")
            for link in re.findall(r'<hyperlink ref="([^"]+)"', xml):
                lc, lr = _colnum(link)
                for c1, r1, c2, r2, ref in boxes:
                    if (c1 <= lc <= c2 and r1 <= lr <= r2
                            and (lc, lr) != (c1, r1)):
                        found.append(f"{name}: hyperlink {link} is inside "
                                     f"merged range {ref}")
    return found


def paren_depth(formula: str) -> int:
    """Open brackets minus closing ones, ignoring anything inside text.

    Excel escapes a quote inside a string by doubling it, so `""` is not
    the end of the string. Deeply nested IF() chains built by string
    concatenation are the one bug this whole file exists to catch: Excel
    and openpyxl both accept a short formula silently, and the buyer
    sees the damage.
    """
    depth = 0
    in_text = False
    i = 0
    while i < len(formula):
        char = formula[i]
        if in_text:
            if char == '"':
                if i + 1 < len(formula) and formula[i + 1] == '"':
                    i += 1
                else:
                    in_text = False
        elif char == '"':
            in_text = True
        elif char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
        i += 1
    return depth


#: divisions that really are meant to read `a + b/c`.  Each one is the
#: formula with its numbers blanked out, so one entry covers the whole
#: column.
PRECEDENCE_OK = {
    # nothing here yet: every `a + b/c` on this workbook is bracketed
    # in the source, which is the honest fix rather than a waiver
}


def _top_level(text, wanted):
    """Positions of `wanted` characters outside brackets and strings."""
    out = []
    depth = 0
    in_text = False
    i = 0
    while i < len(text):
        char = text[i]
        if in_text:
            if char == '"':
                if i + 1 < len(text) and text[i + 1] == '"':
                    i += 1
                else:
                    in_text = False
        elif char == '"':
            in_text = True
        elif char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
        elif depth == 0 and char in wanted:
            before = text[i - 1] if i else ""
            if char != "/" and (before == "" or before in "(,*/+-^<>=&"):
                # a leading sign, not an operator between two values
                i += 1
                continue
            out.append(i)
        i += 1
    return out


def _arguments(formula):
    """Every comma separated argument, at every bracket depth."""
    starts = {0: 0}
    depth = 0
    in_text = False
    i = 0
    while i < len(formula):
        char = formula[i]
        if in_text:
            if char == '"':
                if i + 1 < len(formula) and formula[i + 1] == '"':
                    i += 1
                else:
                    in_text = False
        elif char == '"':
            in_text = True
        elif char == "(":
            depth += 1
            starts[depth] = i + 1
        elif char == ")":
            yield formula[starts.get(depth, 0):i]
            depth -= 1
        elif char == ",":
            yield formula[starts.get(depth, 0):i]
            starts[depth] = i + 1
        i += 1
    yield formula[starts.get(0, 0):]


def precedence_slip(formula):
    """`SUMIFS(..)+SUMIFS(..)/COUNTIFS(..)` - the sum is not divided.

    Excel divides before it adds, so a total built by adding two
    SUMIFS together has to be bracketed before it is divided by
    anything.  The file still opens, every bracket balances and the
    calculation engine returns a number - it is simply the wrong
    number, which is the worst kind of bug to ship.
    """
    for argument in _arguments(formula):
        if "/" not in argument:
            continue
        shape = re.sub(r"\$?[A-Z]{1,3}\$?\d+", "#", argument)
        shape = re.sub(r"\d+", "#", shape)
        if shape in PRECEDENCE_OK:
            continue
        sums = _top_level(argument, "+-")
        if not sums:
            continue
        if any(divide > sums[0] for divide in _top_level(argument, "/")):
            return argument
    return ""


RANGE_RE = re.compile(
    r"^(?:(?:'[^']+'|[A-Za-z0-9_.]+)!)?"
    r"\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)$")


def _span(text: str):
    """(rows, columns) of an A1 range, or None if it is not one."""
    match = RANGE_RE.match(text.strip())
    if not match:
        return None
    c1, r1, c2, r2 = match.groups()

    def num(letters):
        value = 0
        for char in letters:
            value = value * 26 + (ord(char) - 64)
        return value
    return (abs(int(r2) - int(r1)) + 1, abs(num(c2) - num(c1)) + 1)


def _calls(formula: str, names):
    """Yield (function, [argument, ...]) for every call in *formula*."""
    depth = 0
    in_text = False
    stack = []
    i = 0
    while i < len(formula):
        char = formula[i]
        if in_text:
            if char == '"':
                if i + 1 < len(formula) and formula[i + 1] == '"':
                    i += 1
                else:
                    in_text = False
        elif char == '"':
            in_text = True
        elif char == "(":
            head = re.search(r"([A-Z][A-Z0-9\.]*)$", formula[:i])
            stack.append([head.group(1) if head else "", i + 1, [], depth])
            depth += 1
        elif char == ")":
            depth -= 1
            if stack:
                name, start, args, _ = stack.pop()
                args.append(formula[start:i])
                if name in names:
                    yield name, args
                if stack:
                    stack[-1][1] = stack[-1][1]
        elif char == "," and stack and depth == stack[-1][3] + 1:
            frame = stack[-1]
            frame[2].append(formula[frame[1]:i])
            frame[1] = i + 1
        i += 1


def ragged_criteria(formula: str) -> str:
    """COUNTIFS/SUMIFS whose ranges are not all the same shape.

    Excel answers #VALUE! the moment two criteria ranges differ in
    size, and it is an easy mistake to make when one of them is a
    running range that grows row by row.
    """
    for name, args in _calls(formula, ("COUNTIFS", "SUMIFS",
                                       "AVERAGEIFS")):
        first = 1 if name in ("SUMIFS", "AVERAGEIFS") else 0
        ranges = [args[0]] if first else []
        ranges += args[first::2]
        spans = {}
        for text in ranges:
            span = _span(text)
            if span:
                spans.setdefault(span, text)
        if len(spans) > 1:
            return f"{name}: " + " vs ".join(sorted(spans.values()))
    return ""


def collapsed_criteria(formula: str) -> str:
    """A COUNTIFS/SUMIFS criteria *range* that carries its own test.

    A criterion is two arguments - the range, then the thing it is
    compared against.  Writing it as one argument::

        COUNTIFS(prop, "PRP-01", scores="Yes")

    leaves the function an odd number of arguments and makes Excel
    answer #VALUE! on every row.  It reads perfectly naturally, which
    is exactly why it is worth a check: `SUMPRODUCT((range="Yes")*..)`
    really is written that way, so the habit carries over.
    """
    for name, args in _calls(formula, ("COUNTIFS", "SUMIFS",
                                       "AVERAGEIFS")):
        first = 1 if name in ("SUMIFS", "AVERAGEIFS") else 0
        if (len(args) - first) % 2:
            return f"{name}: odd number of range/criteria arguments"
        for text in args[first::2]:
            stripped = re.sub(r'"[^"]*"', '""', text)
            if re.search(r"[<>=]", stripped):
                return f"{name}: criteria range carries a test -> {text}"
    return ""


#: input cells that ship with a formula in them as a helpful default -
#: the buyer is meant to be able to type a fixed value straight over it
SEEDED_INPUTS = {
    "Monthly Reports!B5",        # the month the report opens on
    "Annual Reports!B5",         # the year the report opens on
    "Tax Summary!C5",            # the tax year
    "Booking Calendar!B10",      # the property the calendar shows
    "Booking Calendar!B11",      # the month the calendar shows
    "Turnover Checklist!C5",     # the cleaning job being printed
}

#: the automation chain, link by link: (sheet, column key, what the
#: formula in that column has to still be pointing at)
CHAIN_CHECKS = [
    # ---- properties: the hub every other tab joins to -------------------
    ("Property Database", "id", ("text", "PropertyPrefix")),
    ("Property Database", "bookings", ("sheet", "Booking Tracker")),
    ("Property Database", "nights_booked", ("sheet", "Booking Tracker")),
    ("Property Database", "occupancy", ("column", "Property Database",
                                        "nights_available")),
    ("Property Database", "adr", ("column", "Property Database",
                                  "accommodation")),
    ("Property Database", "revpar", ("column", "Property Database",
                                     "accommodation")),
    ("Property Database", "cleaning_spend", ("sheet",
                                             "Cleaning Schedule")),
    ("Property Database", "maintenance_spend", ("sheet",
                                                "Maintenance Log")),
    ("Property Database", "other_expenses", ("sheet",
                                             "Expense Tracker")),
    ("Property Database", "profit", ("column", "Property Database",
                                     "total_costs")),
    ("Property Database", "margin", ("column", "Property Database",
                                     "net_revenue")),
    ("Property Database", "rating", ("sheet", "Reviews Tracker")),
    ("Property Database", "low_stock", ("sheet", "Supplies Inventory")),
    ("Property Database", "next_arrival", ("sheet", "Booking Tracker")),
    # ---- bookings: the other half of the join ---------------------------
    ("Booking Tracker", "id", ("text", "BookingPrefix")),
    ("Booking Tracker", "property_name", ("sheet", "Property Database")),
    ("Booking Tracker", "guest_name", ("sheet", "Guest Database")),
    ("Booking Tracker", "nights", ("column", "Booking Tracker",
                                   "check_in")),
    ("Booking Tracker", "rate_used", ("sheet", "Property Database")),
    ("Booking Tracker", "accommodation", ("column", "Booking Tracker",
                                          "rate_used")),
    ("Booking Tracker", "extra_fee", ("text", "ExtraGuestFee")),
    ("Booking Tracker", "gross", ("column", "Booking Tracker",
                                  "accommodation")),
    ("Booking Tracker", "fee_pct", ("text", "PlatformCommission")),
    ("Booking Tracker", "pay_pct", ("text", "PlatformPayment")),
    ("Booking Tracker", "platform_fee", ("column", "Booking Tracker",
                                         "fee_pct")),
    ("Booking Tracker", "payment_fee", ("column", "Booking Tracker",
                                        "pay_pct")),
    ("Booking Tracker", "net_revenue", ("column", "Booking Tracker",
                                        "platform_fee_total")),
    ("Booking Tracker", "tax", ("text", "TaxRate")),
    ("Booking Tracker", "clash", ("column", "Booking Tracker",
                                  "ci_num")),
    ("Booking Tracker", "same_day", ("column", "Booking Tracker",
                                     "co_num")),
    ("Booking Tracker", "has_cleaning", ("sheet", "Cleaning Schedule")),
    ("Booking Tracker", "arrival_rank", ("column", "Booking Tracker",
                                         "arr_key")),
    # ---- guests ----------------------------------------------------------
    ("Guest Database", "stays", ("sheet", "Booking Tracker")),
    ("Guest Database", "nights", ("sheet", "Booking Tracker")),
    ("Guest Database", "repeat", ("column", "Guest Database", "stays")),
    ("Guest Database", "net", ("sheet", "Booking Tracker")),
    ("Guest Database", "first_stay", ("sheet", "Booking Tracker")),
    # ---- arrivals --------------------------------------------------------
    ("Check-in & Check-out", "property_name", ("sheet",
                                               "Booking Tracker")),
    ("Check-in & Check-out", "arrival_date", ("sheet",
                                              "Booking Tracker")),
    ("Check-in & Check-out", "steps_done", ("column",
                                            "Check-in & Check-out",
                                            "sent_directions")),
    ("Check-in & Check-out", "readiness", ("column",
                                           "Check-in & Check-out",
                                           "sent_directions")),
    # ---- cleaning: generated from the check-outs -------------------------
    ("Cleaning Schedule", "checkout_date", ("sheet", "Booking Tracker")),
    ("Cleaning Schedule", "clean_date", ("column", "Cleaning Schedule",
                                         "checkout_date")),
    ("Cleaning Schedule", "cleaner_name", ("sheet", "Staff & Vendors")),
    ("Cleaning Schedule", "hours", ("column", "Cleaning Schedule",
                                    "start_time")),
    ("Cleaning Schedule", "completion", ("column", "Cleaning Schedule",
                                         "task_bedroom")),
    ("Cleaning Schedule", "next_arrival", ("sheet", "Booking Tracker")),
    ("Cleaning Schedule", "window_hours", ("sheet",
                                           "Property Database")),
    ("Cleaning Schedule", "total_cost", ("text", "LaundryPerTurn")),
    ("Cleaning Schedule", "clean_cost", ("sheet", "Property Database")),
    ("Cleaning Schedule", "cost_gap", ("column", "Cleaning Schedule",
                                       "fee_charged")),
    # ---- payouts: assumed fees against real ones -------------------------
    ("Platform Revenue & Fees", "exp_net", ("sheet", "Booking Tracker")),
    ("Platform Revenue & Fees", "received", ("column",
                                             "Platform Revenue & Fees",
                                             "amount_in")),
    ("Platform Revenue & Fees", "variance", ("column",
                                             "Platform Revenue & Fees",
                                             "exp_net")),
    ("Platform Revenue & Fees", "alert", ("text", "PayoutChaseDays")),
    # ---- costs -----------------------------------------------------------
    ("Expense Tracker", "is_cleaning", ("text", "Cleaning")),
    ("Expense Tracker", "in_pl", ("text", "CleaningInExpenses")),
    ("Expense Tracker", "pl_amount", ("column", "Expense Tracker",
                                      "in_pl")),
    ("Expense Tracker", "monthly_equiv", ("column", "Expense Tracker",
                                          "frequency")),
    ("Supplies Inventory", "turns_month", ("sheet", "Cleaning Schedule")),
    ("Supplies Inventory", "weeks_cover", ("column", "Supplies Inventory",
                                           "usage_month")),
    ("Supplies Inventory", "status", ("column", "Supplies Inventory",
                                      "min_level")),
    ("Supplies Inventory", "needs_order", ("text", "RestockWeeks")),
    ("Supplies Inventory", "alert", ("column", "Supplies Inventory",
                                     "needs_order")),
    ("Maintenance Log", "vendor_name", ("sheet", "Staff & Vendors")),
    ("Maintenance Log", "cost_used", ("column", "Maintenance Log",
                                      "act_cost")),
    ("Maintenance Log", "is_open", ("column", "Maintenance Log",
                                    "status")),
    ("Maintenance Log", "alert", ("text", "MaintenanceWarnDays")),
    # ---- reviews and vendors ---------------------------------------------
    ("Reviews Tracker", "property", ("sheet", "Booking Tracker")),
    ("Reviews Tracker", "detail_avg", ("column", "Reviews Tracker",
                                       "rate_clean")),
    ("Reviews Tracker", "below_target", ("text", "TargetRating")),
    ("Reviews Tracker", "alert", ("text", "ReviewChaseDays")),
    ("Staff & Vendors", "cleans", ("sheet", "Cleaning Schedule")),
    ("Staff & Vendors", "jobs", ("sheet", "Maintenance Log")),
    ("Staff & Vendors", "paid_total", ("column", "Staff & Vendors",
                                       "paid_cleaning")),
    # ---- the investor tab -------------------------------------------------
    ("Occupancy & Performance", "nights_booked", ("sheet",
                                                  "Booking Tracker")),
    ("Occupancy & Performance", "occupancy", ("column",
                                              "Occupancy & Performance",
                                              "nights_available")),
    ("Occupancy & Performance", "adr", ("column",
                                        "Occupancy & Performance",
                                        "accommodation")),
    ("Occupancy & Performance", "revpar", ("column",
                                           "Occupancy & Performance",
                                           "accommodation")),
    ("Occupancy & Performance", "profit", ("column",
                                           "Occupancy & Performance",
                                           "total_cost")),
    ("Occupancy & Performance", "clean_per_booking",
     ("column", "Occupancy & Performance", "cleaning_cost")),
]

#: cells reached by a defined name that have to hold a live formula -
#: the summary pages are where the value of the file lives, and a name
#: pointing at an empty cell means a page quietly stopped working.
PAGE_FORMULAS = [
    ("TodayDate", "settings: the live date every alert reads"),
    ("ThisMonthKey", "settings: the month key the totals group on"),
    ("PropertyCount", "settings: properties on the books"),
    ("ActiveProperties", "settings: properties currently letting"),
    ("CalStart", "calendar: the first of the month being shown"),
    ("CalGridStart", "calendar: the Monday the grid starts on"),
    ("TaxYearStart", "tax summary: the first day of the tax year"),
    ("TaxYearEnd", "tax summary: the last day of the tax year"),
    ("PrevMonthKey", "monthly reports: last month, for comparison"),
    ("LastYearMonthKey", "monthly reports: the same month last year"),
]

#: cells a buyer has to be able to type in
PAGE_INPUTS = [
    ("HostName", "settings: the business name"),
    ("OwnerName", "settings: the host or owner"),
    ("CurrencySign", "settings: the currency symbol"),
    ("TrackingStart", "settings: the day tracking began"),
    ("DefaultNightly", "settings: the standard nightly rate"),
    ("DefaultCleaningFee", "settings: the cleaning fee charged"),
    ("DefaultCleaningCost", "settings: what a turnover costs you"),
    ("ExtraGuestAfter", "settings: guests included in the rate"),
    ("ExtraGuestFee", "settings: the extra guest charge"),
    ("PetFee", "settings: the pet fee"),
    ("DefaultDeposit", "settings: the security deposit"),
    ("DefaultMinNights", "settings: the minimum stay"),
    ("ManagementFee", "settings: the co-host share"),
    ("TaxLabel", "settings: what the lodging tax is called"),
    ("TaxRate", "settings: the lodging tax rate"),
    ("TaxCollectedByPlatform", "settings: who remits the lodging tax"),
    ("LaundryPerTurn", "settings: laundry cost a turnover"),
    ("SuppliesPerTurn", "settings: consumables a turnover"),
    ("WelcomePerTurn", "settings: the welcome basket"),
    ("CleanerRate", "settings: the default cleaner rate"),
    ("TightTurnHours", "settings: what counts as a tight turnaround"),
    ("CleaningInExpenses", "settings: does cleaning count twice"),
    ("MaintenanceInExpenses", "settings: does maintenance count twice"),
    ("TargetOccupancy", "settings: the occupancy you are aiming at"),
    ("TargetADR", "settings: the rate you are aiming at"),
    ("TargetRevPAR", "settings: the RevPAR you are aiming at"),
    ("TargetMargin", "settings: the margin you are aiming at"),
    ("TargetRating", "settings: the review score you are aiming at"),
    ("TargetRevenueMonth", "settings: the monthly revenue target"),
    ("TargetRevenueYear", "settings: the annual revenue target"),
    ("TargetProfitMonth", "settings: the monthly profit target"),
    ("UpcomingDays", "settings: the arriving-soon window"),
    ("ReviewChaseDays", "settings: when to chase a review"),
    ("PayoutChaseDays", "settings: when to chase a payout"),
    ("MaintenanceWarnDays", "settings: when a repair starts warning"),
    ("RestockWeeks", "settings: when to reorder a supply"),
    ("LongStayNights", "settings: what counts as a long stay"),
    ("PropertyPrefix", "settings: the property ID prefix"),
    ("BookingPrefix", "settings: the booking ID prefix"),
    ("CalProperty", "calendar: the property being shown"),
    ("CalMonth", "calendar: the month being shown"),
    ("JobId", "turnover checklist: the job being printed"),
    ("TaxYear", "tax summary: the year being reported on"),
    ("ReportMonth", "monthly reports: the month being reported on"),
    ("ReportYear", "annual reports: the year being reported on"),
]

#: every drop down on the workbook reads one of these - a missing one
#: means a blank list and a column nobody can fill in
REQUIRED_LIST_NAMES = [
    "PropertyList", "PropertyNameList", "BookingList", "GuestList",
    "VendorList", "VendorNameList", "SupplyList", "PlatformList",
    "PlatformCommission", "PlatformPayment", "BookingStatusList",
    "PaymentStatusList", "PropertyTypeList", "PropertyStatusList",
    "OwnershipList", "CleaningStatusList", "MaintenanceCategoryList",
    "MaintenanceStatusList", "PriorityList", "ExpenseCategoryList",
    "TaxCategoryList", "PaymentMethodList", "FrequencyList",
    "SupplyCategoryList", "StockStatusList", "StaffRoleList",
    "PayBasisList", "ContractStatusList", "AvailabilityList",
    "ReviewStatusList", "GuestTypeList", "TripPurposeList",
    "PayoutStatusList", "YesNoList", "MonthKeyList", "CountryList",
]


def _name_target(wb, name):
    """The one cell a defined name points at, or None."""
    defn = wb.defined_names.get(name)
    if defn is None:
        for ws in wb.worksheets:
            if name in ws.defined_names:
                defn = ws.defined_names[name]
                break
    if defn is None:
        return None
    try:
        sheet, ref = next(defn.destinations)
    except (StopIteration, Exception):          # noqa: B902
        return None
    if sheet not in wb.sheetnames or ":" in ref:
        return None
    return wb[sheet][ref.replace("$", "")]


def check_pages(wb, problems: list, notes: list) -> None:
    """The summary pages, checked through their defined names."""
    live = 0
    for name, what in PAGE_FORMULAS:
        cell = _name_target(wb, name)
        if cell is None:
            problems.append(f"{name} is missing - {what}")
            continue
        value = cell.value
        if not (isinstance(value, str) and value.startswith("=")):
            problems.append(f"{name} no longer holds a formula - {what}")
            continue
        if cell.protection and not cell.protection.locked:
            problems.append(f"{name} is a formula but is unlocked - {what}")
            continue
        live += 1
    typed = 0
    for name, what in PAGE_INPUTS:
        cell = _name_target(wb, name)
        if cell is None:
            problems.append(f"{name} is missing - {what}")
            continue
        if cell.protection and cell.protection.locked:
            problems.append(f"{name} is locked, so nobody can change it "
                            f"- {what}")
            continue
        typed += 1
    notes.append(f"summary pages: {live} live formulas, "
                 f"{typed} typeable inputs verified")


def _specs():
    """Every register specification, keyed by sheet name."""
    from airbnb_rental_spreadsheet.build import REGISTER_MODULES
    from airbnb_rental_spreadsheet.config import Config
    cfg = Config()
    return {module.NAME: module.spec(cfg) for module in REGISTER_MODULES}


def check_chain(wb, problems: list, notes: list) -> None:
    """The enquiry-to-enrolment-to-invoice chain only holds while these do.

    Every one of them is a formula in row 5 of a register, so the check
    is cheap: find the column by its key in the specification, read the
    formula, and make sure it still mentions the thing it depends on.
    """
    specs = _specs()
    checked = 0
    for sheet_name, key, expect in CHAIN_CHECKS:
        spec = specs.get(sheet_name)
        if spec is None or sheet_name not in wb.sheetnames:
            problems.append(f"missing register {sheet_name}")
            continue
        try:
            column = spec.index(key) + 1
        except Exception:
            problems.append(f"{sheet_name}: column '{key}' has gone")
            continue
        formula = wb[sheet_name].cell(row=5, column=column).value
        if not isinstance(formula, str) or not formula.startswith("="):
            problems.append(f"{sheet_name}.{key} is no longer a formula")
            continue
        if expect[0] == "sheet":
            found = expect[1] in formula
            label = f"the {expect[1]} tab"
        elif expect[0] == "text":
            found = expect[1] in formula
            label = expect[1]
        else:
            letter = specs[expect[1]].letter(expect[2])
            label = f"{expect[1]}.{expect[2]} (column {letter})"
            if expect[1] == sheet_name:
                # a link inside the same register is a plain A1 reference
                found = bool(re.search(r"(?<![A-Z$])\$?" + letter +
                                       r"\$?5(?![0-9])", formula))
            else:
                found = f"${letter}$" in formula
        if not found:
            problems.append(
                f"{sheet_name}.{key} no longer points at {label} - "
                f"the automation chain is broken")
        checked += 1
    notes.append(f"automation chain: {checked} links verified")


def check_platform_table(wb, problems: list, notes: list) -> None:
    """The platform fee table a buyer must be able to type over.

    This is the single most important editable thing in the product:
    commission varies by platform, by market, by account type and by
    negotiation, so a workbook that locks these cells is wrong for
    almost everybody who buys it.
    """
    ranges = ("PlatformList", "PlatformCommission", "PlatformPayment",
              "PlatformIsOTA")
    found = 0
    for name in ranges:
        defn = wb.defined_names.get(name)
        if defn is None:
            problems.append(f"the editable platform fee table has lost "
                            f"{name}")
            continue
        try:
            _sheet, ref = next(defn.destinations)
        except Exception:                        # noqa: B902
            problems.append(f"{name} does not point anywhere")
            continue
        if ":" not in ref:
            problems.append(f"{name} should be a range, got {ref}")
            continue
        found += 1

    locked = 0
    cells = 0
    for name in ("PlatformCommission", "PlatformPayment"):
        defn = wb.defined_names.get(name)
        if defn is None:
            continue
        sheet, ref = next(defn.destinations)
        for row in wb[sheet][ref.replace("$", "")]:
            for cell in row:
                cells += 1
                if cell.protection and cell.protection.locked:
                    locked += 1
    if locked:
        problems.append(
            f"{locked} platform fee cells are locked - commission "
            f"rates vary by platform, market and account, so buyers "
            f"have to be able to type their own")
    notes.append(f"platform fee table: {found}/4 named ranges, "
                 f"{cells} editable rate cells"
                 if found else "platform fee table: NOT FOUND")


def check_promises(wb, problems: list, notes: list) -> None:
    """Three promises the product makes in writing.

    1. ADR and RevPAR use accommodation revenue only - counting the
       cleaning fee inflates both and is the single most common error
       in rental spreadsheets sold online.
    2. the workbook never claims to synchronise with a platform.
    3. the tax page carries a jurisdiction disclaimer.
    """
    from airbnb_rental_spreadsheet.build import REGISTER_MODULES
    from airbnb_rental_spreadsheet.config import Config
    cfg = Config()
    specs = {m.NAME: m.spec(cfg) for m in REGISTER_MODULES}

    for sheet_name, keys in (("Property Database", ("adr", "revpar")),
                             ("Occupancy & Performance",
                              ("adr", "revpar"))):
        spec = specs.get(sheet_name)
        if spec is None:
            continue
        for key in keys:
            try:
                column = spec.index(key) + 1
            except Exception:                    # noqa: B902
                problems.append(f"{sheet_name}.{key} has gone")
                continue
            formula = wb[sheet_name].cell(row=5, column=column).value
            if not isinstance(formula, str):
                continue
            if "cleaning" in formula.lower():
                problems.append(
                    f"{sheet_name}.{key} mentions cleaning - ADR and "
                    f"RevPAR must use accommodation revenue only")

    banned = ("sync", "synchronise", "synchronize", "live feed",
              "api connection", "automatically imports")
    hits = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if not isinstance(value, str) or value.startswith("="):
                    continue
                low = value.lower()
                for phrase in banned:
                    if phrase in low and "not" not in low \
                            and "cannot" not in low:
                        hits += 1
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: claims a "
                            f"platform connection -> {value[:70]}")
    disclaimer = False
    for ws in (wb["Tax Summary"], wb["Settings & Lists"],
               wb["Getting Started"]):
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if isinstance(value, str) and "jurisdiction" in \
                        value.lower():
                    disclaimer = True
    if not disclaimer:
        problems.append("no tax disclaimer mentioning jurisdiction")
    notes.append(f"product promises: ADR/RevPAR exclude cleaning, "
                 f"{hits} false sync claims, tax disclaimer "
                 f"{'present' if disclaimer else 'MISSING'}")


def check(path: str) -> int:
    problems: list = list(check_raw_xml(path))
    notes: list = []
    if problems:
        print(f"--- {path}")
        for problem in problems[:40]:
            print(f"    - {problem}")
        return 1
    wb = openpyxl.load_workbook(path, data_only=False)

    # ---------------------------------------------------------- structure
    names = wb.sheetnames
    if names != EXPECTED_SHEETS:
        problems.append(f"sheet order/={names}")
    hidden = [ws.title for ws in wb.worksheets if ws.sheet_state != "visible"]
    if hidden:
        problems.append(f"no tab should be hidden, got {hidden}")

    check_chain(wb, problems, notes)
    check_pages(wb, problems, notes)
    check_platform_table(wb, problems, notes)
    check_promises(wb, problems, notes)

    # --------------------------------------------------------- properties
    props = wb.properties
    if props.creator != "Novality Store":
        problems.append(f"creator is {props.creator!r}")
    if props.lastModifiedBy != "Novality Store":
        problems.append(f"lastModifiedBy is {props.lastModifiedBy!r}")

    # --------------------------------------------------------- protection
    for ws in wb.worksheets:
        prot = ws.protection
        if not prot.sheet:
            problems.append(f"{ws.title}: sheet protection off")
        elif not prot.password:
            problems.append(f"{ws.title}: protected without a password")

    # -------------------------------------------------------- definednames
    defined = set(wb.defined_names)
    for ws in wb.worksheets:
        defined |= set(ws.defined_names)
    for name in REQUIRED_LIST_NAMES:
        if name not in defined:
            problems.append(f"drop down list {name} is missing")
    notes.append(f"{len(defined)} defined names")

    # ------------------------------------------------------------ formulas
    used_names: Counter = Counter()
    unknown_functions: Counter = Counter()
    bare_refs: dict = {}
    bad_functions: Counter = Counter()
    formula_cells = 0
    locked_formula_problem = 0
    unlocked_input_cells = 0
    per_sheet = defaultdict(int)

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if isinstance(value, str) and value.startswith("="):
                    formula_cells += 1
                    per_sheet[ws.title] += 1
                    if "[[" in value or "{" in value or "}" in value:
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: unresolved "
                            f"template token -> {value[:90]}")
                    if value != value.strip():
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: formula has "
                            f"stray whitespace around it, which some "
                            f"calculation engines refuse to parse -> "
                            f"{value[:70]!r}")
                    if "#REF" in value:
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: #REF -> "
                            f"{value[:90]}")
                    where = f"{ws.title}!{cell.coordinate}"
                    if (cell.protection and not cell.protection.locked
                            and where not in SEEDED_INPUTS):
                        locked_formula_problem += 1
                        if locked_formula_problem < 25:
                            problems.append(
                                f"{where}: formula cell is NOT locked")
                    stripped = re.sub(r'"[^"]*"', '""', value)
                    stripped = re.sub(r"'[^']*'!", "SHEET!", stripped)
                    if value.count('"') % 2:
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: odd number of "
                            f"quotes -> {value[:80]}")
                    depth = paren_depth(value)
                    if depth:
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: unbalanced "
                            f"brackets ({depth:+d}) -> {value[:80]}")
                    slip = precedence_slip(value)
                    if slip:
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: a+b/c where "
                            f"(a+b)/c was meant -> {slip[:90]}")
                    collapsed = collapsed_criteria(value)
                    if collapsed:
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: a criteria "
                            f"range and its test were written as one "
                            f"argument, Excel answers #VALUE! -> "
                            f"{collapsed[:90]}")
                    ragged = ragged_criteria(value)
                    if ragged:
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: criteria "
                            f"ranges are different sizes, Excel answers "
                            f"#VALUE! -> {ragged[:90]}")
                    if ",)" in stripped or "(," in stripped or \
                            ",," in stripped:
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: empty argument "
                            f"-> {value[:80]}")
                    for func in FUNCTION_RE.findall(stripped):
                        if func in MODERN_ONLY:
                            bad_functions[func] += 1
                        elif func not in KNOWN_FUNCTIONS:
                            unknown_functions[func] += 1
                    for token in NAME_RE.findall(stripped):
                        if token in KNOWN_FUNCTIONS:
                            continue
                        used_names[token] += 1
                    for token in BARE_COL_RE.findall(stripped):
                        if token in KNOWN_FUNCTIONS:
                            continue
                        bare_refs[f"{ws.title}!{cell.coordinate} {token}"] = \
                            value[:90]
                elif value is not None and cell.protection and \
                        not cell.protection.locked:
                    unlocked_input_cells += 1

    # sheet names used inside formulas must exist
    sheet_refs = Counter()
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if isinstance(value, str) and value.startswith("="):
                    body = re.sub(r'"[^"]*"', '""', value)
                    for ref in re.findall(r"'([^']+)'!", body):
                        sheet_refs[ref] += 1
                    for ref in re.findall(r"(?<![A-Za-z0-9_'])"
                                          r"([A-Za-z][A-Za-z0-9_]*)!", body):
                        sheet_refs[ref] += 1
    for ref, count in sheet_refs.items():
        if ref not in names:
            problems.append(f"formula points at missing sheet "
                            f"'{ref}' ({count}x)")
    notes.append("cross sheet references: " + ", ".join(
        f"{k} {v}" for k, v in sorted(sheet_refs.items(),
                                      key=lambda kv: -kv[1])[:6]))

    for where, formula in list(bare_refs.items())[:15]:
        if where.rsplit(" ", 1)[-1] not in defined:
            problems.append(f"{where}: column letter with no row "
                            f"number -> {formula}")
    for func, count in unknown_functions.items():
        problems.append(f"unknown function {func}() used {count}x")
    for func, count in bad_functions.items():
        problems.append(f"{func}() needs a modern Excel - used {count}x")

    missing = {}
    for token, count in used_names.items():
        if token in defined:
            continue
        if token in KNOWN_FUNCTIONS or token.upper() in KNOWN_FUNCTIONS:
            continue
        if re.match(r"^[A-Z]{1,3}\d+$", token):
            continue
        missing[token] = count
    for token, count in sorted(missing.items(), key=lambda kv: -kv[1])[:25]:
        problems.append(f"undefined name used {count}x: {token}")

    # ----------------------------------------------------- data validation
    for ws in wb.worksheets:
        for dv in ws.data_validations.dataValidation:
            formula = (dv.formula1 or "").lstrip("=")
            if dv.type == "list" and formula and \
                    re.match(r"^[A-Za-z_][A-Za-z0-9_.]*$", formula):
                if formula not in defined:
                    problems.append(
                        f"{ws.title}: drop down uses missing name "
                        f"{formula}")

    # ------------------------------------------------------------- report
    with zipfile.ZipFile(path) as zf:
        size = sum(i.file_size for i in zf.infolist())
    notes.append(f"{formula_cells:,} formula cells, "
                 f"{unlocked_input_cells:,} unlocked value cells")
    notes.append(f"uncompressed xml {size / 1_000_000:.1f} MB")

    print(f"--- {path}")
    for note in notes:
        print(f"  . {note}")
    print("  . formulas per sheet: " + ", ".join(
        f"{k} {v}" for k, v in sorted(per_sheet.items(),
                                      key=lambda kv: -kv[1])[:10]))
    if problems:
        print(f"  ! {len(problems)} problem(s):")
        for problem in problems[:60]:
            print(f"    - {problem}")
        return 1
    print("  OK - no problems found")
    return 0


if __name__ == "__main__":
    target = (sys.argv[1] if len(sys.argv) > 1
              else "output/Airbnb & Short-Term Rental Management Spreadsheet.xlsx")
    raise SystemExit(check(target))
