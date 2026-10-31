"""Static checks on a generated Daycare Business workbook.

Run it with any Python that has openpyxl installed::

    python tools/validate_daycare_workbook.py output/daycare.xlsx

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
from daycare_business_spreadsheet.build import SHEET_ORDER as EXPECTED_SHEETS  # noqa: E402,E501

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


#: input cells that ship with a formula in them as a helpful default -
#: the buyer is meant to be able to type a fixed value straight over it
SEEDED_INPUTS = {
    "Monthly Reports!B5",       # the month the report opens on
    "Annual Reports!B5",        # the year the report opens on
}

#: the automation chain, link by link: (sheet, column key, what the
#: formula in that column has to still be pointing at)
CHAIN_CHECKS = [
    # ---- enrolment: who is on the books and what they are worth ---------
    ("Child Enrollment", "age_months", ("text", "DATEDIF")),
    ("Child Enrollment", "age_band", ("text", "AgeBandList")),
    ("Child Enrollment", "family_seq", ("column", "Child Enrollment",
                                        "family")),
    ("Child Enrollment", "auto_discount", ("text", "SiblingDiscount")),
    ("Child Enrollment", "gross_week", ("column", "Child Enrollment",
                                        "rate")),
    ("Child Enrollment", "weekly_net", ("column", "Child Enrollment",
                                        "gross_week")),
    ("Child Enrollment", "monthly_net", ("column", "Child Enrollment",
                                         "weekly_net")),
    ("Child Enrollment", "billed", ("sheet", "Tuition & Fees")),
    ("Child Enrollment", "balance", ("sheet", "Parent Database")),
    ("Child Enrollment", "days_in", ("sheet", "Attendance")),
    ("Child Enrollment", "attendance_pct", ("column", "Child Enrollment",
                                            "sessions")),
    ("Child Enrollment", "docs_ok", ("sheet", "Document Tracker")),
    ("Child Enrollment", "bday_in", ("column", "Child Enrollment",
                                     "next_bday")),
    # ---- families: one bill, several children ---------------------------
    ("Parent Database", "children", ("sheet", "Child Enrollment")),
    ("Parent Database", "weekly_fees", ("sheet", "Child Enrollment")),
    ("Parent Database", "invoiced", ("sheet", "Invoices & Receipts")),
    ("Parent Database", "paid", ("sheet", "Invoices & Receipts")),
    ("Parent Database", "balance", ("column", "Parent Database", "paid")),
    ("Parent Database", "overdue", ("text", "Overdue")),
    ("Parent Database", "messages", ("sheet", "Parent Communication")),
    # ---- the waiting list ------------------------------------------------
    ("Waitlist", "age_at_start", ("column", "Waitlist", "start_wanted")),
    ("Waitlist", "seats_now", ("sheet", "Classroom Management")),
    ("Waitlist", "converted", ("text", "Enrolled")),
    # ---- the daily register ---------------------------------------------
    ("Attendance", "present", ("text", "Present")),
    ("Attendance", "hours", ("column", "Attendance", "in_time")),
    ("Attendance", "due_in", ("sheet", "Child Enrollment")),
    ("Attendance", "late_in", ("column", "Attendance", "due_in")),
    ("Attendance", "late_fee", ("text", "LatePickupFee")),
    ("Attendance", "auth_list", ("sheet", "Child Enrollment")),
    ("Attendance", "pickup_ok", ("column", "Attendance", "auth_list")),
    ("Attendance", "daykey", ("column", "Attendance", "child")),
    # ---- what each child is charged -------------------------------------
    ("Tuition & Fees", "gross", ("column", "Tuition & Fees", "rate")),
    ("Tuition & Fees", "discount", ("column", "Tuition & Fees",
                                    "discount_pct")),
    ("Tuition & Fees", "net", ("column", "Tuition & Fees", "is_credit")),
    ("Tuition & Fees", "billed", ("column", "Tuition & Fees", "invoice")),
    ("Tuition & Fees", "key", ("column", "Tuition & Fees", "seq")),
    # ---- the family ledger ----------------------------------------------
    ("Invoices & Receipts", "charges", ("sheet", "Tuition & Fees")),
    ("Invoices & Receipts", "total", ("column", "Invoices & Receipts",
                                      "charges")),
    ("Invoices & Receipts", "paid", ("text", "Payment")),
    ("Invoices & Receipts", "balance", ("column", "Invoices & Receipts",
                                        "paid")),
    ("Invoices & Receipts", "status", ("column", "Invoices & Receipts",
                                       "balance")),
    ("Invoices & Receipts", "days_overdue", ("column",
                                             "Invoices & Receipts",
                                             "due_date")),
    ("Invoices & Receipts", "billed", ("text", "Credit note")),
    ("Invoices & Receipts", "collected", ("text", "Refund")),
    # ---- what the parent reads at the end of the day --------------------
    ("Daily Reports", "hours", ("sheet", "Attendance")),
    ("Daily Reports", "meals_logged", ("sheet", "Meal Tracker")),
    ("Daily Reports", "incidents", ("sheet", "Incident Reports")),
    ("Daily Reports", "nap_hours", ("column", "Daily Reports", "nap_from")),
    # ---- food, allergies, accidents -------------------------------------
    ("Meal Tracker", "diet", ("sheet", "Child Enrollment")),
    ("Meal Tracker", "allergy_level", ("sheet", "Child Enrollment")),
    ("Meal Tracker", "refused", ("text", "Refused")),
    ("Allergy & Medical", "days_left", ("text", "TodayDate")),
    ("Allergy & Medical", "doses_logged", ("column", "Allergy & Medical",
                                           "detail")),
    ("Incident Reports", "age_at", ("sheet", "Child Enrollment")),
    ("Incident Reports", "notify_mins", ("column", "Incident Reports",
                                         "notified_time")),
    ("Incident Reports", "complete", ("column", "Incident Reports",
                                      "notified")),
    # ---- the team --------------------------------------------------------
    ("Staff Database", "next_expiry", ("column", "Staff Database",
                                       "first_aid_exp")),
    ("Staff Database", "days_to_expiry", ("column", "Staff Database",
                                          "next_expiry")),
    ("Staff Database", "certs_ok", ("text", "DocUrgentDays")),
    ("Staff Database", "hours_week", ("sheet", "Staff Schedule")),
    ("Staff Database", "pay_month", ("sheet", "Payroll Tracker")),
    ("Staff Database", "children_room", ("sheet", "Child Enrollment")),
    ("Staff Schedule", "hours", ("text", "BreakMinutes")),
    ("Staff Schedule", "actual_hours", ("column", "Staff Schedule",
                                        "clock_in")),
    ("Staff Schedule", "week_to_date", ("column", "Staff Schedule",
                                        "actual_hours")),
    ("Staff Schedule", "ot_hours", ("text", "OvertimeAfter")),
    ("Staff Schedule", "cost", ("text", "OvertimeMultiplier")),
    ("Staff Schedule", "clash", ("column", "Staff Schedule", "finish")),
    ("Payroll Tracker", "basic_hours", ("sheet", "Staff Schedule")),
    ("Payroll Tracker", "ot_hours", ("sheet", "Staff Schedule")),
    ("Payroll Tracker", "gross", ("column", "Payroll Tracker", "ot_pay")),
    ("Payroll Tracker", "employer_cost", ("text", "PayrollBurden")),
    ("Payroll Tracker", "total_cost", ("column", "Payroll Tracker",
                                       "employer_cost")),
    ("Payroll Tracker", "net_pay", ("column", "Payroll Tracker",
                                    "deductions")),
    # ---- rooms, ratios and capacity -------------------------------------
    ("Classroom Management", "enrolled", ("sheet", "Child Enrollment")),
    ("Classroom Management", "seats", ("column", "Classroom Management",
                                       "capacity")),
    ("Classroom Management", "occupancy", ("column",
                                           "Classroom Management",
                                           "enrolled")),
    ("Classroom Management", "waiting", ("sheet", "Waitlist")),
    ("Classroom Management", "present_today", ("sheet", "Attendance")),
    ("Classroom Management", "staff_today", ("sheet", "Staff Schedule")),
    ("Classroom Management", "staff_needed", ("column",
                                              "Classroom Management",
                                              "ratio")),
    ("Classroom Management", "ratio_ok", ("column",
                                          "Classroom Management",
                                          "staff_needed")),
    ("Classroom Management", "revenue", ("sheet", "Child Enrollment")),
    ("Classroom Management", "empty_cost", ("column",
                                            "Classroom Management",
                                            "revenue")),
    # ---- costs, stock, paperwork ----------------------------------------
    ("Expense Tracker", "is_wages", ("text", "Staff wages")),
    ("Expense Tracker", "in_pl", ("text", "PayrollInExpenses")),
    ("Expense Tracker", "pl_amount", ("column", "Expense Tracker",
                                      "in_pl")),
    ("Expense Tracker", "monthly", ("text", "Quarterly")),
    ("Expense Tracker", "budget", ("text", "BudgetAmount")),
    ("Expense Tracker", "child_cost", ("text", "ChildrenEnrolled")),
    ("Supply Inventory", "stock", ("column", "Supply Inventory",
                                   "received")),
    ("Supply Inventory", "weeks_cover", ("column", "Supply Inventory",
                                         "weekly_use")),
    ("Supply Inventory", "order_qty", ("text", "ReorderWeeks")),
    ("Supply Inventory", "status", ("text", "StockWarningWeeks")),
    ("Parent Communication", "days_open", ("column",
                                           "Parent Communication",
                                           "closed_date")),
    ("Parent Communication", "overdue", ("column",
                                         "Parent Communication",
                                         "follow_date")),
    ("Document Tracker", "who_name", ("sheet", "Staff Database")),
    ("Document Tracker", "days_left", ("text", "TodayDate")),
    ("Document Tracker", "state", ("column", "Document Tracker",
                                   "days_left")),
]

#: cells reached by a defined name that have to hold a live formula -
#: the summary pages are where the value of the file lives, and a name
#: pointing at an empty cell means a page quietly stopped working.
PAGE_FORMULAS = [
    ("TodayDate", "settings: the live date every alert reads"),
    ("ThisMonthKey", "settings: the month key the totals group on"),
    ("TrackingWeeks", "settings: weeks of history, for the averages"),
    ("OpenHoursDay", "settings: hours the doors are open each day"),
    ("BudgetMonthlyTotal", "settings: the monthly running cost budget"),
    ("ChildrenEnrolled", "settings: children on the books"),
    ("ChildrenActive", "settings: children actually attending"),
    ("StaffCount", "settings: how many staff are on the payroll"),
    ("RoomCapacity", "settings: places across all the rooms"),
    ("RevenueAll", "settings: everything billed, all time"),
    ("CollectedAll", "settings: everything collected, all time"),
    ("DashKey", "dashboard: the month the tiles report on"),
    ("DashPrevKey", "dashboard: last month, for the comparisons"),
    ("MonthBilled", "dashboard: billed this month"),
    ("MonthCollected", "dashboard: collected this month"),
    ("MonthWages", "dashboard: the pay runs this month"),
    ("MonthExpenses", "dashboard: the bills this month"),
    ("MonthProfit", "dashboard: what is left after the bills"),
    ("ReportPrev", "monthly reports: last month, for the comparisons"),
    ("ReportYearPrev", "annual reports: last year, for the comparisons"),
]

#: cells a buyer has to be able to type in
PAGE_INPUTS = [
    ("CentreName", "settings: the nursery's name"),
    ("OwnerName", "settings: the owner or director"),
    ("CurrencySign", "settings: the currency symbol"),
    ("TrackingStart", "settings: the day tracking began"),
    ("OpenHour", "settings: opening time"),
    ("CloseHour", "settings: closing time"),
    ("OpenDays", "settings: days open a week"),
    ("WeeksOpen", "settings: trading weeks a year"),
    ("LicensedCapacity", "settings: licensed places"),
    ("RatioWarning", "settings: the occupancy a room warns at"),
    ("FullTimeWeek", "settings: the standard full-time weekly fee"),
    ("SiblingDiscount", "settings: the sibling discount"),
    ("PaymentTermsDays", "settings: how long a family has to pay"),
    ("LatePickupFee", "settings: the late collection charge"),
    ("TaxRate", "settings: the tax rate"),
    ("TaxOnTuition", "settings: does tax apply to tuition"),
    ("PayrollBurden", "settings: employer costs on wages"),
    ("OvertimeAfter", "settings: hours a week before overtime"),
    ("PayrollInExpenses", "settings: are wages typed as expenses too"),
    ("RevenueTargetMonth", "settings: the monthly revenue target"),
    ("TargetOccupancy", "settings: the occupancy you are aiming at"),
    ("DocWarningDays", "settings: how early paperwork warns"),
    ("ReportMonth", "monthly reports: the month being reported on"),
    ("ReportYear", "annual reports: the year being reported on"),
    ("PrintInvoice", "monthly reports: the invoice being printed"),
]

#: every drop down on the workbook reads one of these - a missing one
#: means a blank list and a column nobody can fill in
REQUIRED_LIST_NAMES = [
    "ChildList", "FamilyList", "ClassroomList", "StaffList",
    "InvoiceList", "SupplyList", "EnrollmentStatusList",
    "AttendanceStatusList", "ProgramList", "ScheduleTypeList",
    "BillingBasisList", "FeeTypeList", "LedgerTypeList",
    "PaymentMethodList", "ExpenseCategoryList", "DocumentTypeList",
    "IncidentCategoryList", "AllergySeverityList", "DietList",
    "YesNoList", "MonthKeyList", "AgeBandList",
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
    from daycare_business_spreadsheet.build import REGISTER_MODULES
    from daycare_business_spreadsheet.config import Config
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


def check_ratio_engine(wb, problems: list, notes: list) -> None:
    """The ratio table a buyer edits to match their own licensing."""
    ws = wb["Settings & Lists"]
    bands = 0
    for row in ws.iter_rows():
        for cell in row:
            if cell.value in (None, ""):
                continue
            if isinstance(cell.value, str) and cell.value.startswith("="):
                continue
            if cell.column_letter == "B" and isinstance(cell.value, str):
                pass
        bands += 0
    for name in ("AgeBandList", "AgeBandFrom", "AgeBandTo", "AgeBandRatio",
                 "AgeBandMax"):
        defn = wb.defined_names.get(name)
        if defn is None:
            problems.append(f"the editable ratio table has lost {name}")
            continue
        try:
            sheet, ref = next(defn.destinations)
        except Exception:                        # noqa: B902
            problems.append(f"{name} does not point anywhere")
            continue
        if ":" not in ref:
            problems.append(f"{name} should be a range, got {ref}")
            continue
        bands = max(bands, 1)
    locked = 0
    defn = wb.defined_names.get("AgeBandRatio")
    if defn is not None:
        sheet, ref = next(defn.destinations)
        for row in wb[sheet][ref.replace("$", "")]:
            for cell in row:
                if cell.protection and cell.protection.locked:
                    locked += 1
    if locked:
        problems.append(f"{locked} staff-to-child ratio cells are locked - "
                        f"buyers have to be able to match local licensing")
    notes.append("ratio table: editable, five named ranges present"
                 if bands else "ratio table: NOT FOUND")


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
    check_ratio_engine(wb, problems, notes)

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
              else "output/Daycare Business Management Spreadsheet.xlsx")
    raise SystemExit(check(target))
