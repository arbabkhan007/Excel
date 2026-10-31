#!/usr/bin/env python3
"""Does the demo nursery actually hang together?

The validator checks the formulas; this checks the *story* they are
told. It reads only the literal cells a user would have typed - never
a formula column - rebuilds the arithmetic in Python, and complains if
the example data would make the workbook look wrong on the first open:
children in rooms that do not exist, attendance for children who had
left, charges on invoices that were never raised, payments bigger than
the bill, shifts that overlap, stock that goes negative, IDs that
point at nothing.

    python tools/check_daycare_sample.py output/daycare.xlsx
"""

from __future__ import annotations

import datetime as _dt
import os
import sys
from collections import defaultdict

try:
    import openpyxl
except ImportError:                                   # pragma: no cover
    print("openpyxl is required: pip install openpyxl")
    raise SystemExit(2)

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from daycare_business_spreadsheet.build import REGISTER_MODULES  # noqa: E402
from daycare_business_spreadsheet.config import Config, rid      # noqa: E402

MONEY = 0.015           # a cent and a half of rounding slack

#: the column every register is keyed on - a row with this cell empty
#: is an empty row, whatever else is in it
ANCHORS = {
    "Classroom Management": "name",
    "Parent Database": "parent1",
    "Child Enrollment": "name",
    "Waitlist": "child_name",
    "Attendance": "date",
    "Tuition & Fees": "date",
    "Invoices & Receipts": "date",
    "Daily Reports": "date",
    "Meal Tracker": "date",
    "Allergy & Medical": "child",
    "Incident Reports": "date",
    "Staff Database": "name",
    "Staff Schedule": "date",
    "Payroll Tracker": "period_end",
    "Expense Tracker": "date",
    "Supply Inventory": "name",
    "Parent Communication": "date",
    "Document Tracker": "doc_type",
}

#: ledger rows that reduce what a family owes
CREDIT_TYPES = ("Credit note", "Write-off")


def _specs(cfg: Config):
    return {module.NAME: module.spec(cfg) for module in REGISTER_MODULES}


def _rows(ws, spec, key_column: str):
    """Every filled data row as {column key: literal value}."""
    first = spec.first_data_row                      # 1 based
    last = spec.last_data_row
    key_index = spec.index(key_column) + 1
    out = []
    for excel_row in range(first, last + 1):
        anchor = ws.cell(row=excel_row, column=key_index).value
        if anchor in (None, ""):
            continue
        record = {"_row": excel_row - first + 1}
        for index, col in enumerate(spec.columns):
            value = ws.cell(row=excel_row, column=index + 1).value
            if isinstance(value, str) and value.startswith("="):
                continue                              # never trust a formula
            record[col.key] = value
        out.append(record)
    return out


def _month(value) -> str:
    if isinstance(value, _dt.datetime):
        value = value.date()
    return f"{value.year:04d}-{value.month:02d}" if value else ""


def _day(value):
    return value.date() if isinstance(value, _dt.datetime) else value


def _minutes(value) -> int:
    if isinstance(value, _dt.datetime):
        return value.hour * 60 + value.minute
    if isinstance(value, _dt.time):
        return value.hour * 60 + value.minute
    if isinstance(value, (int, float)):
        return int(round(float(value) * 24 * 60))
    return 0


def _num(value) -> float:
    return float(value) if isinstance(value, (int, float)) else 0.0


def check(path: str) -> int:                          # noqa: C901
    cfg = Config()
    specs = _specs(cfg)
    wb = openpyxl.load_workbook(path, data_only=False)
    problems: list = []
    notes: list = []

    data = {}
    for name, anchor in ANCHORS.items():
        if name not in wb.sheetnames:
            problems.append(f"missing sheet {name}")
            continue
        data[name] = _rows(wb[name], specs[name], anchor)
    if problems:
        for line in problems:
            print("    - " + line)
        return 1

    counts = ", ".join(f"{name} {len(rows)}"
                       for name, rows in data.items() if rows)
    notes.append("demo rows: " + counts)
    if not data["Child Enrollment"]:
        print(f"--- {path}\n  . no example data in this build (blank "
              f"version) - nothing to check")
        return 0

    # ---------------------------------------------------- the ID spine --
    room_by_id = {rid("Classroom", row["_row"]): row
                  for row in data["Classroom Management"]}
    family_by_id = {rid("Family", row["_row"]): row
                    for row in data["Parent Database"]}
    child_by_id = {rid("Child", row["_row"]): row
                   for row in data["Child Enrollment"]}
    staff_by_id = {rid("Staff", row["_row"]): row
                   for row in data["Staff Database"]}
    supply_by_id = {rid("Supply", row["_row"]): row
                    for row in data["Supply Inventory"]}
    ledger_by_id = {rid("Invoice", row["_row"]): row
                    for row in data["Invoices & Receipts"]}

    room_names = {row.get("name") for row in data["Classroom Management"]}

    def known(table, value, where, kind):
        if value in (None, ""):
            return False
        if value not in table:
            problems.append(f"{where}: {kind} '{value}' does not exist")
            return False
        return True

    for row in data["Child Enrollment"]:
        where = f"Child Enrollment row {row['_row']}"
        known(family_by_id, row.get("family"), where, "family")
        if row.get("room") and row["room"] not in room_names:
            problems.append(f"{where}: room '{row['room']}' is not a room")
    for row in data["Attendance"]:
        where = f"Attendance row {row['_row']}"
        known(child_by_id, row.get("child"), where, "child")
    for row in data["Tuition & Fees"]:
        where = f"Tuition & Fees row {row['_row']}"
        known(child_by_id, row.get("child"), where, "child")
        if row.get("invoice"):
            known(ledger_by_id, row["invoice"], where, "invoice")
    for row in data["Invoices & Receipts"]:
        where = f"Invoices & Receipts row {row['_row']}"
        known(family_by_id, row.get("family"), where, "family")
        if row.get("child"):
            known(child_by_id, row["child"], where, "child")
        if row.get("against"):
            known(ledger_by_id, row["against"], where, "invoice")
    for name, key in (("Daily Reports", "child"), ("Meal Tracker", "child"),
                      ("Allergy & Medical", "child"),
                      ("Incident Reports", "child")):
        for row in data[name]:
            known(child_by_id, row.get(key), f"{name} row {row['_row']}",
                  "child")
    for row in data["Staff Schedule"]:
        where = f"Staff Schedule row {row['_row']}"
        known(staff_by_id, row.get("staff"), where, "staff member")
        if row.get("room") and row["room"] not in room_names:
            problems.append(f"{where}: room '{row['room']}' is not a room")
    for row in data["Payroll Tracker"]:
        known(staff_by_id, row.get("staff"),
              f"Payroll Tracker row {row['_row']}", "staff member")
    for row in data["Parent Communication"]:
        where = f"Parent Communication row {row['_row']}"
        known(family_by_id, row.get("family"), where, "family")
        if row.get("child"):
            known(child_by_id, row["child"], where, "child")
    for row in data["Document Tracker"]:
        where = f"Document Tracker row {row['_row']}"
        holder, who = row.get("holder"), row.get("who")
        if holder == "Child" and who:
            known(child_by_id, who, where, "child")
        if holder == "Staff" and who:
            known(staff_by_id, who, where, "staff member")
    notes.append(f"ID spine: {len(child_by_id)} children, "
                 f"{len(family_by_id)} families, {len(staff_by_id)} staff, "
                 f"{len(room_by_id)} rooms, {len(supply_by_id)} supplies")

    # --------------------------------------------- multi-child families --
    per_family = defaultdict(list)
    for key, row in child_by_id.items():
        if row.get("family"):
            per_family[row["family"]].append(key)
    siblings = sum(1 for kids in per_family.values() if len(kids) > 1)
    if siblings < 2:
        problems.append("the demo needs at least two families with more "
                        "than one child - that is the headline feature")
    notes.append(f"families: {siblings} with siblings, "
                 f"{len(per_family)} with a child on the books")

    # ------------------------------------------------ rooms and capacity --
    in_room = defaultdict(int)
    for row in child_by_id.values():
        if row.get("status") in ("Active", "On hold", "Starting soon"):
            in_room[row.get("room")] += 1
    for row in data["Classroom Management"]:
        places = _num(row.get("capacity"))
        taken = in_room.get(row.get("name"), 0)
        if places and taken > places:
            problems.append(
                f"room '{row['name']}' holds {places:.0f} but "
                f"{taken} children are enrolled in it")
    notes.append("rooms: " + ", ".join(
        f"{row['name']} {in_room.get(row['name'], 0)}/"
        f"{_num(row.get('capacity')):.0f}"
        for row in data["Classroom Management"]))

    # -------------------------------------------------------- attendance --
    seen = set()
    for row in data["Attendance"]:
        day, child = _day(row.get("date")), row.get("child")
        if (day, child) in seen:
            problems.append(f"Attendance row {row['_row']}: {child} is "
                            f"marked twice on {day}")
        seen.add((day, child))
        enrolled = child_by_id.get(child, {})
        start = _day(enrolled.get("start_date"))
        end = _day(enrolled.get("end_date"))
        if start and day and day < start:
            problems.append(f"Attendance row {row['_row']}: {child} "
                            f"attended on {day}, before they started")
        if end and day and day > end:
            problems.append(f"Attendance row {row['_row']}: {child} "
                            f"attended on {day}, after they left")
        in_time, out_time = row.get("in_time"), row.get("out_time")
        if in_time and out_time and _minutes(out_time) <= _minutes(in_time):
            problems.append(f"Attendance row {row['_row']}: collected "
                            f"before they arrived")
    present = sum(1 for row in data["Attendance"]
                  if row.get("status") in ("Present", "Late",
                                           "Early pickup"))
    sessions = sum(1 for row in data["Attendance"]
                   if row.get("status") != "Holiday / closed")
    if sessions:
        rate = present / sessions
        if not 0.7 <= rate <= 0.99:
            problems.append(f"attendance runs at {rate:.0%} - between "
                            f"70% and 99% looks like a real nursery")
        notes.append(f"attendance: {len(data['Attendance'])} rows, "
                     f"{rate:.0%} present")
    missing_out = sum(1 for row in data["Attendance"]
                      if row.get("in_time") and not row.get("out_time"))
    if not missing_out:
        problems.append("no child is left without a check-out time - the "
                        "missing check-out alert has nothing to show")
    notes.append(f"check-out alerts: {missing_out} open")

    # ------------------------------------------------- charges and money --
    charge_total = defaultdict(float)
    for row in data["Tuition & Fees"]:
        qty = row.get("qty")
        qty = 1.0 if qty in (None, "") else float(qty)
        gross = round(qty * _num(row.get("rate")), 2)
        net = round(gross * (1 - _num(row.get("discount_pct"))), 2)
        if row.get("invoice"):
            charge_total[row["invoice"]] += net
    invoiced = paid = credited = 0.0
    for row in data["Invoices & Receipts"]:
        kind = row.get("type")
        amount = _num(row.get("amount"))
        key = rid("Invoice", row["_row"])
        if kind == "Invoice":
            lines = charge_total.get(key, 0.0)
            invoiced += lines if lines else amount
            if lines and amount and abs(lines - amount) > MONEY:
                problems.append(
                    f"invoice {key} totals {lines:,.2f} from the charge "
                    f"lines but {amount:,.2f} was typed in")
        elif kind == "Payment":
            paid += amount
        elif kind == "Refund":
            paid -= amount
        elif kind in CREDIT_TYPES:
            credited += amount
        label = f"{row.get('period') or ''} {row.get('description') or ''}"
        deposit = "deposit" in label.lower()
        if (kind in ("Payment", "Refund") and not row.get("against")
                and not deposit):
            # a deposit is paid before there is anything to pay it
            # against - everything else has to name its invoice
            problems.append(f"Invoices & Receipts row {row['_row']}: a "
                            f"{kind.lower()} with no invoice to match")
    owed = invoiced - credited - paid
    if invoiced <= 0:
        problems.append("nothing has been invoiced - the money pages "
                        "will all read zero")
    else:
        collection = paid / invoiced
        if not 0.6 <= collection <= 1.0:
            problems.append(
                f"the demo has collected {collection:.0%} of what it "
                f"billed - between 60% and 100% tells a better story")
        notes.append(f"money: billed {invoiced:,.0f}, collected "
                     f"{paid:,.0f} ({paid / invoiced:.0%}), credited "
                     f"{credited:,.0f}, outstanding {owed:,.0f}")
    over = defaultdict(float)
    for row in data["Invoices & Receipts"]:
        if row.get("type") == "Payment" and row.get("against"):
            over[row["against"]] += _num(row.get("amount"))
    for key, amount in over.items():
        target = charge_total.get(key)
        if target is None:
            invoice = ledger_by_id.get(key, {})
            target = _num(invoice.get("amount"))
        if target and amount - target > MONEY:
            problems.append(f"invoice {key} is overpaid: {amount:,.2f} "
                            f"against {target:,.2f}")
    unbilled = sum(1 for row in data["Tuition & Fees"]
                   if not row.get("invoice"))
    notes.append(f"charges: {len(data['Tuition & Fees'])} lines, "
                 f"{unbilled} not yet on an invoice")

    # ------------------------------------------------------- the rota ----
    rota = defaultdict(list)
    for row in data["Staff Schedule"]:
        if not row.get("start") or not row.get("finish"):
            continue
        start, finish = _minutes(row["start"]), _minutes(row["finish"])
        if finish <= start:
            problems.append(f"Staff Schedule row {row['_row']}: the shift "
                            f"finishes before it starts")
            continue
        rota[(row.get("staff"), _day(row.get("date")))].append(
            (start, finish, row["_row"]))
    clashes = 0
    for (who, day), shifts in rota.items():
        shifts.sort()
        for first, second in zip(shifts, shifts[1:]):
            if second[0] < first[1]:
                clashes += 1
                if clashes < 6:
                    problems.append(f"double shift: {who} on {day}, rows "
                                    f"{first[2]} and {second[2]}")
    hours = sum((f - s) / 60.0 for shifts in rota.values()
                for s, f, _ in shifts)
    notes.append(f"rota: {len(data['Staff Schedule'])} shifts, "
                 f"{hours:,.0f} hours, {clashes} clashes")

    # ------------------------------------------------------- the team ----
    expiring = 0
    for row in data["Staff Database"]:
        for key in ("first_aid_exp", "cpr_exp", "safeguard_exp",
                    "check_exp"):
            when = _day(row.get(key))
            if when and when <= cfg.today:
                expiring += 1
    if not expiring:
        problems.append("every certificate is in date - the expiry "
                        "reminder has nothing to show")
    paid_staff = [row for row in data["Staff Database"]
                  if row.get("status") == "Active"]
    notes.append(f"team: {len(paid_staff)} active, {expiring} lapsed "
                 f"certificates on file")

    # --------------------------------------------------------- the stock -
    for row in data["Supply Inventory"]:
        stock = (_num(row.get("opening")) + _num(row.get("received"))
                 - _num(row.get("used")) - _num(row.get("waste"))
                 + _num(row.get("adjust")))
        if stock < 0:
            problems.append(f"supply '{row.get('name')}' goes to "
                            f"{stock:,.0f} - stock cannot be negative")

    # ------------------------------------------------ the monthly shape --
    by_month = defaultdict(float)
    for row in data["Invoices & Receipts"]:
        if row.get("type") == "Invoice":
            key = rid("Invoice", row["_row"])
            amount = charge_total.get(key) or _num(row.get("amount"))
            by_month[_month(row.get("date"))] += amount
    if by_month:
        average = sum(by_month.values()) / len(by_month)
        target = cfg.revenue_target_month
        if average < target * 0.5 or average > target * 1.6:
            problems.append(
                f"the demo bills {average:,.0f} a month against a target "
                f"of {target:,.0f} - the dashboard will look odd")
        notes.append("billing by month: " + ", ".join(
            f"{key} {value:,.0f}" for key, value in sorted(by_month.items())))

    # ------------------------------------------------------- paperwork ---
    overdue_docs = sum(1 for row in data["Document Tracker"]
                       if _day(row.get("expires"))
                       and _day(row["expires"]) < cfg.today)
    if not overdue_docs:
        problems.append("no document has expired - the renewal alert has "
                        "nothing to show")
    notes.append(f"paperwork: {len(data['Document Tracker'])} documents, "
                 f"{overdue_docs} past their renewal date")

    print(f"--- {path}")
    for note in notes:
        print(f"  . {note}")
    if problems:
        print(f"  ! {len(problems)} problem(s):")
        for line in problems[:30]:
            print("    - " + line)
        return 1
    print("  OK - the example nursery adds up")
    return 0


if __name__ == "__main__":
    target = (sys.argv[1] if len(sys.argv) > 1
              else "output/Daycare Business Management Spreadsheet.xlsx")
    raise SystemExit(check(target))
