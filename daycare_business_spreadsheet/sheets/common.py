"""Small building blocks shared by the register sheets."""

from __future__ import annotations

from typing import Any, Dict, List

from ..config import (ABSENT_STATUSES, CREDIT_FEE_TYPES,
                      NON_SESSION_STATUSES, OCCUPYING_STATUSES,
                      PRESENT_STATUSES)
from ..layout import HEADER_ROW, Col
from ..writer import cf_text, yes_no

# Everything below assumes the standard geometry: headers on row 4, data from
# row 5 - ROW()-4 therefore gives a 1 based record number.
ROW_OFFSET = HEADER_ROW


# ---------------------------------------------------------------------------
# Standard columns
# ---------------------------------------------------------------------------

def id_col(trigger: str, prefix: str, digits: int = 4, width: float = 11,
           header: str = "ID", note: str = "") -> Col:
    """Auto numbered ID column (blank until the trigger column is filled).

    `prefix` is a formula fragment - either a defined name such as
    ``ChildPrefix`` or a quoted literal such as ``'"INC-"'``.
    """
    zeros = "0" * digits
    return Col(
        "id", header, width, "id",
        formula=(f'=IF({{{trigger}}}{{r}}="","",'
                 f'{prefix}&TEXT(ROW()-{ROW_OFFSET},"{zeros}"))'),
        note=note or "Generated automatically from the row number. "
                     "Filter the list rather than re-sorting it so the IDs "
                     "keep matching your other tabs.")


def month_col(date_key: str, key: str = "month", header: str = "Month",
              width: float = 10, hidden: bool = True) -> Col:
    """`yyyy-mm` text column - the key every monthly roll-up groups on."""
    return Col(key, header, width, "text",
               formula=(f'=IF({{{date_key}}}{{r}}="","",'
                        f'TEXT({{{date_key}}}{{r}},"yyyy-mm"))'),
               hidden=hidden,
               note="Automatic. Every monthly report groups on this.")


def year_col(date_key: str, key: str = "year", header: str = "Year",
             width: float = 8) -> Col:
    return Col(key, header, width, "int",
               formula=(f'=IF({{{date_key}}}{{r}}="","",'
                        f'YEAR({{{date_key}}}{{r}}))'),
               hidden=True)


def weekday_col(date_key: str, key: str = "weekday", header: str = "Day",
                width: float = 8, hidden: bool = False) -> Col:
    return Col(key, header, width, "text",
               formula=(f'=IF({{{date_key}}}{{r}}="","",'
                        f'TEXT({{{date_key}}}{{r}},"ddd"))'),
               hidden=hidden)


def week_col(date_key: str, key: str = "week", header: str = "Week",
             width: float = 7) -> Col:
    """Week number of the year - the rota and the payroll group on it."""
    return Col(key, header, width, "int",
               formula=(f'=IF({{{date_key}}}{{r}}="","",'
                        f'WEEKNUM({{{date_key}}}{{r}},2))'),
               hidden=True)


def week_start_col(date_key: str, key: str = "week_start",
                   header: str = "Week beginning", width: float = 12) -> Col:
    """The Monday of that week - a far safer grouping key than WEEKNUM."""
    return Col(key, header, width, "date",
               formula=(f'=IF({{{date_key}}}{{r}}="","",'
                        f'{{{date_key}}}{{r}}-WEEKDAY({{{date_key}}}{{r}},3))'),
               hidden=True,
               note="The Monday of that week. Weekly totals group on it.")


# ---------------------------------------------------------------------------
# Conditional formatting sets
# ---------------------------------------------------------------------------

def alert_cf() -> List[Dict[str, Any]]:
    """Traffic light colouring for the 'Action' / 'Check' columns."""
    return [
        cf_text("OK", "good", "begins with"),
        cf_text("Overdue", "bad"),
        cf_text("Expired", "bad"),
        cf_text("Missing", "bad"),
        cf_text("Over capacity", "bad"),
        cf_text("Over ratio", "bad"),
        cf_text("Under ratio", "bad"),
        cf_text("Unsafe", "bad"),
        cf_text("Allergy", "bad"),
        cf_text("Not notified", "bad"),
        cf_text("No check-out", "bad"),
        cf_text("Unexplained", "bad"),
        cf_text("Duplicate", "bad"),
        cf_text("Loss", "bad"),
        cf_text("Unpaid", "bad"),
        cf_text("Not authorised", "bad"),
        cf_text("Chase", "warn"),
        cf_text("Due", "warn"),
        cf_text("Check", "warn"),
        cf_text("Action", "warn"),
        cf_text("Expiring", "warn"),
        cf_text("Nearly full", "warn"),
        cf_text("Order", "warn"),
        cf_text("Low stock", "warn"),
        cf_text("Confirm", "warn"),
        cf_text("Add", "warn"),
        cf_text("Invoice", "warn"),
        cf_text("Follow up", "warn"),
        cf_text("Today", "info"),
        cf_text("Tomorrow", "info"),
        cf_text("Upcoming", "info"),
        cf_text("Birthday", "info"),
        cf_text("Starting", "info"),
        cf_text("Space", "info"),
    ]


def status_cf() -> List[Dict[str, Any]]:
    """Colouring for the enrollment, attendance, invoice and staff statuses."""
    return [
        cf_text("Active", "good"),
        cf_text("Present", "good"),
        cf_text("Paid", "good"),
        cf_text("On file", "good"),
        cf_text("Signed", "good"),
        cf_text("Received", "good"),
        cf_text("Enrolled", "good"),
        cf_text("Offer accepted", "good"),
        cf_text("In stock", "good"),
        cf_text("Complete", "good"),
        cf_text("Closed", "good"),
        cf_text("Yes", "good"),
        cf_text("Pending start", "info"),
        cf_text("Sent", "info"),
        cf_text("Tour booked", "info"),
        cf_text("Toured", "info"),
        cf_text("Place offered", "info"),
        cf_text("New enquiry", "info"),
        cf_text("Probation", "info"),
        cf_text("Scheduled", "info"),
        cf_text("Meeting booked", "info"),
        cf_text("Late", "warn"),
        cf_text("Part paid", "warn"),
        cf_text("Draft", "warn"),
        cf_text("Requested", "warn"),
        cf_text("Chased", "warn"),
        cf_text("Waitlisted", "warn"),
        cf_text("On hold", "warn"),
        cf_text("Notice period", "warn"),
        cf_text("Notice given", "warn"),
        cf_text("On leave", "warn"),
        cf_text("Open", "warn"),
        cf_text("Low stock", "warn"),
        cf_text("Early pickup", "warn"),
        cf_text("Absent", "bad"),
        cf_text("Overdue", "bad"),
        cf_text("Expired", "bad"),
        cf_text("Missing", "bad"),
        cf_text("Written off", "bad"),
        cf_text("Cancelled", "bad"),
        cf_text("Out of stock", "bad"),
        cf_text("Offer declined", "bad"),
        cf_text("Lost", "bad"),
        cf_text("Withdrawn", "muted"),
        cf_text("Graduated", "muted"),
        cf_text("Inactive", "muted"),
        cf_text("Left", "muted"),
        cf_text("Not required", "muted"),
        cf_text("Not scheduled", "muted"),
        cf_text("Centre closed", "muted"),
    ]


def severity_cf() -> List[Dict[str, Any]]:
    """Allergy and incident severity - the one place red really matters."""
    return [
        cf_text("Anaphylaxis", "bold_bad"),
        cf_text("Emergency services", "bold_bad"),
        cf_text("Severe", "bad"),
        cf_text("Serious", "bad"),
        cf_text("Moderate", "warn"),
        cf_text("Mild", "info"),
        cf_text("Minor", "info"),
        cf_text("No injury", "good"),
        cf_text("None recorded", "muted"),
    ]


def bar_cf(color: str = "#4D96AE") -> List[Dict[str, Any]]:
    return [{"type": "data_bar", "bar_color": color}]


def scale_cf() -> List[Dict[str, Any]]:
    return [{"type": "3_color_scale", "min_color": "#F8DAD9",
             "mid_color": "#FDEDD2", "max_color": "#DBEDD7"}]


def money_cf() -> List[Dict[str, Any]]:
    return [{"type": "cell", "criteria": ">", "value": 0, "fmt": "bold_good"}]


def loss_cf() -> List[Dict[str, Any]]:
    """Red when a profit column goes negative."""
    return [{"type": "cell", "criteria": "<", "value": 0, "fmt": "bold_bad"},
            {"type": "cell", "criteria": ">", "value": 0, "fmt": "good"}]


def owing_cf() -> List[Dict[str, Any]]:
    """Amber on anything still owed, blue on a credit balance."""
    return [{"type": "cell", "criteria": ">", "value": 0.004, "fmt": "warn"},
            {"type": "cell", "criteria": "<", "value": -0.004, "fmt": "info"}]


def dup_cf() -> List[Dict[str, Any]]:
    return [cf_text("Duplicate", "bad"), cf_text("OK", "good")]


def flag_yes_no_cf() -> List[Dict[str, Any]]:
    return [cf_text("No", "warn"), cf_text("Yes", "good")]


def flag_no_bad_cf() -> List[Dict[str, Any]]:
    """For the ones where 'No' is a real problem, not a preference."""
    return [cf_text("No", "bad"), cf_text("Yes", "good")]


def target_cf(target_name: str) -> List[Dict[str, Any]]:
    """Green at or above a named target, amber below it."""
    return [
        {"type": "cell", "criteria": ">=", "value": f"={target_name}",
         "fmt": "good"},
        {"type": "cell", "criteria": "<", "value": f"={target_name}",
         "fmt": "warn"},
    ]


def yesno(note: str = "") -> Dict[str, Any]:
    return yes_no(note)


def score_validation(maximum: int = 5) -> Dict[str, Any]:
    return {
        "validate": "integer",
        "criteria": "between",
        "minimum": 1,
        "maximum": maximum,
        "error_title": f"Score 1-{maximum}",
        "error_message": f"Enter a whole number from 1 (poor) to {maximum} "
                         f"(excellent).",
    }


def blank(trigger: str, body: str) -> str:
    """`=IF(trigger="","",body)` with the standard row placeholders."""
    return f'=IF({{{trigger}}}{{r}}="","",{body})'


# ---------------------------------------------------------------------------
# Cross sheet lookups.  Every one is blank-row safe and says what to do
# when the ID does not exist, because a typo in an ID is the single most
# common thing that goes wrong in a workbook like this.
# ---------------------------------------------------------------------------

def lookup(sheet: str, target: str, key_col: str, key: str,
           missing: str = "") -> str:
    """Generic `INDEX/MATCH` against any register, blank row safe."""
    return (f'=IF({{{key}}}{{r}}="","",IFERROR(INDEX([[{sheet}.'
            f'{target}]],MATCH({{{key}}}{{r}},'
            f'[[{sheet}.{key_col}]],0)),"{missing}"))')


def lookup_num(sheet: str, target: str, key_col: str, key: str) -> str:
    """Numeric `INDEX/MATCH` - 0 when the ID is missing or the cell empty.

    `INDEX` on an empty cell returns 0, not "", so numeric lookups are
    coerced with `*1` inside the `IFERROR` rather than compared to "".
    """
    return (f'=IF({{{key}}}{{r}}="","",IFERROR(INDEX([[{sheet}.'
            f'{target}]],MATCH({{{key}}}{{r}},'
            f'[[{sheet}.{key_col}]],0))*1,0))')


def lookup_child(target: str, key: str = "child",
                 missing: str = "Not in Child Enrollment") -> str:
    return lookup("Child Enrollment", target, "id", key, missing)


def lookup_child_num(target: str, key: str = "child") -> str:
    return lookup_num("Child Enrollment", target, "id", key)


def lookup_family(target: str, key: str = "family",
                  missing: str = "Not in Parent Database") -> str:
    return lookup("Parent Database", target, "id", key, missing)


def lookup_staff(target: str, key: str = "staff",
                 missing: str = "Not in Staff Database") -> str:
    return lookup("Staff Database", target, "id", key, missing)


def lookup_staff_num(target: str, key: str = "staff") -> str:
    return lookup_num("Staff Database", target, "id", key)


def lookup_room(target: str, key: str = "room",
                missing: str = "Not in Classroom Management") -> str:
    return lookup("Classroom Management", target, "name", key, missing)


def lookup_supply(target: str, key: str = "supply",
                  missing: str = "Not in Supply Inventory") -> str:
    return lookup("Supply Inventory", target, "id", key, missing)


# ---------------------------------------------------------------------------
# Shared formula fragments
# ---------------------------------------------------------------------------

def in_list(cell: str, values: List[str]) -> str:
    """`OR(cell="a",cell="b",...)` written as a COUNTIF so it stays short.

    `OR()` evaluates every argument, which makes long chains both slow
    and unsafe next to a lookup, so membership tests are done with an
    arithmetic sum of equality tests instead.
    """
    tests = "+".join(f'({cell}="{v}")' for v in values)
    return f"({tests})>0"


def is_present(cell: str) -> str:
    """True when an attendance status means the child was in the building."""
    return in_list(cell, PRESENT_STATUSES)


def is_absent(cell: str) -> str:
    """True when a scheduled session was missed."""
    return in_list(cell, ABSENT_STATUSES)


def is_session(cell: str) -> str:
    """True when a session was expected at all (the attendance base)."""
    tests = "+".join(f'({cell}="{v}")' for v in NON_SESSION_STATUSES)
    return f'IF({cell}="",FALSE,({tests})=0)'


def is_enrolled(cell: str) -> str:
    """True when an enrollment status means the child holds a place."""
    return in_list(cell, OCCUPYING_STATUSES)


def is_credit(cell: str) -> str:
    """True when a fee type reduces the bill instead of adding to it."""
    return in_list(cell, CREDIT_FEE_TYPES)


#: `SUMIFS` across the attendance register for one child, counting only
#: the statuses that mean they were actually here.
def count_present(child_cell: str, extra: str = "") -> str:
    parts = "+".join(
        f'COUNTIFS([[Attendance.child]],{child_cell},'
        f'[[Attendance.status]],"{status}"{extra})'
        for status in PRESENT_STATUSES)
    return parts


def count_absent(child_cell: str, extra: str = "") -> str:
    parts = "+".join(
        f'COUNTIFS([[Attendance.child]],{child_cell},'
        f'[[Attendance.status]],"{status}"{extra})'
        for status in ABSENT_STATUSES)
    return parts


def age_months(dob_cell: str, on_cell: str = "TodayDate") -> str:
    """Whole months between a date of birth and a date.

    `DATEDIF` is the one Excel function that gets this right across a
    leap year, and it has existed since Excel 5, so it is safe here.
    """
    return f'IF({dob_cell}="","",DATEDIF({dob_cell},MAX({dob_cell},{on_cell}),"m"))'


def age_years_text(dob_cell: str, on_cell: str = "TodayDate") -> str:
    """`3y 7m` - how a nursery actually says a child's age."""
    return (f'IF({dob_cell}="","",'
            f'DATEDIF({dob_cell},MAX({dob_cell},{on_cell}),"y")&"y "&'
            f'DATEDIF({dob_cell},MAX({dob_cell},{on_cell}),"ym")&"m")')


def next_birthday(dob_cell: str) -> str:
    """The next anniversary of `dob_cell` on or after today."""
    return (f'IF({dob_cell}="","",'
            f'IF(DATE(YEAR(TodayDate),MONTH({dob_cell}),DAY({dob_cell}))'
            f'<TodayDate,'
            f'DATE(YEAR(TodayDate)+1,MONTH({dob_cell}),DAY({dob_cell})),'
            f'DATE(YEAR(TodayDate),MONTH({dob_cell}),DAY({dob_cell}))))')


def hours_between(start: str, finish: str) -> str:
    """Decimal hours between two time cells, tolerant of an overnight end."""
    return (f'IF({start}="","",IF({finish}="","",'
            f'ROUND((MOD({finish}-{start},1))*24,2)))')


def rank_desc(key: str, guard: str = '{name}{r}=""',
              header: str = "Rank", width: float = 9,
              note: str = "") -> Col:
    """Hidden dense rank, 1 = the biggest value in `key`.

    Ties are broken by row order, so the ranks come out 1, 2, 3 ... with
    no gaps and no duplicates.  Pages then read the nth row with an
    ordinary `MATCH(n, rank, 0)` - no array formula, so nothing in the
    workbook needs Ctrl+Shift+Enter.
    """
    return Col(f"{key}_rank", header, width, "int",
               formula=(f'=IF({guard},"",COUNTIFS({{{key}_first}}:'
                        f'{{{key}_last}},">"&{{{key}}}{{r}})+'
                        f'COUNTIFS({{{key}_first}}:{{{key}}}{{r}},'
                        f'{{{key}}}{{r}}))'),
               hidden=True, note=note or f"1 = the largest {key}.")


def rank_asc(key: str, guard: str = '{name}{r}=""',
             header: str = "Rank", width: float = 9,
             note: str = "") -> Col:
    """Hidden dense rank, 1 = the smallest value in `key`."""
    return Col(f"{key}_rank", header, width, "int",
               formula=(f'=IF({guard},"",COUNTIFS({{{key}_first}}:'
                        f'{{{key}_last}},"<"&{{{key}}}{{r}})+'
                        f'COUNTIFS({{{key}_first}}:{{{key}}}{{r}},'
                        f'{{{key}}}{{r}}))'),
               hidden=True, note=note or f"1 = the smallest {key}.")


def min_number(value_range: str, cond: str, cap: int = 999999) -> str:
    """Smallest positive number in `value_range` where `cond` holds.

    Excel 2010 has no MINIFS, and a plain SUMPRODUCT(MIN(...)) trips over
    the empty strings that formula columns leave behind, so the values are
    coerced with IFERROR(...*1,0) and the non matching rows are swapped for
    a big number.  Returns `cap` when nothing matches - test for that.
    """
    safe = f"IFERROR({value_range}*1,0)"
    hit = f"(({cond})*({safe}>0))"
    return f"SUMPRODUCT(MIN({hit}*{safe}+(1-{hit})*{cap}))"


def seq_key(parent: str, trigger: str = "") -> List[Col]:
    """Hidden `1, 2, 3 ...` within a parent, plus a `parent|n` key.

    Printables and per-parent summaries use `MATCH(parent & "|" & n, key, 0)`
    to pull the nth line of one invoice without an array formula.
    """
    trig = trigger or parent
    return [
        Col("seq", "Line", 6, "int",
            formula=blank(
                trig,
                f'COUNTIFS({{{parent}_first}}:{{{parent}}}{{r}},'
                f'{{{parent}}}{{r}})'),
            hidden=True),
        Col("key", "Key", 14, "text",
            formula=blank(trig, f'{{{parent}}}{{r}}&"|"&{{seq}}{{r}}'),
            hidden=True),
    ]


# ---------------------------------------------------------------------------
# Money rules used in more than one place
# ---------------------------------------------------------------------------

def tax_on(amount: str, is_tuition: str) -> str:
    """`amount` x the tax rate, if that kind of line is taxable.

    Childcare is tax exempt in a great many places, which is why the
    default is zero and both switches are set on the Settings tab rather
    than assumed here.
    """
    return (f'ROUND({amount}*IF({is_tuition},'
            f'IF(TaxOnTuition="Yes",TaxRate,0),'
            f'IF(TaxOnExtras="Yes",TaxRate,0)),2)')


def card_fee(amount: str, method_cell: str) -> str:
    """What the card processor keeps out of a payment."""
    return (f'IF({method_cell}="Card",ROUND({amount}*CardFeePercent+'
            f'CardFeeFixed,2),0)')


def late_pickup_charge(minutes_cell: str) -> str:
    """The late collection charge for N minutes past closing.

    Free for the grace period, then a flat first block, then a per minute
    rate - which is how most settings actually word it in the handbook.
    """
    return (f'IF({minutes_cell}<=LatePickupGrace,0,'
            f'ROUND(LatePickupFee+MAX(0,{minutes_cell}-LatePickupGrace-'
            f'LatePickupBlock)*LatePickupPerMin,2))')
