"""Small building blocks shared by the register sheets."""

from __future__ import annotations

from typing import Any, Dict, List

from ..layout import HEADER_ROW, Col
from ..writer import cf_text, yes_no

# Everything below assumes the standard geometry: headers on row 4, data from
# row 5 - ROW()-4 therefore gives a 1 based record number.
ROW_OFFSET = HEADER_ROW


def id_col(trigger: str, prefix: str, digits: int = 4, width: float = 11,
           header: str = "ID", note: str = "") -> Col:
    """Auto numbered ID column (blank until the trigger column is filled).

    `prefix` is a formula fragment - either a defined name such as
    ``ClientPrefix`` or a quoted literal such as ``'"APT-"'``.
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
              width: float = 10) -> Col:
    """`yyyy-mm` text column - the key every monthly roll-up groups on."""
    return Col(key, header, width, "text",
               formula=(f'=IF({{{date_key}}}{{r}}="","",'
                        f'TEXT({{{date_key}}}{{r}},"yyyy-mm"))'),
               hidden=True,
               note="Automatic. Every monthly report groups on this.")


def weekday_col(date_key: str, key: str = "weekday", header: str = "Day",
                width: float = 8) -> Col:
    return Col(key, header, width, "text",
               formula=(f'=IF({{{date_key}}}{{r}}="","",'
                        f'TEXT({{{date_key}}}{{r}},"ddd"))'))


def week_col(date_key: str, key: str = "week", header: str = "Week",
             width: float = 7) -> Col:
    """Week number of the year."""
    return Col(key, header, width, "int",
               formula=(f'=IF({{{date_key}}}{{r}}="","",'
                        f'WEEKNUM({{{date_key}}}{{r}},2))'),
               hidden=True)


def alert_cf() -> List[Dict[str, Any]]:
    """Traffic light colouring for the 'Alert' / 'Check' columns."""
    return [
        cf_text("OK", "good", "begins with"),
        cf_text("Overdue", "bad"),
        cf_text("Late", "bad"),
        cf_text("Missing", "bad"),
        cf_text("Clash", "bad"),
        cf_text("Conflict", "bad"),
        cf_text("Duplicate", "bad"),
        cf_text("Unpaid", "bad"),
        cf_text("Loss", "bad"),
        cf_text("Below", "bad"),
        cf_text("Out of stock", "bad"),
        cf_text("No show", "bad"),
        cf_text("Action", "warn"),
        cf_text("Due", "warn"),
        cf_text("Check", "warn"),
        cf_text("Chase", "warn"),
        cf_text("Assign", "warn"),
        cf_text("Confirm", "warn"),
        cf_text("Order", "warn"),
        cf_text("Rebook", "warn"),
        cf_text("Add", "warn"),
        cf_text("Today", "info"),
        cf_text("Tomorrow", "info"),
        cf_text("Upcoming", "info"),
    ]


def status_cf() -> List[Dict[str, Any]]:
    """Colouring for the booking, payment, stock and staff statuses."""
    return [
        cf_text("Completed", "good"),
        cf_text("Complete", "good"),
        cf_text("Paid", "good"),
        cf_text("Settled", "good"),
        cf_text("Arrived", "good"),
        cf_text("Active", "good"),
        cf_text("Regular", "good"),
        cf_text("VIP", "good"),
        cf_text("In stock", "good"),
        cf_text("Received", "good"),
        cf_text("Approved", "good"),
        cf_text("Yes", "good"),
        cf_text("Booked", "info"),
        cf_text("Confirmed", "info"),
        cf_text("Scheduled", "info"),
        cf_text("New", "info"),
        cf_text("Ordered", "info"),
        cf_text("Running", "info"),
        cf_text("Part paid", "warn"),
        cf_text("Pending", "warn"),
        cf_text("Enquiry", "warn"),
        cf_text("Rescheduled", "warn"),
        cf_text("Low stock", "warn"),
        cf_text("Lapsing", "warn"),
        cf_text("Planned", "warn"),
        cf_text("On hold", "warn"),
        cf_text("Draft", "warn"),
        cf_text("Cancelled", "bad"),
        cf_text("No show", "bad"),
        cf_text("Refunded", "bad"),
        cf_text("Out of stock", "bad"),
        cf_text("Overdue", "bad"),
        cf_text("Written off", "bad"),
        cf_text("Inactive", "muted"),
        cf_text("Archived", "muted"),
        cf_text("Finished", "muted"),
        cf_text("Discontinued", "muted"),
    ]


def bar_cf(color: str = "#9A5C86") -> List[Dict[str, Any]]:
    return [{"type": "data_bar", "bar_color": color}]


def scale_cf() -> List[Dict[str, Any]]:
    return [{"type": "3_color_scale", "min_color": "#F8DAD9",
             "mid_color": "#FDEDD2", "max_color": "#DCEBD8"}]


def money_cf() -> List[Dict[str, Any]]:
    return [{"type": "cell", "criteria": ">", "value": 0, "fmt": "bold_good"}]


def loss_cf() -> List[Dict[str, Any]]:
    """Red when a profit column goes negative."""
    return [{"type": "cell", "criteria": "<", "value": 0, "fmt": "bold_bad"},
            {"type": "cell", "criteria": ">", "value": 0, "fmt": "good"}]


def owing_cf() -> List[Dict[str, Any]]:
    """Amber on anything still owed, green on a cleared balance."""
    return [{"type": "cell", "criteria": ">", "value": 0.004, "fmt": "warn"},
            {"type": "cell", "criteria": "<", "value": -0.004, "fmt": "info"}]


def dup_cf() -> List[Dict[str, Any]]:
    return [cf_text("Duplicate", "bad"), cf_text("OK", "good")]


def flag_yes_no_cf() -> List[Dict[str, Any]]:
    return [cf_text("No", "warn"), cf_text("Yes", "good")]


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


def lookup_client(target: str, key: str = "client",
                  missing: str = "Not in Clients") -> str:
    return lookup("Clients", target, "id", key, missing)


def lookup_service(target: str, key: str = "service",
                   missing: str = "Not in Services") -> str:
    return lookup("Services", target, "id", key, missing)


def lookup_staff(target: str, key: str = "staff",
                 missing: str = "Not in Staff") -> str:
    return lookup("Staff", target, "id", key, missing)


def lookup_product(target: str, key: str = "product",
                   missing: str = "Not in Inventory") -> str:
    return lookup("Inventory", target, "id", key, missing)


def lookup_appointment(target: str, key: str = "appointment",
                       missing: str = "Not in Appointments") -> str:
    return lookup("Appointments", target, "id", key, missing)


#: how each rounding rule on the Settings tab is applied.  Every branch
#: rounds *up* to the nearest charm price so a rounded price can never
#: drop below the price the maths asked for.
ROUNDING_BRANCHES = [
    ('"None"', "{x}"),
    ('"Nearest 0.99"', 'MAX(0.99,CEILING({x},1)-0.01)'),
    ('"Nearest 0.95"', 'MAX(0.95,CEILING({x},1)-0.05)'),
    ('"Nearest 0.50"', 'CEILING({x},0.5)'),
    ('"Nearest 1"', 'CEILING({x},1)'),
    ('"Nearest 5"', 'CEILING({x},5)'),
    ('"Up to whole 1"', 'CEILING({x},1)'),
    ('"Up to whole 5"', 'CEILING({x},5)'),
]


def round_price(cell: str, rule: str = "RoundingRule") -> str:
    """Apply the rounding rule chosen on Settings to a price `cell`.

    Pass a single cell reference, not an expression - the reference is
    repeated once per rule, and a long expression would be copied eight
    times into every row of the service menu.
    """
    out = f"ROUND({cell},2)"
    for test, body in reversed(ROUNDING_BRANCHES):
        out = f'IF({rule}={test},{body.format(x=cell)},{out})'
    return out


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
    to pull the nth line of one order without an array formula.
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

#: Sales tax on a line, honouring the two switches on Settings: tax is
#: added to retail, to services, to both or to neither.
def tax_on(amount: str, kind_is_retail: str) -> str:
    """`amount` x the tax rate, if that kind of line is taxable."""
    return (f'ROUND({amount}*IF({kind_is_retail},'
            f'IF(TaxOnRetail="Yes",TaxRate,0),'
            f'IF(TaxOnServices="Yes",TaxRate,0)),2)')


#: What the card machine keeps out of a payment.
def card_fee(amount: str, method_cell: str) -> str:
    return (f'IF({method_cell}="Card",ROUND({amount}*CardFeePercent+'
            f'CardFeeFixed,2),0)')
