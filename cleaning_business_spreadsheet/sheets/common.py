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
    ``ClientPrefix`` or a quoted literal such as ``'"JOB-"'``.
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
               note="Automatic. Every monthly report groups on this.")


def weekday_col(date_key: str, key: str = "weekday", header: str = "Day",
                width: float = 9) -> Col:
    return Col(key, header, width, "text",
               formula=(f'=IF({{{date_key}}}{{r}}="","",'
                        f'TEXT({{{date_key}}}{{r}},"ddd"))'))


def week_col(date_key: str, key: str = "week", header: str = "Week",
             width: float = 7) -> Col:
    """Week number, honouring the week start day on Settings."""
    return Col(key, header, width, "int",
               formula=(f'=IF({{{date_key}}}{{r}}="","",'
                        f'WEEKNUM({{{date_key}}}{{r}},'
                        f'IF(WeekStart="Sunday",1,2)))'),
               note="Week number of the year, using the week start day set "
                    "on the Settings tab.")


def alert_cf() -> List[Dict[str, Any]]:
    """Traffic light colouring for the 'Alert' / 'Check' columns."""
    return [
        cf_text("OK", "good", "begins with"),
        cf_text("Overdue", "bad"),
        cf_text("Late", "bad"),
        cf_text("Missing", "bad"),
        cf_text("Conflict", "bad"),
        cf_text("Clash", "bad"),
        cf_text("Expired", "bad"),
        cf_text("Duplicate", "bad"),
        cf_text("Unpaid", "bad"),
        cf_text("Loss", "bad"),
        cf_text("Below", "bad"),
        cf_text("Action", "warn"),
        cf_text("Due", "warn"),
        cf_text("Check", "warn"),
        cf_text("Chase", "warn"),
        cf_text("Assign", "warn"),
        cf_text("Confirm", "warn"),
        cf_text("Order", "warn"),
        cf_text("Today", "info"),
        cf_text("Upcoming", "info"),
        cf_text("Tomorrow", "info"),
    ]


def status_cf() -> List[Dict[str, Any]]:
    """Colouring for the job, quote, invoice and stock status columns."""
    return [
        cf_text("Completed", "good"),
        cf_text("Complete", "good"),
        cf_text("Done", "good"),
        cf_text("Paid", "good"),
        cf_text("Accepted", "good"),
        cf_text("Won", "good"),
        cf_text("Active", "good"),
        cf_text("In stock", "good"),
        cf_text("Resolved", "good"),
        cf_text("Signed", "good"),
        cf_text("Yes", "good"),
        cf_text("Booked", "info"),
        cf_text("Confirmed", "info"),
        cf_text("Scheduled", "info"),
        cf_text("In progress", "info"),
        cf_text("Sent", "info"),
        cf_text("Invoiced", "info"),
        cf_text("Quoted", "info"),
        cf_text("Prospect", "info"),
        cf_text("Draft", "warn"),
        cf_text("Enquiry", "warn"),
        cf_text("Pending", "warn"),
        cf_text("Part paid", "warn"),
        cf_text("Followed up", "warn"),
        cf_text("Low stock", "warn"),
        cf_text("Due soon", "warn"),
        cf_text("On hold", "warn"),
        cf_text("Unassigned", "warn"),
        cf_text("Overdue", "bad"),
        cf_text("Cancelled", "bad"),
        cf_text("Declined", "bad"),
        cf_text("Lost", "bad"),
        cf_text("No access", "bad"),
        cf_text("Out of stock", "bad"),
        cf_text("Written off", "bad"),
        cf_text("Expired", "bad"),
        cf_text("Inactive", "muted"),
        cf_text("Archived", "muted"),
        cf_text("Not needed", "muted"),
    ]


def rate_cf(target_name: str = "TargetMargin") -> List[Dict[str, Any]]:
    """Green at or above the target, amber below it."""
    return [
        {"type": "cell", "criteria": ">=", "value": f"={target_name}",
         "fmt": "good"},
        {"type": "cell", "criteria": "<", "value": f"={target_name}",
         "fmt": "warn"},
    ]


def bar_cf(color: str = "#2F8F9D") -> List[Dict[str, Any]]:
    return [{"type": "data_bar", "bar_color": color}]


def scale_cf() -> List[Dict[str, Any]]:
    return [{"type": "3_color_scale", "min_color": "#FBD9DD",
             "mid_color": "#FFF0CE", "max_color": "#D7F2E3"}]


def money_cf() -> List[Dict[str, Any]]:
    return [{"type": "cell", "criteria": ">", "value": 0, "fmt": "bold_good"}]


def loss_cf() -> List[Dict[str, Any]]:
    """Red when a profit column goes negative."""
    return [{"type": "cell", "criteria": "<", "value": 0, "fmt": "bold_bad"},
            {"type": "cell", "criteria": ">", "value": 0, "fmt": "good"}]


def dup_cf() -> List[Dict[str, Any]]:
    return [cf_text("Duplicate", "bad"), cf_text("OK", "good")]


def flag_yes_no_cf() -> List[Dict[str, Any]]:
    return [cf_text("No", "warn"), cf_text("Yes", "good")]


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


def lookup_client(target: str, key: str = "client",
                  missing: str = "Not found") -> str:
    """Pull a column of the Clients tab for the client ID in `key`."""
    return (f'=IF({{{key}}}{{r}}="","",IFERROR(INDEX([[Clients.'
            f'{target}]],MATCH({{{key}}}{{r}},'
            f'[[Clients.id]],0)),"{missing}"))')


def lookup_property(target: str, key: str = "property",
                    missing: str = "Not found") -> str:
    """Pull a column of the Properties tab for the property ID in `key`."""
    return (f'=IF({{{key}}}{{r}}="","",IFERROR(INDEX([[Properties.'
            f'{target}]],MATCH({{{key}}}{{r}},'
            f'[[Properties.id]],0)),"{missing}"))')


def lookup_service(target: str, key: str = "service",
                   missing: str = "Add to Services") -> str:
    """Pull a column of the Services catalogue for the service in `key`."""
    return (f'=IF({{{key}}}{{r}}="","",IFERROR(INDEX([[Services.'
            f'{target}]],MATCH({{{key}}}{{r}},'
            f'[[Services.name]],0)),"{missing}"))')


def lookup_job(target: str, key: str = "job", missing: str = "") -> str:
    """Pull a column of the Jobs tab for the job ID in `key`."""
    return (f'=IF({{{key}}}{{r}}="","",IFERROR(INDEX([[Jobs.'
            f'{target}]],MATCH({{{key}}}{{r}},[[Jobs.id]],0)),"{missing}"))')


def lookup_employee(target: str, key: str = "employee",
                    missing: str = "") -> str:
    """Pull a column of the Employees tab for the name in `key`."""
    return (f'=IF({{{key}}}{{r}}="","",IFERROR(INDEX([[Employees.'
            f'{target}]],MATCH({{{key}}}{{r}},'
            f'[[Employees.name]],0)),"{missing}"))')


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
