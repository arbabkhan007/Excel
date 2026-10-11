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
    ``ContentPrefix`` or a quoted literal such as ``'"IDEA-"'``.
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
    """ISO-ish week number, honouring the week start day on Settings."""
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
        cf_text("Expired", "bad"),
        cf_text("Duplicate", "bad"),
        cf_text("Action", "warn"),
        cf_text("Due", "warn"),
        cf_text("Check", "warn"),
        cf_text("Add", "warn"),
        cf_text("Today", "info"),
        cf_text("Upcoming", "info"),
    ]


def status_cf() -> List[Dict[str, Any]]:
    """Colouring for the content workflow and every other status column."""
    return [
        cf_text("Published", "good"),
        cf_text("Approved", "good"),
        cf_text("Complete", "good"),
        cf_text("Active", "good"),
        cf_text("Agreed", "good"),
        cf_text("Yes", "good"),
        cf_text("Won", "good"),
        cf_text("Scheduled", "info"),
        cf_text("In production", "info"),
        cf_text("Running", "info"),
        cf_text("Sent", "info"),
        cf_text("Testing", "info"),
        cf_text("Idea", "warn"),
        cf_text("Brief", "warn"),
        cf_text("Draft", "warn"),
        cf_text("Awaiting review", "warn"),
        cf_text("Pending", "warn"),
        cf_text("On hold", "warn"),
        cf_text("Revision needed", "bad"),
        cf_text("Rejected", "bad"),
        cf_text("Cancelled", "bad"),
        cf_text("Missed", "bad"),
        cf_text("Overdue", "bad"),
        cf_text("Declined", "bad"),
        cf_text("Archived", "muted"),
        cf_text("Paused", "muted"),
        cf_text("Not started", "muted"),
    ]


def rate_cf(target_name: str = "TargetEngagement") -> List[Dict[str, Any]]:
    """Green above the target rate, amber below it."""
    return [
        {"type": "cell", "criteria": ">=", "value": f"={target_name}",
         "fmt": "good"},
        {"type": "cell", "criteria": "<", "value": f"={target_name}",
         "fmt": "warn"},
    ]


def bar_cf(color: str = "#6350A8") -> List[Dict[str, Any]]:
    return [{"type": "data_bar", "bar_color": color}]


def scale_cf() -> List[Dict[str, Any]]:
    return [{"type": "3_color_scale", "min_color": "#FBD9DD",
             "mid_color": "#FFF1CF", "max_color": "#D9F2E4"}]


def money_cf() -> List[Dict[str, Any]]:
    return [{"type": "cell", "criteria": ">", "value": 0, "fmt": "bold_good"}]


def dup_cf() -> List[Dict[str, Any]]:
    return [cf_text("Duplicate", "bad"), cf_text("OK", "good")]


def flag_yes_no_cf() -> List[Dict[str, Any]]:
    return [cf_text("No", "warn"), cf_text("Yes", "good")]


def yesno(note: str = "") -> Dict[str, Any]:
    return yes_no(note)


def score_validation(maximum: int = 10) -> Dict[str, Any]:
    return {
        "validate": "integer",
        "criteria": "between",
        "minimum": 1,
        "maximum": maximum,
        "error_title": f"Score 1-{maximum}",
        "error_message": f"Enter a whole number from 1 (low) to {maximum} "
                         f"(high).",
    }


def blank(trigger: str, body: str) -> str:
    """`=IF(trigger="","",body)` with the standard row placeholders."""
    return f'=IF({{{trigger}}}{{r}}="","",{body})'


def lookup_content(target: str, key: str = "content") -> str:
    """Pull a column of the Content Calendar for the content ID in `key`."""
    return (f'=IF({{{key}}}{{r}}="","",IFERROR(INDEX([[Content Calendar.'
            f'{target}]],MATCH({{{key}}}{{r}},'
            f'[[Content Calendar.id]],0)),"Not found"))')
