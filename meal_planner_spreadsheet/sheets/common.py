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
    ``RecipePrefix`` or a quoted literal such as ``'"MEAL-"'``.
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
    """Colouring for the meal, shopping and prep status columns."""
    return [
        cf_text("Prepared", "good"),
        cf_text("Cooked", "good"),
        cf_text("Eaten", "good"),
        cf_text("Bought", "good"),
        cf_text("Purchased", "good"),
        cf_text("Done", "good"),
        cf_text("Complete", "good"),
        cf_text("In stock", "good"),
        cf_text("Yes", "good"),
        cf_text("Planned", "info"),
        cf_text("Scheduled", "info"),
        cf_text("Shopping", "info"),
        cf_text("On the list", "info"),
        cf_text("Prepping", "info"),
        cf_text("Frozen", "info"),
        cf_text("To do", "warn"),
        cf_text("Not started", "warn"),
        cf_text("Low stock", "warn"),
        cf_text("Use soon", "warn"),
        cf_text("Defrosting", "warn"),
        cf_text("Substituted", "warn"),
        cf_text("Out of stock", "bad"),
        cf_text("Expired", "bad"),
        cf_text("Skipped", "bad"),
        cf_text("Binned", "bad"),
        cf_text("Wasted", "bad"),
        cf_text("Cancelled", "bad"),
        cf_text("Archived", "muted"),
        cf_text("Not needed", "muted"),
    ]


def rate_cf(target_name: str = "CalorieTarget") -> List[Dict[str, Any]]:
    """Green above the target rate, amber below it."""
    return [
        {"type": "cell", "criteria": ">=", "value": f"={target_name}",
         "fmt": "good"},
        {"type": "cell", "criteria": "<", "value": f"={target_name}",
         "fmt": "warn"},
    ]


def bar_cf(color: str = "#4E8A63") -> List[Dict[str, Any]]:
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


def lookup_recipe(target: str, key: str = "recipe") -> str:
    """Pull a column of the Recipes tab for the recipe ID in `key`."""
    return (f'=IF({{{key}}}{{r}}="","",IFERROR(INDEX([[Recipes.'
            f'{target}]],MATCH({{{key}}}{{r}},'
            f'[[Recipes.id]],0)),"Not found"))')


def lookup_ingredient(target: str, key: str = "ingredient") -> str:
    """Pull a column of the Ingredients price book for the name in `key`."""
    return (f'=IF({{{key}}}{{r}}="","",IFERROR(INDEX([[Ingredients.'
            f'{target}]],MATCH({{{key}}}{{r}},'
            f'[[Ingredients.name]],0)),"Add to Ingredients"))')


#: grams / millilitres / each - the base units everything is converted to
def to_base(qty_key: str, unit_key: str) -> str:
    """Convert a quantity in any unit into the base unit for its type.

    The factor comes from the unit table on the Lists tab, so adding a
    unit there is all it takes to support it everywhere.
    """
    return (f'=IF({{{qty_key}}}{{r}}="","",IFERROR({{{qty_key}}}{{r}}*'
            f'INDEX(UnitFactors,MATCH({{{unit_key}}}{{r}},UnitNames,0)),""))')


def base_unit_of(unit_key: str) -> str:
    """`g`, `ml` or `each` for the unit chosen in `unit_key`."""
    return (f'=IF({{{unit_key}}}{{r}}="","",IFERROR('
            f'INDEX(UnitBases,MATCH({{{unit_key}}}{{r}},UnitNames,0)),"?"))')


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
