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
    """Week number of the year."""
    return Col(key, header, width, "int",
               formula=(f'=IF({{{date_key}}}{{r}}="","",'
                        f'WEEKNUM({{{date_key}}}{{r}},2))'))


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
        cf_text("Discontinued", "muted"),
        cf_text("Idea", "muted"),
        cf_text("Prototype", "warn"),
        cf_text("Made to order", "info"),
        cf_text("Seasonal", "info"),
        cf_text("Review", "warn"),
        cf_text("Raise price", "warn"),
        cf_text("Inactive", "muted"),
        cf_text("Archived", "muted"),
        cf_text("Not needed", "muted"),
    ]


def rate_cf(target_name: str = "TargetRetailMargin") -> List[Dict[str, Any]]:
    """Green at or above the target, amber below it."""
    return [
        {"type": "cell", "criteria": ">=", "value": f"={target_name}",
         "fmt": "good"},
        {"type": "cell", "criteria": "<", "value": f"={target_name}",
         "fmt": "warn"},
    ]


def bar_cf(color: str = "#B5704F") -> List[Dict[str, Any]]:
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


def lookup_product(target: str, key: str = "product",
                   missing: str = "Not found") -> str:
    """Pull a column of the Products catalog for the ID in `key`."""
    return (f'=IF({{{key}}}{{r}}="","",IFERROR(INDEX([[Products.'
            f'{target}]],MATCH({{{key}}}{{r}},'
            f'[[Products.id]],0)),"{missing}"))')


def lookup_material(target: str, key: str = "material",
                    missing: str = "Add to Materials") -> str:
    """Pull a column of the Materials tab for the material in `key`."""
    return (f'=IF({{{key}}}{{r}}="","",IFERROR(INDEX([[Materials.'
            f'{target}]],MATCH({{{key}}}{{r}},'
            f'[[Materials.id]],0)),"{missing}"))')


def lookup_channel(target: str, key: str = "channel",
                   missing: str = "Add to Channels") -> str:
    """Pull a column of the Channels tab for the sales channel in `key`."""
    return (f'=IF({{{key}}}{{r}}="","",IFERROR(INDEX([[Channels.'
            f'{target}]],MATCH({{{key}}}{{r}},'
            f'[[Channels.name]],0)),"{missing}"))')


def lookup(sheet: str, target: str, key_col: str, key: str,
           missing: str = "") -> str:
    """Generic `INDEX/MATCH` against any register, blank row safe."""
    return (f'=IF({{{key}}}{{r}}="","",IFERROR(INDEX([[{sheet}.'
            f'{target}]],MATCH({{{key}}}{{r}},'
            f'[[{sheet}.{key_col}]],0)),"{missing}"))')


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
    times into every row of the catalog.
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


# ---------------------------------------------------------------------------
# The six specialist calculators share a tail
# ---------------------------------------------------------------------------

def craft_tail(trigger: str = "name") -> List[Col]:
    """Cost, suggested price, margin and hourly pay for a craft tab.

    Every specialist calculator works out its own `mat_cost`,
    `labor_cost`, `other_cost` and `minutes`; from there the sums are
    identical, so they live here once.  The suggested price uses the
    default channel's fees and the target margin from Settings, which
    makes each tab usable on its own even for a piece that is not in
    the catalog yet.
    """
    return [
        Col("total_cost", "Cost to make", 12, "money",
            formula=blank(
                trigger,
                'ROUND({mat_cost}{r}+{labor_cost}{r}+{other_cost}{r},2)'),
            note="Materials + labour + extras. Packaging, postage and "
                 "overhead are added in the catalog, which knows which "
                 "profile and channel this product uses."),
        Col("fee_pct", "Fee rate", 9, "percent",
            formula=blank(
                trigger,
                'IFERROR(INDEX([[Channels.pct_total]],MATCH(DefaultChannel,'
                '[[Channels.name]],0)),0)'),
            hidden=True),
        Col("fee_fixed", "Fixed fees", 9, "money",
            formula=blank(
                trigger,
                'IFERROR(INDEX([[Channels.fixed_total]],MATCH(DefaultChannel,'
                '[[Channels.name]],0)),0)'),
            hidden=True),
        Col("denom", "Denominator", 9, "number",
            formula=blank(trigger, '1-TargetRetailMargin-{fee_pct}{r}'),
            hidden=True),
        Col("suggested_raw", "Unrounded", 10, "money",
            formula=blank(
                trigger,
                'IF({denom}{r}<=0,"",({total_cost}{r}+{fee_fixed}{r})/'
                '{denom}{r})'),
            hidden=True),
        Col("suggested", "Suggested price", 13, "money",
            formula=blank(
                trigger,
                'IF({suggested_raw}{r}="","Check the settings",ROUND(' +
                round_price("{suggested_raw}{r}") + ',2))'),
            note="(Cost + fixed fees) / (1 - target margin - fee rate), "
                 "priced for the default channel on Settings."),
        Col("price", "Price used", 11, "money",
            formula=blank(
                trigger,
                'IF({price_input}{r}<>"",{price_input}{r},'
                'IF(ISNUMBER({suggested}{r}),{suggested}{r},""))')),
        Col("profit", "Profit", 11, "money",
            formula=blank(
                trigger,
                'IF({price}{r}="","",ROUND({price}{r}-{total_cost}{r}-'
                '({price}{r}*{fee_pct}{r}+{fee_fixed}{r}),2))'),
            cf=loss_cf()),
        Col("margin", "Margin", 9, "percent",
            formula=blank(
                trigger,
                'IF(OR({price}{r}="",{price}{r}=0),"",{profit}{r}/'
                '{price}{r})'),
            cf=[{"type": "cell", "criteria": "<", "value": 0, "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": "=MarginWarning",
                 "fmt": "warn"},
                {"type": "cell", "criteria": ">=", "value": "=MarginWarning",
                 "fmt": "good"}]),
        Col("hours", "Hours", 9, "number",
            formula=blank(trigger, 'ROUND({minutes}{r}/60,3)'),
            hidden=True),
        Col("wage_hour", "You earn an hour", 12, "money",
            formula=blank(
                trigger,
                'IF(OR({hours}{r}=0,{hours}{r}="",{profit}{r}=""),"",'
                'ROUND(({profit}{r}+{labor_cost}{r})/{hours}{r},2))'),
            cf=[{"type": "cell", "criteria": "<", "value": "=MinimumWage",
                 "fmt": "bad"}],
            note="Wage plus profit over the hours it takes. The honest "
                 "test of whether this piece is worth making."),
    ]


def craft_check(trigger: str = "name", needs: str = "mat_cost") -> Col:
    """The shared 'is this row sane' column for a specialist tab."""
    return Col(
        "check", "Check", 30, "status",
        formula=blank(
            trigger,
            'IF(AND({product}{r}<>"",ISERROR(MATCH({product}{r},'
            '[[Products.id]],0))),"Check - no such product ID",'
            f'IF({{{needs}}}{{r}}=0,"Add - no material cost yet",'
            'IF({minutes}{r}=0,"Add - how long does it take?",'
            'IF(AND({product}{r}<>"",COUNTIFS([[Material Usage.product]],'
            '{product}{r})>0),'
            '"Check - this product also has a bill of materials",'
            'IF(AND({product}{r}<>"",COUNTIFS([[Labor.product]],'
            '{product}{r})>0),'
            '"Check - this product also has labour lines",'
            'IF({wage_hour}{r}<MinimumWage,'
            '"Below - pays under minimum wage","OK")))))) '),
        cf=alert_cf())
