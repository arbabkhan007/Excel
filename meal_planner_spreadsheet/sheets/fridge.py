"""`Fridge & Freezer` - what is in the cold bits, and how long you have.

The Pantry tab holds everything; this page pulls out the fridge, the
freezer and the leftovers and sorts them by urgency, because that is
the list you actually want at five o'clock on a Tuesday.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Fridge & Freezer"
LAST_COL = 8
ROWS = 16


def _pick(location: str, n: int) -> str:
    """Nth item in `location`, soonest to expire first."""
    return f'MATCH({n},[[Pantry.{location.lower()}_rank]],0)'


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_mid)
    ws = s.ws

    s.banner("Fridge & Freezer",
             "The cold storage list, sorted by what needs eating "
             "first.", LAST_COL)
    s.set_col(0, 0, 6)
    s.set_col(1, 1, 26)
    for c in range(2, LAST_COL + 1):
        s.set_col(c, c, 15)

    row = 4
    kpis = [
        ("In the fridge",
         '=COUNTIFS([[Pantry.location]],"Fridge")', "int",
         "items logged"),
        ("In the freezer",
         '=COUNTIFS([[Pantry.location]],"Freezer")', "int",
         "items logged"),
        ("Needs eating this week",
         '=COUNTIF([[Pantry.status]],"Use soon")', "int",
         f"within {cfg.expiry_warning_days} days of the date"),
        ("Value in cold storage",
         '=SUMIFS([[Pantry.value]],[[Pantry.location]],"Fridge")+'
         'SUMIFS([[Pantry.value]],[[Pantry.location]],"Freezer")',
         "money", "at price book prices"),
    ]
    for i, (label, formula, style, note) in enumerate(kpis):
        s.kpi(row, i * 2, label, s.reg.render(formula), style=style,
              span=2, note=note)
    row += 4

    for location, blurb in [
            ("Fridge", "Soonest to expire at the top. Plan a meal "
                       "round the first three and the rest looks "
                       "after itself."),
            ("Freezer", "The freezer is where food goes to be "
                        "forgotten. Label everything, and work "
                        "through the top of this list once a month.")]:
        row = s.section(row, location, LAST_COL)
        row = s.note(row, blurb, LAST_COL, height=24)
        row = s.table_head(row, 0, ["#", "Item", "How much", "Unit",
                                    "Use by", "Days left", "Value",
                                    "Status"],
                           [5, 26, 12, 10, 13, 12, 12, 14])
        first = row
        for n in range(ROWS):
            pick = _pick(location, n + 1)
            s.write(row, 0, n + 1, fmts.chrome("label"))
            for col, key, style in [(1, "ingredient", "text"),
                                    (2, "qty", "number1"),
                                    (3, "unit", "text"),
                                    (4, "est_expiry", "date"),
                                    (5, "days_left", "int"),
                                    (6, "value", "money"),
                                    (7, "status", "status")]:
                ws.write_formula(
                    row, col,
                    s.reg.render(f'=IFERROR(INDEX([[Pantry.{key}]],'
                                 f'{pick}),"")'), fmts.calc(style))
            row += 1
        last = row - 1
        ws.conditional_format(first, 5, last, 5,
                              {"type": "cell", "criteria": "<",
                               "value": 0, "format": fmts.cf("bad")})
        ws.conditional_format(first, 5, last, 5,
                              {"type": "cell", "criteria": "between",
                               "minimum": 0,
                               "maximum": cfg.expiry_warning_days,
                               "format": fmts.cf("warn")})
        ws.conditional_format(first, 7, last, 7,
                              {"type": "text", "criteria": "containing",
                               "value": "Expired",
                               "format": fmts.cf("bad")})
        row += 1

    # ------------------------------------------------------- leftovers
    row = s.section(row, "Leftovers waiting to be eaten", LAST_COL)
    row = s.table_head(row, 0, ["#", "What it is", "Portions", "Where",
                                "Cooked", "Eat by", "Days left",
                                "Worth"],
                       [5, 28, 10, 12, 13, 13, 12, 12])
    first = row
    for n in range(8):
        pick = f'MATCH({n + 1},[[Leftovers.queue_rank]],0)'
        s.write(row, 0, n + 1, fmts.chrome("label"))
        for col, key, style in [(1, "item", "text"),
                                (2, "portions", "number1"),
                                (3, "location", "text"),
                                (4, "cooked", "date"),
                                (5, "use_by", "date"),
                                (6, "days_left", "int"),
                                (7, "value", "money")]:
            ws.write_formula(
                row, col,
                s.reg.render(f'=IFERROR(INDEX([[Leftovers.{key}]],'
                             f'{pick}),"")'), fmts.calc(style))
        row += 1
    ws.conditional_format(first, 6, row - 1, 6,
                          {"type": "cell", "criteria": "<=", "value": 1,
                           "format": fmts.cf("warn")})
    row += 1

    # ------------------------------------------------- defrost planner
    row = s.section(row, "Taking something out of the freezer", LAST_COL)
    row = s.table_head(row, 0, ["Size", "In the fridge",
                                "In cold water", "Never"],
                       [22, 22, 24, 30])
    defrost = [
        ("Small portion (1-2 servings)", "Overnight",
         "1-2 hours, change the water", "On the worktop all day"),
        ("Family batch (4-6 servings)", "24 hours",
         "Not really practical", "In the sun or a warm kitchen"),
        ("Whole chicken", "24-36 hours",
         "2-3 hours, change the water", "Refreezing once thawed"),
        ("Bread, cake, pastry", "A couple of hours out of the freezer",
         "-", "Worrying about it - it is fine"),
        ("Soup, stew, sauce", "Overnight",
         "Straight into the pan on a low heat",
         "Reheating twice"),
    ]
    for size, fridge, water, never in defrost:
        s.write(row, 0, size, fmts.chrome("label_bold"))
        s.write(row, 1, fridge, fmts.chrome("label"))
        s.merge(row, 2, 3, water, fmts.chrome("label"))
        s.merge(row, 4, LAST_COL, never, fmts.chrome("note"))
        row += 1
    row += 1

    s.note(row, "Days left comes from the use-by date you typed in, or "
                "from the shelf life on the Ingredients tab if you "
                "left it blank. Both are estimates - the packet and "
                "your own senses outrank this page every time.",
           LAST_COL, height=30)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
