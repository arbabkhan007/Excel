"""
✅ Checklists - the four lists from the spec:

  1. Event Prep (before the big day)
  2. Shopping Run
  3. Day-Of
  4. End-of-Event

Copy a block per event if you like, or just tick through them for the
event you are working on - the counters at the top update live.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "checklists"
LAST_COL = "H"

BLOCKS = [
    ("prep", "  \U0001F4CB  1 \u00B7 EVENT PREP  \u2014  the week before", [
        ("Confirm final guest count with client", "7 days out"),
        ("Confirm menu + dietary requirements in writing", "7 days out"),
        ("Confirm venue access time, power & parking", "5 days out"),
        ("Staff roster locked and briefed", "4 days out"),
        ("Equipment reserved and checked against tracker", "3 days out"),
        ("Supplier orders placed (lead times respected)", "3 days out"),
        ("Prep started: marinades, sauces, desserts", "2 days out"),
        ("Day-of timeline printed and shared with the team", "1 day out"),
    ]),
    ("shop", "  \U0001F6D2  2 \u00B7 SHOPPING RUN", [
        ("Shopping List tab reviewed (To buy column)", "day before"),
        ("Cold chain planned: coolers, ice packs", "day before"),
        ("Proteins bought first, kept cold", "morning of"),
        ("Produce picked over for quality", "morning of"),
        ("Receipts kept for the \U0001F4B8 Expenses tab", "every trip"),
        ("\U0001F4E6 Inventory quantities updated", "same day"),
        ("Crates labelled per event", "on return"),
        ("Perishables chilled / frozen immediately", "on return"),
    ]),
    ("day", "  \U0001F525  3 \u00B7 DAY-OF", [
        ("Load-out checked against Equipment tracker", "departure"),
        ("Venue walkthrough with the client contact", "on arrival"),
        ("Chafers / ranges set up and test-fired", "setup"),
        ("Hand-wash & sanitising station out", "setup"),
        ("Team briefing: allergies, timings, roles", "1 hr before"),
        ("Food temps logged (hot >63\u00B0C, cold <5\u00B0C)", "before service"),
        ("Backup supplies within reach", "before service"),
        ("Service start time confirmed with venue", "before service"),
    ]),
    ("end", "  \U0001F3C1  4 \u00B7 END-OF-EVENT", [
        ("Leftovers packed, labelled, handed to client", "at close"),
        ("Equipment counted back into the van", "at close"),
        ("Venue left as found - final walk with contact", "at close"),
        ("Gas off, power off, temps logged", "at close"),
        ("Wages & tips settled \u2192 \U0001F477 Staff tab", "same night"),
        ("Actual costs logged \u2192 \U0001F4B8 Expenses tab", "same night"),
        ("Event marked \U0001F389 Completed on Events tab", "same night"),
        ("Final invoice sent + review/referral requested", "next day"),
    ]),
]

BLOCK_TOP = 9


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY, {"A": 2.2, "B": 4.5, "C": 22, "D": 22, "E": 18, "F": 13,
                    "G": 8, "H": 20})
    n_items = sum(len(items) for _k, _t, items in BLOCKS)
    foot = BLOCK_TOP + len(BLOCKS) * 2 + n_items + 8
    bk.paint(KEY, 0, 0, foot, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \u2705  The Four Checklists",
        "  Prep \u2192 Shop \u2192 Day-of \u2192 Close.  Tick as you go; "
        "the counters above each block update live.",
        LAST_COL)

    # ------------------------------------------------------------------
    # counters (one chip per block)
    # ------------------------------------------------------------------
    tick_col = "G"
    row = BLOCK_TOP
    ranges = {}
    for bkey, _title, items in BLOCKS:
        first = row + 1
        last = first + len(items) - 1
        ranges[bkey] = (first, last)
        row = last + 2

    colors = {"prep": "primary", "shop": "ok", "day": "accent", "end": "info"}
    short = {"prep": "\U0001F4CB Prep", "shop": "\U0001F6D2 Shop",
             "day": "\U0001F525 Day-of", "end": "\U0001F3C1 Close"}
    spans = [2, 2, 2, 1]
    col = 1
    for (bkey, _title, items), span in zip(BLOCKS, spans):
        first, last = ranges[bkey]
        rng = "$%s$%d:$%s$%d" % (tick_col, first, tick_col, last)
        done = len(m.checks.get(bkey, ()))
        fmt = S.pill(th.soft(colors[bkey]), getattr(th, colors[bkey]),
                     size=10.5, bold=True, align="left")
        if span > 1:
            ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS),
                           col + span - 1, "", fmt)
        ws.write_formula(
            r(C.ROW_STATS), col,
            '="%s: "&COUNTIF(%s,"%s")&"/%d"' % (short[bkey], rng, C.TICK,
                                                len(items)),
            fmt, "%s: %d/%d" % (short[bkey], done, len(items)))
        bk.stats["formulas"] += 1
        col += span
    if col < ci(LAST_COL):
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), ci(LAST_COL),
                       "", S.canvas)
    elif col == ci(LAST_COL):
        ws.write_blank(r(C.ROW_STATS), col, None, S.canvas)

    # ------------------------------------------------------------------
    # the four blocks
    # ------------------------------------------------------------------
    item_fmt = S.f(**S.base(font_size=10.5, font_color=th.ink,
                            bg_color=th.card, align="left", valign="vcenter",
                            border=1, border_color=th.border, indent=1))
    item_fmt_alt = S.f(**S.base(font_size=10.5, font_color=th.ink,
                                bg_color=th.alt, align="left",
                                valign="vcenter", border=1,
                                border_color=th.border, indent=1))
    when_fmt = S.f(**S.base(font_size=9.5, italic=True, font_color=th.muted,
                            bg_color=th.card, align="center",
                            valign="vcenter", border=1,
                            border_color=th.border))
    when_fmt_alt = S.f(**S.base(font_size=9.5, italic=True,
                                font_color=th.muted, bg_color=th.alt,
                                align="center", valign="vcenter", border=1,
                                border_color=th.border))
    block_colors = {"prep": th.primary, "shop": th.ok, "day": th.accent,
                    "end": th.info}
    for bi, (bkey, title, items) in enumerate(BLOCKS):
        first, last = ranges[bkey]
        hrow = first - 1
        ws.set_row(r(hrow), 24)
        ws.merge_range(r(hrow), 1, r(hrow), ci(LAST_COL), title,
                       S.f(**S.base(font_name=th.title_font, font_size=12,
                                    bold=True, font_color=th.white,
                                    bg_color=block_colors[bkey],
                                    align="left", valign="vcenter",
                                    indent=1)))
        checked = m.checks.get(bkey, set())
        for i, (text, when) in enumerate(items):
            row = first + i
            a = i % 2
            ws.set_row(r(row), 20)
            ws.write(r(row), ci("B"), i + 1, S.idx(a))
            ws.merge_range(r(row), ci("C"), r(row), ci("E"), text,
                           item_fmt_alt if a else item_fmt)
            ws.write(r(row), ci("F"), when, when_fmt_alt if a else when_fmt)
            tick = C.TICK if i in checked else ""
            ws.write(r(row), ci("G"), tick, S.cell("tick", a))
            ws.write_blank(r(row), ci("H"), None,
                           S.cell("text", a))
        # tick validation + glow for this block
        ws.data_validation(r(first), ci("G"), r(last), ci("G"), {
            "validate": "list", "source": "=Tick", "ignore_blank": True,
            "input_title": "Done?", "show_input": True,
            "input_message": "Pick \u2713 from the dropdown when done."})
        bk.stats["validations"] += 1
        bk.cond(KEY, first, ci("C"), last, ci("G"), {
            "type": "formula",
            "criteria": '=$G%d="%s"' % (first, C.TICK),
            "format": S.cf(bg=th.ok_soft, fg=th.ok, bold=True,
                           border=th.border)})
        bk.cond(KEY, first, ci("G"), last, ci("G"), {
            "type": "formula",
            "criteria": '=$G%d="%s"' % (first, C.TICK),
            "format": S.cf(bg=th.ok_soft, fg=th.ok, bold=True, size=13,
                           border=th.border)})

    # ------------------------------------------------------------------
    # footer
    # ------------------------------------------------------------------
    frow = row + 1
    ws.set_row(r(frow), 20)
    ws.set_row(r(frow + 1), 20)
    ws.merge_range(r(frow), 1, r(frow + 1), ci(LAST_COL),
                   "  \U0001F4A1  Running several events a week? Duplicate "
                   "this sheet (right-click the tab \u2192 Move or Copy "
                   "\u2192 Create a copy) and rename it per event. The "
                   "counters on each copy work independently.", S.note)

    nav = frow + 3
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)
    bk.page(KEY, LAST_COL, nav + 1, landscape=False, zoom=100)
