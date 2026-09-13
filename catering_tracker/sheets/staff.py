"""
\U0001F477 Staff & Labor Calculator - one row per person per shift.

Total pay = hours x rate + overtime hours x rate x 1.5.  The Paid tick
feeds the "unpaid shifts" alert on the dashboard.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "staff"
LAST_COL = "K"

COLUMNS = [
    ("n", "#", "idx", None),
    ("name", "Team member", "text", None),
    ("role", "Role", "center", None),
    ("event", "Event ID", "center", None),
    ("hours", "Hours", "qty", None),
    ("rate", "Rate / hr", "money", None),
    ("ot", "OT hrs", "qty", None),
    ("total", "Total pay", "calc_money", "primary_2"),
    ("paid", "\u2705 Paid", "tick", None),
    ("notes", "Notes", "text", None),
]


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY)
    bk.paint(KEY, 0, 0, C.last_row(KEY) + 24, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F477  Staff & Labor Calculator",
        "  One row per person per shift.  Overtime is paid at 1.5x "
        "automatically; the total feeds your labor cost picture.",
        LAST_COL)

    # ------------------------------------------------------------------
    # quick-stat strip
    # ------------------------------------------------------------------
    tot = bk.rng(KEY, "total")
    paid_c = bk.rng(KEY, "paid")
    name_c = bk.rng(KEY, "name")
    shifts_cached = len(m.staff)
    hours_cached = sum(s["hours"] + s["ot"] for s in m.staff)
    cost_cached = sum(s["total"] for s in m.staff)
    paid_cached = sum(s["total"] for s in m.staff if s["paid"] == C.TICK)
    chips = [
        ('\U0001F477 Shifts logged: "&COUNTIF(%s,"?*")' % name_c,
         "primary", "Shifts logged: %d" % shifts_cached),
        ('\u23F1 Hours worked: "&TEXT(SUM($F$%d:$F$%d)+SUM($H$%d:$H$%d),"0")'
         % (C.ROW_FIRST, C.last_row(KEY), C.ROW_FIRST, C.last_row(KEY)),
         "info", "Hours worked: %d" % round(hours_cached)),
        ('\U0001F4B0 Labor booked: "&Currency&TEXT(SUM(%s),"#,##0")' % tot,
         "primary_2", "Labor booked: %s" % m.money(cost_cached)),
        ('\u2705 Paid out: "&Currency&TEXT(SUMIF(%s,"%s",%s),"#,##0")'
         % (paid_c, C.TICK, tot), "ok",
         "Paid out: %s" % m.money(paid_cached)),
        ('\U0001F534 Unpaid shifts: "&%s' % bk.kpi("staff_unpaid"),
         "bad", "Unpaid shifts: %d" % m.agg.get("staff_unpaid", 0)),
    ]
    col = 1
    for formula, color, cached in chips:
        fmt = S.pill(th.soft(color), getattr(th, color), size=10.5,
                     bold=True, align="left")
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), col + 1, "", fmt)
        ws.write_formula(r(C.ROW_STATS), col, '="' + formula, fmt, cached)
        bk.stats["formulas"] += 1
        col += 2
    if col < ci(LAST_COL):
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), ci(LAST_COL),
                       "", S.canvas)
    elif col == ci(LAST_COL):
        ws.write_blank(r(C.ROW_STATS), col, None, S.canvas)

    # ------------------------------------------------------------------
    # table
    # ------------------------------------------------------------------
    K.table_frame(bk, KEY, COLUMNS, height=20)
    for i in range(K.n_rows(KEY)):
        rownum = C.ROW_FIRST + i
        values = {
            "n": i + 1,
            "total": ('=IF(OR($C%d="",$G%d=""),"",'
                      '$F%d*$G%d+$H%d*$G%d*1.5)'
                      % (rownum, rownum, rownum, rownum, rownum, rownum)),
        }
        cached = {"total": ""}
        if i < len(m.staff):
            s = m.staff[i]
            values.update({k: s[k] for k in
                           ("name", "role", "event", "hours", "rate", "ot",
                            "paid", "notes")})
            cached["total"] = s["total"]
        K.write_row(bk, KEY, COLUMNS, rownum, values, cached)

    ws.autofilter(r(C.ROW_HEADER), ci("B"), r(C.last_row(KEY)), ci(LAST_COL))
    ws.freeze_panes(r(C.ROW_FIRST), ci("C"))

    # ------------------------------------------------------------------
    # validation
    # ------------------------------------------------------------------
    K.list_dv(bk, KEY, "role", "staff_roles", title="Role",
              message="Head Chef, Server, Driver\u2026 editable on "
                      "\u2699\uFE0F Setup.")
    bk.validate(KEY, C.ROW_FIRST, ci(bk.col(KEY, "event")),
                C.last_row(KEY), ci(bk.col(KEY, "event")),
                "=EventList", title="Event ID",
                message="Which event is this shift for?")
    K.whole_dv(bk, KEY, "hours", minimum=0, maximum=24)
    K.whole_dv(bk, KEY, "ot", minimum=0, maximum=24)
    K.money_dv(bk, KEY, ("rate",))
    K.tick_dv(bk, KEY, ("paid",))

    # ------------------------------------------------------------------
    # conditional formatting
    # ------------------------------------------------------------------
    K.tick_cf(bk, KEY, ("paid",))
    # unpaid highlight once a total exists
    bk.cond(KEY, C.ROW_FIRST, ci(bk.col(KEY, "total")),
            C.last_row(KEY), ci(bk.col(KEY, "total")), {
        "type": "formula",
        "criteria": '=AND($I%d<>"",$J%d<>"%s")' % (C.ROW_FIRST, C.ROW_FIRST,
                                                   C.TICK),
        "format": S.cf(bg=th.bad_soft, fg=th.bad, bold=True)})

    # ------------------------------------------------------------------
    # totals + notes
    # ------------------------------------------------------------------
    trow = C.last_row(KEY) + 2
    K.totals_row(bk, KEY, trow, {
        "hours": ("=SUM($F$%d:$F$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                  "#,##0", sum(s["hours"] for s in m.staff)),
        "ot": ("=SUM($H$%d:$H$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
               "#,##0", sum(s["ot"] for s in m.staff)),
        "total": ("=SUM($I$%d:$I$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                  "#,##0.00", cost_cached),
    }, label="TOTALS  \u2192", label_span=("B", "E"), last_col=LAST_COL)

    K.note_block(
        bk, KEY, trow + 2, "B", "K",
        ["Overtime hours are paid at 1.5x the hourly rate automatically.",
         "Red Total = the shift has not been ticked Paid yet.",
         "When you actually pay wages, also log them on the \U0001F4B8 "
         "Expenses tab (Staff / Labor) so the P&L stays honest."],
        title="\U0001F4A1 How to use this tab")

    nav = trow + 8
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)
    bk.page(KEY, LAST_COL, nav + 1, freeze=(C.ROW_FIRST, 2),
            title_rows=(C.ROW_HEADER, C.ROW_HEADER))
