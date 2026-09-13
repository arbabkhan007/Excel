"""
\U0001F373 Equipment Tracker - what you own, what's out, what's broken.

Available = owned - reserved - damaged.  Status flags anything damaged or
with maintenance due inside 30 days.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "equipment"
LAST_COL = "L"

COLUMNS = [
    ("n", "#", "idx", None),
    ("item", "Equipment", "text", None),
    ("category", "Category", "center", None),
    ("owned", "Owned", "qty", None),
    ("reserved", "On events", "qty", None),
    ("damaged", "Damaged", "qty", None),
    ("available", "Available", "calc_num", "primary_2"),
    ("unit_value", "Unit value", "money", None),
    ("value", "Total value", "calc_money", "primary_2"),
    ("maintenance", "Next service", "date", None),
    ("status", "Status", "calc_c", "primary_2"),
]

CATEGORIES = "Serving,Furniture,Linens,Tableware,Beverage,Kitchen,Transport,Other"


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY)
    bk.paint(KEY, 0, 0, C.last_row(KEY) + 24, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F373  Equipment Tracker",
        "  Know exactly what you own, what is committed to events, what is "
        "damaged and what needs servicing \u2014 before you promise it to "
        "a client.",
        LAST_COL)

    # ------------------------------------------------------------------
    # quick-stat strip
    # ------------------------------------------------------------------
    item_c = bk.rng(KEY, "item")
    damaged_c = bk.rng(KEY, "damaged")
    chips = [
        ('\U0001F373 Equipment lines: "&COUNTIF(%s,"?*")' % item_c,
         "primary", "Equipment lines: %d" % len(m.equipment)),
        ('\U0001F4B0 Fleet value: "&Currency&TEXT(%s,"#,##0")'
         % bk.kpi("equip_value"), "primary_2",
         "Fleet value: %s" % m.money(m.agg.get("equip_value", 0))),
        ('\U0001F527 Needs attention: "&%s' % bk.kpi("equip_service"),
         "warn", "Needs attention: %d" % m.agg.get("equip_service", 0)),
        ('\U0001F4A5 Damaged units: "&SUM(%s)' % damaged_c, "bad",
         "Damaged units: %d" % sum(q["damaged"] for q in m.equipment)),
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
            "available": '=IF($C%d="","",$E%d-$F%d-$G%d)'
                         % (rownum, rownum, rownum, rownum),
            "value": '=IF(OR($C%d="",$I%d=""),"",$E%d*$I%d)'
                     % (rownum, rownum, rownum, rownum),
            "status": ('=IF($C%d="","",IF($G%d>0,"\U0001F527 Service",'
                       'IF(AND($K%d<>"",$K%d<=TODAY()+30),"\U0001F7E1 Soon",'
                       '"\U0001F7E2 Ready")))'
                       % (rownum, rownum, rownum, rownum)),
        }
        cached = {"available": "", "value": "", "status": ""}
        if i < len(m.equipment):
            q = m.equipment[i]
            values.update({k: q[k] for k in
                           ("item", "category", "owned", "reserved",
                            "damaged", "unit_value", "maintenance", "notes")
                           if k in q})
            cached.update({"available": q["available"], "value": q["value"],
                           "status": q["status"]})
        K.write_row(bk, KEY, COLUMNS, rownum, values, cached)

    ws.autofilter(r(C.ROW_HEADER), ci("B"), r(C.last_row(KEY)), ci(LAST_COL))
    ws.freeze_panes(r(C.ROW_FIRST), ci("C"))

    # ------------------------------------------------------------------
    # validation
    # ------------------------------------------------------------------
    bk.validate(KEY, C.ROW_FIRST, ci(bk.col(KEY, "category")),
                C.last_row(KEY), ci(bk.col(KEY, "category")),
                '"%s"' % CATEGORIES, title="Category")
    K.whole_dv(bk, KEY, ("owned", "reserved", "damaged"), minimum=0,
               maximum=10000)
    K.money_dv(bk, KEY, ("unit_value",))
    K.date_dv(bk, KEY, ("maintenance",),
              message="Next service / inspection date.")

    # ------------------------------------------------------------------
    # conditional formatting
    # ------------------------------------------------------------------
    K.status_cf(bk, KEY, "status", {
        "\U0001F527 Service": (th.bad_soft, th.bad),
        "\U0001F7E1 Soon": (th.warn_soft, th.warn),
        "\U0001F7E2 Ready": (th.ok_soft, th.ok),
    })
    K.deadline_cf(bk, KEY, "maintenance")
    # available == 0 → red number
    bk.cond(KEY, C.ROW_FIRST, ci("H"), C.last_row(KEY), ci("H"), {
        "type": "cell", "criteria": "<=", "value": 0,
        "format": S.cf(bg=th.bad_soft, fg=th.bad, bold=True)})

    # ------------------------------------------------------------------
    # totals + notes
    # ------------------------------------------------------------------
    trow = C.last_row(KEY) + 2
    K.totals_row(bk, KEY, trow, {
        "owned": ("=SUM($E$%d:$E$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                  "#,##0", sum(q["owned"] for q in m.equipment)),
        "reserved": ("=SUM($F$%d:$F$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                     "#,##0", sum(q["reserved"] for q in m.equipment)),
        "damaged": ("=SUM($G$%d:$G$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                    "#,##0", sum(q["damaged"] for q in m.equipment)),
        "available": ("=SUM($H$%d:$H$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                      "#,##0", sum(q["available"] for q in m.equipment)),
        "value": ("=SUM($J$%d:$J$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                  "#,##0.00", m.agg.get("equip_value", 0)),
    }, label="TOTALS  \u2192", label_span=("B", "D"), last_col=LAST_COL)

    K.note_block(
        bk, KEY, trow + 2, "B", "L",
        ["\U0001F527 Service = something is damaged. \U0001F7E1 Soon = "
         "maintenance due inside 30 days. \U0001F7E2 Ready = good to go.",
         "\u201COn events\u201D is how many units are committed right now "
         "- update it as bookings come and go.",
         "Fleet value feeds your balance sheet (and your insurance "
         "conversation)."],
        title="\U0001F4A1 How to use this tab")

    nav = trow + 8
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)
    bk.page(KEY, LAST_COL, nav + 1, freeze=(C.ROW_FIRST, 2),
            title_rows=(C.ROW_HEADER, C.ROW_HEADER))
