"""
\U0001F69A Vendor / Supplier Database.

A simple, searchable rolodex: who supplies what, at what price, with what
lead time and terms.  The Supplier dropdowns elsewhere read this list.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "suppliers"
LAST_COL = "M"

COLUMNS = [
    ("n", "#", "idx", None),
    ("supplier", "Supplier", "text", None),
    ("contact", "Contact person", "text", None),
    ("phone", "Phone", "center", None),
    ("email", "Email", "text", None),
    ("category", "Supplies", "center", None),
    ("item", "Key items", "text", None),
    ("price", "Typical unit price", "money", None),
    ("min_order", "Min order", "money0", None),
    ("delivery", "Delivery time", "center", None),
    ("terms", "Payment terms", "center", None),
    ("notes", "Notes", "text", None),
]


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY)
    bk.paint(KEY, 0, 0, C.last_row(KEY) + 22, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F69A  Supplier Database",
        "  Your wholesale rolodex.  Names typed here feed the Supplier "
        "dropdowns on the Inventory and Shopping tabs.",
        LAST_COL)

    # ------------------------------------------------------------------
    # quick-stat strip
    # ------------------------------------------------------------------
    sup_c = bk.rng(KEY, "supplier")
    chips = [
        ('\U0001F69A Suppliers: "&COUNTIF(%s,"?*")' % sup_c,
         "primary", "Suppliers: %d" % len(m.suppliers)),
        ('\U0001F4E6 Ingredients tracked: "&COUNTA(%s)'
         % bk.rng("inventory", "ingredient"),
         "primary_2", "Ingredients tracked: %d" % len(m.inventory)),
        ('\U0001F534 Below minimum: "&%s&" item(s)"' % bk.kpi("low_stock"),
         "bad", "Below minimum: %d item(s)" % m.agg.get("low_stock", 0)),
    ]
    col = 1
    for formula, color, cached in chips:
        fmt = S.pill(th.soft(color), getattr(th, color), size=10.5,
                     bold=True, align="left")
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), col + 2, "", fmt)
        ws.write_formula(r(C.ROW_STATS), col, '="' + formula, fmt, cached)
        bk.stats["formulas"] += 1
        col += 3
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
        values = {"n": i + 1}
        if i < len(m.suppliers):
            values.update(m.suppliers[i])
        K.write_row(bk, KEY, COLUMNS, rownum, values)

    ws.autofilter(r(C.ROW_HEADER), ci("B"), r(C.last_row(KEY)), ci(LAST_COL))
    ws.freeze_panes(r(C.ROW_FIRST), ci("C"))

    # ------------------------------------------------------------------
    # validation
    # ------------------------------------------------------------------
    K.list_dv(bk, KEY, "category", "ingredient_categories",
              title="Supplies what?",
              message="Broad category of what this supplier sells.")
    K.money_dv(bk, KEY, ("price", "min_order"))

    # ------------------------------------------------------------------
    # conditional formatting
    # ------------------------------------------------------------------
    K.databar(bk, KEY, "price", color=th.info)

    # ------------------------------------------------------------------
    # notes
    # ------------------------------------------------------------------
    trow = C.last_row(KEY) + 2
    K.note_block(
        bk, KEY, trow, "B", "M",
        ["Keep at least two suppliers for your critical ingredients - "
         "wedding weekends do not forgive stock-outs.",
         "\u201CDelivery time\u201D is the lead time from order to door: "
         "plan shopping runs around it.",
         "Terms like Net 15 mean you have 15 days to pay - useful for "
         "cash-flow timing before big events."],
        title="\U0001F4A1 How to use this tab")

    nav = trow + 6
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)
    bk.page(KEY, LAST_COL, nav + 1, freeze=(C.ROW_FIRST, 2),
            title_rows=(C.ROW_HEADER, C.ROW_HEADER))
