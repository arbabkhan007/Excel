"""
\U0001F4B8 Expense Tracker - every rand/dollar/rupee that leaves the till.

Categories are the eleven from the spec (Ingredients ... Miscellaneous),
editable on Setup.  Feeds the P&L, the tax tab and the dashboard pie.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "expenses"
LAST_COL = "J"

COLUMNS = [
    ("n", "#", "idx", None),
    ("date", "Date", "date", None),
    ("vendor", "Vendor / payee", "text", None),
    ("category", "Category", "center", None),
    ("desc", "Description", "text", None),
    ("amount", "Amount", "money", None),
    ("method", "Paid by", "center", None),
    ("event", "Event ID", "center", None),
    ("receipt", "\U0001F9FE Receipt filed", "tick", None),
]


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY)
    bk.paint(KEY, 0, 0, C.last_row(KEY) + 24, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F4B8  Expense Tracker",
        "  Log every business cost as it happens \u2014 the P&L, tax "
        "summary and dashboard pie all read this table.",
        LAST_COL)

    # ------------------------------------------------------------------
    # quick-stat strip
    # ------------------------------------------------------------------
    amt = bk.rng(KEY, "amount")
    cat = bk.rng(KEY, "category")
    dt = bk.rng(KEY, "date")
    this_month = ('SUMIFS(%s,%s,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),'
                  '%s,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1))'
                  % (amt, dt, dt))
    month_cached = 0
    if bk.demo:
        from ..demo import TODAY as _now
        month_cached = sum(x["amount"] for x in m.expenses
                           if x["date"].month == _now.month
                           and x["date"].year == _now.year)
    chips = [
        ('\U0001F4B8 Total expenses: "&Currency&TEXT(%s,"#,##0")'
         % bk.kpi("expenses"), "bad",
         "Total expenses: %s" % m.money(m.agg.get("expenses", 0))),
        ('\U0001F4C5 This month: "&Currency&TEXT(%s,"#,##0")' % this_month,
         "warn", "This month: %s" % m.money(month_cached)),
        ('\U0001F373 Food + packaging: "&Currency&TEXT(%s,"#,##0")'
         % bk.kpi("food_cost"), "primary_2",
         "Food + packaging: %s" % m.money(m.agg.get("food_cost", 0))),
        ('\U0001F477 Staff / labor: "&Currency&TEXT(%s,"#,##0")'
         % bk.kpi("labor_cost"), "info",
         "Staff / labor: %s" % m.money(m.agg.get("labor_cost", 0))),
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
        values = {"n": i + 1}
        if i < len(m.expenses):
            values.update(m.expenses[i])
        K.write_row(bk, KEY, COLUMNS, rownum, values)

    ws.autofilter(r(C.ROW_HEADER), ci("B"), r(C.last_row(KEY)), ci(LAST_COL))
    ws.freeze_panes(r(C.ROW_FIRST), ci("B"))

    # ------------------------------------------------------------------
    # validation
    # ------------------------------------------------------------------
    K.date_dv(bk, KEY, "date")
    K.list_dv(bk, KEY, "category", "expense_categories", title="Category",
              message="The eleven built-in categories - editable on "
                      "\u2699\uFE0F Setup.",
              error="Pick a category (or add your own on Setup first).")
    K.list_dv(bk, KEY, "method", "payment_methods", title="Paid by")
    K.money_dv(bk, KEY, ("amount",))
    bk.validate(KEY, C.ROW_FIRST, ci(bk.col(KEY, "event")),
                C.last_row(KEY), ci(bk.col(KEY, "event")),
                "=EventList", title="Event ID",
                message="Optional: tie this expense to an event.")
    K.tick_dv(bk, KEY, ("receipt",))

    # ------------------------------------------------------------------
    # conditional formatting
    # ------------------------------------------------------------------
    K.tick_cf(bk, KEY, ("receipt",))
    K.databar(bk, KEY, "amount", color=th.accent)
    # missing receipts glow red once the row has an amount
    bk.cond(KEY, C.ROW_FIRST, ci(bk.col(KEY, "receipt")),
            C.last_row(KEY), ci(bk.col(KEY, "receipt")), {
        "type": "formula",
        "criteria": '=AND($G%d<>"",$J%d<>"%s")' % (C.ROW_FIRST, C.ROW_FIRST,
                                                   C.TICK),
        "format": S.cf(bg=th.bad_soft, fg=th.bad, bold=True)})

    # ------------------------------------------------------------------
    # totals + notes
    # ------------------------------------------------------------------
    trow = C.last_row(KEY) + 2
    K.totals_row(bk, KEY, trow, {
        "amount": ("=SUM($G$%d:$G$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                   "#,##0.00", sum(x["amount"] for x in m.expenses)),
    }, label="TOTAL SPENT  \u2192", label_span=("B", "F"), last_col=LAST_COL)

    K.note_block(
        bk, KEY, trow + 2, "B", "J",
        ["Log the expense the day it happens - future-you doing tax season "
         "will say thanks.",
         "The Event ID column is optional, but tagging event purchases "
         "makes each event's true cost visible.",
         "Red tick column = money spent with no receipt filed yet."],
        title="\U0001F4A1 How to use this tab")

    nav = trow + 8
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)
    bk.page(KEY, LAST_COL, nav + 1, freeze=(C.ROW_FIRST, 0),
            title_rows=(C.ROW_HEADER, C.ROW_HEADER))
