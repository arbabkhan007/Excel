"""
\U0001F4C5 Event / Order Tracker - the pipeline.

💬 Inquiry → 📝 Quote Sent → 💰 Deposit Paid → ✅ Confirmed → 🎉 Completed
(❌ Cancelled).  White cells are inputs; Profit, Margin and Balance are
formulas and locked.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "events"
LAST_COL = "R"

COLUMNS = [
    ("n", "#", "idx", None),
    ("id", "Event ID", "text", None),
    ("client", "Client", "text", None),
    ("date", "Event date", "date", None),
    ("type", "Type", "center", None),
    ("guests", "Guests", "qty", None),
    ("menu", "Menu / package", "wrap", None),
    ("staff_req", "Staff req.", "qty", None),
    ("equipment_req", "Equipment req.", "wrap", None),
    ("cost", "Total cost", "money", None),
    ("price", "Price quoted", "money", None),
    ("profit", "Profit", "calc_money", "primary_2"),
    ("margin", "Margin", "calc_pct", "primary_2"),
    ("deposit", "Deposit held", "money", None),
    ("balance", "Balance", "calc_money", "primary_2"),
    ("status", "Pipeline status", "center", None),
    ("notes", "Notes", "text", None),
]

STATUS_COLORS = {
    C.ES_INQUIRY: ("info_soft", "info"),
    C.ES_QUOTED: ("plum_soft", "plum"),
    C.ES_DEPOSIT: ("warn_soft", "warn"),
    C.ES_CONFIRMED: ("primary_soft", "primary"),
    C.ES_DONE: ("ok_soft", "ok"),
    C.ES_CANCEL: ("bad_soft", "bad"),
}


def _formulas(row):
    return {
        "n": row - C.ROW_FIRST + 1,
        "profit": '=IF(OR($K%d="",$L%d=""),"",$L%d-$K%d)' % (row, row, row, row),
        "margin": '=IF(OR($L%d="",$L%d=0),"",$M%d/$L%d)' % (row, row, row, row),
        "balance": '=IF($L%d="","",$L%d-$O%d)' % (row, row, row),
    }


def _cached(m, i):
    if i < len(m.events):
        e = m.events[i]
        return {"profit": e["profit"], "margin": e["margin"],
                "balance": e["balance"]}
    return {}


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY)
    bk.paint(KEY, 0, 0, C.last_row(KEY) + 26, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F4C5  Event & Order Tracker",
        "  The pipeline: \U0001F4AC Inquiry \u2192 \U0001F4DD Quote Sent "
        "\u2192 \U0001F4B0 Deposit Paid \u2192 \u2705 Confirmed \u2192 "
        "\U0001F389 Completed.  White cells are yours; tinted cells "
        "calculate.",
        LAST_COL)

    # ------------------------------------------------------------------
    # quick-stat strip
    # ------------------------------------------------------------------
    price = bk.rng(KEY, "price")
    profit = bk.rng(KEY, "profit")
    status = bk.rng(KEY, "status")
    not_cancel = '"<>%s"' % C.ES_CANCEL
    booked = sum(e["price"] for e in m.events if e["status"] != C.ES_CANCEL)
    projected = sum(e["profit"] for e in m.events
                    if e["status"] != C.ES_CANCEL)
    chips = [
        ('\U0001F4C5 Events: "&%s' % bk.kpi("events_total"),
         "primary", "Events: %d" % m.agg.get("events_total", 0)),
        ('\u2705 Confirmed + deposit: "&%s' % bk.kpi("events_confirmed"),
         "ok", "Confirmed + deposit: %d" % m.agg.get("events_confirmed", 0)),
        ('\U0001F389 Completed: "&%s&"  \u274C Cancelled: "&%s'
         % (bk.kpi("events_done"), bk.kpi("events_cancelled")),
         "gold", "Completed: %d  Cancelled: %d"
         % (m.agg.get("events_done", 0), m.agg.get("events_cancelled", 0))),
        ('\U0001F4B0 Booked value: "&Currency&TEXT(SUMIF(%s,%s,%s),"#,##0")'
         % (status, not_cancel, price),
         "primary_2", "Booked value: %s" % m.money(booked)),
        ('\U0001F4C8 Projected profit: "&Currency&TEXT(SUMIF(%s,%s,%s),"#,##0")'
         % (status, not_cancel, profit),
         "accent", "Projected profit: %s" % m.money(projected)),
    ]
    col = 1
    for formula, color, cached in chips:
        fmt = S.pill(th.soft(color), getattr(th, color), size=10.5,
                     bold=True, align="left")
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), col + 2, "", fmt)
        ws.write_formula(r(C.ROW_STATS), col, '="' + formula, fmt, cached)
        bk.stats["formulas"] += 1
        col += 3
    if col <= ci(LAST_COL):
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), ci(LAST_COL),
                       "", S.canvas)

    # ------------------------------------------------------------------
    # table
    # ------------------------------------------------------------------
    K.table_frame(bk, KEY, COLUMNS, height=34)
    for i in range(K.n_rows(KEY)):
        rownum = C.ROW_FIRST + i
        values = _formulas(rownum)
        cached = {"profit": "", "margin": "", "balance": ""}
        if i < len(m.events):
            e = m.events[i]
            values.update({k: e[k] for k in
                           ("id", "client", "date", "type", "guests", "menu",
                            "staff_req", "equipment_req", "cost", "price",
                            "deposit", "status", "notes")})
            cached = _cached(m, i)
        K.write_row(bk, KEY, COLUMNS, rownum, values, cached)

    ws.autofilter(r(C.ROW_HEADER), ci("B"), r(C.last_row(KEY)), ci(LAST_COL))
    ws.freeze_panes(r(C.ROW_FIRST), ci("D"))

    # ------------------------------------------------------------------
    # validation
    # ------------------------------------------------------------------
    bk.validate(KEY, C.ROW_FIRST, ci(bk.col(KEY, "client")),
                C.last_row(KEY), ci(bk.col(KEY, "client")),
                "=ClientsList", title="Client",
                message="Pick a client from the \U0001F465 Clients tab, or "
                        "type a new name.")
    K.date_dv(bk, KEY, ("date",),
              message="When is the event? Type a date or pick from the "
                      "calendar.")
    K.list_dv(bk, KEY, "type", "event_types", title="Event type",
              message="Wedding, corporate, birthday\u2026 editable on "
                      "\u2699\uFE0F Setup.")
    K.whole_dv(bk, KEY, "guests", minimum=0, maximum=100000)
    K.whole_dv(bk, KEY, "staff_req", minimum=0, maximum=500)
    K.money_dv(bk, KEY, ("cost", "price", "deposit"))
    K.fixed_dv(bk, KEY, "status", "EventStatuses", title="Pipeline status",
               message="\U0001F4AC Inquiry \u2192 \U0001F4DD Quote Sent "
                       "\u2192 \U0001F4B0 Deposit Paid \u2192 \u2705 "
                       "Confirmed \u2192 \U0001F389 Completed.",
               error="Pick one of the six pipeline statuses \u2014 the "
                     "colours and the dashboard depend on them.")

    # ------------------------------------------------------------------
    # conditional formatting
    # ------------------------------------------------------------------
    K.status_cf(bk, KEY, "status",
                {text: (getattr(th, bg), getattr(th, fg))
                 for text, (bg, fg) in STATUS_COLORS.items()})
    K.deadline_cf(bk, KEY, "date")
    L = bk.col(KEY, "price")
    bk.cond(KEY, C.ROW_FIRST, ci(L), C.last_row(KEY), ci(L), {
        "type": "formula", "criteria": '=$%s%d<>""' % (bk.col(KEY, "id"),
                                                       C.ROW_FIRST),
        "format": S.cf(bg=th.gold_soft, fg=th.ink, bold=True)})
    # cancelled rows: strike the client + price
    for field in ("client", "price"):
        col = bk.col(KEY, field)
        bk.cond(KEY, C.ROW_FIRST, ci(col), C.last_row(KEY), ci(col), {
            "type": "formula",
            "criteria": '=$%s%d="%s"' % (bk.col(KEY, "status"), C.ROW_FIRST,
                                         C.ES_CANCEL),
            "format": S.cf(strike=True, fg=th.muted)})

    # ------------------------------------------------------------------
    # totals
    # ------------------------------------------------------------------
    trow = C.last_row(KEY) + 2
    K.totals_row(bk, KEY, trow, {
        "guests": ("=SUM($G$%d:$G$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                   "#,##0", sum(e["guests"] for e in m.events)),
        "cost": ("=SUM($K$%d:$K$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                 "#,##0.00", sum(e["cost"] for e in m.events)),
        "price": ("=SUM($L$%d:$L$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                  "#,##0.00", sum(e["price"] for e in m.events)),
        "profit": ("=SUM($M$%d:$M$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                   "#,##0.00", sum(e["profit"] for e in m.events)),
        "deposit": ("=SUM($O$%d:$O$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                    "#,##0.00", sum(e["deposit"] for e in m.events)),
    }, label="TOTALS  (all events incl. cancelled)  \u2192",
        label_span=("B", "F"), last_col=LAST_COL)

    K.note_block(
        bk, KEY, trow + 2, "B", "R",
        ["Fill the white cells - the tinted Profit / Margin / Balance "
         "columns do the maths.",
         "Need a price? Build it on the \U0001F9EE Quote Calculator first, "
         "then copy the recommended price into 'Price quoted'.",
         "Cancelled events stay in the list (greyed + struck through) so "
         "your history is complete - they are excluded from every total "
         "and average."],
        title="\U0001F4A1 How to use this tab")

    nav = trow + 8
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)
    bk.page(KEY, LAST_COL, nav + 1, freeze=(C.ROW_FIRST, 3),
            title_rows=(C.ROW_HEADER, C.ROW_HEADER))
