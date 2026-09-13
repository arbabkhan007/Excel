"""
\U0001F465 Client Tracker - the address book + latest booking per client.

One row per client with contact details and the money picture of their
current/next event (quote, deposit, balance, payment status).
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "clients"
LAST_COL = "O"

COLUMNS = [
    ("n", "#", "idx", None),
    ("name", "Client name", "text", None),
    ("phone", "Phone", "center", None),
    ("email", "Email", "text", None),
    ("event_date", "Next / last event", "date", None),
    ("event_type", "Event type", "center", None),
    ("guests", "Guests", "qty", None),
    ("venue", "Venue", "text", None),
    ("package", "Package", "center", None),
    ("quote", "Quote value", "money", None),
    ("deposit", "Deposit held", "money", None),
    ("balance", "Balance", "calc_money", "primary_2"),
    ("status", "Payment status", "center", None),
    ("notes", "Notes", "text", None),
]

PACKAGES = "Custom,Bronze,Silver,Gold,Platinum"


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY)
    bk.paint(KEY, 0, 0, C.last_row(KEY) + 24, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F465  Client Tracker",
        "  One row per client: contacts, their next (or latest) event and "
        "the money picture.  Repeat clients are gold for your business "
        "\u2014 the dashboard counts them.",
        LAST_COL)

    # ------------------------------------------------------------------
    # quick-stat strip
    # ------------------------------------------------------------------
    name_rng = bk.rng(KEY, "name")
    quote_rng = bk.rng(KEY, "quote")
    dep_rng = bk.rng(KEY, "deposit")
    ev_client = bk.rng("events", "client")
    chips = [
        ('\U0001F465 Clients: "&COUNTIF(%s,"?*")' % name_rng,
         "primary", "Clients: %d" % len(m.clients)),
        ('\U0001F501 Repeat customers: "&TEXT(%s,"0%%")'
         % bk.kpi("repeat_pct"), "gold",
         "Repeat customers: %d%%" % round(m.agg.get("repeat_pct", 0) * 100)),
        ('\U0001F4B0 Total quoted: "&Currency&TEXT(SUM(%s),"#,##0")'
         % quote_rng, "primary_2",
         "Total quoted: %s" % m.money(sum(c["quote"] for c in m.clients))),
        ('\U0001F4B5 Deposits held: "&Currency&TEXT(SUM(%s),"#,##0")'
         % dep_rng, "ok",
         "Deposits held: %s" % m.money(sum(c["deposit"] for c in m.clients))),
    ]
    col = 1
    for formula, color, cached in chips:
        fmt = S.pill(th.soft(color), getattr(th, color), size=10.5,
                     bold=True, align="left")
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), col + 2, "", fmt)
        ws.write_formula(r(C.ROW_STATS), col, '="' + formula, fmt, cached)
        bk.stats["formulas"] += 1
        col += 3
    ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), ci(LAST_COL),
                   "", S.canvas)

    # ------------------------------------------------------------------
    # table
    # ------------------------------------------------------------------
    K.table_frame(bk, KEY, COLUMNS, height=22)
    for i in range(K.n_rows(KEY)):
        rownum = C.ROW_FIRST + i
        values = {
            "n": i + 1,
            "balance": '=IF($K%d="","",$K%d-$L%d)' % (rownum, rownum, rownum),
        }
        cached = {"balance": ""}
        if i < len(m.clients):
            cl = m.clients[i]
            values.update({k: cl[k] for k in
                           ("name", "phone", "email", "event_date",
                            "event_type", "guests", "venue", "package",
                            "quote", "deposit", "status", "notes")})
            cached["balance"] = cl["quote"] - cl["deposit"]
        K.write_row(bk, KEY, COLUMNS, rownum, values, cached)

    ws.autofilter(r(C.ROW_HEADER), ci("B"), r(C.last_row(KEY)), ci(LAST_COL))
    ws.freeze_panes(r(C.ROW_FIRST), ci("D"))

    # ------------------------------------------------------------------
    # validation
    # ------------------------------------------------------------------
    K.date_dv(bk, KEY, ("event_date",))
    K.list_dv(bk, KEY, "event_type", "event_types", title="Event type")
    K.whole_dv(bk, KEY, "guests", minimum=0, maximum=100000)
    bk.validate(KEY, C.ROW_FIRST, ci(bk.col(KEY, "package")),
                C.last_row(KEY), ci(bk.col(KEY, "package")),
                '"%s"' % PACKAGES, title="Package",
                message="Your service tiers - edit the list on \u2699\uFE0F "
                        "Setup if you named them differently.")
    K.money_dv(bk, KEY, ("quote", "deposit"))
    K.fixed_dv(bk, KEY, "status", "PaymentStatuses", title="Payment status",
               message="\U0001F534 Unpaid \u2192 \U0001F7E0 Partial \u2192 "
                       "\U0001F7E2 Paid in full.",
               error="Pick one of the three payment statuses.")

    # ------------------------------------------------------------------
    # conditional formatting
    # ------------------------------------------------------------------
    K.status_cf(bk, KEY, "status", {
        C.PS_UNPAID: (th.bad_soft, th.bad),
        C.PS_PART: (th.warn_soft, th.warn),
        C.PS_PAID: (th.ok_soft, th.ok),
    })
    # repeat customers get a golden name
    bk.cond(KEY, C.ROW_FIRST, ci("C"), C.last_row(KEY), ci("C"), {
        "type": "formula",
        "criteria": '=AND($C%d<>"",COUNTIF(%s,$C%d)>1)'
                   % (C.ROW_FIRST, ev_client, C.ROW_FIRST),
        "format": S.cf(bg=th.gold_soft, fg=th.gold, bold=True)})
    K.deadline_cf(bk, KEY, "event_date")

    # ------------------------------------------------------------------
    # totals + notes
    # ------------------------------------------------------------------
    trow = C.last_row(KEY) + 2
    K.totals_row(bk, KEY, trow, {
        "guests": ("=SUM($H$%d:$H$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                   "#,##0", sum(c["guests"] for c in m.clients)),
        "quote": ("=SUM($K$%d:$K$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                  "#,##0.00", sum(c["quote"] for c in m.clients)),
        "deposit": ("=SUM($L$%d:$L$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                    "#,##0.00", sum(c["deposit"] for c in m.clients)),
        "balance": ("=SUM($M$%d:$M$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                    "#,##0.00",
                    sum(c["quote"] - c["deposit"] for c in m.clients)),
    }, label="TOTALS  \u2192", label_span=("B", "G"), last_col=LAST_COL)

    K.note_block(
        bk, KEY, trow + 2, "B", "O",
        ["Names typed here feed the Client dropdown on the Events, "
         "Payments and Invoice tabs.",
         "A golden name means the client appears on more than one event "
         "- repeat business!",
         "Quote / Deposit are inputs; Balance calculates itself. For the "
         "full payment schedule use the \U0001F4B0 Payments tab."],
        title="\U0001F4A1 How to use this tab")

    nav = trow + 8
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)
    bk.page(KEY, LAST_COL, nav + 1, freeze=(C.ROW_FIRST, 3),
            title_rows=(C.ROW_HEADER, C.ROW_HEADER))
