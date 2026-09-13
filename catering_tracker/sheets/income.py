"""
\U0001F4B0 Income & Payments - one row per invoice.

Inputs: invoice #, client, event date, invoice amount, deposit, payment 1,
payment 2, due date.  Automatic: received, balance, payment status.
Overdue balances flag themselves red - the OVERDUE auto-flag from the spec.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "income"
LAST_COL = "M"

COLUMNS = [
    ("n", "#", "idx", None),
    ("invoice", "Invoice #", "center", None),
    ("client", "Client", "text", None),
    ("event_date", "Event date", "date", None),
    ("amount", "Invoice total", "money", None),
    ("deposit", "Deposit", "money", None),
    ("pay1", "Payment 1", "money", None),
    ("pay2", "Payment 2", "money", None),
    ("received", "Received", "calc_money", "primary_2"),
    ("balance", "Balance due", "calc_money", "primary_2"),
    ("due", "Due date", "date", None),
    ("status", "Status", "calc_c", "primary_2"),
]


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY)
    bk.paint(KEY, 0, 0, C.last_row(KEY) + 24, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F4B0  Income & Payments",
        "  One row per invoice.  Log deposits and payments as they arrive "
        "\u2014 Received, Balance and Status calculate themselves, and "
        "overdue money turns red.",
        LAST_COL)

    # ------------------------------------------------------------------
    # quick-stat strip
    # ------------------------------------------------------------------
    amt = bk.rng(KEY, "amount")
    recv = bk.rng(KEY, "received")
    bal = bk.rng(KEY, "balance")
    due = bk.rng(KEY, "due")
    invoiced_cached = sum(i["amount"] for i in m.income)
    chips = [
        ('\U0001F4B0 Cash received: "&Currency&TEXT(%s,"#,##0")'
         % bk.kpi("revenue"), "ok",
         "Cash received: %s" % m.money(m.agg.get("revenue", 0))),
        ('\U0001F9FE Invoiced: "&Currency&TEXT(SUM(%s),"#,##0")' % amt,
         "primary", "Invoiced: %s" % m.money(invoiced_cached)),
        ('\u23F3 Outstanding: "&Currency&TEXT(SUM(%s),"#,##0")' % bal,
         "warn", "Outstanding: %s" % m.money(m.agg.get("outstanding", 0))),
        ('\U0001F534 Overdue: "&%s&" invoice(s)"' % bk.kpi("overdue"),
         "bad", "Overdue: %d invoice(s)" % m.agg.get("overdue", 0)),
        ('\U0001F4C2 Open invoices: "&%s' % bk.kpi("invoices_open"),
         "info", "Open invoices: %d" % m.agg.get("invoices_open", 0)),
    ]
    col = 1
    for formula, color, cached in chips:
        fmt = S.pill(th.soft(color), getattr(th, color), size=10.5,
                     bold=True, align="left")
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), col + 1, "", fmt)
        ws.write_formula(r(C.ROW_STATS), col, '="' + formula, fmt, cached)
        bk.stats["formulas"] += 1
        col += 2
    if col <= ci(LAST_COL):
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), ci(LAST_COL),
                       "", S.canvas)

    # ------------------------------------------------------------------
    # table
    # ------------------------------------------------------------------
    K.table_frame(bk, KEY, COLUMNS, height=20)
    for i in range(K.n_rows(KEY)):
        rownum = C.ROW_FIRST + i
        values = {
            "n": i + 1,
            "received": '=IF($F%d="","",SUM($G%d:$I%d))'
                        % (rownum, rownum, rownum),
            "balance": '=IF($F%d="","",$F%d-$J%d)' % (rownum, rownum, rownum),
            "status": ('=IF($F%d="","",IF($J%d>=$F%d,"%s",IF($J%d>0,"%s","%s")))'
                       % (rownum, rownum, rownum, C.PS_PAID,
                          rownum, C.PS_PART, C.PS_UNPAID)),
        }
        cached = {"received": "", "balance": "", "status": ""}
        if i < len(m.income):
            inv = m.income[i]
            values.update({k: inv[k] for k in
                           ("invoice", "client", "event_date", "amount",
                            "deposit", "pay1", "pay2", "due")})
            cached.update({"received": inv["received"],
                           "balance": inv["balance"], "status": inv["status"]})
        K.write_row(bk, KEY, COLUMNS, rownum, values, cached)

    ws.autofilter(r(C.ROW_HEADER), ci("B"), r(C.last_row(KEY)), ci(LAST_COL))
    ws.freeze_panes(r(C.ROW_FIRST), ci("D"))

    # ------------------------------------------------------------------
    # validation
    # ------------------------------------------------------------------
    bk.validate(KEY, C.ROW_FIRST, ci(bk.col(KEY, "client")),
                C.last_row(KEY), ci(bk.col(KEY, "client")),
                "=ClientsList", title="Client",
                message="Pick the client (from the \U0001F465 Clients tab).")
    K.date_dv(bk, KEY, ("event_date", "due"))
    K.money_dv(bk, KEY, ("amount", "deposit", "pay1", "pay2"))

    # ------------------------------------------------------------------
    # conditional formatting
    # ------------------------------------------------------------------
    K.status_cf(bk, KEY, "status", {
        C.PS_UNPAID: (th.bad_soft, th.bad),
        C.PS_PART: (th.warn_soft, th.warn),
        C.PS_PAID: (th.ok_soft, th.ok),
    })
    # OVERDUE auto-flag: balance > 0 and due date passed
    L = bk.col(KEY, "due")
    Kb = bk.col(KEY, "balance")
    for field in ("due", "balance", "status"):
        colL = bk.col(KEY, field)
        bk.cond(KEY, C.ROW_FIRST, ci(colL), C.last_row(KEY), ci(colL), {
            "type": "formula",
            "criteria": '=AND($%s%d<>"",$%s%d>0,$%s%d<TODAY())'
                       % (L, C.ROW_FIRST, Kb, C.ROW_FIRST,
                          L, C.ROW_FIRST),
            "format": S.cf(bg=th.bad_soft, fg=th.bad, bold=True)})
        bk.cond(KEY, C.ROW_FIRST, ci(colL), C.last_row(KEY), ci(colL), {
            "type": "formula",
            "criteria": '=AND($%s%d<>"",$%s%d>0,$%s%d>=TODAY(),'
                       '$%s%d-TODAY()<=DueSoonDays)'
                       % (L, C.ROW_FIRST, Kb, C.ROW_FIRST,
                          L, C.ROW_FIRST, L, C.ROW_FIRST),
            "format": S.cf(bg=th.warn_soft, fg=th.warn, bold=True)})

    # ------------------------------------------------------------------
    # totals + notes
    # ------------------------------------------------------------------
    trow = C.last_row(KEY) + 2
    K.totals_row(bk, KEY, trow, {
        "amount": ("=SUM($F$%d:$F$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                   "#,##0.00", invoiced_cached),
        "deposit": ("=SUM($G$%d:$G$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                    "#,##0.00", sum(i["deposit"] for i in m.income)),
        "pay1": ("=SUM($H$%d:$H$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                 "#,##0.00", sum(i["pay1"] for i in m.income)),
        "pay2": ("=SUM($I$%d:$I$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                 "#,##0.00", sum(i["pay2"] for i in m.income)),
        "received": ("=SUM($J$%d:$J$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                     "#,##0.00", m.agg.get("revenue", 0)),
        "balance": ("=SUM($K$%d:$K$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                    "#,##0.00", m.agg.get("outstanding", 0)),
    }, label="TOTALS  \u2192", label_span=("B", "E"), last_col=LAST_COL)

    K.note_block(
        bk, KEY, trow + 2, "B", "M",
        ["Invoice total, deposit and the two payment columns are the only "
         "money you type - Received, Balance and Status are formulas.",
         "RED row ending = the due date has passed and money is still "
         "owed. Chase it. AMBER = due inside your \u201Cdue soon\u201D "
         "window (Setup).",
         "Need a pretty document? The \U0001F5A8\uFE0F Invoice tab turns "
         "any row into a printable invoice."],
        title="\U0001F4A1 How to use this tab")

    nav = trow + 8
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)
    bk.page(KEY, LAST_COL, nav + 1, freeze=(C.ROW_FIRST, 3),
            title_rows=(C.ROW_HEADER, C.ROW_HEADER))
