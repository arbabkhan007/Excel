"""📦 Made & Stocked - production batches and live inventory per batch."""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "production"
LAST_COL = "L"

COLUMNS = [
    ("n", "#", "idx", None),
    ("product", "Product", "text", None),
    ("date", "Date made", "date", None),
    ("made", "Made", "qty", None),
    ("reserved", "Reserved", "qty", None),
    ("taken", "Taken to fair", "qty", None),
    ("sold", "Sold", "qty", None),
    ("returned", "Returned", "qty", None),
    ("damaged", "Damaged", "qty", None),
    ("current", "Current stock", "calc_num", "primary_2"),
    ("available", "Available", "calc_num", "primary_2"),
    ("notes", "Notes", "wrap", None),
]


def build(bk):
    th = bk.th
    ws = bk.ws(KEY)
    m = bk.demo
    K.sheet_head(bk, KEY, LAST_COL,
                 "\U0001F4E6  Made & Stocked",
                 "One row per batch: Available = Made + Returned \u2212 Sold "
                 "\u2212 Damaged \u2212 Reserved, automatically.")
    K.table_frame(bk, KEY, COLUMNS, height=22)

    for i in range(C.CAP[KEY]):
        row = C.ROW_FIRST + i
        b = m.production[i] if m and i < len(m.production) else None
        values = {"n": i + 1}
        if b:
            values.update({"product": b["product"], "date": b["date"],
                           "made": b["made"], "reserved": b["reserved"],
                           "taken": b["taken"], "sold": b["sold"],
                           "returned": b["returned"],
                           "damaged": b["damaged"], "notes": b["notes"]})
        values["current"] = ('=IF($B%d="","",$D%d+$H%d-$G%d-$I%d)'
                             % (row, row, row, row, row))
        values["available"] = '=IF($B%d="","",$J%d-$E%d)' % (row, row, row)
        cached = {"current": b["current"], "available": b["available"]} \
            if b else {}
        K.write_row(bk, KEY, COLUMNS, row, values, cached)

    K.list_dv(bk, KEY, "product", "products", title="Product")
    K.date_dv(bk, KEY, ("date",))
    K.whole_dv(bk, KEY, ("made", "reserved", "taken", "sold", "returned",
                         "damaged"), minimum=0, maximum=100000)
    ws.conditional_format(
        r(C.ROW_FIRST), ci(bk.col(KEY, "available")),
        r(C.last_row(KEY)), ci(bk.col(KEY, "available")),
        {"type": "cell", "criteria": "<=", "value": 0,
         "format": bk.S.cf(bg=th.warn_soft, fg=th.warn, bold=True)})
    bk.stats["cond_formats"] += 1

    tot = C.last_row(KEY) + 1
    K.totals_row(bk, KEY, tot, {
        "made": ("=SUM(D%d:D%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                 "#,##0", m.agg.get("made_total", 0) if m else 0),
        "sold": ("=SUM(G%d:G%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                 "#,##0", m.agg.get("units", 0) if m else 0),
        "current": ("=SUM(J%d:J%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                    "#,##0", sum(b["current"] for b in m.production)
                    if m else 0),
        "available": ("=SUM(K%d:K%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                      "#,##0", sum(b["available"] for b in m.production)
                      if m else 0)},
        first_col="B", last_col="L")

    chips = [
        ('="\U0001F9F6 Batches: "&COUNTA(%s)' % bk.rng(KEY, "product"),
         "primary", "Batches: %d" % (len(m.production) if m else 0), 3),
        ('="\U0001F4E6 Made: "&%s' % bk.kpi("made_total"), "info",
         "Made: %d" % (m.agg.get("made_total", 0) if m else 0), 3),
        ('="\U0001F4B0 Sold: "&%s' % bk.kpi("units"), "ok",
         "Sold: %d" % (m.agg.get("units", 0) if m else 0), 3),
        ('="\U0001F4C8 Sell-through: "&TEXT(%s,"0%%")'
         % bk.kpi("sell_through"), "gold",
         "Sell-through: %.0f%%" % (100 * (m.agg.get("sell_through", 0)
                                          if m else 0)), 3),
    ]
    K.chips(bk, KEY, chips, LAST_COL)
    K.footer_nav(bk, KEY, tot + 2, LAST_COL)
