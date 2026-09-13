"""📅 Monthly Summary - the year at a glance, month by month."""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "monthly"
LAST_COL = "K"

COLUMNS = [
    ("month", "Month", "center", None),
    ("revenue", "Revenue", "calc_money", "ok"),
    ("cogs", "Product costs", "calc_money", None),
    ("fees", "Fair fees", "calc_money", None),
    ("net", "Net profit", "calc_money", "primary_2"),
    ("margin", "Margin", "calc_pct", "primary_2"),
    ("units", "Units sold", "calc_num", None),
    ("aov", "Avg sale", "calc_money", None),
    ("per_item", "Profit / item", "calc_money", None),
    ("markets", "Markets", "calc_num", None),
]


def build(bk):
    th = bk.th
    ws = bk.ws(KEY)
    m = bk.demo
    K.sheet_head(bk, KEY, LAST_COL,
                 "\U0001F4C5  Monthly Summary",
                 "Revenue, costs, profit and markets for every month of "
                 "the year set on \u2699\uFE0F Lists & Settings.")
    # header + rows
    from .common import header_row, data_rows
    cols = [("n", "#", "idx", None)] + [(f, l, k, c)
                                        for f, l, k, c in
                                        [(c[0], c[1], c[2], c[3])
                                         for c in COLUMNS]]
    header_row(bk, KEY, cols)
    data_rows(bk, KEY, height=20)

    d = bk.q("data")
    months = (m.agg.get("months") or []) if m else []
    for i in range(12):
        row = C.ROW_FIRST + i
        dr = C.DATA_MONTH_FIRST + i
        values = {"n": i + 1, "month": C.MONTH_NAMES[i][:3]}
        for field, col in (("revenue", "I"), ("cogs", "J"), ("fees", "K"),
                           ("net", "L"), ("units", "M"),
                           ("markets", "N")):
            values[field] = "=%s!$%s$%d" % (d, col, dr)
        values["margin"] = '=IF($C%d=0,"",IFERROR($F%d/$C%d,0))' % (
            row, row, row)
        values["aov"] = '=IF(%s!$O$%d=0,"",IFERROR($C%d/%s!$O$%d,0))' % (
            d, dr, row, d, dr)
        values["per_item"] = '=IF($H%d=0,"",IFERROR($F%d/$H%d,0))' % (
            row, row, row)
        cached = {}
        if i < len(months):
            _, rev, cogs, fees, net, units, markets, txns = months[i]
            cached = {"revenue": rev, "cogs": cogs, "fees": fees,
                      "net": net, "units": units, "markets": markets,
                      "margin": (net / rev) if rev else "",
                      "aov": (rev / txns) if txns else "",
                      "txns": txns,
                      "per_item": (net / units) if units else ""}
        K.write_row(bk, KEY, cols, row, values, cached)

    tot = C.ROW_FIRST + 12
    K.totals_row(bk, KEY, tot, {
        "revenue": ("=SUM(C%d:C%d)" % (C.ROW_FIRST, C.ROW_FIRST + 11),
                    "#,##0.00", m.agg.get("revenue", 0) if m else 0),
        "cogs": ("=SUM(D%d:D%d)" % (C.ROW_FIRST, C.ROW_FIRST + 11),
                 "#,##0.00", m.agg.get("cogs", 0) if m else 0),
        "fees": ("=SUM(E%d:E%d)" % (C.ROW_FIRST, C.ROW_FIRST + 11),
                 "#,##0.00", m.agg.get("fees", 0) if m else 0),
        "net": ("=SUM(F%d:F%d)" % (C.ROW_FIRST, C.ROW_FIRST + 11),
                "#,##0.00", m.agg.get("profit", 0) if m else 0),
        "units": ("=SUM(H%d:H%d)" % (C.ROW_FIRST, C.ROW_FIRST + 11),
                  "#,##0", m.agg.get("units", 0) if m else 0),
        "markets": ("=SUM(K%d:K%d)" % (C.ROW_FIRST, C.ROW_FIRST + 11),
                    "#,##0", m.agg.get("events_done", 0) if m else 0)},
        first_col="B", last_col=LAST_COL)

    ws.conditional_format(
        r(C.ROW_FIRST), ci(bk.col(KEY, "net")),
        r(C.ROW_FIRST + 11), ci(bk.col(KEY, "net")),
        {"type": "cell", "criteria": "<", "value": 0,
         "format": bk.S.cf(bg=th.bad_soft, fg=th.bad, bold=True)})
    bk.stats["cond_formats"] += 1

    best = max(months, key=lambda x: x[4]) if months else None
    chips = [
        ('="\U0001F4C5 Year: "&TEXT(ReportYear,"0")', "primary",
         "Year: %d" % (m.settings["year"] if m else 2026), 2),
        ('="\U0001F4B0 Best month: "&INDEX($B$%d:$B$%d,MATCH(MAX($F$%d:'
         '$F$%d),$F$%d:$F$%d,0))'
         % (C.ROW_FIRST, C.ROW_FIRST + 11, C.ROW_FIRST, C.ROW_FIRST + 11,
            C.ROW_FIRST, C.ROW_FIRST + 11), "ok",
         "Best month: %s" % (C.MONTH_NAMES[best[0] - 1][:3]
                             if best else ""), 3),
        ('="\U0001F3EA Markets worked: "&%s' % bk.kpi("events_done"),
         "info", "Markets worked: %d" % (m.agg.get("events_done", 0)
                                         if m else 0), 3),
        ('="\U0001F4C8 Year margin: "&TEXT(%s,"0.0%%")' % bk.kpi("margin"),
         "gold", "Year margin: %.1f%%" % (100 * (m.agg.get("margin", 0)
                                                 if m else 0)), 3),
    ]
    K.chips(bk, KEY, chips, LAST_COL)
    K.footer_nav(bk, KEY, tot + 2, LAST_COL)
