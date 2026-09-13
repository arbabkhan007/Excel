"""🏪 Craft Fairs - one record per market, with live profit per fair."""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "events"
LAST_COL = "R"

COLUMNS = [
    ("n", "#", "idx", None),
    ("name", "Fair / event", "text", None),
    ("date", "Date", "date", None),
    ("location", "Location", "text", None),
    ("organizer", "Organizer", "text", None),
    ("booth", "Booth fee", "money0", None),
    ("travel", "Travel", "money0", None),
    ("parking", "Parking", "money0", None),
    ("food", "Food / other", "money0", None),
    ("display", "Display", "money0", None),
    ("total", "Total expenses", "calc_money", "bad"),
    ("taken", "Items taken", "qty", None),
    ("sold", "Items sold", "calc_num", "primary_2"),
    ("sales", "Total sales", "calc_money", "primary_2"),
    ("cogs", "Product cost", "calc_money", "primary_2"),
    ("net", "Net profit", "calc_money", "ok"),
    ("margin", "Margin", "calc_pct", "ok"),
    ("best", "Best seller", "calc_wrap", None),
]


def build(bk):
    th = bk.th
    ws = bk.ws(KEY)
    m = bk.demo
    K.sheet_head(bk, KEY, LAST_COL,
                 "\U0001F3EA  Craft Fairs",
                 "Every market you work - sales, costs and best seller pull "
                 "straight from the \U0001F4B0 Sales Log.")
    K.table_frame(bk, KEY, COLUMNS, height=26)

    sale_ev = bk.rng("sales", "event")
    sale_qty = bk.rng("sales", "qty")
    sale_tot = bk.rng("sales", "total")
    sale_cost = bk.rng("sales", "cost")
    sale_prod = bk.rng("sales", "product")
    sale_pair = bk.rng("sales", "pair")
    for i in range(C.CAP[KEY]):
        row = C.ROW_FIRST + i
        e = m.events[i] if m and i < len(m.events) else None
        values = {"n": i + 1}
        if e:
            values.update({"name": e["name"], "date": e["date"],
                           "location": e["location"],
                           "organizer": e["organizer"], "booth": e["booth"],
                           "travel": e["travel"], "parking": e["parking"],
                           "food": e["food"], "display": e["display"],
                           "taken": e["taken"]})
        values["total"] = '=IF($B%d="","",SUM($F%d:$J%d))' % (row, row, row)
        values["sold"] = '=IF($B%d="","",SUMIFS(%s,%s,$B%d))' % (
            row, sale_qty, sale_ev, row)
        values["sales"] = '=IF($B%d="","",SUMIFS(%s,%s,$B%d))' % (
            row, sale_tot, sale_ev, row)
        values["cogs"] = '=IF($B%d="","",SUMIFS(%s,%s,$B%d))' % (
            row, sale_cost, sale_ev, row)
        values["net"] = '=IF($N%d="","",$N%d-$K%d-$O%d)' % (row, row, row,
                                                            row)
        values["margin"] = '=IF($N%d="","",IF($N%d=0,"",$P%d/$N%d))' % (
            row, row, row, row)
        values["best"] = (
            '=IF($B%d="","",IFERROR(INDEX(%s,MATCH(1,INDEX((%s=$B%d)*'
            '(%s=SUMPRODUCT(MAX((%s=$B%d)*%s))),0),0)),""))'
            % (row, sale_prod, sale_ev, row, sale_pair, sale_ev, row,
               sale_pair))
        cached = {}
        if e:
            cached = {"total": e["total"], "sold": e["sold"],
                      "sales": e["sales"], "cogs": e["cogs"],
                      "net": e["net"], "margin": e["margin"],
                      "best": e["best"]}
        K.write_row(bk, KEY, COLUMNS, row, values, cached)

    K.date_dv(bk, KEY, ("date",))
    K.money_dv(bk, KEY, ("booth", "travel", "parking", "food", "display"))
    K.databar(bk, KEY, "net", color=th.ok)
    ws.conditional_format(
        r(C.ROW_FIRST), ci(bk.col(KEY, "net")),
        r(C.last_row(KEY)), ci(bk.col(KEY, "net")),
        {"type": "cell", "criteria": "<", "value": 0,
         "format": bk.S.cf(bg=th.bad_soft, fg=th.bad, bold=True)})
    bk.stats["cond_formats"] += 1

    tot = C.last_row(KEY) + 1
    K.totals_row(bk, KEY, tot, {
        "total": ("=SUM(K%d:K%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                  "#,##0.00", m.agg.get("fees", 0) if m else 0),
        "sales": ("=SUM(N%d:N%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                  "#,##0.00", m.agg.get("revenue", 0) if m else 0),
        "net": ("=SUM(P%d:P%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                "#,##0.00", sum(e["net"] for e in m.events) if m else 0)},
        first_col="B", last_col="R")

    done = [e for e in (m.events if m else []) if e["sales"]]
    chips = [
        ('="\U0001F3EA Fairs logged: "&COUNTA(%s)' % bk.rng(KEY, "name"),
         "primary", "Fairs logged: %d" % (len(m.events) if m else 0), 3),
        ('="\u2705 Worked: "&%s' % bk.kpi("events_done"), "ok",
         "Worked: %d" % (m.agg.get("events_done", 0) if m else 0), 3),
        ('="\U0001F4B0 Takings: "&Currency&TEXT(SUM(%s),"#,##0")'
         % bk.rng(KEY, "sales"), "info",
         "Takings: %s" % (m.money(m.agg.get("revenue", 0)) if m else "$0"),
         3),
        ('="\U0001F3C6 Best fair: "&%s' % bk.kpi("best_event"), "gold",
         "Best fair: %s" % (m.agg.get("best_event", "") if m else ""), 5),
    ]
    K.chips(bk, KEY, chips, LAST_COL)
    K.footer_nav(bk, KEY, tot + 2, LAST_COL, tip=
                 "  \U0001F4A1  Wondering \u0022should I do this market "
                 "again?\u0022  Sort by Net profit, or check the ranking on "
                 "the \U0001F4CA Dashboard.")
