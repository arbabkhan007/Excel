"""💰 Sales Log - every sale, with event & product dropdowns and live cost."""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "sales"
LAST_COL = "M"

COLUMNS = [
    ("n", "#", "idx", None),
    ("date", "Date", "date", None),
    ("event", "Fair / event", "text", None),
    ("product", "Product", "text", None),
    ("qty", "Qty", "qty", None),
    ("unit", "Unit price", "money", None),
    ("discount", "Discount", "money", None),
    ("total", "Sale total", "calc_money", "ok"),
    ("method", "Payment method", "center", None),
    ("ref", "Txn / ref #", "text", None),
    ("notes", "Notes", "wrap", None),
    ("pair", "Pair units", "calc_num", "primary_2"),
    ("cost", "Cost of sale", "calc_money", "primary_2"),
]


def build(bk):
    th = bk.th
    ws = bk.ws(KEY)
    m = bk.demo
    K.sheet_head(bk, KEY, LAST_COL,
                 "\U0001F4B0  Sales Log",
                 "Log each sale once - the fair totals, dashboard, monthly "
                 "summary and best-sellers all read this page.")
    K.table_frame(bk, KEY, COLUMNS, height=22)

    first, last = C.ROW_FIRST, C.last_row(KEY)
    cat_name = bk.rng("catalog", "name")
    cat_cost = bk.rng("catalog", "unit_cost")
    for i in range(C.CAP[KEY]):
        row = C.ROW_FIRST + i
        s = m.sales[i] if m and i < len(m.sales) else None
        values = {"n": i + 1}
        if s:
            values.update({"date": s["date"], "event": s["event"],
                           "product": s["product"], "qty": s["qty"],
                           "unit": s["unit"], "discount": s["discount"],
                           "method": s["method"], "ref": s["ref"],
                           "notes": s["notes"]})
        values["total"] = ('=IF($D%d="","",IF($E%d="","",E%d*F%d)-'
                           'IF($G%d="",0,$G%d))'
                           % (row, row, row, row, row, row))
        values["pair"] = ('=IF($C%d="","",SUMIFS($E$%d:$E$%d,$C$%d:$C$%d,'
                          '$C%d,$D$%d:$D$%d,$D%d))'
                          % (row, first, last, first, last, row, first,
                             last, row))
        values["cost"] = ('=IF($D%d="","",$E%d*IFERROR(INDEX(%s,MATCH('
                          '$D%d,%s,0)),0))' % (row, row, cat_cost, row,
                                               cat_name))
        cached = {}
        if s:
            uc = _uc(m, s["product"])
            cached = {"total": round(s["qty"] * s["unit"] - s["discount"],
                                     2),
                      "pair": _pair(m, s),
                      "cost": round(s["qty"] * uc, 2)}
        K.write_row(bk, KEY, COLUMNS, row, values, cached)

    K.list_dv(bk, KEY, "event", "events", title="Fair / event")
    K.list_dv(bk, KEY, "product", "products", title="Product")
    K.list_dv(bk, KEY, "method", "payments", title="Payment method")
    K.date_dv(bk, KEY, ("date",))
    K.money_dv(bk, KEY, ("unit", "discount"))
    K.whole_dv(bk, KEY, ("qty",), minimum=0, maximum=100000)

    tot = C.last_row(KEY) + 1
    K.totals_row(bk, KEY, tot, {
        "qty": ("=SUM(E%d:E%d)" % (first, last), "#,##0",
                m.agg.get("units", 0) if m else 0),
        "total": ("=SUM(H%d:H%d)" % (first, last), "#,##0.00",
                  m.agg.get("revenue", 0) if m else 0),
        "cost": ("=SUM(M%d:M%d)" % (first, last), "#,##0.00",
                 m.agg.get("cogs", 0) if m else 0)},
        first_col="B", last_col=LAST_COL)

    chips = [
        ('="\U0001F4B0 Takings: "&Currency&TEXT(SUM(%s),"#,##0.00")'
         % bk.rng(KEY, "total"), "ok",
         "Takings: %s" % (m.money(m.agg.get("revenue", 0), 2)
                          if m else "$0.00"), 4),
        ('="\U0001F9F6 Units: "&%s' % bk.kpi("units"), "primary",
         "Units: %d" % (m.agg.get("units", 0) if m else 0), 2),
        ('="\U0001F9FE Sales logged: "&%s' % bk.kpi("txns"), "info",
         "Sales logged: %d" % (m.agg.get("txns", 0) if m else 0), 3),
        ('="\U0001F3F7\uFE0F Discounts: "&Currency&TEXT(SUM(%s),"#,##0.00")'
         % bk.rng(KEY, "discount"), "gold",
         "Discounts: %s" % (m.money(m.agg.get("discounts", 0), 2)
                            if m else "$0.00"), 3),
    ]
    K.chips(bk, KEY, chips, LAST_COL)
    K.footer_nav(bk, KEY, tot + 2, LAST_COL)


def _uc(m, product):
    for p in m.products:
        if p["name"] == product:
            return p["unit_cost"]
    return 0.0


def _pair(m, sale):
    return sum(s["qty"] for s in m.sales
               if s["event"] == sale["event"]
               and s["product"] == sale["product"])
