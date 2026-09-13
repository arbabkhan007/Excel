"""🧶 Product Catalog - master list of everything you make and sell."""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "catalog"
LAST_COL = "P"


def _columns(has_prod):
    return [
        ("n", "#", "idx", None),
        ("sku", "SKU", "text", None),
        ("name", "Product name", "text", None),
        ("category", "Category", "center", None),
        ("desc", "Description", "wrap", None),
        ("price", "Selling price", "money", None),
        ("yarn_cost", "Yarn / material", "money", None),
        ("pack_cost", "Packaging", "money", None),
        ("unit_cost", "Unit cost", "calc_money", "primary_2"),
        ("profit", "Profit / item", "calc_money", "primary_2"),
        ("margin", "Margin %", "calc_pct", "primary_2"),
        ("time_h", "Hours", "qty1", None),
        ("stock", "In stock", "calc_num" if has_prod else "qty", None),
        ("min", "Min level", "qty", None),
        ("status", "Reorder status", "calc_c", None),
        ("link", "Photo / link", "link", None),
    ]


def build(bk):
    th = bk.th
    ws = bk.ws(key := KEY)
    m = bk.demo
    has_prod = bk.has("production")
    columns = _columns(has_prod)

    K.sheet_head(bk, KEY, LAST_COL,
                 "\U0001F9F6  Product Catalog",
                 "Everything you make, what it costs you and what it earns "
                 "you - stock counts flow in from \U0001F4E6 Made & Stocked.")
    K.table_frame(bk, KEY, columns, height=32)

    for i in range(C.CAP[KEY]):
        row = C.ROW_FIRST + i
        prod = m.products[i] if m and i < len(m.products) else None
        values = {"n": i + 1}
        if prod:
            values.update({
                "sku": prod["sku"], "name": prod["name"],
                "category": prod["category"], "desc": prod["desc"],
                "price": prod["price"], "yarn_cost": prod["yarn"],
                "pack_cost": prod["pack"], "time_h": prod["hours"],
                "min": prod["min"], "link": prod["link"]})
        values["unit_cost"] = '=IF($C%d="","",$G%d+$H%d)' % (row, row, row)
        values["profit"] = '=IF($F%d="","",$F%d-$I%d)' % (row, row, row)
        values["margin"] = '=IF($F%d="","",$J%d/$F%d)' % (row, row, row)
        values["status"] = ('=IF($C%d="","",IF($M%d<=0,"%s",IF($M%d<$N%d,'
                            '"%s","%s")))'
                            % (row, row, C.ST_OUT, row, row, C.ST_LOW,
                               C.ST_OK))
        if has_prod:
            values["stock"] = ('=SUMIFS(%s,%s,$C%d)'
                               % (bk.rng("production", "current"),
                                  bk.rng("production", "product"), row))
        elif prod:
            values["stock"] = prod["stock"]
        cached = {}
        if prod:
            cached = {"unit_cost": prod["unit_cost"],
                      "profit": prod["profit"], "margin": prod["margin"],
                      "status": {"out": C.ST_OUT, "low": C.ST_LOW,
                                 "ok": C.ST_OK}[prod["status"]],
                      "stock": prod["stock"]}
        K.write_row(bk, KEY, columns, row, values, cached)

    K.list_dv(bk, KEY, "category", "categories", title="Product category")
    K.money_dv(bk, KEY, ("price", "yarn_cost", "pack_cost"))
    K.whole_dv(bk, KEY, ("min",), minimum=0, maximum=100000)
    K.status_cf(bk, KEY, "status",
                {C.ST_OUT: (th.bad_soft, th.bad),
                 C.ST_LOW: (th.warn_soft, th.warn),
                 C.ST_OK: (th.ok_soft, th.ok)})
    ws.conditional_format(
        r(C.ROW_FIRST), ci(bk.col(KEY, "margin")),
        r(C.last_row(KEY)), ci(bk.col(KEY, "margin")),
        {"type": "cell", "criteria": "<", "value": 0.3,
         "format": bk.S.cf(bg=th.warn_soft, fg=th.warn, bold=True)})
    bk.stats["cond_formats"] += 1

    chips = [
        ('="\U0001F9F6 Products: "&COUNTA(%s)' % bk.rng(KEY, "name"),
         "primary", "Products: %d" % (len(m.products) if m else 0), 3),
        ('="\U0001F4B0 Stock value: "&Currency&TEXT(SUMPRODUCT(IFERROR(%s*1,0)*IFERROR(%s*1,0)),'
         '"#,##0.00")' % (bk.rng(KEY, "stock"), bk.rng(KEY, "unit_cost")),
         "info", "Stock value: %s"
         % (m.money(m.agg["inv_value"], 2) if m else "$0.00"), 4),
        ('="\U0001F534 Below minimum: "&%s' % bk.kpi("low_products"),
         "bad", "Below minimum: %d" % (m.agg.get("low_products", 0)
                                       if m else 0), 4),
        ('="\U0001F4C8 Avg margin: "&TEXT(IFERROR(AVERAGE(%s),0),"0%%")'
         % bk.rng(KEY, "margin"), "ok",
         "Avg margin: %.0f%%" % (100 * _avg_margin(m) if m else 0), 3),
    ]
    K.chips(bk, KEY, chips, LAST_COL)
    K.footer_nav(bk, KEY, C.last_row(KEY) + 3, LAST_COL)


def _avg_margin(m):
    if not m or not m.products:
        return 0.0
    return sum(p["margin"] for p in m.products) / len(m.products)
