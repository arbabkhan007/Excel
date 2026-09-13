"""🧵 Yarn & Materials - what is on the shelf and what needs reordering."""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "materials"
LAST_COL = "N"

COLUMNS = [
    ("n", "#", "idx", None),
    ("name", "Material / yarn", "text", None),
    ("color", "Colour", "text", None),
    ("brand", "Brand", "text", None),
    ("weight", "Weight / type", "center", None),
    ("purchased", "Purchased", "qty", None),
    ("unit", "Unit", "center", None),
    ("cost_unit", "Cost / unit", "money", None),
    ("total", "Total cost", "calc_money", "primary_2"),
    ("used", "Used", "qty1", None),
    ("remaining", "Remaining", "calc_qty1", "primary_2"),
    ("supplier", "Supplier", "text", None),
    ("threshold", "Reorder at", "qty1", None),
    ("reorder", "Reorder?", "calc_c", None),
]


def build(bk):
    th = bk.th
    ws = bk.ws(KEY)
    m = bk.demo
    K.sheet_head(bk, KEY, LAST_COL,
                 "\U0001F9F5  Yarn & Materials",
                 "Because \u0022I have five tote bags\u0022 matters less than "
                 "\u0022I am almost out of the cotton to make more\u0022.")
    K.table_frame(bk, KEY, COLUMNS, height=22)

    for i in range(C.CAP[KEY]):
        row = C.ROW_FIRST + i
        mat = m.materials[i] if m and i < len(m.materials) else None
        values = {"n": i + 1}
        if mat:
            values.update({
                "name": mat["name"], "color": mat["color"],
                "brand": mat["brand"], "weight": mat["weight"],
                "purchased": mat["purchased"], "unit": mat["unit"],
                "cost_unit": mat["cost_unit"], "used": mat["used"],
                "supplier": mat["supplier"],
                "threshold": mat["threshold"]})
        values["total"] = '=IF($B%d="","",$F%d*$H%d)' % (row, row, row)
        values["remaining"] = '=IF($B%d="","",$F%d-$J%d)' % (row, row, row)
        values["reorder"] = ('=IF($B%d="","",IF($K%d<=$M%d,"%s","%s"))'
                             % (row, row, row, C.RE_YES, C.RE_NO))
        cached = {}
        if mat:
            cached = {"total": mat["total"],
                      "remaining": mat["remaining"],
                      "reorder": C.RE_YES if mat["remaining"]
                      <= mat["threshold"] else C.RE_NO}
        K.write_row(bk, KEY, COLUMNS, row, values, cached)

    K.list_dv(bk, KEY, "weight", "weights", title="Yarn weight")
    K.list_dv(bk, KEY, "unit", "units", title="Unit")
    K.list_dv(bk, KEY, "supplier", "suppliers", title="Supplier")
    K.money_dv(bk, KEY, ("cost_unit",), label="cost per unit")
    K.status_cf(bk, KEY, "reorder",
                {C.RE_YES: (th.bad_soft, th.bad),
                 C.RE_NO: (th.ok_soft, th.ok)})
    ws.conditional_format(
        r(C.ROW_FIRST), ci(bk.col(KEY, "remaining")),
        r(C.last_row(KEY)), ci(bk.col(KEY, "remaining")),
        {"type": "cell", "criteria": "<=",
         "value": "=$M%d" % C.ROW_FIRST,
         "format": bk.S.cf(bg=th.bad_soft, fg=th.bad, bold=True)})
    bk.stats["cond_formats"] += 1

    chips = [
        ('="\U0001F9F5 Materials: "&COUNTA(%s)' % bk.rng(KEY, "name"),
         "primary", "Materials: %d" % (len(m.materials) if m else 0), 3),
        ('="\U0001F4B0 Shelf value: "&Currency&TEXT(SUM(%s),"#,##0.00")'
         % bk.rng(KEY, "total"), "info",
         "Shelf value: %s" % (m.money(sum(x["total"] for x in m.materials),
                                      2) if m else "$0.00"), 4),
        ('="\U0001F534 Reorder now: "&%s' % bk.kpi("low_materials"), "bad",
         "Reorder now: %d" % (m.agg.get("low_materials", 0) if m else 0), 3),
        ('="\u2705 Healthy lines: "&COUNTIF(%s,"%s")'
         % (bk.rng(KEY, "reorder"), C.RE_NO), "ok",
         "Healthy lines: %d" % (sum(1 for x in (m.materials if m else [])
                                    if x["remaining"] > x["threshold"])), 4),
    ]
    K.chips(bk, KEY, chips, LAST_COL)
    K.footer_nav(bk, KEY, C.last_row(KEY) + 3, LAST_COL)
