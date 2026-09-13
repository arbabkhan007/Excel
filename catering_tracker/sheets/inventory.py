"""
\U0001F4E6 Ingredient Inventory - with LOW STOCK conditional formatting.

Qty, minimum level and unit cost are inputs; Stock value and Reorder status
are formulas.  🔴 Reorder / 🟡 Low / 🟢 OK colour themselves.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "inventory"
LAST_COL = "L"

COLUMNS = [
    ("n", "#", "idx", None),
    ("ingredient", "Ingredient", "text", None),
    ("category", "Category", "center", None),
    ("unit", "Unit", "center", None),
    ("qty", "Qty on hand", "qty1", None),
    ("min", "Minimum level", "qty1", None),
    ("unit_cost", "Unit cost", "money", None),
    ("value", "Stock value", "calc_money", "primary_2"),
    ("supplier", "Supplier", "text", None),
    ("status", "Reorder status", "calc_c", "primary_2"),
    ("notes", "Notes", "text", None),
]

UNITS = "kg,g,L,mL,pc,tray,box,bag,bottle"


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY)
    bk.paint(KEY, 0, 0, C.last_row(KEY) + 24, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F4E6  Ingredient Inventory",
        "  What's on the shelf right now.  Stock below its minimum lights "
        "up \U0001F534 red - the shopping list and recipe calculator read "
        "these prices.",
        LAST_COL)

    # ------------------------------------------------------------------
    # quick-stat strip
    # ------------------------------------------------------------------
    ing_c = bk.rng(KEY, "ingredient")
    status_c = bk.rng(KEY, "status")
    chips = [
        ('\U0001F4E6 Ingredients: "&COUNTIF(%s,"?*")' % ing_c,
         "primary", "Ingredients: %d" % len(m.inventory)),
        ('\U0001F4B0 Stock value: "&Currency&TEXT(%s,"#,##0.00")'
         % bk.kpi("inv_value"), "primary_2",
         "Stock value: %s" % m.money(m.agg.get("inv_value", 0), 2)),
        ('\U0001F534 Reorder now: "&COUNTIF(%s,"\U0001F534 Reorder")'
         % status_c, "bad",
         "Reorder now: %d" % sum(1 for i in m.inventory
                                 if i["status"] == "\U0001F534 Reorder")),
        ('\U0001F7E1 Running low: "&COUNTIF(%s,"\U0001F7E1 Low")' % status_c,
         "warn", "Running low: %d"
         % sum(1 for i in m.inventory if i["status"] == "\U0001F7E1 Low")),
        ('\U0001F7E2 Well stocked: "&COUNTIF(%s,"\U0001F7E2 OK")' % status_c,
         "ok", "Well stocked: %d"
         % sum(1 for i in m.inventory if i["status"] == "\U0001F7E2 OK")),
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
        values = {
            "n": i + 1,
            "value": '=IF(OR($C%d="",$F%d=""),"",$F%d*$H%d)'
                     % (rownum, rownum, rownum, rownum),
            "status": ('=IF($C%d="","",IF($F%d<$G%d,"\U0001F534 Reorder",'
                       'IF($F%d<$G%d*1.5,"\U0001F7E1 Low","\U0001F7E2 OK")))'
                       % (rownum, rownum, rownum, rownum, rownum)),
        }
        cached = {"value": "", "status": ""}
        if i < len(m.inventory):
            it = m.inventory[i]
            values.update({k: it[k] for k in
                           ("ingredient", "category", "unit", "qty", "min",
                            "unit_cost", "supplier", "notes")})
            cached.update({"value": it["value"], "status": it["status"]})
        K.write_row(bk, KEY, COLUMNS, rownum, values, cached)

    ws.autofilter(r(C.ROW_HEADER), ci("B"), r(C.last_row(KEY)), ci(LAST_COL))
    ws.freeze_panes(r(C.ROW_FIRST), ci("C"))

    # ------------------------------------------------------------------
    # validation
    # ------------------------------------------------------------------
    K.list_dv(bk, KEY, "category", "ingredient_categories",
              title="Ingredient category")
    bk.validate(KEY, C.ROW_FIRST, ci(bk.col(KEY, "unit")),
                C.last_row(KEY), ci(bk.col(KEY, "unit")),
                '"%s"' % UNITS, title="Unit",
                message="kg, L, tray\u2026 or type your own.")
    bk.validate(KEY, C.ROW_FIRST, ci(bk.col(KEY, "supplier")),
                C.last_row(KEY), ci(bk.col(KEY, "supplier")),
                "=SuppliersList", title="Supplier",
                message="Pick from the \U0001F69A Suppliers tab, or type a "
                        "new one.")
    ws.data_validation(r(C.ROW_FIRST), ci(bk.col(KEY, "qty")),
                       r(C.last_row(KEY)), ci(bk.col(KEY, "min")),
                       {"validate": "decimal", "criteria": ">=", "value": 0,
                        "input_title": "Quantity",
                        "input_message": "Zero or more (decimals are fine).",
                        "show_input": True, "ignore_blank": True})
    bk.stats["validations"] += 1
    K.money_dv(bk, KEY, ("unit_cost",))

    # ------------------------------------------------------------------
    # conditional formatting
    # ------------------------------------------------------------------
    K.status_cf(bk, KEY, "status", {
        "\U0001F534 Reorder": (th.bad_soft, th.bad),
        "\U0001F7E1 Low": (th.warn_soft, th.warn),
        "\U0001F7E2 OK": (th.ok_soft, th.ok),
    })
    # the whole qty cell goes red when below minimum
    bk.cond(KEY, C.ROW_FIRST, ci("F"), C.last_row(KEY), ci("F"), {
        "type": "formula",
        "criteria": '=AND($C%d<>"",$F%d<$G%d)' % (C.ROW_FIRST, C.ROW_FIRST,
                                                  C.ROW_FIRST),
        "format": S.cf(bg=th.bad_soft, fg=th.bad, bold=True)})
    K.databar(bk, KEY, "value", color=th.primary_2)

    # ------------------------------------------------------------------
    # totals + notes
    # ------------------------------------------------------------------
    trow = C.last_row(KEY) + 2
    K.totals_row(bk, KEY, trow, {
        "value": ("=SUM($I$%d:$I$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                  "#,##0.00", m.agg.get("inv_value", 0)),
    }, label="TOTAL STOCK VALUE  \u2192", label_span=("B", "H"),
        last_col=LAST_COL)

    K.note_block(
        bk, KEY, trow + 2, "B", "L",
        ["\U0001F534 Reorder = below minimum. \U0001F7E1 Low = below 1.5x "
         "the minimum. \U0001F7E2 OK = comfortable.",
         "Unit costs typed here power the \U0001F37D\uFE0F Recipe "
         "Calculator and the \U0001F6D2 Shopping List estimates.",
         "Do a quick shelf count before every big event - ten minutes now "
         "saves an emergency run later."],
        title="\U0001F4A1 How to use this tab")

    nav = trow + 8
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)
    bk.page(KEY, LAST_COL, nav + 1, freeze=(C.ROW_FIRST, 2),
            title_rows=(C.ROW_HEADER, C.ROW_HEADER))
