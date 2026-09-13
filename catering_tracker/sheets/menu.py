"""
\U0001F37D️ Menu & Recipe Costing.

Top: the priced menu (cost per portion, price per portion, profit and
margin - with a red flag under 30%).
Bottom: the RECIPE COST CALCULATOR - list up to six ingredients with the
quantity per portion; unit prices are looked up from Inventory, the cost
per portion and a suggested selling price (at your default margin) build
themselves.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "menu"
LAST_COL = "K"

COLUMNS = [
    ("n", "#", "idx", None),
    ("item", "Menu item", "text", None),
    ("category", "Category", "center", None),
    ("ingredients", "Main ingredients", "wrap", None),
    ("portion", "Portion", "center", None),
    ("cost", "Cost / portion", "money", None),
    ("price", "Price / portion", "money", None),
    ("profit", "Profit", "calc_money", "primary_2"),
    ("margin", "Margin", "calc_pct", "primary_2"),
    ("notes", "Notes", "text", None),
]




def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY)
    trow = C.last_row(KEY) + 2
    rc_top = trow + 9
    last = rc_top + 16
    bk.paint(KEY, 0, 0, last + 12, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F37D\uFE0F  Menu & Recipe Costing",
        "  Know the true cost of every dish.  Margins under 30% light up "
        "red \u2014 reprice them or retire them.",
        LAST_COL)

    # ------------------------------------------------------------------
    # quick-stat strip
    # ------------------------------------------------------------------
    item_c = bk.rng(KEY, "item")
    marg_c = bk.rng(KEY, "margin")
    chips = [
        ('\U0001F37D\uFE0F Items priced: "&%s' % bk.kpi("menu_items"),
         "primary", "Items priced: %d" % m.agg.get("menu_items", 0)),
        ('\U0001F4C8 Average margin: "&TEXT(%s,"0%%")' % bk.kpi("menu_margin"),
         "ok", "Average margin: %d%%"
         % round(m.agg.get("menu_margin", 0) * 100)),
        ('\u26A0\uFE0F Under 30%% margin: "&TEXT(COUNTIF(%s,">0")-'
         'COUNTIF(%s,">=0.3"),"0")' % (marg_c, marg_c), "bad",
         "Under 30%% margin: %d"
         % sum(1 for x in m.menu if x["margin"] < 0.30)),
    ]
    col = 1
    for formula, color, cached in chips:
        fmt = S.pill(th.soft(color), getattr(th, color), size=10.5,
                     bold=True, align="left")
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), col + 2, "", fmt)
        ws.write_formula(r(C.ROW_STATS), col, '="' + formula, fmt, cached)
        bk.stats["formulas"] += 1
        col += 3
    if col < ci(LAST_COL):
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), ci(LAST_COL),
                       "", S.canvas)
    elif col == ci(LAST_COL):
        ws.write_blank(r(C.ROW_STATS), col, None, S.canvas)

    # ------------------------------------------------------------------
    # menu table
    # ------------------------------------------------------------------
    K.table_frame(bk, KEY, COLUMNS, height=34)
    for i in range(K.n_rows(KEY)):
        rownum = C.ROW_FIRST + i
        values = {
            "n": i + 1,
            "profit": '=IF(OR($C%d="",$H%d=""),"",$H%d-$G%d)'
                      % (rownum, rownum, rownum, rownum),
            "margin": '=IF(OR($C%d="",$H%d="",$H%d=0),"",$I%d/$H%d)'
                      % (rownum, rownum, rownum, rownum, rownum),
        }
        cached = {"profit": "", "margin": ""}
        if i < len(m.menu):
            it = m.menu[i]
            values.update({k: it[k] for k in
                           ("item", "category", "ingredients", "portion",
                            "cost", "price", "notes")})
            cached.update({"profit": it["profit"], "margin": it["margin"]})
        K.write_row(bk, KEY, COLUMNS, rownum, values, cached)

    ws.autofilter(r(C.ROW_HEADER), ci("B"), r(C.last_row(KEY)), ci(LAST_COL))
    ws.freeze_panes(r(C.ROW_FIRST), ci("C"))

    K.list_dv(bk, KEY, "category", "menu_categories", title="Menu category")
    K.money_dv(bk, KEY, ("cost", "price"))
    bk.cond(KEY, C.ROW_FIRST, ci("J"), C.last_row(KEY), ci("J"), {
        "type": "formula",
        "criteria": '=AND($C%d<>"",$J%d<0.3)' % (C.ROW_FIRST, C.ROW_FIRST),
        "format": S.cf(bg=th.bad_soft, fg=th.bad, bold=True)})
    bk.cond(KEY, C.ROW_FIRST, ci("J"), C.last_row(KEY), ci("J"), {
        "type": "formula",
        "criteria": '=AND($C%d<>"",$J%d>=0.6)' % (C.ROW_FIRST, C.ROW_FIRST),
        "format": S.cf(bg=th.ok_soft, fg=th.ok, bold=True)})
    K.databar(bk, KEY, "price", color=th.gold)

    K.totals_row(bk, KEY, trow, {
        "cost": ("=IFERROR(AVERAGE($G$%d:$G$%d),0)"
                 % (C.ROW_FIRST, C.last_row(KEY)), "#,##0.00",
                 sum(x["cost"] for x in m.menu) / max(1, len(m.menu))),
        "price": ("=IFERROR(AVERAGE($H$%d:$H$%d),0)"
                  % (C.ROW_FIRST, C.last_row(KEY)), "#,##0.00",
                  sum(x["price"] for x in m.menu) / max(1, len(m.menu))),
        "margin": ("=IFERROR(AVERAGE($J$%d:$J$%d),0)"
                   % (C.ROW_FIRST, C.last_row(KEY)), "0%",
                   m.agg.get("menu_margin", 0)),
    }, label="AVERAGES  \u2192", label_span=("B", "F"), last_col=LAST_COL)

    # ------------------------------------------------------------------
    # recipe cost calculator
    # ------------------------------------------------------------------
    sec = rc_top
    ws.set_row(r(sec), 26)
    ws.merge_range(r(sec), 1, r(sec), ci(LAST_COL),
                   "  \U0001F9EE  RECIPE COST CALCULATOR  \u2014  price a "
                   "new dish before it goes on the menu", S.section_accent)

    dish_row = sec + 1
    ws.set_row(r(dish_row), 32)
    lbl = S.f(**S.base(font_size=11, bold=True, font_color=th.ink,
                       bg_color=th.card, align="left", valign="vcenter",
                       border=1, border_color=th.border, indent=1))
    inp = S.f(**S.base(font_size=11, bold=True, font_color=th.primary,
                       bg_color=th.gold_soft, align="center",
                       valign="vcenter", border=2, border_color=th.gold,
                       locked=False))
    inp_l = S.f(**S.base(font_size=11, bold=True, font_color=th.primary,
                         bg_color=th.gold_soft, align="left",
                         valign="vcenter", border=2, border_color=th.gold,
                         indent=1, locked=False))
    calc = S.f(**S.base(font_size=11, bold=True, font_color=th.primary,
                        bg_color=th.primary_soft, align="center",
                        valign="vcenter", border=1, border_color=th.border,
                        num_format="#,##0.00"))
    ws.merge_range(r(dish_row), 1, r(dish_row), 2, "Dish being costed", lbl)
    ws.merge_range(r(dish_row), 3, r(dish_row), 6, "", inp_l)
    dish_name = m.menu[0]["item"] if m.menu else ""
    ws.write(r(dish_row), 3, dish_name, inp_l)
    ws.merge_range(r(dish_row), 7, r(dish_row), ci(LAST_COL),
                   "Unit prices below are looked up from \U0001F4E6 "
                   "Inventory - add the ingredient there first.",
                   S.note_plain)

    hdr_row = dish_row + 1
    ws.set_row(r(hdr_row), 20)
    for c1, c2, text in ((1, 1, "#"), (2, 3, "Ingredient"), (4, 5, "Qty per portion"),
                         (6, 7, "Unit"), (8, 9, "Unit cost (auto)"),
                         (10, 10, "Line cost")):
        if c1 == c2:
            ws.write(r(hdr_row), c1, text, S.header(th.primary_2))
        else:
            ws.merge_range(r(hdr_row), c1, r(hdr_row), c2, text,
                           S.header(th.primary_2))

    demo_recipe = []
    if m.menu:
        demo_recipe = [("Chicken Breast", 0.22, "kg", 12.00),
                       ("Butter", 0.03, "kg", 6.40),
                       ("Olive Oil", 0.02, "L", 8.20),
                       ("Fresh Herbs", 0.01, "kg", 9.00)]

    ing_first = hdr_row + 1
    for i in range(6):
        row = ing_first + i
        ws.set_row(r(row), 20)
        a = i % 2
        ws.write(r(row), 1, i + 1, S.idx(a))
        ws.merge_range(r(row), 2, r(row), 3, "", S.cell("text", a))
        ws.merge_range(r(row), 4, r(row), 5, "", S.cell("qty1", a))
        ws.merge_range(r(row), 6, r(row), 7, "", S.cell("center", a))
        ws.merge_range(r(row), 8, r(row), 9, "", S.cell("calc_money", a))
        ws.write_formula(
            r(row), 8,
            '=IF($C%d="","",IFERROR(SUMIF(%s,$C%d,%s),0))'
            % (row, bk.rng("inventory", "ingredient"), row,
               bk.rng("inventory", "unit_cost")),
            S.cell("calc_money", a),
            demo_recipe[i][3] if i < len(demo_recipe) else "")
        bk.stats["formulas"] += 1
        ws.write_formula(
            r(row), 10, '=IF(OR($C%d="",$E%d=""),"",$E%d*$I%d)'
            % (row, row, row, row), S.cell("calc_money", a),
            round(demo_recipe[i][1] * demo_recipe[i][3], 4)
            if i < len(demo_recipe) else "")
        bk.stats["formulas"] += 1

    ing_last = ing_first + 5
    # demo ingredients for the signature dish
    if m.menu:
        for i, (ing, qty, unit, _uc) in enumerate(demo_recipe):
            row = ing_first + i
            ws.write(r(row), 2, ing, S.cell("text", i % 2))
            ws.write(r(row), 4, qty, S.cell("qty1", i % 2))
            ws.write(r(row), 6, unit, S.cell("center", i % 2))
    demo_cost = round(sum(q * u for _i, q, _un, u in demo_recipe), 4)

    out_row = ing_last + 1
    margin_d = m.settings["margin"] or 0.35
    demo_price = round(demo_cost / (1 - margin_d), 4) if demo_cost else 0
    outs = [
        ("Cost per portion",
         "=SUM($K$%d:$K$%d)" % (ing_first, ing_last), demo_cost),
        ("Suggested price at your default margin",
         "=IFERROR($D$%d/(1-DefaultMargin),\"\")" % out_row, demo_price),
        ("Profit per portion at that price",
         "=IFERROR($D$%d-$D$%d,\"\")" % (out_row + 1, out_row),
         round(demo_price - demo_cost, 4)),
    ]
    hints = ["Add the finished dish to the menu table above when you're "
             "happy with the numbers.",
             "DefaultMargin comes from \u2699\uFE0F Setup - change it once, "
             "every suggested price follows.",
             "Ingredient names must match the \U0001F4E6 Inventory spelling "
             "for the unit-cost lookup to work."]
    for j, (label, formula, cached) in enumerate(outs):
        row = out_row + j
        ws.set_row(r(row), 24)
        ws.merge_range(r(row), 1, r(row), 2, label, lbl)
        ws.merge_range(r(row), 3, r(row), 4, "", calc)
        ws.write_formula(r(row), 3, formula, calc, cached)
        bk.stats["formulas"] += 1
        ws.merge_range(r(row), 5, r(row), ci(LAST_COL), hints[j],
                       S.note_plain)

    nav = out_row + 4
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)
    bk.page(KEY, LAST_COL, nav + 1, freeze=(C.ROW_FIRST, 2),
            title_rows=(C.ROW_HEADER, C.ROW_HEADER))
