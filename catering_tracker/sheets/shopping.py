"""
\U0001F6D2 Shopping List - auto-built from what events need vs what the
shelf already holds.

Required qty is the only real input besides the ingredient name:
Available = SUMIF over the inventory, To buy = MAX(0, required-available),
Est. cost = to-buy x the inventory unit cost.  Tick 🛒 Purchased when it's
in the van.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "shopping"
LAST_COL = "L"

COLUMNS = [
    ("n", "#", "idx", None),
    ("event", "Event ID", "center", None),
    ("ingredient", "Ingredient", "text", None),
    ("required", "Qty needed", "qty1", None),
    ("available", "In stock", "calc_qty1", "primary_2"),
    ("to_buy", "To buy", "calc_qty1", "primary_2"),
    ("supplier", "Supplier", "text", None),
    ("est_cost", "Est. cost", "calc_money", "primary_2"),
    ("purchased", "\U0001F6D2 Got it", "tick", None),
    ("date", "Bought on", "date", None),
    ("notes", "Notes", "text", None),
]


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY)
    bk.paint(KEY, 0, 0, C.last_row(KEY) + 24, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F6D2  Shopping List",
        "  What each event needs vs what the shelf holds \u2014 To buy and "
        "Est. cost fill themselves in from the \U0001F4E6 Inventory.",
        LAST_COL)

    # ------------------------------------------------------------------
    # quick-stat strip
    # ------------------------------------------------------------------
    ing_c = bk.rng(KEY, "ingredient")
    tick_c = bk.rng(KEY, "purchased")
    est_c = bk.rng(KEY, "est_cost")
    chips = [
        ('\U0001F6D2 Lines: "&COUNTIF(%s,"?*")' % ing_c,
         "primary", "Lines: %d" % len(m.shopping)),
        ('\U0001F4B0 Still to buy: "&%s&" line(s)"' % bk.kpi("shop_lines"),
         "warn", "Still to buy: %d line(s)" % m.agg.get("shop_lines", 0)),
        ('\U0001F4B5 Estimated cost to buy: "&Currency&TEXT(%s,"#,##0.00")'
         % bk.kpi("shop_cost"), "accent",
         "Estimated cost to buy: %s" % m.money(m.agg.get("shop_cost", 0), 2)),
        ('\u2705 Purchased: "&COUNTIF(%s,"%s")&" line(s)"'
         % (tick_c, C.TICK), "ok",
         "Purchased: %d line(s)"
         % sum(1 for s in m.shopping if s["purchased"] == C.TICK)),
        ('\U0001F9FE Spent on this list: "&Currency&TEXT(SUMIF(%s,"%s",%s),'
         '"#,##0.00")' % (tick_c, C.TICK, est_c), "primary_2",
         "Spent on this list: %s"
         % m.money(sum(s["est_cost"] for s in m.shopping
                       if s["purchased"] == C.TICK), 2)),
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
    inv_ing = bk.rng("inventory", "ingredient")
    inv_qty = bk.rng("inventory", "qty")
    inv_cost = bk.rng("inventory", "unit_cost")
    K.table_frame(bk, KEY, COLUMNS, height=20)
    for i in range(K.n_rows(KEY)):
        rownum = C.ROW_FIRST + i
        values = {
            "n": i + 1,
            "available": '=IF($D%d="",0,SUMIF(%s,$D%d,%s))'
                         % (rownum, inv_ing, rownum, inv_qty),
            "to_buy": '=IF($D%d="","",MAX(0,$E%d-$F%d))'
                      % (rownum, rownum, rownum),
            "est_cost": '=IF($G%d="","",ROUND($G%d*SUMIF(%s,$D%d,%s),2))'
                        % (rownum, rownum, inv_ing, rownum, inv_cost),
        }
        cached = {"available": 0, "to_buy": "", "est_cost": ""}
        if i < len(m.shopping):
            s = m.shopping[i]
            values.update({k: s[k] for k in
                           ("event", "ingredient", "required", "supplier",
                            "purchased", "date", "notes")})
            cached.update({"available": s["available"],
                           "to_buy": s["to_buy"], "est_cost": s["est_cost"]})
        K.write_row(bk, KEY, COLUMNS, rownum, values, cached)

    ws.autofilter(r(C.ROW_HEADER), ci("B"), r(C.last_row(KEY)), ci(LAST_COL))
    ws.freeze_panes(r(C.ROW_FIRST), ci("C"))

    # ------------------------------------------------------------------
    # validation
    # ------------------------------------------------------------------
    bk.validate(KEY, C.ROW_FIRST, ci(bk.col(KEY, "event")),
                C.last_row(KEY), ci(bk.col(KEY, "event")),
                "=EventList", title="Event ID",
                message="Which event needs this?")
    bk.validate(KEY, C.ROW_FIRST, ci(bk.col(KEY, "ingredient")),
                C.last_row(KEY), ci(bk.col(KEY, "ingredient")),
                "=OFFSET(%s!$C$%d,0,0,MAX(1,COUNTA(%s)),1)"
                % (bk.q("inventory"), C.ROW_FIRST, inv_ing),
                title="Ingredient",
                message="Match the spelling on the \U0001F4E6 Inventory tab "
                        "so In stock and Est. cost can look themselves up.")
    bk.validate(KEY, C.ROW_FIRST, ci(bk.col(KEY, "supplier")),
                C.last_row(KEY), ci(bk.col(KEY, "supplier")),
                "=SuppliersList", title="Supplier")
    ws.data_validation(r(C.ROW_FIRST), ci(bk.col(KEY, "required")),
                       r(C.last_row(KEY)), ci(bk.col(KEY, "required")),
                       {"validate": "decimal", "criteria": ">=", "value": 0,
                        "input_title": "Qty needed",
                        "input_message": "How much does the event need?",
                        "show_input": True, "ignore_blank": True})
    bk.stats["validations"] += 1
    K.tick_dv(bk, KEY, ("purchased",))
    K.date_dv(bk, KEY, ("date",))

    # ------------------------------------------------------------------
    # conditional formatting
    # ------------------------------------------------------------------
    K.tick_cf(bk, KEY, ("purchased",))
    # To buy > 0 and not purchased yet → red; purchased → whole row calms down
    bk.cond(KEY, C.ROW_FIRST, ci("G"), C.last_row(KEY), ci("G"), {
        "type": "formula",
        "criteria": '=AND($G%d>0,$J%d<>"%s")' % (C.ROW_FIRST, C.ROW_FIRST,
                                                 C.TICK),
        "format": S.cf(bg=th.bad_soft, fg=th.bad, bold=True)})
    bk.cond(KEY, C.ROW_FIRST, ci("G"), C.last_row(KEY), ci("G"), {
        "type": "formula",
        "criteria": '=AND($G%d=0,$D%d<>"")' % (C.ROW_FIRST, C.ROW_FIRST),
        "format": S.cf(bg=th.ok_soft, fg=th.ok, bold=True)})
    bk.cond(KEY, C.ROW_FIRST, ci("C"), C.last_row(KEY), ci("L"), {
        "type": "formula",
        "criteria": '=$J%d="%s"' % (C.ROW_FIRST, C.TICK),
        "format": S.cf(fg=th.muted, strike=True)})

    # ------------------------------------------------------------------
    # totals + notes
    # ------------------------------------------------------------------
    trow = C.last_row(KEY) + 2
    K.totals_row(bk, KEY, trow, {
        "required": ("=SUM($E$%d:$E$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                     "#,##0.0", sum(s["required"] for s in m.shopping)),
        "to_buy": ("=SUM($G$%d:$G$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                   "#,##0.0", sum(s["to_buy"] for s in m.shopping)),
        "est_cost": ("=SUM($I$%d:$I$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                     "#,##0.00",
                     sum(s["est_cost"] for s in m.shopping)),
    }, label="TOTALS  \u2192", label_span=("B", "D"), last_col=LAST_COL)

    K.note_block(
        bk, KEY, trow + 2, "B", "L",
        ["Add one line per ingredient per event - In stock, To buy and Est. "
         "cost fill themselves in.",
         "Red 'To buy' = still needs purchasing.  Green = the shelf already "
         "covers it.  Struck-through = bought and done.",
         "When you buy, log the real spend on \U0001F4B8 Expenses (category "
         "Ingredients) so the P&L stays true."],
        title="\U0001F4A1 How to use this tab")

    nav = trow + 8
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)
    bk.page(KEY, LAST_COL, nav + 1, freeze=(C.ROW_FIRST, 2),
            title_rows=(C.ROW_HEADER, C.ROW_HEADER))
