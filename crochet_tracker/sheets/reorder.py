"""🔄 Reorder List - auto-pulled low products & materials, red when urgent."""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "reorder"
LAST_COL = "J"
PROD_SLOTS = C.RO_PROD_LAST - C.RO_PROD_FIRST + 1
MAT_SLOTS = C.RO_MAT_LAST - C.RO_MAT_FIRST + 1
PROD_HEAD = C.RO_PROD_FIRST - 1
MAT_HEAD = C.RO_MAT_FIRST - 1

COLUMNS = [
    ("item", "Item", 26, "text"),
    ("current", "Current", 9, "qty1"),
    ("min", "Minimum", 8, "qty1"),
    ("suggest", "Suggest order", 10, "qty1"),
    ("supplier", "Supplier", 18, "text"),
    ("est", "Est. cost", 10, "money"),
    ("priority", "Priority", 13, "center"),
    ("ordered", "Ordered?", 10, "tick"),
    ("ordered_date", "Order date", 12, "date"),
]


def build(bk):
    th = bk.th
    ws = bk.ws(key := KEY)
    m = bk.demo
    premium = bk.has("materials")
    K.sheet_head(bk, KEY, LAST_COL,
                 "\U0001F504  Reorder List",
                 "Low products and yarn pull themselves onto this list - "
                 "urgent lines glow red until you tick them ordered.")
    bk.paint(KEY, 0, 0, MAT_HEAD + MAT_SLOTS + 8, ci(LAST_COL),
             bk.S.canvas)

    lab = bk.S.f(**bk.S.base(font_size=10.5, font_color=th.ink,
                             bg_color=th.card, align="left",
                             valign="vcenter", indent=1, border=1,
                             border_color=th.border))
    calc = bk.S.f(**bk.S.base(font_size=10.5, font_color=th.primary,
                              bg_color=th.alt, align="center",
                              valign="vcenter", border=1,
                              border_color=th.border,
                              num_format="#,##0.0#"))
    money = bk.S.f(**bk.S.base(font_size=10.5, font_color=th.primary,
                               bg_color=th.alt, align="right",
                               valign="vcenter", border=1,
                               border_color=th.border,
                               num_format="#,##0.00"))
    prio = bk.S.f(**bk.S.base(font_size=10.5, bold=True, font_color=th.ink,
                              bg_color=th.card, align="center",
                              valign="vcenter", border=1,
                              border_color=th.border))
    tick = bk.S.f(**bk.S.base(font_size=13, bold=True, font_color=th.ok,
                              bg_color=th.card, align="center",
                              valign="vcenter", border=1,
                              border_color=th.border, locked=False))
    datef = bk.S.f(**bk.S.base(font_size=10.5, font_color=th.ink,
                               bg_color=th.card, align="center",
                               valign="vcenter", border=1,
                               border_color=th.border,
                               num_format="dd mmm yyyy", locked=False))

    cat = bk.q("catalog")
    d = bk.q("data")
    seqp = "%s!$AK$%d:$AK$%d" % (d, C.DATA_SEQP_FIRST,
                                 C.DATA_SEQP_FIRST + C.CAP["catalog"] - 1)
    seqm = "%s!$AN$%d:$AN$%d" % (d, C.DATA_SEQM_FIRST,
                                 C.DATA_SEQM_FIRST + C.CAP["materials"] - 1)
    cl = C.ROW_FIRST
    clast = C.last_row("catalog")
    mlast = C.last_row("materials")

    def header(row, emoji, text):
        ws.set_row(r(row), 22)
        ws.merge_range(r(row), 1, r(row), ci(LAST_COL),
                       "  %s  %s" % (emoji, text), bk.S.section_soft)
        ws.set_row(r(row + 1), 20)
        for label, col in (("Item", 1), ("Current", 2), ("Minimum", 3),
                           ("Suggest", 4), ("Supplier", 5), ("Est. cost", 6),
                           ("Priority", 7), ("Ordered?", 8),
                           ("Order date", 9)):
            ws.write(r(row + 1), col, label, bk.S.thead)

    def slot(row, k, kind):
        ws.set_row(r(row), 20)
        if kind == "product":
            seq, sheet, name_col, cur_col, min_col, cost_col, sup = (
                seqp, cat, "C", "M", "N", "I", None)
            last = clast
        else:
            seq, sheet, name_col, cur_col, min_col, cost_col, sup = (
                seqm, bk.q("materials"), "B", "K", "M", "H", "L")
            last = mlast
        match = "MATCH(%d,%s,0)" % (k, seq)
        ws.write_formula(r(row), 1,
                         '=IFERROR(INDEX(%s!$%s$%d:$%s$%d,%s),"")'
                         % (sheet, name_col, cl, name_col, last, match),
                         lab, "")
        bk.stats["formulas"] += 1
        for col, src in ((2, cur_col), (3, min_col)):
            ws.write_formula(r(row), col,
                             '=IF($B%d="","",INDEX(%s!$%s$%d:$%s$%d,%s))'
                             % (row, sheet, src, cl, src, last, match),
                             calc, "")
            bk.stats["formulas"] += 1
        ws.write_formula(r(row), 4,
                         '=IF($B%d="","",($D%d*2)-$C%d)' % (row, row, row),
                         calc, "")
        bk.stats["formulas"] += 1
        if sup:
            ws.write_formula(r(row), 5,
                             '=IF($B%d="","",INDEX(%s!$%s$%d:$%s$%d,%s))'
                             % (row, sheet, sup, cl, sup, last, match),
                             lab, "")
            bk.stats["formulas"] += 1
        else:
            ws.write_formula(r(row), 5,
                             '=IF($B%d="","","Make in-house")' % row,
                             lab, "")
            bk.stats["formulas"] += 1
        ws.write_formula(r(row), 6,
                         '=IF($B%d="","",$E%d*INDEX(%s!$%s$%d:$%s$%d,%s))'
                         % (row, row, sheet, cost_col, cl, cost_col, last,
                            match),
                         money, "")
        bk.stats["formulas"] += 1
        ws.write_formula(r(row), 7,
                         '=IF($B%d="","",IF($C%d<=0,"%s",IF($C%d<=$D%d/2,'
                         '"%s","%s")))'
                         % (row, row, C.PR_URGENT, row, row, C.PR_HIGH,
                            C.PR_NORMAL),
                         prio, "")
        bk.stats["formulas"] += 1
        ws.write(r(row), 8, "", tick)
        ws.write(r(row), 9, "", datef)

    header(PROD_HEAD - 1, "\U0001F9F6", "PRODUCTS RUNNING LOW")
    for k in range(1, PROD_SLOTS + 1):
        slot(PROD_HEAD + k, k, "product")
    if premium:
        header(MAT_HEAD - 1, "\U0001F9F5", "MATERIALS RUNNING LOW")
        for k in range(1, MAT_SLOTS + 1):
            slot(MAT_HEAD + k, k, "material")

    # conditional formatting: priority colours + ticked green
    first_p, last_p = PROD_HEAD + 1, PROD_HEAD + PROD_SLOTS
    first_m, last_m = MAT_HEAD + 1, MAT_HEAD + MAT_SLOTS
    for txt, bg, fg in ((C.PR_URGENT, th.bad_soft, th.bad),
                        (C.PR_HIGH, th.warn_soft, th.warn),
                        (C.PR_NORMAL, th.ok_soft, th.ok)):
        for a, b in ((first_p, last_p), (first_m, last_m)):
            bk.cond(KEY, a, 7, b, 7, {
                "type": "formula",
                "criteria": '=$H%d="%s"' % (a, txt),
                "format": bk.S.cf(bg=bg, fg=fg, bold=True)})
    for a, b in ((first_p, last_p), (first_m, last_m)):
        bk.cond(KEY, a, 8, b, 8, {
            "type": "formula",
            "criteria": '=$I%d="%s"' % (a, C.TICK),
            "format": bk.S.cf(bg=th.ok_soft, fg=th.ok, bold=True)})
    bk.validate(KEY, first_p, 8, last_m if premium else last_p, 8,
                "=Tick", title="Ordered?")
    K.date_dv(bk, KEY, ("ordered_date",), first=first_p,
              last=last_m if premium else last_p)

    # chips
    allb = "$B$%d:$B$%d" % (first_p, last_m if premium else last_p)
    alli = "$I$%d:$I$%d" % (first_p, last_m if premium else last_p)
    allg = "$G$%d:$G$%d" % (first_p, last_m if premium else last_p)
    allh = "$H$%d:$H$%d" % (first_p, last_m if premium else last_p)
    chips = [
        ('="\U0001F534 Open lines: "&SUMPRODUCT((%s<>"")*(%s<>"%s"))'
         % (allb, alli, C.TICK), "bad",
         "Open lines: %d" % ((m.agg.get("low_products", 0)
                              + m.agg.get("low_materials", 0)) if m else 0),
         4),
        ('="\u26A1 Urgent: "&COUNTIF(%s,"%s")' % (allh, C.PR_URGENT),
         "warn", "Urgent: %d" % (sum(
             1 for p in (m.products if m else []) if p["stock"] <= 0)
             + sum(1 for x in (m.materials if m else [])
                   if x["remaining"] <= 0 and x["remaining"]
                   <= x["threshold"])), 3),
        ('="\U0001F4B0 Budget to reorder: "&Currency&TEXT(SUMPRODUCT((%s'
         '<>"")*(%s<>"%s")*IFERROR(%s*1,0)),"#,##0.00")'
         % (allb, alli, C.TICK, allg),
         "info", "Budget to reorder: %s"
         % (m.money(m.agg.get("reorder_cost", 0), 2) if m else "$0.00"), 4),
        ('="\u2705 Ordered lines: "&COUNTIF(%s,"%s")' % (alli, C.TICK),
         "ok", "Ordered lines: 0", 3),
    ]
    K.chips(bk, KEY, chips, LAST_COL)
    K.footer_nav(bk, KEY, (last_m if premium else last_p) + 3, LAST_COL)
