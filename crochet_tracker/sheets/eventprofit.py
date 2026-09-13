"""🧮 Event Profit - pick a fair, see exactly what it made (or lost)."""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "eventprofit"
LAST_COL = "J"


def build(bk):
    th = bk.th
    ws = bk.ws(KEY)
    m = bk.demo
    ep = m.agg.get("event_profit", {}) if m else {}
    K.sheet_head(bk, KEY, LAST_COL,
                 "\U0001F9EE  Event Profit",
                 "Pick any fair and see revenue, every cost line, net "
                 "profit, margin, break-even and ROI - instantly.")
    bk.paint(KEY, 0, 0, C.EP_LAST_ROW + 8, ci(LAST_COL), S_can(bk))

    lab = bk.S.f(**bk.S.base(font_size=11, font_color=bk.th.ink,
                             bg_color=bk.th.card, align="left",
                             valign="vcenter", indent=1, border=1,
                             border_color=bk.th.border))
    val = bk.S.f(**bk.S.base(font_size=12, bold=True,
                             font_color=bk.th.primary,
                             bg_color=bk.th.gold_soft, align="right",
                             valign="vcenter", border=1,
                             border_color=bk.th.border_strong,
                             num_format="#,##0.00"))
    val0 = bk.S.f(**bk.S.base(font_size=12, bold=True,
                              font_color=bk.th.primary,
                              bg_color=bk.th.gold_soft, align="right",
                              valign="vcenter", border=1,
                              border_color=bk.th.border_strong,
                              num_format="#,##0"))
    valp = bk.S.f(**bk.S.base(font_size=12, bold=True,
                              font_color=bk.th.primary,
                              bg_color=bk.th.gold_soft, align="right",
                              valign="vcenter", border=1,
                              border_color=bk.th.border_strong,
                              num_format="0.0%"))
    big = bk.S.f(**bk.S.base(font_size=15, bold=True,
                             font_color=bk.th.white, bg_color=bk.th.ok,
                             align="right", valign="vcenter", border=1,
                             border_color=bk.th.ok,
                             num_format="#,##0.00"))

    # picker
    ws.set_row(r(C.EP_EVENT_ROW), 26)
    ws.merge_range(r(C.EP_EVENT_ROW), 1, r(C.EP_EVENT_ROW), 4,
                   "  \U0001F3EA  Pick a fair to analyse", lab)
    ws.merge_range(r(C.EP_EVENT_ROW), 5, r(C.EP_EVENT_ROW), ci(LAST_COL),
                   "", bk.S.f(**bk.S.base(font_size=12, bold=True,
                                          font_color=bk.th.primary,
                                          bg_color=bk.th.gold_soft,
                                          border=1,
                                          border_color=bk.th.border_strong,
                                          align="left", valign="vcenter",
                                          indent=1)))
    bk.validate(KEY, C.EP_EVENT_ROW, 5, C.EP_EVENT_ROW, 5,
                "=" + bk.listname("events"), title="Pick a fair")
    sel = "$E$%d" % C.EP_EVENT_ROW
    if m:
        ws.write(r(C.EP_EVENT_ROW), 5, m.profit_event,
                 bk.S.f(**bk.S.base(font_size=12, bold=True,
                                    font_color=bk.th.primary,
                                    bg_color=bk.th.gold_soft, border=1,
                                    border_color=bk.th.border_strong,
                                    align="left", valign="vcenter",
                                    indent=1)))

    ev = bk.q("events")
    sale_ev = bk.rng("sales", "event")

    def section(row, emoji, text):
        ws.set_row(r(row), 22)
        ws.merge_range(r(row), 1, r(row), ci(LAST_COL),
                       "  %s  %s" % (emoji, text), bk.S.section)

    def line(row, label, formula, fmt, cached):
        ws.set_row(r(row), 20)
        ws.merge_range(r(row), 1, r(row), 5, "  " + label, lab)
        ws.merge_range(r(row), 6, r(row), ci(LAST_COL), "", fmt)
        ws.write_formula(r(row), 6, formula, fmt, cached)
        bk.stats["formulas"] += 1

    section(10, "\U0001F4B0", "REVENUE")
    line(C.EP_REV["units"], "Total units sold",
         '=SUMIFS(%s,%s,%s)' % (bk.rng("sales", "qty"), sale_ev, sel),
         val0, ep.get("units", 0))
    line(C.EP_REV["gross"], "Gross sales",
         '=SUMIFS(%s,%s,%s)' % (bk.rng("sales", "total"), sale_ev, sel),
         val, ep.get("gross", 0))
    line(C.EP_REV["discounts"], "Discounts given",
         '=SUMIFS(%s,%s,%s)' % (bk.rng("sales", "discount"), sale_ev, sel),
         val, ep.get("discounts", 0))

    section(15, "\U0001F9FE", "COSTS")
    line(C.EP_COST["cogs"], "Cost of products sold",
         '=SUMIFS(%s,%s,%s)' % (bk.rng("sales", "cost"), sale_ev, sel),
         val, ep.get("cogs", 0))
    for i, (field, label) in enumerate((("booth", "Booth fee"),
                                        ("travel", "Travel"),
                                        ("parking", "Parking"),
                                        ("food", "Food & drinks"),
                                        ("display", "Display & decor"))):
        line(C.EP_COST[field], label,
             '=IFERROR(INDEX(%s!$%s$%d:$%s$%d,MATCH(%s,%s!$B$%d:$B$%d,0)),0)'
             % (ev, C.COLS["events"][field], C.ROW_FIRST,
                C.COLS["events"][field], C.last_row("events"), sel, ev,
                C.ROW_FIRST, C.last_row("events")),
             val, ep.get(field, 0))

    section(23, "\U0001F3C1", "RESULTS")
    line(C.EP_OUT["gross_profit"], "Gross profit (sales \u2212 product cost)",
         "=$F$%d-$F$%d" % (C.EP_REV["gross"], C.EP_COST["cogs"]),
         val, ep.get("gross_profit", 0))
    line(C.EP_OUT["total_cost"], "Total costs (product + fair)",
         "=SUM($F$%d:$F$%d)" % (C.EP_COST["cogs"], C.EP_COST["display"]),
         val, ep.get("total_cost", 0))
    ws.set_row(r(C.EP_OUT["net"]), 26)
    ws.merge_range(r(C.EP_OUT["net"]), 1, r(C.EP_OUT["net"]), 5,
                   "  NET EVENT PROFIT", bk.S.f(**bk.S.base(
                       bold=True, font_size=12, font_color=bk.th.white,
                       bg_color=bk.th.ok, align="left", valign="vcenter",
                       indent=1)))
    ws.merge_range(r(C.EP_OUT["net"]), 6, r(C.EP_OUT["net"]), ci(LAST_COL),
                   "", big)
    ws.write_formula(r(C.EP_OUT["net"]), 6,
                     "=$F$%d-$F$%d" % (C.EP_REV["gross"],
                                      C.EP_OUT["total_cost"]),
                     big, ep.get("net", 0))
    bk.stats["formulas"] += 1
    line(C.EP_OUT["margin"], "Profit margin %",
         '=IFERROR($F$%d/$F$%d,0)' % (C.EP_OUT["net"], C.EP_REV["gross"]),
         valp, ep.get("margin", 0))
    line(C.EP_OUT["avg_sale"], "Average sale value",
         '=IFERROR($F$%d/COUNTIFS(%s,%s),0)'
         % (C.EP_REV["gross"], sale_ev, sel), val, ep.get("avg_sale", 0))
    line(C.EP_OUT["breakeven"], "Break-even sales",
         '=IFERROR($F$%d/($F$%d/$F$%d),0)'
         % (C.EP_OUT["total_cost"], C.EP_OUT["gross_profit"],
            C.EP_REV["gross"]), val, ep.get("breakeven", 0))
    line(C.EP_OUT["roi"], "Return on costs (ROI %)",
         '=IFERROR($F$%d/$F$%d,0)' % (C.EP_OUT["net"],
                                      C.EP_OUT["total_cost"]),
         valp, ep.get("roi", 0))

    K.note_block(bk, KEY, C.EP_LAST_ROW - 2, "B", LAST_COL, [
        "How to read this page:",
        "1.  Gross profit = takings minus what the sold items cost you in "
        "yarn and packaging.",
        "2.  Net profit also subtracts every fair cost - booth, travel, "
        "parking, food and display.",
        "3.  Break-even sales = the takings you needed to cover all costs "
        "at this margin.  Beat it next time!",
    ], title="  \U0001F4DA  READING YOUR NUMBERS")
    K.footer_nav(bk, KEY, C.EP_LAST_ROW + 4, LAST_COL, landscape=False)


def S_can(bk):
    return bk.S.canvas
