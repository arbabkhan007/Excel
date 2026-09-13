"""
\U0001F4C8 P&L & Reports.

  * Annual P&L card: revenue, COGS, gross profit, operating expenses,
    net profit, margins.
  * Month-by-month table (all formulas read the hidden _Data pools, which
    read the Payments / Expenses / Events tabs) + combo chart.
  * Revenue by event type + bar chart.
  * Top clients and best-selling menu items tables.
"""

from .. import config as C
from ..book import r, ci

KEY = "reports"
LAST_COL = "N"
PL_FIRST = 10               # monthly table first row (12 rows)


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo
    d = bk.q("data")
    months = m.agg.get("months") or [("", 0, 0, 0, 0)] * 12

    bk.widths(KEY, {"A": 2.2, "B": 13, "C": 13, "D": 13, "E": 13, "F": 11,
                    "G": 3, "H": 20, "I": 13, "J": 3, "K": 22, "L": 12,
                    "M": 3, "N": 4})
    bk.paint(KEY, 0, 0, 74, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F4C8  Profit & Loss \u00B7 Reports",
        "  The year at a glance, month by month \u2014 everything below is "
        "calculated from your tabs (reporting year: \u2699\uFE0F Setup).",
        LAST_COL)

    lbl = S.f(**S.base(font_size=10.5, bold=True, font_color=th.ink,
                       bg_color=th.card, align="left", valign="vcenter",
                       border=1, border_color=th.border, indent=1))
    val = S.f(**S.base(font_size=12, bold=True, font_color=th.primary,
                       bg_color=th.primary_soft, align="center",
                       valign="vcenter", border=1, border_color=th.border,
                       num_format="#,##0.00"))
    val_pct = S.f(**S.base(font_size=12, bold=True, font_color=th.primary,
                           bg_color=th.primary_soft, align="center",
                           valign="vcenter", border=1,
                           border_color=th.border, num_format="0.0%"))
    hero = S.f(**S.base(font_name=th.title_font, font_size=15, bold=True,
                        font_color=th.white, bg_color=th.ok, align="center",
                        valign="vcenter", border=1, border_color=th.ok,
                        num_format="#,##0.00"))

    # ------------------------------------------------------------------
    # annual P&L card
    # ------------------------------------------------------------------
    mf, ml = C.DATA_MONTH_FIRST, C.DATA_MONTH_FIRST + 11
    mrev = "%s!$I$%d:$I$%d" % (d, mf, ml)
    mexp = "%s!$J$%d:$J$%d" % (d, mf, ml)
    rev_c = m.agg.get("revenue", 0)
    cogs_c = m.agg.get("food_cost", 0)
    opex_c = m.agg.get("expenses", 0) - cogs_c
    net_c = rev_c - m.agg.get("expenses", 0)
    ws.set_row(r(8), 22)
    ws.merge_range(r(8), 1, r(8), ci("F"), "", S.section_soft)
    ws.write_formula(r(8), 1,
                     '="  \U0001F4CA  ANNUAL P&L  \u00B7  "&ReportYear',
                     S.section_soft,
                     "  \U0001F4CA  ANNUAL P&L  \u00B7  %d"
                     % m.settings["year"])
    bk.stats["formulas"] += 1
    pl_rows = [
        ("Revenue (cash received)", "=SUM(%s)" % mrev, rev_c, val),
        ("Cost of goods sold (ingredients + packaging)", "=%s"
         % bk.kpi("food_cost"), cogs_c, val),
        ("GROSS PROFIT", "=IF($D$%d=\"\",\"\",$D$%d-$D$%d)"
         % (PL_FIRST - 1, PL_FIRST - 1, PL_FIRST), None, hero),
        ("Operating expenses (everything else)", "=%s-%s"
         % (bk.kpi("expenses"), bk.kpi("food_cost")), opex_c, val),
        ("NET PROFIT", "=IF($D$%d=\"\",\"\",$D$%d-$D$%d)"
         % (PL_FIRST + 2, PL_FIRST, PL_FIRST + 2), None, hero),
        ("Gross margin %", "=IFERROR($D$%d/$D$%d,0)"
         % (PL_FIRST + 1, PL_FIRST - 1), None, val_pct),
        ("Net margin %", "=IFERROR($D$%d/$D$%d,0)"
         % (PL_FIRST + 3, PL_FIRST - 1), None, val_pct),
    ]
    gross_c = rev_c - cogs_c
    cached_pl = [rev_c, cogs_c, gross_c, opex_c, net_c,
                 gross_c / rev_c if rev_c else 0,
                 net_c / rev_c if rev_c else 0]
    for j, ((label, formula, _c, fmt), cached) in enumerate(zip(pl_rows,
                                                                cached_pl)):
        row = PL_FIRST - 1 + j        # starts at row 9
        big = fmt is hero
        ws.set_row(r(row), 28 if big else 22)
        ws.merge_range(r(row), 1, r(row), 2, label,
                       lbl if not big else S.f(**S.base(
                           font_size=11, bold=True, font_color=th.white,
                           bg_color=th.ok, align="left", valign="vcenter",
                           border=1, border_color=th.ok, indent=1)))
        ws.merge_range(r(row), 3, r(row), 5, "", fmt)
        ws.write_formula(r(row), 3, formula, fmt, cached)
        bk.stats["formulas"] += 1
        ws.write_blank(r(row), 6, None, S.canvas)
    annual_end = PL_FIRST - 1 + len(pl_rows) - 1        # row 15

    # ------------------------------------------------------------------
    # monthly table
    # ------------------------------------------------------------------
    mtop = annual_end + 2
    ws.set_row(r(mtop), 22)
    ws.merge_range(r(mtop), 1, r(mtop), ci("F"),
                   "  \U0001F4C5  MONTH BY MONTH", S.section_soft)
    hrow = mtop + 1
    ws.set_row(r(hrow), 26)
    for j, head in enumerate(["Month", "Cash in", "Expenses", "Profit",
                              "Events"]):
        ws.write(r(hrow), 1 + j, head, S.header(th.primary))
    ws.write_blank(r(hrow), 6, None, S.canvas)
    m_first = hrow + 1
    for i in range(12):
        row = m_first + i
        drow = C.DATA_MONTH_FIRST + i
        a = i % 2
        ws.set_row(r(row), 18)
        ws.write_formula(r(row), 1, "=%s!$H$%d" % (d, drow),
                         S.cell("center", a), months[i][0])
        ws.write_formula(r(row), 2, "=%s!$I$%d" % (d, drow),
                         S.cell("calc_money", a), months[i][1])
        bk.stats["formulas"] += 2
        ws.write_formula(r(row), 3, "=%s!$J$%d" % (d, drow),
                         S.cell("calc_money", a), months[i][2])
        ws.write_formula(r(row), 4, "=%s!$K$%d" % (d, drow),
                         S.cell("calc_money", a), months[i][3])
        ws.write_formula(r(row), 5, "=%s!$L$%d" % (d, drow),
                         S.cell("calc_num", a), months[i][4])
        bk.stats["formulas"] += 3
        ws.write_blank(r(row), 6, None, S.canvas)
    m_last = m_first + 11
    trow = m_last + 1
    ws.set_row(r(trow), 22)
    ws.write(r(trow), 1, "YEAR", S.header(th.accent))
    year_cached = [sum(mm[1] for mm in months), sum(mm[2] for mm in months),
                   sum(mm[3] for mm in months)]
    for j, colL in enumerate(["B", "C", "D"]):
        ws.write_formula(r(trow), 2 + j,
                         "=SUM($%s$%d:$%s$%d)" % (colL, m_first, colL,
                                                  m_last),
                         S.f(**S.base(font_size=11, bold=True,
                                      font_color=th.white, bg_color=th.accent,
                                      align="right", valign="vcenter",
                                      border=1, border_color=th.accent,
                                      num_format="#,##0.00")),
                         year_cached[j])
        bk.stats["formulas"] += 1
    ws.write_formula(r(trow), 5,
                     "=SUM($E$%d:$E$%d)" % (m_first, m_last),
                     S.f(**S.base(font_size=11, bold=True,
                                  font_color=th.white, bg_color=th.accent,
                                  align="center", valign="vcenter",
                                  border=1, border_color=th.accent,
                                  num_format="0")),
                     sum(mm[4] for mm in months))
    bk.stats["formulas"] += 1
    ws.write_blank(r(trow), 6, None, S.canvas)

    # CF: profit column red when negative
    bk.cond(KEY, m_first, 4, m_last, 4, {
        "type": "cell", "criteria": "<", "value": 0,
        "format": S.cf(bg=th.bad_soft, fg=th.bad, bold=True)})

    # ------------------------------------------------------------------
    # revenue by event type (right column) + top clients + menu items
    # ------------------------------------------------------------------
    ws.set_row(r(mtop), 22)
    ws.merge_range(r(mtop), 7, r(mtop), ci("I"),
                   "  \U0001F3AF  REVENUE BY EVENT TYPE", S.section_soft)
    ws.set_row(r(hrow), 26)
    ws.write(r(hrow), 7, "Event type", S.header(th.primary))
    ws.write(r(hrow), 8, "Booked value", S.header(th.primary))
    tf, tl = C.DATA_TYPE_FIRST, C.DATA_TYPE_FIRST + len(C.EVENT_TYPES) - 1
    types_cached = dict(m.agg.get("types") or [])
    for i in range(len(C.EVENT_TYPES)):
        row = m_first + i
        a = i % 2
        ws.set_row(r(row), 18)
        ws.write_formula(r(row), 7, "=%s!$N$%d" % (d, tf + i),
                         S.cell("center", a), C.EVENT_TYPES[i])
        ws.write_formula(r(row), 8, "=%s!$O$%d" % (d, tf + i),
                         S.cell("calc_money", a),
                         types_cached.get(C.EVENT_TYPES[i], 0))
        bk.stats["formulas"] += 2
    t_last_type = m_first + len(C.EVENT_TYPES) - 1

    ctop = t_last_type + 2
    ws.set_row(r(ctop), 22)
    ws.merge_range(r(ctop), 7, r(ctop), ci("I"),
                   "  \U0001F451  TOP CLIENTS (cash received)",
                   S.section_soft)
    ws.set_row(r(ctop + 1), 26)
    ws.write(r(ctop + 1), 7, "Client", S.header(th.primary))
    ws.write(r(ctop + 1), 8, "Received", S.header(th.primary))
    cf_, cl_ = C.DATA_CLIENT_FIRST, C.DATA_CLIENT_FIRST + 9
    clients_cached = (m.agg.get("clients_pool") or [])[:10]
    for i in range(10):
        row = ctop + 2 + i
        a = i % 2
        ws.set_row(r(row), 18)
        ws.write_formula(r(row), 7, "=%s!$Q$%d" % (d, cf_ + i),
                         S.cell("text", a),
                         clients_cached[i][0] if i < len(clients_cached)
                         else "")
        ws.write_formula(r(row), 8, "=%s!$R$%d" % (d, cf_ + i),
                         S.cell("calc_money", a),
                         clients_cached[i][1] if i < len(clients_cached)
                         else 0)
        bk.stats["formulas"] += 2
    c_last = ctop + 11

    # best-selling menu items (columns K..L)
    ws.merge_range(r(mtop), 10, r(mtop), ci("L"),
                   "  \U0001F37D\uFE0F  MOST-BOOKED DISHES", S.section_soft)
    ws.set_row(r(hrow), 26)
    ws.write(r(hrow), 10, "Menu item", S.header(th.primary))
    ws.write(r(hrow), 11, "On menus", S.header(th.primary))
    mf_, ml_ = C.DATA_MENU_FIRST, C.DATA_MENU_FIRST + 13
    menu_cached = (m.agg.get("menu_pool") or [])[:14]
    for i in range(14):
        row = m_first + i
        a = i % 2
        ws.set_row(r(row), 18)
        ws.write_formula(r(row), 10, "=%s!$T$%d" % (d, mf_ + i),
                         S.cell("text", a),
                         menu_cached[i][0] if i < len(menu_cached) else "")
        ws.write_formula(r(row), 11, "=%s!$U$%d" % (d, mf_ + i),
                         S.cell("calc_num", a),
                         menu_cached[i][1] if i < len(menu_cached) else 0)
        bk.stats["formulas"] += 2
    menu_last = m_first + 13

    # ------------------------------------------------------------------
    # charts
    # ------------------------------------------------------------------
    ch_top = max(c_last, menu_last, m_last, trow) + 2
    qname = bk.q(KEY)
    cats = "=%s!$B$%d:$B$%d" % (qname, m_first, m_last)

    ch = bk.chart("column")
    ch.add_series({
        "name": "Cash in", "categories": cats,
        "values": "=%s!$C$%d:$C$%d" % (qname, m_first, m_last),
        "fill": {"color": th.primary}, "border": {"color": th.primary},
        "gap": 60})
    ch.add_series({
        "name": "Expenses", "categories": cats,
        "values": "=%s!$D$%d:$D$%d" % (qname, m_first, m_last),
        "fill": {"color": th.accent}, "border": {"color": th.accent}})
    line = bk.chart("line")
    line.add_series({
        "name": "Profit", "categories": cats,
        "values": "=%s!$E$%d:$E$%d" % (qname, m_first, m_last),
        "line": {"color": th.ok, "width": 2.5},
        "marker": {"type": "circle", "size": 5,
                   "fill": {"color": th.ok}}})
    ch.combine(line)
    ch.set_title({"name": "Cash in vs expenses vs profit, by month",
                  "name_font": {"size": 12, "bold": True, "color": th.primary,
                                "name": th.body_font}})
    ch.set_y_axis({"num_format": "#,##0", "num_font": {"size": 9},
                   "major_gridlines": {"visible": True,
                                       "line": {"color": th.border}}})
    ch.set_x_axis({"num_font": {"size": 9}})
    ch.set_size({"width": 620, "height": 300})
    ch.set_legend({"position": "bottom", "font": {"size": 9}})
    ws.insert_chart(r(ch_top), 1, ch, {"x_offset": 6, "y_offset": 6})

    ch2 = bk.chart("bar")
    ch2.add_series({
        "name": "Booked value",
        "categories": "=%s!$H$%d:$H$%d" % (qname, m_first, t_last_type),
        "values": "=%s!$I$%d:$I$%d" % (qname, m_first, t_last_type),
        "fill": {"color": th.gold}, "border": {"color": th.gold},
        "gap": 45,
        "data_labels": {"value": True, "font": {"size": 8, "color": th.ink},
                        "num_format": "#,##0"}})
    ch2.set_title({"name": "Revenue by event type",
                   "name_font": {"size": 12, "bold": True,
                                 "color": th.primary, "name": th.title_font}})
    ch2.set_x_axis({"num_format": "#,##0", "num_font": {"size": 9},
                    "major_gridlines": {"visible": True,
                                        "line": {"color": th.border}}})
    ch2.set_y_axis({"num_font": {"size": 9}, "label_position": "low"})
    ch2.set_size({"width": 560, "height": 300})
    ch2.set_legend({"none": True})
    ws.insert_chart(r(ch_top), 7, ch2, {"x_offset": 6, "y_offset": 6})

    # ------------------------------------------------------------------
    # footer
    # ------------------------------------------------------------------
    foot = ch_top + 17
    ws.set_row(r(foot), 20)
    ws.set_row(r(foot + 1), 20)
    ws.merge_range(r(foot), 1, r(foot + 1), ci(LAST_COL),
                   "  \U0001F4A1  Cash basis: \u201CCash in\u201D counts "
                   "money actually received (deposits + payments) in the "
                   "month of the event; expenses count on the day you "
                   "spent them. Switch the reporting year on \u2699\uFE0F "
                   "Setup. For official accounts, always reconcile with "
                   "your accountant.", S.note)
    nav = foot + 3
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)
    bk.page(KEY, LAST_COL, nav + 1, freeze=(0, 0), zoom=85)
