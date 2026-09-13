"""
\U0001F4CA Dashboard - the command centre.

Hero band (business name, next-event countdown, today's summary), two rows
of KPI cards, two progress bars, the six-stage pipeline strip, four charts
(monthly cash-in vs expenses vs profit, expense breakdown, pipeline
doughnut, revenue by event type) and the live "next events" / "money
coming in" panels.

Every number is a formula pointing at the hidden _Data sheet, so the
dashboard is always in sync with the tabs - and every formula ships with a
cached value so the EXAMPLE workbooks look perfect before Excel even
recalculates.
"""

from .. import config as C
from ..book import r, ci
from .events import STATUS_COLORS

KEY = "dashboard"
LAST_COL = "M"
CARDS = [(1, 2), (3, 4), (5, 6), (7, 8), (9, 10), (11, 12)]

ROW_HERO_1 = 2
ROW_HERO_2 = 3
ROW_HERO_3 = 4
ROW_MSG = 5
ROW_SEC_1 = 7
ROW_LBL_1 = 8
ROW_VAL_1 = 9
ROW_LBL_2 = 10
ROW_VAL_2 = 11
ROW_BAR_1 = 12
ROW_BAR_2 = 13
ROW_PIPE_SEC = 15
ROW_PIPE = 16
ROW_CHART_1 = 18
ROW_CHART_2 = 34
ROW_PANEL_SEC = 50
ROW_PANEL = 51
PANEL_ROWS = 6
ROW_FOOT = ROW_PANEL + PANEL_ROWS + 1


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo
    agg = m.agg
    d = bk.q("data")
    prem = bk.has("inventory")

    ws.set_column("A:A", 2.2)
    ws.set_column("B:M", 13)
    bk.paint(KEY, 0, 0, ROW_FOOT + 4, ci(LAST_COL), S.canvas)

    # ------------------------------------------------------------------
    # hero
    # ------------------------------------------------------------------
    ws.set_row(r(1), 8)
    ws.set_row(r(ROW_HERO_1), 42)
    ws.set_row(r(ROW_HERO_2), 32)
    ws.set_row(r(ROW_HERO_3), 20)
    ws.set_row(r(ROW_MSG), 22)

    ws.merge_range(r(ROW_HERO_1), 1, r(ROW_HERO_1), ci(LAST_COL), "",
                   S.hero_title)
    ws.write_formula(
        r(ROW_HERO_1), 1,
        '=IF(BusinessName="","\U0001F37D\uFE0F  Catering Command Center",'
        '"\U0001F37D\uFE0F  "&BusinessName&"   \u2022   Catering Command '
        'Center")', S.hero_title,
        "\U0001F37D\uFE0F  Catering Command Center"
        if not m.settings["business"] else
        "\U0001F37D\uFE0F  %s   \u2022   Catering Command Center"
        % m.settings["business"])
    bk.stats["formulas"] += 1

    ws.merge_range(r(ROW_HERO_2), 1, r(ROW_HERO_2), ci(LAST_COL), "",
                   S.hero_count)
    nxt_formula = _next_event(bk)
    ws.write_formula(r(ROW_HERO_2), 1, "=" + nxt_formula, S.hero_count,
                     _next_event_cached(m))
    bk.stats["formulas"] += 1

    ws.merge_range(r(ROW_HERO_3), 1, r(ROW_HERO_3), ci(LAST_COL), "",
                   S.hero_meta)
    meta = ('=TEXT(TODAY(),"dddd, dd mmmm yyyy")&"   \u2022   "&%s&'
            '" events on the books   \u2022   "&Currency&TEXT(%s,"#,##0")&'
            '" cash received   \u2022   "&Currency&TEXT(%s,"#,##0")&'
            '" still outstanding"'
            % (bk.kpi("events_total"), bk.kpi("revenue"),
               bk.kpi("outstanding")))
    ws.write_formula(
        r(ROW_HERO_3), 1, meta, S.hero_meta,
        "%s   \u2022   %d events on the books   \u2022   %s cash received"
        "   \u2022   %s still outstanding"
        % (_today_text(), agg.get("events_total", 0),
           m.money(agg.get("revenue", 0)), m.money(agg.get("outstanding", 0))))
    bk.stats["formulas"] += 1

    # message strip
    msg_fmt = S.f(**S.base(font_size=11, bold=True, italic=True,
                           font_color=th.white, bg_color=th.gold,
                           align="left", valign="vcenter", indent=1))
    ws.merge_range(r(ROW_MSG), 1, r(ROW_MSG), ci(LAST_COL), "", msg_fmt)
    ws.write_formula(
        r(ROW_MSG), 1,
        '=IF(%s!$C$%d="","","\U0001F4E3  "&%s!$C$%d)'
        % (bk.q("setup"), C.SU_MESSAGE, bk.q("setup"), C.SU_MESSAGE),
        msg_fmt,
        ("\U0001F4E3  " + m.settings["message"]) if m.settings["message"]
        else "")
    bk.stats["formulas"] += 1
    ws.set_row(r(6), 8)

    # ------------------------------------------------------------------
    # KPI cards
    # ------------------------------------------------------------------
    _section(ws, S, ROW_SEC_1, "  \U0001F4B0  BUSINESS AT A GLANCE")
    cards1 = [
        ("CASH RECEIVED", "ok",
         '=Currency&TEXT(%s,"#,##0")' % bk.kpi("revenue"),
         m.money(agg.get("revenue", 0))),
        ("TOTAL EXPENSES", "accent",
         '=Currency&TEXT(%s,"#,##0")' % bk.kpi("expenses"),
         m.money(agg.get("expenses", 0))),
        ("NET PROFIT", "primary",
         '=Currency&TEXT(%s,"#,##0")' % bk.kpi("profit"),
         m.money(agg.get("profit", 0))),
        ("NET MARGIN", "gold", "=%s" % bk.kpi("margin"),
         agg.get("margin", 0)),
        ("EVENTS ON BOOKS", "info", "=%s" % bk.kpi("events_total"),
         agg.get("events_total", 0)),
        ("AVG ORDER VALUE", "primary_2",
         '=Currency&TEXT(%s,"#,##0")' % bk.kpi("avg_order"),
         m.money(agg.get("avg_order", 0))),
    ]
    if prem:
        cards2 = [
            ("UPCOMING BOOKED", "primary",
             "=%s" % bk.kpi("events_confirmed"), agg.get("events_confirmed", 0)),
            ("OUTSTANDING", "warn",
             '=Currency&TEXT(%s,"#,##0")' % bk.kpi("outstanding"),
             m.money(agg.get("outstanding", 0))),
            ("OVERDUE INVOICES", "bad", "=%s" % bk.kpi("overdue"),
             agg.get("overdue", 0)),
            ("LOW-STOCK ITEMS", "bad", "=%s" % bk.kpi("low_stock"),
             agg.get("low_stock", 0)),
            ("UNPAID SHIFTS", "plum", "=%s" % bk.kpi("staff_unpaid"),
             agg.get("staff_unpaid", 0)),
            ("INVENTORY VALUE", "info",
             '=Currency&TEXT(%s,"#,##0")' % bk.kpi("inv_value"),
             m.money(agg.get("inv_value", 0))),
        ]
    else:
        cards2 = [
            ("UPCOMING BOOKED", "primary",
             "=%s" % bk.kpi("events_confirmed"), agg.get("events_confirmed", 0)),
            ("OUTSTANDING", "warn",
             '=Currency&TEXT(%s,"#,##0")' % bk.kpi("outstanding"),
             m.money(agg.get("outstanding", 0))),
            ("OVERDUE INVOICES", "bad", "=%s" % bk.kpi("overdue"),
             agg.get("overdue", 0)),
            ("OPEN INVOICES", "info", "=%s" % bk.kpi("invoices_open"),
             agg.get("invoices_open", 0)),
            ("REPEAT CUSTOMERS", "gold",
             '=TEXT(%s,"0%%")' % bk.kpi("repeat_pct"),
             agg.get("repeat_pct", 0)),
            ("GUESTS CATERED", "primary_2", "=%s" % bk.kpi("guests_total"),
             agg.get("guests_total", 0)),
        ]
    _cards(bk, ROW_LBL_1, ROW_VAL_1, cards1,
           num_formats=[None, None, None, "0%", "0", None])
    _cards(bk, ROW_LBL_2, ROW_VAL_2, cards2,
           num_formats=["0", None, "0", "0", "0", None])

    # ------------------------------------------------------------------
    # progress bars
    # ------------------------------------------------------------------
    bars = [
        (ROW_BAR_1, "  Net margin", "profit", "revenue", "margin",
         agg.get("margin", 0)),
        (ROW_BAR_2, "  Food cost % of cash in", "food_cost", "revenue",
         "food_pct", agg.get("food_pct", 0)),
    ]
    for row, label, num_k, den_k, pct_k, cached in bars:
        ws.set_row(r(row), 22)
        ws.merge_range(r(row), 1, r(row), 2, label, S.bar_label)
        ws.merge_range(r(row), 3, r(row), 9, "", S.bar_text)
        ws.write_formula(r(row), 3,
                         bk.bar(bk.kpi(num_k), bk.kpi(den_k), 34),
                         S.bar_text, bk.bar_static(cached, 1, 34))
        ws.merge_range(r(row), 10, r(row), ci(LAST_COL), "", S.bar_pct)
        ws.write_formula(r(row), 10,
                         '=TEXT(%s,"0.0%%")' % bk.kpi(pct_k),
                         S.bar_pct, "%.1f%%" % (cached * 100))
        bk.stats["formulas"] += 2
    ws.set_row(r(14), 8)

    # ------------------------------------------------------------------
    # pipeline strip
    # ------------------------------------------------------------------
    _section(ws, S, ROW_PIPE_SEC, "  \U0001F4C5  THE PIPELINE",
             style=S.section_accent)
    ws.set_row(r(ROW_PIPE), 30)
    ev_status = bk.rng("events", "status")
    counts = (agg.get("status_counts") or {})
    col = 1
    for st in C.EVENT_STATUSES:
        bg, fg = STATUS_COLORS[st]
        fmt = S.pill(getattr(th, bg), getattr(th, fg), size=11, bold=True,
                     align="center")
        ws.merge_range(r(ROW_PIPE), col, r(ROW_PIPE), col + 1, "", fmt)
        ws.write_formula(
            r(ROW_PIPE), col,
            '="%s: "&COUNTIF(%s,"%s")' % (st, ev_status, st), fmt,
            "%s: %d" % (st, counts.get(st, 0)))
        bk.stats["formulas"] += 1
        col += 2

    # ------------------------------------------------------------------
    # charts
    # ------------------------------------------------------------------
    _charts(bk)

    # ------------------------------------------------------------------
    # live panels: next events + money coming in
    # ------------------------------------------------------------------
    _panels(bk)

    # ------------------------------------------------------------------
    # footer
    # ------------------------------------------------------------------
    foot = ROW_FOOT + 1
    ws.set_row(r(foot), 30)
    ws.merge_range(r(foot), 1, r(foot), ci(LAST_COL),
                   "  \U0001F4A1  Everything on this page is calculated "
                   "\u2014 type your business into the tabs and watch it "
                   "update.  Start with the \U0001F4D6 Start Here guide, "
                   "set up \u2699\uFE0F Setup, then log \U0001F4C5 Events "
                   "and \U0001F4B0 Payments.",
                   S.f(**S.base(font_size=10, italic=True, font_color=th.ink,
                                bg_color=th.gold_soft, align="left",
                                valign="vcenter", indent=1)))
    nav = foot + 2
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)
    bk.page(KEY, LAST_COL, nav + 1, landscape=False, zoom=85)


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def _section(ws, S, row, text, style=None):
    ws.set_row(r(row), 22)
    ws.merge_range(r(row), 1, r(row), ci(LAST_COL), text, style or S.section)


def _cards(bk, label_row, value_row, cards, num_formats=None):
    ws, S, th = bk.ws(KEY), bk.S, bk.th
    num_formats = num_formats or [None] * len(cards)
    ws.set_row(r(label_row), 18)
    ws.set_row(r(value_row), 36)
    for i, (label, color, formula, cached) in enumerate(cards):
        c1, c2 = CARDS[i]
        lbl_fmt = S.kpi_label(getattr(th, color), size=9.5, align="center")
        ws.merge_range(r(label_row), c1, r(label_row), c2, label, lbl_fmt)
        val_fmt = S.kpi_value(getattr(th, color),
                              num_format=num_formats[i], size=20)
        ws.merge_range(r(value_row), c1, r(value_row), c2, "", val_fmt)
        ws.write_formula(r(value_row), c1, formula, val_fmt, cached)
        bk.stats["formulas"] += 1


def _next_event(bk):
    """Hero line: next upcoming event + countdown."""
    d = bk.q("data")
    pool_d = "%s!$AA$%d:$AA$%d" % (d, C.DATA_POOL_FIRST,
                                   C.DATA_POOL_FIRST + C.DATA_POOL_ROWS - 1)
    pool_l = "%s!$AB$%d:$AB$%d" % (d, C.DATA_POOL_FIRST,
                                   C.DATA_POOL_FIRST + C.DATA_POOL_ROWS - 1)
    nxt = "IFERROR(SMALL(%s,1),\"\")" % pool_d
    label = 'IFERROR(INDEX(%s,MATCH(%s,%s,0)),"")' % (pool_l, nxt, pool_d)
    return ('IF(%s="","\U0001F334  No upcoming events \u2014 add one on the '
            '\U0001F4C5 Events tab!","\U0001F525  Next event: "&%s&"  '
            '\u2014  in "&(%s-TODAY())&" day(s)")' % (nxt, label, nxt))


def _next_event_cached(m):
    up = m.agg.get("upcoming") or []
    if not up:
        return ("\U0001F334  No upcoming events \u2014 add one on the "
                "\U0001F4C5 Events tab!")
    from ..demo import TODAY
    when, label, _st = up[0]
    return "\U0001F525  Next event: %s  \u2014  in %d day(s)" % (
        label, (when - TODAY).days)


def _today_text():
    from ..demo import TODAY
    return TODAY.strftime("%A, %d %B %Y")


def _charts(bk):
    ws, th = bk.ws(KEY), bk.th
    d = bk.q("data")
    mf, ml = C.DATA_MONTH_FIRST, C.DATA_MONTH_FIRST + 11
    cats = "=%s!$H$%d:$H$%d" % (d, mf, ml)

    # 1 - monthly combo: cash in / expenses columns + profit line
    ch = bk.chart("column")
    ch.add_series({
        "name": "Cash in", "categories": cats,
        "values": "=%s!$I$%d:$I$%d" % (d, mf, ml),
        "fill": {"color": th.primary}, "border": {"color": th.primary},
        "gap": 60})
    ch.add_series({
        "name": "Expenses", "categories": cats,
        "values": "=%s!$J$%d:$J$%d" % (d, mf, ml),
        "fill": {"color": th.accent}, "border": {"color": th.accent}})
    line = bk.chart("line")
    line.add_series({
        "name": "Profit", "categories": cats,
        "values": "=%s!$K$%d:$K$%d" % (d, mf, ml),
        "line": {"color": th.ok, "width": 2.5},
        "marker": {"type": "circle", "size": 5, "fill": {"color": th.ok}}})
    ch.combine(line)
    ch.set_title({"name": "Money by month (reporting year on Setup)",
                  "name_font": {"size": 12, "bold": True,
                                "color": th.primary, "name": th.title_font}})
    ch.set_y_axis({"num_format": "#,##0", "num_font": {"size": 9},
                   "major_gridlines": {"visible": True,
                                       "line": {"color": th.border}}})
    ch.set_x_axis({"num_font": {"size": 9}})
    ch.set_size({"width": 555, "height": 285})
    ch.set_legend({"position": "bottom", "font": {"size": 9}})
    ws.insert_chart(r(ROW_CHART_1), ci("B"), ch, {"x_offset": 8,
                                                  "y_offset": 6})

    # 2 - expense breakdown doughnut
    ef, el = C.DATA_EXP_FIRST, C.DATA_EXP_FIRST + len(C.EXPENSE_CATEGORIES) - 1
    ch = bk.chart("doughnut")
    palette = [th.primary, th.accent, th.gold, th.info, th.plum, th.ok,
               th.primary_2, th.warn, th.bad, "#8C7B6B", "#6B8C7B"]
    ch.add_series({
        "name": "Expenses",
        "categories": "=%s!$N$%d:$N$%d" % (d, ef, el),
        "values": "=%s!$O$%d:$O$%d" % (d, ef, el),
        "points": [{"fill": {"color": palette[i % len(palette)]}}
                   for i in range(el - ef + 1)],
        "data_labels": {"percentage": True,
                        "font": {"size": 8, "color": th.white}},
    })
    ch.set_title({"name": "Where the money goes",
                  "name_font": {"size": 12, "bold": True,
                                "color": th.primary, "name": th.title_font}})
    ch.set_hole_size(58)
    ch.set_size({"width": 555, "height": 285})
    ch.set_legend({"position": "right", "font": {"size": 8}})
    ws.insert_chart(r(ROW_CHART_1), ci("H"), ch, {"x_offset": 8,
                                                  "y_offset": 6})

    # 3 - pipeline doughnut
    sf, sl = C.DATA_STATUS_FIRST, C.DATA_STATUS_FIRST + 5
    ch = bk.chart("doughnut")
    ch.add_series({
        "name": "Pipeline",
        "categories": "=%s!$N$%d:$N$%d" % (d, sf, sl),
        "values": "=%s!$O$%d:$O$%d" % (d, sf, sl),
        "points": [{"fill": {"color": c}} for c in
                   [th.info, th.plum, th.warn, th.primary, th.ok, th.bad]],
        "data_labels": {"value": True,
                        "font": {"size": 9, "color": th.white}},
    })
    ch.set_title({"name": "Event pipeline",
                  "name_font": {"size": 12, "bold": True,
                                "color": th.primary, "name": th.title_font}})
    ch.set_hole_size(58)
    ch.set_size({"width": 555, "height": 285})
    ch.set_legend({"position": "right", "font": {"size": 8}})
    ws.insert_chart(r(ROW_CHART_2), ci("B"), ch, {"x_offset": 8,
                                                  "y_offset": 6})

    # 4 - revenue by event type
    tf, tl = C.DATA_TYPE_FIRST, C.DATA_TYPE_FIRST + len(C.EVENT_TYPES) - 1
    ch = bk.chart("bar")
    ch.add_series({
        "name": "Booked value",
        "categories": "=%s!$N$%d:$N$%d" % (d, tf, tl),
        "values": "=%s!$O$%d:$O$%d" % (d, tf, tl),
        "fill": {"color": th.gold}, "border": {"color": th.gold},
        "gap": 45})
    ch.set_title({"name": "Revenue by event type",
                  "name_font": {"size": 12, "bold": True,
                                "color": th.primary, "name": th.title_font}})
    ch.set_x_axis({"num_format": "#,##0", "num_font": {"size": 9},
                   "major_gridlines": {"visible": True,
                                       "line": {"color": th.border}}})
    ch.set_y_axis({"num_font": {"size": 8}, "label_position": "low"})
    ch.set_size({"width": 555, "height": 285})
    ch.set_legend({"none": True})
    ws.insert_chart(r(ROW_CHART_2), ci("H"), ch, {"x_offset": 8,
                                                  "y_offset": 6})


def _panels(bk):
    ws, S, th, m = bk.ws(KEY), bk.S, bk.th, bk.demo
    d = bk.q("data")
    row = ROW_PANEL_SEC
    ws.set_row(r(row), 22)
    ws.merge_range(r(row), 1, r(row), 5,
                   "  \U0001F4C5  NEXT EVENTS", S.section_soft)
    ws.merge_range(r(row), 6, r(row), ci(LAST_COL),
                   "  \U0001F4B5  MONEY COMING IN", S.section_soft)

    def mk(**props):
        base = dict(font_size=10, font_color=th.ink, bg_color=th.card,
                    align="left", valign="vcenter", border=1,
                    border_color=th.border, indent=1)
        base.update(props)
        alt_props = dict(base)
        alt_props["bg_color"] = th.alt
        return (S.f(**S.base(**base)), S.f(**S.base(**alt_props)))

    t_fmts = mk()
    c_fmts = mk(align="center", bold=True)
    m_fmts = mk(align="right", bold=True, font_color=th.primary,
                num_format="#,##0.00")

    upcoming = (m.agg.get("upcoming") or [])[:PANEL_ROWS]
    dues = (m.agg.get("dues") or [])[:PANEL_ROWS]

    for i in range(PANEL_ROWS):
        prow = ROW_PANEL + i
        ws.set_row(r(prow), 20)
        a = 1 if i % 2 else 0
        drow = C.DATA_POOL_FIRST + i
        due_row = C.DATA_DUE_FIRST + i

        # left panel: date | event | stage
        ws.write_formula(
            r(prow), 1,
            '=IF(%s!$AA$%d="","\u2014",TEXT(%s!$AA$%d,"ddd dd mmm"))'
            % (d, drow, d, drow), c_fmts[a],
            upcoming[i][0].strftime("%a %d %b") if i < len(upcoming)
            else "\u2014")
        bk.stats["formulas"] += 1
        ws.merge_range(r(prow), 2, r(prow), 3, "", t_fmts[a])
        ws.write_formula(
            r(prow), 2,
            '=IF(%s!$AB$%d="","(nothing booked yet)",%s!$AB$%d)'
            % (d, drow, d, drow), t_fmts[a],
            upcoming[i][1] if i < len(upcoming) else "(nothing booked yet)")
        bk.stats["formulas"] += 1
        ws.merge_range(r(prow), 4, r(prow), 5, "", c_fmts[a])
        ws.write_formula(
            r(prow), 4, '=IF(%s!$AC$%d="","",%s!$AC$%d)'
            % (d, drow, d, drow), c_fmts[a],
            upcoming[i][2] if i < len(upcoming) else "")
        bk.stats["formulas"] += 1

        # right panel: due | invoice | balance
        ws.write_formula(
            r(prow), 6,
            '=IF(%s!$AA$%d="","\u2014",TEXT(%s!$AA$%d,"dd mmm"))'
            % (d, due_row, d, due_row), c_fmts[a],
            dues[i][0].strftime("%d %b") if i < len(dues) else "\u2014")
        bk.stats["formulas"] += 1
        ws.merge_range(r(prow), 7, r(prow), 9, "", t_fmts[a])
        ws.write_formula(
            r(prow), 7,
            '=IF(%s!$AB$%d="","(all square \u2014 nothing owed)",%s!$AB$%d)'
            % (d, due_row, d, due_row), t_fmts[a],
            dues[i][1] if i < len(dues)
            else "(all square \u2014 nothing owed)")
        bk.stats["formulas"] += 1
        ws.merge_range(r(prow), 10, r(prow), ci(LAST_COL), "", m_fmts[a])
        ws.write_formula(
            r(prow), 10,
            '=IF(%s!$AA$%d="","",%s!$AC$%d)' % (d, due_row, d, due_row),
            m_fmts[a], dues[i][2] if i < len(dues) else "")
        bk.stats["formulas"] += 1

        # overdue rows glow red
        bk.cond(KEY, prow, 6, prow, ci(LAST_COL), {
            "type": "formula",
            "criteria": '=AND(%s!$AA$%d<>"",%s!$AA$%d<TODAY())'
                       % (d, due_row, d, due_row),
            "format": S.cf(bg=th.bad_soft, fg=th.bad, bold=True,
                           border=th.border)})
        # events inside the due-soon window glow amber
        bk.cond(KEY, prow, 1, prow, 5, {
            "type": "formula",
            "criteria": '=AND(%s!$AA$%d<>"",%s!$AA$%d-TODAY()<=DueSoonDays)'
                       % (d, drow, d, drow),
            "format": S.cf(bg=th.warn_soft, fg=th.warn, bold=True,
                           border=th.border)})
