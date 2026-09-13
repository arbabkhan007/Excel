"""📊 Dashboard - KPIs, bars, six charts, best-sellers and low-stock panels."""

from .. import config as C
from ..book import r, ci

LAST_COL = "M"
ROW_HERO_1 = 2
ROW_HERO_2 = 3
ROW_HERO_3 = 4
ROW_MSG = 5
ROW_SECTION = 7
ROW_CARD1_L = 8
ROW_CARD1_V = 9
ROW_CARD2_L = 10
ROW_CARD2_V = 11
ROW_BAR1 = 12
ROW_BAR2 = 13
ROW_CH1 = 16
ROW_CH2 = 32
ROW_CH3 = 48
ROW_PANELS = 65
ROW_FOOT = 73


def build(bk):
    key = "dashboard"
    ws = bk.ws(key)
    S, th, m = bk.S, bk.th, bk.demo
    agg = m.agg if m else {}
    premium = bk.has("materials")

    bk.widths(key, C.WIDTHS[key])
    bk.paint(key, 0, 0, ROW_FOOT + 6, ci(LAST_COL), S.canvas)

    # ------------------------------------------------------------------
    # hero
    # ------------------------------------------------------------------
    ws.set_row(r(1), 8)
    ws.set_row(r(ROW_HERO_1), 34)
    ws.set_row(r(ROW_HERO_2), 26)
    ws.set_row(r(ROW_HERO_3), 20)
    ws.set_row(r(ROW_MSG), 22)

    ws.merge_range(r(ROW_HERO_1), 1, r(ROW_HERO_1), ci(LAST_COL), "",
                   S.hero_title)
    ws.write_formula(
        r(ROW_HERO_1), 1,
        '=IF(BusinessName="","\U0001F9F6  Craft Fair Command Center",'
        '"\U0001F9F6  "&BusinessName&"   \u2022   Craft Fair Command '
        'Center")', S.hero_title,
        "\U0001F9F6  Craft Fair Command Center"
        if not (m and m.settings["business"]) else
        "\U0001F9F6  %s   \u2022   Craft Fair Command Center"
        % m.settings["business"])
    bk.stats["formulas"] += 1

    ws.merge_range(r(ROW_HERO_2), 1, r(ROW_HERO_2), ci(LAST_COL), "",
                   S.hero_count)
    nxt = _next_event(bk)
    ws.write_formula(r(ROW_HERO_2), 1, "=" + nxt, S.hero_count,
                     _next_event_cached(m))
    bk.stats["formulas"] += 1

    ws.merge_range(r(ROW_HERO_3), 1, r(ROW_HERO_3), ci(LAST_COL), "",
                   S.hero_meta)
    ws.write_formula(
        r(ROW_HERO_3), 1,
        '=TEXT(TODAY(),"dddd, dd mmmm yyyy")&"   \u2022   "&%s&" fairs on '
        'the books   \u2022   "&%s&" taken at the till   \u2022   "&%s&" '
        'net profit so far"'
        % (bk.kpi("events_total"), bk.money(bk.kpi("revenue"), "#,##0"),
           bk.money(bk.kpi("profit"), "#,##0")),
        S.hero_meta,
        "   \u2022   %d fairs on the books   \u2022   %s taken at the till"
        "   \u2022   %s net profit so far"
        % (agg.get("events_total", 0), agg.get("revenue", 0),
           agg.get("profit", 0)) if m else "")
    bk.stats["formulas"] += 1

    ws.merge_range(r(ROW_MSG), 1, r(ROW_MSG), ci(LAST_COL), "",
                   S.f(**S.base(bg_color=th.gold, font_color=th.white,
                                bold=True, font_size=11, align="left",
                                valign="vcenter", indent=1)))
    ws.write_formula(
        r(ROW_MSG), 1,
        '=IF(%s!$C$%d="","","\U0001F4E3  "&%s!$C$%d)'
        % (bk.q("setup"), C.SU_MESSAGE, bk.q("setup"), C.SU_MESSAGE),
        S.f(**S.base(bg_color=th.gold, font_color=th.white, bold=True,
                     font_size=11, align="left", valign="vcenter",
                     indent=1)),
        "\U0001F4E3  " + m.settings["message"] if m else "")
    bk.stats["formulas"] += 1

    # ------------------------------------------------------------------
    # KPI cards
    # ------------------------------------------------------------------
    ws.merge_range(r(ROW_SECTION), 1, r(ROW_SECTION), ci(LAST_COL),
                   "  \U0001F9F6  BUSINESS AT A GLANCE", S.section)
    ws.set_row(r(ROW_SECTION), 22)

    cards1 = [
        ("TOTAL SALES", "ok", bk.money(bk.kpi("revenue"), "#,##0"),
         agg.get("revenue", 0)),
        ("NET PROFIT", "primary", bk.money(bk.kpi("profit"), "#,##0"),
         agg.get("profit", 0)),
        ("PROFIT MARGIN", "gold", "=TEXT(%s,\u00220%%\u0022)"
         % bk.kpi("margin"), _p0(agg.get("margin", 0))),
        ("UNITS SOLD", "info", "=%s" % bk.kpi("units"),
         agg.get("units", 0)),
        ("FAIRS WORKED", "primary_2", "=%s" % bk.kpi("events_done"),
         agg.get("events_done", 0)),
        ("AVERAGE SALE", "accent", bk.money(bk.kpi("aov"), "#,##0.00"),
         agg.get("aov", 0)),
    ]
    cards2 = [
        ("INVENTORY VALUE", "info", bk.money(bk.kpi("inv_value"),
                                             "#,##0.00"),
         agg.get("inv_value", 0)),
        ("LOW-STOCK ITEMS", "bad", "=%s+%s"
         % (bk.kpi("low_products"), bk.kpi("low_materials")),
         agg.get("low_products", 0) + agg.get("low_materials", 0)),
        ("SELL-THROUGH", "gold", "=TEXT(%s,\u00220%%\u0022)"
         % bk.kpi("sell_through"),
         _p0(agg.get("sell_through", 0))),
        ("BEST-SELLER", "ok", "=%s" % bk.kpi("bs_units"),
         agg.get("bs_units", "")),
        ("TOP FAIR", "accent", "=%s" % bk.kpi("best_event"),
         agg.get("best_event", "")),
        ("REORDER BUDGET", "bad", bk.money(bk.kpi("reorder_cost"),
                                           "#,##0.00"),
         agg.get("reorder_cost", 0)),
    ]
    _cards(bk, cards1, ROW_CARD1_L, ROW_CARD1_V)
    _cards(bk, cards2, ROW_CARD2_L, ROW_CARD2_V)

    # ------------------------------------------------------------------
    # progress bars
    # ------------------------------------------------------------------
    _bar(bk, ROW_BAR1, "Net margin", bk.kpi("profit"), bk.kpi("revenue"),
         agg.get("margin", 0), "margin")
    _bar(bk, ROW_BAR2, "Sell-through (sold \u00f7 made)", bk.kpi("units"),
         bk.kpi("made_total"), agg.get("sell_through", 0), "pct")

    # ------------------------------------------------------------------
    # charts
    # ------------------------------------------------------------------
    _charts(bk)

    # ------------------------------------------------------------------
    # panels: best-sellers + low stock
    # ------------------------------------------------------------------
    _panels(bk)

    # ------------------------------------------------------------------
    # footer + nav
    # ------------------------------------------------------------------
    ws.set_row(r(ROW_FOOT), 30)
    ws.merge_range(r(ROW_FOOT), 1, r(ROW_FOOT), ci(LAST_COL),
                   "  \U0001F4A1  Everything on this page is calculated "
                   "\u2014 type your makes and sales into the tabs and "
                   "watch it update.  Start with the \U0001F4D6 Start Here "
                   "guide and \u2699\uFE0F Lists & Settings.",
                   S.f(**S.base(font_size=10, italic=True, font_color=th.ink,
                                bg_color=th.gold_soft, align="left",
                                valign="vcenter", indent=1)))
    nav = ROW_FOOT + 2
    ws.set_row(r(nav), 24)
    bk.nav_row(key, nav, max_col=LAST_COL)
    bk.page(key, LAST_COL, nav + 1, landscape=True, zoom=85)


# ===========================================================================
def _cards(bk, cards, row_l, row_v):
    S, th = bk.S, bk.th
    ws = bk.ws("dashboard")
    pairs = [(1, 2), (3, 4), (5, 6), (7, 8), (9, 10), (11, 12)]
    for (label, colour, formula, cached), (c1, c2) in zip(cards, pairs):
        bg = getattr(th, colour)
        lbl = S.f(**S.base(bold=True, font_size=9.5, font_color=th.white,
                           bg_color=bg, align="left", valign="vcenter",
                           indent=1))
        val = S.f(**S.base(bold=True, font_size=15,
                           font_color=bg if colour != "gold" else th.gold,
                           bg_color=th.card, align="center",
                           valign="vcenter"))
        ws.merge_range(r(row_l), c1, r(row_l), c2, "  " + label, lbl)
        ws.merge_range(r(row_v), c1, r(row_v), c2, "", val)
        if formula.startswith("="):
            ws.write_formula(r(row_v), c1, formula, val, cached)
        else:
            ws.write_formula(r(row_v), c1, "=" + formula, val, cached)
        bk.stats["formulas"] += 1
    ws.set_row(r(row_l), 18)
    ws.set_row(r(row_v), 26)


def _bar(bk, row, label, num_k, den_k, cached_pct, kind=None):
    S, th = bk.S, bk.th
    ws = bk.ws("dashboard")
    ws.set_row(r(row), 20)
    ws.merge_range(r(row), 1, r(row), 2, "  " + label, S.bar_label)
    ws.merge_range(r(row), 3, r(row), 9, "", S.bar_text)
    ws.write_formula(r(row), 3, bk.bar(num_k, den_k, 34), S.bar_text,
                     _bar_cached(cached_pct, 34))
    bk.stats["formulas"] += 1
    ws.merge_range(r(row), 10, r(row), 12, "", S.bar_pct)
    ws.write_formula(r(row), 10,
                     '=TEXT(IFERROR(%s/%s,0),"0.0%%")' % (num_k, den_k),
                     S.bar_pct, _p1(cached_pct))
    bk.stats["formulas"] += 1


def _p0(x):
    return "%.0f%%" % (100.0 * float(x or 0))


def _p1(x):
    return "%.1f%%" % (100.0 * float(x or 0))


def _bar_cached(pct, blocks):
    filled = int(round(min(1.0, max(0.0, float(pct or 0))) * blocks))
    return "\u2588" * filled + "\u2591" * (blocks - filled)


# ===========================================================================
def _charts(bk):
    S, th = bk.S, bk.th
    wb = bk.wb
    d = bk.q("data").strip("'")
    mr, ml = C.DATA_MONTH_FIRST, C.DATA_MONTH_FIRST + 11
    er, el = C.DATA_EVENT_FIRST, C.DATA_EVENT_FIRST + C.CAP["events"] - 1
    pr, pl = C.DATA_PROD_FIRST, C.DATA_PROD_FIRST + C.CAP["catalog"] - 1

    def base_chart(ctype, title):
        ch = wb.add_chart({"type": ctype})
        ch.set_title({"name": title, "name_font": {"size": 11,
                                                  "color": th.primary,
                                                  "bold": True}})
        ch.set_legend({"position": "bottom",
                       "font": {"size": 9, "color": th.ink}})
        ch.set_chartarea({"border": {"color": th.border},
                          "fill": {"color": th.card}})
        ch.set_plotarea({"fill": {"color": th.card}})
        ch.set_size({"width": 555, "height": 250})
        ch.show_hidden_data()
        bk.stats["charts"] += 1
        return ch

    # 1 - monthly revenue + profit
    combo = base_chart("column", "Monthly revenue & net profit")
    combo.add_series({
        "name": "Revenue",
        "categories": "=%s!$H$%d:$H$%d" % (d, mr, ml),
        "values": "=%s!$I$%d:$I$%d" % (d, mr, ml),
        "fill": {"color": th.ok},
        "border": {"color": th.ok},
    })
    line = wb.add_chart({"type": "line"})
    line.add_series({
        "name": "Net profit",
        "categories": "=%s!$H$%d:$H$%d" % (d, mr, ml),
        "values": "=%s!$L$%d:$L$%d" % (d, mr, ml),
        "line": {"color": th.accent, "width": 2.5},
        "marker": {"type": "circle", "size": 6,
                   "fill": {"color": th.accent}},
    })
    combo.combine(line)
    bk.ws("dashboard").insert_chart(r(ROW_CH1), 2, combo)

    # 2 - sales & profit by event
    ev = base_chart("column", "Sales & net profit by fair")
    ev.add_series({
        "name": "Sales",
        "categories": "=%s!$AC$%d:$AC$%d" % (d, er, el),
        "values": "=%s!$AD$%d:$AD$%d" % (d, er, el),
        "fill": {"color": th.info},
        "border": {"color": th.info},
    })
    ev2 = wb.add_chart({"type": "line"})
    ev2.add_series({
        "name": "Net profit",
        "categories": "=%s!$AC$%d:$AC$%d" % (d, er, el),
        "values": "=%s!$AE$%d:$AE$%d" % (d, er, el),
        "line": {"color": th.gold, "width": 2.5},
        "marker": {"type": "circle", "size": 6,
                   "fill": {"color": th.gold}},
    })
    ev.combine(ev2)
    bk.ws("dashboard").insert_chart(r(ROW_CH1), 8, ev)

    # 3 - units sold by product
    prod = base_chart("bar", "Units sold by product")
    prod.add_series({
        "name": "Units",
        "categories": "=%s!$V$%d:$V$%d" % (d, pr, pl),
        "values": "=%s!$W$%d:$W$%d" % (d, pr, pl),
        "fill": {"color": th.accent},
        "border": {"color": th.accent},
    })
    bk.ws("dashboard").insert_chart(r(ROW_CH2), 2, prod)

    # 4 - inventory remaining
    inv = base_chart("bar", "Inventory remaining (units)")
    inv.add_series({
        "name": "In stock",
        "categories": "=%s!$V$%d:$V$%d" % (d, pr, pl),
        "values": "=%s!$Y$%d:$Y$%d" % (d, pr, pl),
        "fill": {"color": th.primary_2},
        "border": {"color": th.primary_2},
    })
    bk.ws("dashboard").insert_chart(r(ROW_CH2), 8, inv)

    # 5 - takings by payment method
    pay = base_chart("doughnut", "Takings by payment method")
    pay.add_series({
        "name": "Taken",
        "categories": "=%s!$P$%d:$P$%d" % (d, C.DATA_PAY_FIRST,
                                           C.DATA_PAY_FIRST + 5),
        "values": "=%s!$Q$%d:$Q$%d" % (d, C.DATA_PAY_FIRST,
                                       C.DATA_PAY_FIRST + 5),
    })
    bk.ws("dashboard").insert_chart(r(ROW_CH3), 2, pay)

    # 6 - fair expenses by category
    exp = base_chart("doughnut", "Fair expenses by category")
    exp.add_series({
        "name": "Spent",
        "categories": "=%s!$S$%d:$S$%d" % (d, C.DATA_EXP_FIRST,
                                           C.DATA_EXP_FIRST + 4),
        "values": "=%s!$T$%d:$T$%d" % (d, C.DATA_EXP_FIRST,
                                       C.DATA_EXP_FIRST + 4),
    })
    bk.ws("dashboard").insert_chart(r(ROW_CH3), 8, exp)


# ===========================================================================
def _panels(bk):
    S, th, m = bk.S, bk.th, bk.demo
    ws = bk.ws("dashboard")
    agg = m.agg if m else {}
    ws.set_row(r(ROW_PANELS), 22)
    ws.merge_range(r(ROW_PANELS), 1, r(ROW_PANELS), 6,
                   "  \U0001F3C6  BEST-SELLERS", S.section_soft)
    ws.merge_range(r(ROW_PANELS), 7, r(ROW_PANELS), ci(LAST_COL),
                   "  \U0001F534  LOW STOCK RIGHT NOW", S.section_soft)

    left = [("Most units sold", "bs_units"),
            ("Highest revenue", "bs_revenue"),
            ("Highest profit", "bs_profit"),
            ("Highest margin", "bs_margin"),
            ("Slowest mover", "bs_slow")]
    lab = S.f(**S.base(font_size=10, font_color=th.muted, bg_color=th.card,
                       align="left", valign="vcenter", indent=1,
                       border=1, border_color=th.border))
    val = S.f(**S.base(bold=True, font_size=10.5, font_color=th.primary,
                       bg_color=th.card, align="left", valign="vcenter",
                       indent=1, border=1, border_color=th.border))
    for i, (label, kpi) in enumerate(left):
        row = ROW_PANELS + 1 + i
        ws.set_row(r(row), 18)
        ws.merge_range(r(row), 1, r(row), 2, "  " + label, lab)
        ws.merge_range(r(row), 3, r(row), 6, "", val)
        ws.write_formula(r(row), 3, "=%s" % bk.kpi(kpi), val,
                         agg.get(kpi, ""))
        bk.stats["formulas"] += 1

    # right: low stock preview (products then materials)
    rk = S.f(**S.base(font_size=10, font_color=th.ink, bg_color=th.card,
                      align="left", valign="vcenter", indent=1,
                      border=1, border_color=th.border))
    rv = S.f(**S.base(bold=True, font_size=10.5, font_color=th.bad,
                      bg_color=th.card, align="center", valign="vcenter",
                      border=1, border_color=th.border))
    cat = bk.q("catalog")
    seqp = "$AK$%d:$AK$%d" % (C.DATA_SEQP_FIRST,
                              C.DATA_SEQP_FIRST + C.CAP["catalog"] - 1)
    for i in range(5):
        row = ROW_PANELS + 1 + i
        k = i + 1
        ws.merge_range(r(row), 7, r(row), 10, "", rk)
        ws.write_formula(
            r(row), 7,
            '=IFERROR(INDEX(%s!$C$%d:$C$%d,MATCH(%d,%s,0)),"")'
            % (cat, C.ROW_FIRST, C.last_row("catalog"), k,
               bk.q("data") + "!" + seqp),
            rk, "")
        bk.stats["formulas"] += 1
        ws.merge_range(r(row), 11, r(row), ci(LAST_COL), "", rv)
        ws.write_formula(
            r(row), 11,
            '=IFERROR(INDEX(%s!$M$%d:$M$%d,MATCH(%d,%s,0))&" left (min "&'
            'INDEX(%s!$N$%d:$N$%d,MATCH(%d,%s,0))&")","")'
            % (cat, C.ROW_FIRST, C.last_row("catalog"), k,
               bk.q("data") + "!" + seqp, cat, C.ROW_FIRST,
               C.last_row("catalog"), k, bk.q("data") + "!" + seqp),
            rv, "")
        bk.stats["formulas"] += 1


# ===========================================================================
def _next_event(bk):
    ev = bk.q("events")
    d = bk.q("data")
    pool = "%s!$AG$%d:$AG$%d" % (d, C.DATA_EVENT_FIRST,
                                 C.DATA_EVENT_FIRST + C.CAP["events"] - 1)
    return ('IFERROR(TEXT(MIN(%s),"dd mmm")&"  \u2014  "&INDEX(%s!$B$%d:'
            '$B$%d,MATCH(MIN(%s),%s!$C$%d:$C$%d,0))&"  is next \u2014  '
            'pack the van!","\U0001F334  No upcoming fairs logged \u2014 '
            'add one on the \U0001F3EA Craft Fairs tab!")') % (
        pool, ev, C.ROW_FIRST, C.last_row("events"), pool,
        ev, C.ROW_FIRST, C.last_row("events"))


def _next_event_cached(m):
    if not m:
        return ""
    import datetime
    future = [e for e in m.events if e["date"] >= datetime.date(2026, 9, 13)]
    if not future:
        return ("\U0001F334  No upcoming fairs logged \u2014 add one on the "
                "\U0001F3EA Craft Fairs tab!")
    nxt = min(future, key=lambda e: e["date"])
    return ("%s  \u2014  %s  is next \u2014  pack the van!"
            % (nxt["date"].strftime("%d %b"), nxt["name"]))
