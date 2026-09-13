"""_Data - the hidden engine room: pools, KPI cells and chart fuel."""

from .. import config as C
from ..book import r, ci


def build(bk):
    key = "data"
    ws = bk.ws(key)
    for col, width in (("A", 3), ("H", 9), ("I", 11), ("J", 11), ("K", 11),
                       ("L", 11), ("M", 9), ("N", 9), ("O", 9),
                       ("P", 16), ("Q", 11), ("S", 17), ("T", 11),
                       ("V", 24), ("W", 9), ("X", 11), ("Y", 9), ("Z", 9),
                       ("AA", 9), ("AB", 11), ("AC", 24), ("AD", 10),
                       ("AE", 11), ("AF", 9), ("AG", 12), ("AH", 28),
                       ("AI", 12), ("AK", 9), ("AL", 24), ("AN", 9),
                       ("AO", 24)):
        ws.set_column(ci(col), ci(col), width)
    S, th, m = bk.S, bk.th, bk.demo
    agg = m.agg if m else {}
    premium = bk.has("materials")

    hdr = S.f(**S.base(bold=True, font_size=9, font_color=th.muted,
                       bg_color=th.bg, align="left", valign="vcenter"))
    txt = S.f(**S.base(font_size=9, font_color=th.ink, bg_color=th.bg))
    num = S.f(**S.base(font_size=9, font_color=th.ink, bg_color=th.bg,
                       num_format="#,##0.00"))
    integer = S.f(**S.base(font_size=9, font_color=th.ink, bg_color=th.bg,
                           num_format="0"))

    def q(k):
        return bk.q(k)

    sales = q("sales") if bk.has("sales") else None
    events = q("events") if bk.has("events") else None
    catalog = q("catalog")
    prod = q("production") if bk.has("production") else None
    mats = q("materials") if premium else None
    pricing = q("pricing")

    n = [0]

    def fcount():
        n[0] += 1
        bk.stats["formulas"] += 1

    # ------------------------------------------------------------------
    # month table H..N rows 2-13
    # ------------------------------------------------------------------
    for col, label in (("H", "Month"), ("I", "Revenue"), ("J", "COGS"),
                       ("K", "Event fees"), ("L", "Net"), ("M", "Units"),
                       ("N", "Markets"), ("O", "Sales lines")):
        ws.write(r(1), ci(col), label, hdr)
    months = (agg.get("months") or [])
    for i in range(12):
        row = C.DATA_MONTH_FIRST + i
        mon = i + 1
        ws.write(r(row), ci("H"), C.MONTH_NAMES[i], txt)
        d0 = "DATE(ReportYear,%d,1)" % mon
        d1 = "EDATE(DATE(ReportYear,%d,1),1)" % mon
        cached = months[i] if i < len(months) else None
        if sales:
            ws.write_formula(
                r(row), ci("I"),
                '=SUMIFS(%s,%s,">="&%s,%s,"<"&%s)'
                % (bk.rng("sales", "total"), bk.rng("sales", "date"), d0,
                   bk.rng("sales", "date"), d1),
                num, cached[1] if cached else 0); fcount()
            ws.write_formula(
                r(row), ci("J"),
                '=SUMIFS(%s,%s,">="&%s,%s,"<"&%s)'
                % (bk.rng("sales", "cost"), bk.rng("sales", "date"), d0,
                   bk.rng("sales", "date"), d1),
                num, cached[2] if cached else 0); fcount()
            ws.write_formula(
                r(row), ci("M"),
                '=SUMIFS(%s,%s,">="&%s,%s,"<"&%s)'
                % (bk.rng("sales", "qty"), bk.rng("sales", "date"), d0,
                   bk.rng("sales", "date"), d1),
                integer, cached[5] if cached else 0); fcount()
            ws.write_formula(
                r(row), ci("O"),
                '=COUNTIFS(%s,">="&%s,%s,"<"&%s)'
                % (bk.rng("sales", "date"), d0, bk.rng("sales", "date"),
                   d1),
                integer, cached[7] if cached and len(cached) > 7 else 0)
            fcount()
        else:
            for col in ("I", "J", "M"):
                ws.write(r(row), ci(col), 0, num)
        if events:
            ws.write_formula(
                r(row), ci("K"),
                '=SUMIFS(%s,%s,">="&%s,%s,"<"&%s)'
                % (bk.rng("events", "total"), bk.rng("events", "date"), d0,
                   bk.rng("events", "date"), d1),
                num, cached[3] if cached else 0); fcount()
            ws.write_formula(
                r(row), ci("N"),
                '=COUNTIFS(%s,">="&%s,%s,"<"&%s)'
                % (bk.rng("events", "date"), d0, bk.rng("events", "date"),
                   d1),
                integer, cached[6] if cached else 0); fcount()
        else:
            ws.write(r(row), ci("K"), 0, num)
            ws.write(r(row), ci("N"), 0, integer)
        ws.write_formula(r(row), ci("L"), "=$I%d-$J%d-$K%d" % (row, row, row),
                         num, cached[4] if cached else 0); fcount()

    # ------------------------------------------------------------------
    # payment methods P..Q rows 2-7
    # ------------------------------------------------------------------
    ws.write(r(1), ci("P"), "Method", hdr)
    ws.write(r(1), ci("Q"), "Taken", hdr)
    pays = agg.get("payments") or []
    for i, method in enumerate(C.PAYMENT_METHODS):
        row = C.DATA_PAY_FIRST + i
        ws.write(r(row), ci("P"), method, txt)
        if sales:
            ws.write_formula(r(row), ci("Q"),
                             "=SUMIFS(%s,%s,$P%d)"
                             % (bk.rng("sales", "total"),
                                bk.rng("sales", "method"), row),
                             num, pays[i][1] if i < len(pays) else 0); fcount()
        else:
            ws.write(r(row), ci("Q"), 0, num)

    # ------------------------------------------------------------------
    # event expense categories S..T rows 2-6
    # ------------------------------------------------------------------
    ws.write(r(1), ci("S"), "Category", hdr)
    ws.write(r(1), ci("T"), "Spent", hdr)
    exp_cols = (("Booth fees", "booth"), ("Travel", "travel"),
                ("Parking", "parking"), ("Food & drinks", "food"),
                ("Display & decor", "display"))
    exps = dict(agg.get("exp_cats") or [])
    for i, (label, field) in enumerate(exp_cols):
        row = C.DATA_EXP_FIRST + i
        ws.write(r(row), ci("S"), label, txt)
        if events:
            ws.write_formula(r(row), ci("T"), "=SUM(%s)"
                             % bk.rng("events", field),
                             num, exps.get(label, 0)); fcount()
        else:
            ws.write(r(row), ci("T"), 0, num)

    # ------------------------------------------------------------------
    # product pool V..AA rows 2-25
    # ------------------------------------------------------------------
    for col, label in (("V", "Product"), ("W", "Units"), ("X", "Revenue"),
                       ("Y", "In stock"), ("Z", "Margin adj"),
                       ("AA", "Slow adj"), ("AB", "Profit")):
        ws.write(r(1), ci(col), label, hdr)
    stats = agg.get("product_stats") or []
    last = C.ROW_FIRST + C.CAP["catalog"] - 1
    for i in range(C.CAP["catalog"]):
        row = C.DATA_PROD_FIRST + i
        crow = C.ROW_FIRST + i
        ws.write_formula(r(row), ci("V"),
                         '=IF(%s!$C$%d="","",%s!$C$%d)'
                         % (catalog, crow, catalog, crow),
                         txt, stats[i][0] if i < len(stats) else ""); fcount()
        if sales:
            ws.write_formula(r(row), ci("W"),
                             '=IF($V%d="",0,SUMIFS(%s,%s,$V%d))'
                             % (row, bk.rng("sales", "qty"),
                                bk.rng("sales", "product"), row),
                             integer, stats[i][1] if i < len(stats) else 0)
            fcount()
            ws.write_formula(r(row), ci("X"),
                             '=IF($V%d="",0,SUMIFS(%s,%s,$V%d))'
                             % (row, bk.rng("sales", "total"),
                                bk.rng("sales", "product"), row),
                             num, stats[i][2] if i < len(stats) else 0)
            fcount()
        else:
            ws.write(r(row), ci("W"), 0, integer)
            ws.write(r(row), ci("X"), 0, num)
        ws.write_formula(r(row), ci("Y"), '=IF($V%d="","",%s!$M$%d)'
                         % (row, catalog, crow),
                         integer, stats[i][3] if i < len(stats) else "")
        fcount()
        ws.write_formula(r(row), ci("Z"),
                         '=IF($W%d>0,%s!$K$%d,0)' % (row, catalog, crow),
                         S.f(**S.base(font_size=9, font_color=th.ink,
                                      bg_color=th.bg, num_format="0.0%")),
                         stats[i][4] if i < len(stats) else 0); fcount()
        ws.write_formula(r(row), ci("AA"), '=IF($V%d="","",$W%d)'
                         % (row, row),
                         integer, stats[i][5] if i < len(stats) else "")
        fcount()
        ws.write_formula(r(row), ci("AB"),
                         '=IFERROR(ROUND($W%d*%s!$J$%d,2),0)' % (row, catalog, crow),
                         num, round(stats[i][1] * _profit_of(
                             m, stats[i][0]), 2) if i < len(stats) else 0)
        fcount()

    # ------------------------------------------------------------------
    # event pool AC..AF rows 2-25
    # ------------------------------------------------------------------
    for col, label in (("AC", "Event"), ("AD", "Sales"), ("AE", "Net"),
                       ("AF", "Units")):
        ws.write(r(1), ci(col), label, hdr)
    ev_stats = agg.get("event_stats") or []
    for i in range(C.CAP["events"]):
        row = C.DATA_EVENT_FIRST + i
        erow = C.ROW_FIRST + i
        ws.write_formula(r(row), ci("AC"),
                         '=IF(%s!$B$%d="","",%s!$B$%d)'
                         % (events, erow, events, erow),
                         txt, ev_stats[i][0] if i < len(ev_stats) else "")
        fcount()
        ws.write_formula(r(row), ci("AD"), '=IF($AC%d="","",%s!$N$%d)'
                         % (row, events, erow),
                         num, ev_stats[i][1] if i < len(ev_stats) else "")
        fcount()
        ws.write_formula(r(row), ci("AE"), '=IF($AC%d="","",%s!$P$%d)'
                         % (row, events, erow),
                         num, ev_stats[i][2] if i < len(ev_stats) else "")
        fcount()
        ws.write_formula(r(row), ci("AF"), '=IF($AC%d="","",%s!$M$%d)'
                         % (row, events, erow),
                         integer, ev_stats[i][3] if i < len(ev_stats) else "")
        fcount()

    # ------------------------------------------------------------------
    # upcoming fair dates AG rows 2-25 (hero line fuel)
    # ------------------------------------------------------------------
    ws.write(r(1), ci("AG"), "Upcoming", hdr)
    for i in range(C.CAP["events"]):
        row = C.DATA_EVENT_FIRST + i
        erow = C.ROW_FIRST + i
        ws.write_formula(r(row), ci("AG"),
                         '=IF(AND(%s!$C$%d<>"",%s!$C$%d>=TODAY()),%s!$C$%d,"")'
                         % (events, erow, events, erow, events, erow),
                         txt, "")
        fcount()

    # ------------------------------------------------------------------
    # KPI cells AH/AI
    # ------------------------------------------------------------------
    ws.write(r(1), ci("AH"), "KPI", hdr)
    ws.write(r(1), ci("AI"), "Value", hdr)
    F = {}
    mr = C.DATA_MONTH_FIRST
    ml = C.DATA_MONTH_FIRST + 11
    F["revenue"] = ("=SUM($I$%d:$I$%d)" % (mr, ml), "#,##0.00")
    F["discounts"] = ("=SUM(%s)" % bk.rng("sales", "discount"), "#,##0.00")
    F["cogs"] = ("=SUMPRODUCT(IFERROR(%s*1,0))" % bk.rng("sales", "cost"),
                 "#,##0.00")
    F["fees"] = ('=SUMIFS(%s,%s,"<="&TODAY())'
                 % (bk.rng("events", "total"), bk.rng("events", "date")),
                 "#,##0.00")
    F["profit"] = ("=$AI$%d-$AI$%d-$AI$%d"
                   % (C.KPI_ROW["revenue"], C.KPI_ROW["cogs"],
                      C.KPI_ROW["fees"]), "#,##0.00")
    F["margin"] = ('=IFERROR($AI$%d/$AI$%d,0)'
                   % (C.KPI_ROW["profit"], C.KPI_ROW["revenue"]), "0.0%")
    F["units"] = ("=SUM(%s)" % bk.rng("sales", "qty"), "0")
    F["txns"] = ("=COUNT(%s)" % bk.rng("sales", "total"), "0")
    F["aov"] = ('=IFERROR($AI$%d/$AI$%d,0)'
                % (C.KPI_ROW["revenue"], C.KPI_ROW["txns"]), "#,##0.00")
    F["per_item"] = ('=IFERROR($AI$%d/$AI$%d,0)'
                     % (C.KPI_ROW["profit"], C.KPI_ROW["units"]),
                     "#,##0.00")
    F["events_total"] = ("=COUNTA(%s)" % bk.rng("events", "name"), "0")
    F["events_done"] = ('=COUNTIFS(%s,"<="&TODAY(),%s,">0")'
                        % (bk.rng("events", "date"),
                           bk.rng("events", "sales")), "0")
    F["inv_value"] = ("=SUMPRODUCT(IFERROR(%s*1,0)*IFERROR(%s*1,0))"
                      % (bk.rng("catalog", "stock"),
                         bk.rng("catalog", "unit_cost")), "#,##0.00")
    F["low_products"] = ('=COUNTIF(%s,"%s")+COUNTIF(%s,"%s")'
                         % (bk.rng("catalog", "status"), C.ST_LOW,
                            bk.rng("catalog", "status"), C.ST_OUT), "0")
    F["low_materials"] = ('=COUNTIF(%s,"%s")'
                          % (bk.rng("materials", "reorder"), C.RE_YES)
                          if premium else "=0", "0")
    ro = bk.q("reorder")
    if premium:
        F["reorder_cost"] = ("=SUM(%s!$G$%d:$G$%d)+SUM(%s!$G$%d:$G$%d)"
                             % (ro, C.RO_PROD_FIRST, C.RO_PROD_LAST, ro,
                                C.RO_MAT_FIRST, C.RO_MAT_LAST),
                             "#,##0.00")
    else:
        F["reorder_cost"] = (
            "=SUMPRODUCT((%s<=%s)*IFERROR((%s*2-%s)*%s,0))"
            % (bk.rng("catalog", "stock"), bk.rng("catalog", "min"),
               bk.rng("catalog", "min"), bk.rng("catalog", "stock"),
               bk.rng("catalog", "unit_cost")),
            "#,##0.00")
    F["sell_through"] = ('=IFERROR($AI$%d/$AI$%d,0)'
                         % (C.KPI_ROW["units"], C.KPI_ROW["made_total"]),
                         "0.0%")
    F["made_total"] = ("=SUM(%s)" % bk.rng("production", "made"), "0")
    pr = C.DATA_PROD_FIRST
    pl = C.DATA_PROD_FIRST + C.CAP["catalog"] - 1
    F["best_product"] = ('=IFERROR(INDEX($V$%d:$V$%d,MATCH(MAX($X$%d:$X$%d),'
                         '$X$%d:$X$%d,0)),"")' % (pr, pl, pr, pl, pr, pl), "@")
    er = C.DATA_EVENT_FIRST
    el = C.DATA_EVENT_FIRST + C.CAP["events"] - 1
    F["best_event"] = ('=IFERROR(INDEX($AC$%d:$AC$%d,MATCH(MAX($AE$%d:'
                       '$AE$%d),$AE$%d:$AE$%d,0)),"")'
                       % (er, el, er, el, er, el), "@")
    F["bs_units"] = ('=IFERROR(INDEX($V$%d:$V$%d,MATCH(MAX($W$%d:$W$%d),'
                     '$W$%d:$W$%d,0)),"")' % (pr, pl, pr, pl, pr, pl), "@")
    F["bs_revenue"] = F["best_product"]
    F["bs_profit"] = ('=IFERROR(INDEX($V$%d:$V$%d,MATCH(MAX($AB$%d:$AB$%d),'
                      '$AB$%d:$AB$%d,0)),"")' % (pr, pl, pr, pl, pr, pl), "@")
    F["bs_margin"] = ('=IFERROR(INDEX($V$%d:$V$%d,MATCH(MAX($Z$%d:$Z$%d),'
                      '$Z$%d:$Z$%d,0)),"")' % (pr, pl, pr, pl, pr, pl), "@")
    F["bs_slow"] = ('=IFERROR(INDEX($V$%d:$V$%d,MATCH(MIN($AA$%d:$AA$%d),'
                    '$AA$%d:$AA$%d,0)),"")' % (pr, pl, pr, pl, pr, pl), "@")
    F["price_suggest"] = ("=%s!$J$%d" % (pricing, C.PR_OUT["price"]),
                          "#,##0.00")
    for name in ("revenue", "discounts", "cogs", "fees", "profit", "margin",
                 "units", "txns", "aov", "per_item", "events_total",
                 "events_done", "inv_value", "low_products", "low_materials",
                 "reorder_cost", "sell_through", "made_total"):
        if not bk.has("production") and name in ("made_total",
                                                 "sell_through",
                                                 "inv_value"):
            F[name] = ("=0", C.KPI_FMT[name])
    for key, (formula, fmt) in F.items():
        row = C.KPI_ROW[key]
        ws.write(r(row), ci("AH"), C.KPI_LABEL[key], hdr)
        cached = agg.get(key, 0 if fmt != "@" else "")
        ws.write_formula(r(row), ci("AI"), formula,
                         S.f(**S.base(font_size=9, font_color=th.ink,
                                      bg_color=th.bg, num_format=fmt)),
                         cached)
        fcount()

    # ------------------------------------------------------------------
    # reorder sequence pools AK/AL (products) AN/AO (materials)
    # ------------------------------------------------------------------
    ws.write(r(1), ci("AK"), "Prod seq", hdr)
    ws.write(r(1), ci("AL"), "Prod row", hdr)
    ws.write(r(1), ci("AN"), "Mat seq", hdr)
    ws.write(r(1), ci("AO"), "Mat row", hdr)
    st = bk.rng("catalog", "status")
    seq_cached_p = _seq_cache([srow[3] < _min_of(m, srow[0])
                               for srow in (agg.get("product_stats") or [])],
                              m)
    for i in range(C.CAP["catalog"]):
        row = C.DATA_SEQP_FIRST + i
        crow = C.ROW_FIRST + i
        flag = 'OR(%s!$O$%d="%s",%s!$O$%d="%s")' % (catalog, crow, C.ST_LOW,
                                                   catalog, crow, C.ST_OUT)
        if i == 0:
            body = "1"
        else:
            body = ('1+SUMPRODUCT((%s!$O$%d:$O$%d="%s")+(%s!$O$%d:$O$%d="%s"))'
                    % (catalog, C.ROW_FIRST, crow - 1, C.ST_LOW,
                       catalog, C.ROW_FIRST, crow - 1, C.ST_OUT))
        ws.write_formula(r(row), ci("AK"),
                         '=IF(%s,%s,"")' % (flag, body),
                         integer, seq_cached_p[i] if i < len(seq_cached_p)
                         else "")
        fcount()
        ws.write(r(row), ci("AL"), crow, integer)
    if premium:
        seq_cached_m = _seq_cache(
            [mrow[1] <= mrow[2] for mrow in
             [(mm["remaining"], mm["remaining"], mm["threshold"])
              for mm in (m.materials if m else [])]], None)
        for i in range(C.CAP["materials"]):
            row = C.DATA_SEQM_FIRST + i
            mrow = C.ROW_FIRST + i
            flag = 'AND(%s!$B$%d<>"",%s!$K$%d<=%s!$M$%d)' % (mats, mrow,
                                                            mats, mrow,
                                                            mats, mrow)
            if i == 0:
                body = "1"
            else:
                body = ('1+SUMPRODUCT((%s!$B$%d:$B$%d<>"")*'
                        '(%s!$K$%d:$K$%d<=%s!$M$%d:$M$%d))'
                        % (mats, C.ROW_FIRST, mrow - 1, mats,
                           C.ROW_FIRST, mrow - 1, mats, C.ROW_FIRST,
                           mrow - 1))
            ws.write_formula(r(row), ci("AN"),
                             '=IF(%s,%s,"")' % (flag, body),
                             integer, seq_cached_m[i]
                             if i < len(seq_cached_m) else "")
            fcount()
            ws.write(r(row), ci("AO"), mrow, integer)
    return None


def _seq_cache(flags, _):
    out, k = [], 0
    for f in flags:
        if f:
            k += 1
            out.append(k)
        else:
            out.append("")
    return out


def _min_of(demo, name):
    if not demo:
        return 0
    for p in demo.products:
        if p["name"] == name:
            return p["min"]
    return 0


def _profit_of(demo, name):
    if not demo:
        return 0.0
    for p in demo.products:
        if p["name"] == name:
            return p["profit"]
    return 0.0
