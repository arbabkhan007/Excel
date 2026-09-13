"""
The hidden ``_Data`` worksheet - the calculation engine.

Layout (see config.py):

  H..L   rows  2-13   monthly pool  (month, cash in, expenses, profit, events)
  N..O   rows  2-12   expense-category pool   (pie chart)
  N..O   rows 16-24   revenue by event type   (bar chart)
  N..O   rows 28-33   event status counts     (doughnut chart)
  Q..R   rows 26-45   revenue by client
  T..U   rows 26-45   menu popularity (times a dish appears on an event menu)
  AA..AC rows  2-41   upcoming-events pool    (dashboard strip)
  AA..AC rows 42-101  outstanding-payments pool (dashboard strip)
  AE..AF rows  2-36   the KPI table - every headline number in the workbook

Keeping the maths on one hidden sheet keeps the visible tabs clean and the
dashboard formulas readable (``=_Data!$AF$5`` instead of a nine-line
SUMPRODUCT), and gives the charts contiguous, pre-computed ranges.
"""

from .. import config as C
from ..book import r, ci

TODAY = "TODAY()"


def _month_start(m):
    return "DATE(ReportYear,%d,1)" % m


def _month_end(m):
    return "DATE(ReportYear,%d,1)" % (m + 1) if m < 12 \
        else "DATE(ReportYear+1,1,1)"


def build(bk):
    ws = bk.ws("data")
    S, th = bk.S, bk.th
    agg = (bk.demo.agg if bk.demo else {})
    prem = bk.has("inventory")

    hdr = S.f(**S.base(font_size=9, bold=True, font_color=th.white,
                       bg_color=th.muted, align="center", valign="vcenter",
                       border=1, border_color=th.muted))
    txt = S.f(**S.base(font_size=9, font_color=th.ink, align="left"))
    num = S.f(**S.base(font_size=9, num_format="#,##0.00"))
    integer = S.f(**S.base(font_size=9, num_format="0"))
    dt = S.f(**S.base(font_size=9, num_format="dd mmm"))
    pct = S.f(**S.base(font_size=9, num_format="0.0%"))

    ws.set_column("A:A", 14)
    ws.set_column("B:G", 2)
    ws.set_column("H:H", 10)
    ws.set_column("I:L", 12)
    ws.set_column("M:M", 2)
    ws.set_column("N:N", 20)
    ws.set_column("O:O", 12)
    ws.set_column("P:P", 2)
    ws.set_column("Q:Q", 22)
    ws.set_column("R:R", 12)
    ws.set_column("S:S", 2)
    ws.set_column("T:T", 26)
    ws.set_column("U:U", 9)
    ws.set_column("V:Z", 2)
    ws.set_column("AA:AA", 11)
    ws.set_column("AB:AB", 40)
    ws.set_column("AC:AC", 13)
    ws.set_column("AD:AD", 2)
    ws.set_column("AE:AE", 32)
    ws.set_column("AF:AF", 14)

    def fcount():
        bk.stats["formulas"] += 1

    # ------------------------------------------------------------------
    # monthly pool  (H..L rows 2-13) - cash basis, by event month
    # ------------------------------------------------------------------
    inc = bk.q("income")
    exp = bk.q("events")
    inc_recv = bk.rng("income", "received")
    inc_date = bk.rng("income", "event_date")
    inc_amt = bk.rng("income", "amount")
    exp_amt = bk.rng("expenses", "amount")
    exp_date = bk.rng("expenses", "date")
    ev_date = bk.rng("events", "date")
    ev_id = bk.rng("events", "id")

    ws.write(r(1), ci("H"), "Month", hdr)
    for col, label in (("H", "Month"), ("I", "Cash in"), ("J", "Expenses"),
                       ("K", "Profit"), ("L", "Events")):
        ws.write(r(1), ci(col), label, hdr)
    for i in range(12):
        row = C.DATA_MONTH_FIRST + i
        m = i + 1
        cached = agg.get("months", [("", 0, 0, 0, 0)] * 12)[i]
        ws.write(r(row), ci("H"), cached[0] or
                 ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug",
                  "Sep", "Oct", "Nov", "Dec"][i], txt)
        ws.write_formula(
            r(row), ci("I"),
            '=SUMIFS(%s,%s,">="&%s,%s,"<"&%s)'
            % (inc_recv, inc_date, _month_start(m), inc_date, _month_end(m)),
            num, cached[1]); fcount()
        ws.write_formula(
            r(row), ci("J"),
            '=SUMIFS(%s,%s,">="&%s,%s,"<"&%s)'
            % (exp_amt, exp_date, _month_start(m), exp_date, _month_end(m)),
            num, cached[2]); fcount()
        ws.write_formula(r(row), ci("K"), "=I%d-J%d" % (row, row),
                         num, cached[3]); fcount()
        ws.write_formula(
            r(row), ci("L"),
            '=COUNTIFS(%s,">="&%s,%s,"<"&%s,%s,"<>")'
            % (ev_date, _month_start(m), ev_date, _month_end(m), ev_id),
            integer, cached[4]); fcount()

    # ------------------------------------------------------------------
    # expense category pool (N..O rows 2-12) - pie chart
    # ------------------------------------------------------------------
    exp_cat = bk.rng("expenses", "category")
    ws.write(r(1), ci("N"), "Expense category", hdr)
    ws.write(r(1), ci("O"), "Spend", hdr)
    cats = C.EXPENSE_CATEGORIES
    cached_cat = agg.get("exp_cats")
    for i, cat in enumerate(cats):
        row = C.DATA_EXP_FIRST + i
        ws.write(r(row), ci("N"), cat, txt)
        cached = 0.0
        if cached_cat:
            cached = cached_cat.get(cat, 0.0)
        ws.write_formula(
            r(row), ci("O"), '=SUMIF(%s,$N%d,%s)' % (exp_cat, row, exp_amt),
            num, cached); fcount()

    # ------------------------------------------------------------------
    # revenue by event type (N..O rows 16-24)
    # ------------------------------------------------------------------
    ev_type = bk.rng("events", "type")
    ev_price = bk.rng("events", "price")
    ev_status = bk.rng("events", "status")
    ws.write(r(15), ci("N"), "Event type", hdr)
    ws.write(r(15), ci("O"), "Booked value", hdr)
    cached_types = dict(agg.get("types") or [])
    for i, t in enumerate(C.EVENT_TYPES):
        row = C.DATA_TYPE_FIRST + i
        ws.write(r(row), ci("N"), t, txt)
        ws.write_formula(
            r(row), ci("O"),
            '=SUMIFS(%s,%s,$N%d,%s,"<>%s")'
            % (ev_price, ev_type, row, ev_status, C.ES_CANCEL),
            num, cached_types.get(t, 0)); fcount()

    # ------------------------------------------------------------------
    # event status counts (N..O rows 28-33) - doughnut
    # ------------------------------------------------------------------
    ws.write(r(27), ci("N"), "Pipeline stage", hdr)
    ws.write(r(27), ci("O"), "Events", hdr)
    cached_status = agg.get("status_counts")
    for i, st in enumerate(C.EVENT_STATUSES):
        row = C.DATA_STATUS_FIRST + i
        ws.write(r(row), ci("N"), st, txt)
        cached = 0
        if cached_status:
            cached = cached_status.get(st, 0)
        ws.write_formula(
            r(row), ci("O"), '=COUNTIF(%s,$N%d)' % (ev_status, row),
            integer, cached); fcount()

    # ------------------------------------------------------------------
    # revenue by client (Q..R rows 26-45)
    # ------------------------------------------------------------------
    ws.write(r(25), ci("Q"), "Client", hdr)
    ws.write(r(25), ci("R"), "Cash received", hdr)
    ws.write(r(25), ci("T"), "Menu item", hdr)
    ws.write(r(25), ci("U"), "Booked", hdr)
    cl = bk.q("clients")
    cl_first = C.ROW_FIRST
    cached_clients = dict(agg.get("clients_pool") or [])
    cached_menu = dict(agg.get("menu_pool") or [])
    client_names = [c["name"] for c in (bk.demo.clients if bk.demo else [])]
    menu_names = [x["item"] for x in (bk.demo.menu if bk.demo else [])]
    for i in range(20):
        row = C.DATA_CLIENT_FIRST + i
        setup_row = cl_first + i
        cname = client_names[i] if i < len(client_names) else ""
        ws.write_formula(
            r(row), ci("Q"), "=IF(%s!$C$%d=\"\",\"\",%s!$C$%d)"
            % (cl, setup_row, cl, setup_row), txt, cname); fcount()
        ws.write_formula(
            r(row), ci("R"),
            '=IF($Q%d="",0,SUMIF(%s,$Q%d,%s))'
            % (row, bk.rng("income", "client"), row, inc_recv),
            num, cached_clients.get(cname, 0)); fcount()

    # menu popularity (T..U rows 26-45) - premium only
    if bk.has("menu"):
        mn = bk.q("menu")
        ev_menu = bk.rng("events", "menu")
        for i in range(20):
            row = C.DATA_MENU_FIRST + i
            menu_row = cl_first + i
            mname = menu_names[i] if i < len(menu_names) else ""
            ws.write_formula(
                r(row), ci("T"), "=IF(%s!$C$%d=\"\",\"\",%s!$C$%d)"
                % (mn, menu_row, mn, menu_row), txt, mname); fcount()
            ws.write_formula(
                r(row), ci("U"),
                '=IF($T%d="",0,COUNTIF(%s,"*"&$T%d&"*"))'
                % (row, ev_menu, row),
                integer, cached_menu.get(mname, 0)); fcount()
    else:
        for i in range(20):
            row = C.DATA_MENU_FIRST + i
            ws.write(r(row), ci("T"), "", txt)
            ws.write(r(row), ci("U"), "", integer)

    # ------------------------------------------------------------------
    # upcoming-events pool (AA..AC rows 2-41)
    # ------------------------------------------------------------------
    ws.write(r(1), ci("AA"), "Date", hdr)
    ws.write(r(1), ci("AB"), "Event", hdr)
    ws.write(r(1), ci("AC"), "Stage", hdr)
    ws.write(r(41), ci("AA"), "Due", hdr)
    ws.write(r(41), ci("AB"), "Invoice", hdr)
    ws.write(r(41), ci("AC"), "Balance", hdr)
    ev_client = bk.col("events", "client")
    ev_typec = bk.col("events", "type")
    ev_datec = bk.col("events", "date")
    ev_statusc = bk.col("events", "status")
    cached_up = agg.get("upcoming") or []
    for i in range(C.DATA_POOL_ROWS):
        row = C.DATA_POOL_FIRST + i
        er = C.ROW_FIRST + i                     # matching events row
        e = "%s!$%s$%d" % (exp, ev_datec, er)
        st = "%s!$%s$%d" % (exp, ev_statusc, er)
        cl_ref = "%s!$%s$%d" % (exp, ev_client, er)
        ty_ref = "%s!$%s$%d" % (exp, ev_typec, er)
        keep = ('AND(%s<>"",%s>=%s,%s<>"%s",%s<>"%s")'
                % (e, e, TODAY, st, C.ES_CANCEL, st, C.ES_DONE))
        ws.write_formula(r(row), ci("AA"),
                         "=IF(%s,%s,\"\")" % (keep, e), dt,
                         cached_up[i][0] if i < len(cached_up) else ""); fcount()
        ws.write_formula(r(row), ci("AB"),
                         '=IF($AA%d="","",%s&" \u2022 "&%s)' % (row, cl_ref, ty_ref),
                         txt, cached_up[i][1] if i < len(cached_up) else ""); fcount()
        ws.write_formula(r(row), ci("AC"),
                         '=IF($AA%d="","",%s)' % (row, st),
                         txt, cached_up[i][2] if i < len(cached_up) else ""); fcount()

    # outstanding payments pool (AA..AC rows 42-101)
    inc_client = bk.col("income", "client")
    inc_invoice = bk.col("income", "invoice")
    inc_due = bk.col("income", "due")
    inc_bal = bk.col("income", "balance")
    cached_dues = agg.get("dues") or []
    for i in range(C.DATA_DUE_ROWS):
        row = C.DATA_DUE_FIRST + i
        ir = C.ROW_FIRST + i                     # matching income row
        bal = "%s!$%s$%d" % (inc, inc_bal, ir)
        due = "%s!$%s$%d" % (inc, inc_due, ir)
        cli = "%s!$%s$%d" % (inc, inc_client, ir)
        invn = "%s!$%s$%d" % (inc, inc_invoice, ir)
        ws.write_formula(r(row), ci("AA"),
                         '=IF(AND(ISNUMBER(%s),IFERROR(%s*1,0)>0),%s,"")'
                         % (due, bal, due), dt,
                         cached_dues[i][0] if i < len(cached_dues) else ""); fcount()
        ws.write_formula(r(row), ci("AB"),
                         '=IF($AA%d="","",%s&" ("&%s&")")' % (row, cli, invn),
                         txt, cached_dues[i][1] if i < len(cached_dues) else ""); fcount()
        ws.write_formula(r(row), ci("AC"),
                         '=IF($AA%d="","",%s)' % (row, bal),
                         num, cached_dues[i][2] if i < len(cached_dues) else ""); fcount()

    # ------------------------------------------------------------------
    # the KPI table (AE..AF)
    # ------------------------------------------------------------------
    ws.write(r(1), ci(C.KPI_COL_LABEL), "KPI", hdr)
    ws.write(r(1), ci(C.KPI_COL_VALUE), "Value", hdr)
    K = _kpi_formulas(bk, prem)
    for key, label in C.KPI_LABEL.items():
        row = C.KPI_ROW[key]
        ws.write(r(row), ci(C.KPI_COL_LABEL), label, txt)
        formula, cached = K[key]
        fmt = {"#,##0": integer, "#,##0.00": num, "0": integer,
               "0%": pct, "@": txt}[C.KPI_FMT[key]]
        ws.write_formula(r(row), ci(C.KPI_COL_VALUE), formula, fmt,
                         agg.get(key, cached)); fcount()


# ---------------------------------------------------------------------------
# every KPI as (formula, fallback cached value)
# ---------------------------------------------------------------------------
def _kpi_formulas(bk, prem):
    A = lambda k: "$AF$%d" % C.KPI_ROW[k]          # noqa: E731
    inc = bk.q("income")
    recv = bk.rng("income", "received")
    amt = bk.rng("income", "amount")
    bal = bk.rng("income", "balance")
    due = bk.rng("income", "due")
    e_amt = bk.rng("expenses", "amount")
    e_cat = bk.rng("expenses", "category")
    ev = bk.q("events")
    ev_status = bk.rng("events", "status")
    ev_price = bk.rng("events", "price")
    ev_cost = bk.rng("events", "cost")
    ev_guests = bk.rng("events", "guests")
    ev_client = bk.rng("events", "client")
    cl_names = bk.rng("clients", "name")
    q = bk.q("quote")
    F = {}

    F["revenue"] = ("=SUM(%s)" % recv, 0)
    F["expenses"] = ("=SUM(%s)" % e_amt, 0)
    F["profit"] = ("=%s-%s" % (A("revenue"), A("expenses")), 0)
    F["margin"] = ("=IFERROR(%s/%s,0)" % (A("profit"), A("revenue")), 0)
    F["events_total"] = ("=COUNTA(%s)" % bk.rng("events", "id"), 0)
    F["events_done"] = ('=COUNTIF(%s,"%s")' % (ev_status, C.ES_DONE), 0)
    F["events_confirmed"] = (
        '=COUNTIF(%s,"%s")+COUNTIF(%s,"%s")'
        % (ev_status, C.ES_DEPOSIT, ev_status, C.ES_CONFIRMED), 0)
    F["events_cancelled"] = ('=COUNTIF(%s,"%s")' % (ev_status, C.ES_CANCEL), 0)
    F["avg_order"] = (
        '=IFERROR(AVERAGEIF(%s,"<>%s",%s),0)'
        % (ev_status, C.ES_CANCEL, ev_price), 0)
    F["avg_guests"] = (
        '=IFERROR(AVERAGEIF(%s,"<>%s",%s),0)'
        % (ev_status, C.ES_CANCEL, ev_guests), 0)
    F["avg_profit"] = (
        '=IFERROR(AVERAGEIFS(%s,%s,"<>%s"),0)'
        % (bk.rng("events", "profit"), ev_status, C.ES_CANCEL), 0)
    F["guests_total"] = (
        '=SUMIF(%s,"%s",%s)' % (ev_status, C.ES_DONE, ev_guests), 0)
    F["outstanding"] = ("=SUMIF(%s,\">0\")" % bal, 0)
    F["overdue"] = (
        '=SUMPRODUCT((%s>0)*(%s<>"")*(%s<TODAY()))' % (bal, due, due), 0)
    F["invoices_open"] = ('=COUNTIF(%s,">0")' % bal, 0)
    F["food_cost"] = (
        '=SUMIF(%s,"Ingredients",%s)+SUMIF(%s,"Packaging",%s)'
        % (e_cat, e_amt, e_cat, e_amt), 0)
    F["labor_cost"] = ('=SUMIF(%s,"Staff / Labor",%s)' % (e_cat, e_amt), 0)
    F["food_pct"] = ("=IFERROR(%s/%s,0)" % (A("food_cost"), A("revenue")), 0)
    F["labor_pct"] = ("=IFERROR(%s/%s,0)" % (A("labor_cost"), A("revenue")), 0)
    F["repeat_pct"] = (
        '=IFERROR(SUMPRODUCT((%s<>"")*(COUNTIF(%s,%s)>1))'
        '/COUNTIF(%s,"?*"),0)'
        % (cl_names, ev_client, cl_names, cl_names), 0)

    if prem:
        inv_val = bk.rng("inventory", "value")
        inv_status = bk.rng("inventory", "status")
        F["inv_value"] = ("=SUM(%s)" % inv_val, 0)
        F["low_stock"] = ('=COUNTIF(%s,"\U0001F534 Reorder")' % inv_status, 0)
        s_buy = bk.rng("shopping", "to_buy")
        s_tick = bk.rng("shopping", "purchased")
        s_cost = bk.rng("shopping", "est_cost")
        F["shop_lines"] = (
            '=COUNTIFS(%s,">0",%s,"<>%s")' % (s_buy, s_tick, C.TICK), 0)
        F["shop_cost"] = (
            '=SUMIF(%s,"<>%s",%s)' % (s_tick, C.TICK, s_cost), 0)
        st_name = bk.rng("staff", "name")
        st_paid = bk.rng("staff", "paid")
        F["staff_unpaid"] = (
            '=SUMPRODUCT((%s<>"")*(%s<>"%s"))' % (st_name, st_paid, C.TICK), 0)
        F["equip_value"] = ("=SUM(%s)" % bk.rng("equipment", "value"), 0)
        eq_status = bk.rng("equipment", "status")
        F["equip_service"] = (
            '=COUNTIF(%s,"\U0001F527 Service")+COUNTIF(%s,"\U0001F7E1 Soon")'
            % (eq_status, eq_status), 0)
        F["menu_items"] = ("=COUNTA(%s)" % bk.rng("menu", "item"), 0)
        F["menu_margin"] = (
            "=IFERROR(AVERAGE(%s),0)" % bk.rng("menu", "margin"), 0)
        F["quote_price"] = (
            "=%s!$J$%d" % (q, C.QUOTE_OUT["price"]), "")
        F["quote_per_guest"] = (
            "=%s!$J$%d" % (q, C.QUOTE_OUT["per_guest"]), "")
    else:
        for k in ("inv_value", "low_stock", "shop_lines", "shop_cost",
                  "staff_unpaid", "equip_value", "equip_service",
                  "menu_items", "menu_margin"):
            F[k] = ("=0", 0)
        F["quote_price"] = ("=0", 0)
        F["quote_per_guest"] = ("=0", 0)

    F["tax_collected"] = ("=ROUND(SUM(%s)*TaxRate,2)" % amt, 0)
    if bk.has("tax"):
        t = bk.q("tax")
        F["tax_due"] = ("=%s!$G$%d" % (t, C.TAX_SUM["due"]), 0)
    else:
        F["tax_due"] = ("=%s" % A("tax_collected"), 0)

    n_first, n_last = C.DATA_TYPE_FIRST, C.DATA_TYPE_FIRST \
        + len(C.EVENT_TYPES) - 1
    F["best_type"] = (
        '=IF(MAX($O$%d:$O$%d)<=0,"-",INDEX($N$%d:$N$%d,'
        'MATCH(MAX($O$%d:$O$%d),$O$%d:$O$%d,0)))'
        % (n_first, n_last, n_first, n_last, n_first, n_last,
           n_first, n_last), "-")
    c_first, c_last = C.DATA_CLIENT_FIRST, C.DATA_CLIENT_FIRST + 19
    F["best_client"] = (
        '=IF(MAX($R$%d:$R$%d)<=0,"-",INDEX($Q$%d:$Q$%d,'
        'MATCH(MAX($R$%d:$R$%d),$R$%d:$R$%d,0)))'
        % (c_first, c_last, c_first, c_last, c_first, c_last,
           c_first, c_last), "-")
    return F
