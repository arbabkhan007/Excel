"""
\U0001F9EE Quote / Pricing Calculator - the headline selling point.

Type the guest count and your five cost buckets; the right-hand card turns
them into total cost, recommended price (at your target margin), price per
guest, profit, deposit and balance due.  Copy the recommended price onto
the Events tab or the printable Invoice.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "quote"
LAST_COL = "J"
QR = C.QUOTE_ROWS
QO = C.QUOTE_OUT


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo
    q = m.quote

    bk.widths(KEY, {"A": 2.2, "B": 17, "C": 17, "D": 17, "E": 16, "F": 3,
                    "G": 17, "H": 15, "I": 15, "J": 17})
    bk.paint(KEY, 0, 0, C.QUOTE_LAST_ROW + 8, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F9EE  Quote & Pricing Calculator",
        "  Never guess a price again: costs in \u2192 recommended quote "
        "out, at the margin YOU decide.",
        LAST_COL)
    ws.set_row(r(5), 26)
    ws.merge_range(r(5), 1, r(5), ci(LAST_COL), "", S.canvas)
    ws.set_row(r(8), 8)

    lbl = S.f(**S.base(font_size=11, bold=True, font_color=th.ink,
                       bg_color=th.card, align="left", valign="vcenter",
                       border=1, border_color=th.border, indent=1))
    note_f = S.f(**S.base(font_size=9.5, italic=True, font_color=th.muted,
                          bg_color=th.card, align="left", valign="vcenter",
                          text_wrap=True, border=1, border_color=th.border,
                          indent=1))
    inp = S.f(**S.base(font_size=12, bold=True, font_color=th.primary,
                       bg_color=th.gold_soft, align="center",
                       valign="vcenter", border=2, border_color=th.gold,
                       locked=False))
    inp_n = S.f(**S.base(font_size=12, bold=True, font_color=th.primary,
                         bg_color=th.gold_soft, align="center",
                         valign="vcenter", border=2, border_color=th.gold,
                         num_format="#,##0.00", locked=False))
    inp_i = S.f(**S.base(font_size=12, bold=True, font_color=th.primary,
                         bg_color=th.gold_soft, align="center",
                         valign="vcenter", border=2, border_color=th.gold,
                         num_format="#,##0", locked=False))
    inp_p = S.f(**S.base(font_size=12, bold=True, font_color=th.primary,
                         bg_color=th.gold_soft, align="center",
                         valign="vcenter", border=2, border_color=th.gold,
                         num_format="0%", locked=False))
    inp_d = S.f(**S.base(font_size=12, bold=True, font_color=th.primary,
                         bg_color=th.gold_soft, align="center",
                         valign="vcenter", border=2, border_color=th.gold,
                         num_format="dd mmm yyyy", locked=False))
    out = S.f(**S.base(font_size=12, bold=True, font_color=th.primary,
                       bg_color=th.primary_soft, align="center",
                       valign="vcenter", border=1, border_color=th.border,
                       num_format="#,##0.00"))
    out_hero = S.f(**S.base(font_name=th.title_font, font_size=17, bold=True,
                            font_color=th.white, bg_color=th.ok,
                            align="center", valign="vcenter", border=1,
                            border_color=th.ok, num_format="#,##0.00"))
    out_pct = S.f(**S.base(font_size=12, bold=True, font_color=th.primary,
                           bg_color=th.primary_soft, align="center",
                           valign="vcenter", border=1,
                           border_color=th.border, num_format="0.0%"))

    def left(row, label, value, fmt, cached=0, height=26):
        from datetime import date as _date
        ws.set_row(r(row), height)
        ws.merge_range(r(row), 1, r(row), 3, label, lbl)
        if isinstance(value, str) and value.startswith("="):
            ws.write_formula(r(row), 4, value, fmt, cached)
            bk.stats["formulas"] += 1
        elif isinstance(value, _date):
            ws.write_datetime(r(row), 4, value, fmt)
        else:
            ws.write(r(row), 4, value, fmt)
        ws.write_blank(r(row), 5, None, S.canvas)
        return

    def right(row, label, formula, cached, fmt=None, height=26, hero=False):
        ws.set_row(r(row), height)
        ws.merge_range(r(row), 6, r(row), 8, label,
                       lbl if not hero else S.f(**S.base(
                           font_size=11, bold=True, font_color=th.white,
                           bg_color=th.ok, align="left", valign="vcenter",
                           border=1, border_color=th.ok, indent=1)))
        ws.write_formula(r(row), 9, formula, fmt or (out_hero if hero else out),
                         cached)
        bk.stats["formulas"] += 1

    # ------------------------------------------------------------------
    # left card: inputs
    # ------------------------------------------------------------------
    ws.set_row(r(9), 24)
    ws.merge_range(r(9), 1, r(9), 4, "  1 \u00B7 THE EVENT", S.section_soft)
    ws.merge_range(r(9), 6, r(9), ci(LAST_COL), "", S.canvas)
    left(QR["client"], "Client / enquiry", q.get("client", ""), inp)
    left(QR["date"], "Event date", q.get("date", ""), inp_d)
    left(QR["guests"], "Guest count", q.get("guests", ""), inp_i)

    ws.set_row(r(13), 24)
    ws.merge_range(r(13), 1, r(13), 4, "  2 \u00B7 YOUR COSTS",
                   S.section_soft)
    ws.merge_range(r(13), 6, r(13), ci(LAST_COL), "", S.canvas)
    left(QR["food_pp"], "Food cost per guest", q.get("food_pp", ""), inp_n)
    left(QR["labor"], "Labor total", q.get("labor", ""), inp_n)
    left(QR["equipment"], "Equipment / rental", q.get("equipment", ""),
         inp_n)
    left(QR["transport"], "Transport", q.get("transport", ""), inp_n)
    left(QR["other"], "Other costs", q.get("other", ""), inp_n)

    ws.set_row(r(19), 24)
    ws.merge_range(r(19), 1, r(19), 4, "  3 \u00B7 PRICING POLICY",
                   S.section_soft)
    left(QR["margin"], "Target margin %", "=DefaultMargin", inp_p,
         cached=m.settings["margin"])
    left(QR["deposit_pct"], "Deposit %", "=DepositPct", inp_p,
         cached=m.settings["deposit"])
    ws.set_row(r(22), 8)

    # ------------------------------------------------------------------
    # right card: outputs
    # ------------------------------------------------------------------
    ws.merge_range(r(9), 6, r(12), ci(LAST_COL), "", S.canvas)
    right(QO["food_total"], "Food cost (guests x per guest)",
          '=IF(OR($E$%d="",$E$%d=""),"",$E$%d*$E$%d)'
          % (QR["guests"], QR["food_pp"], QR["guests"], QR["food_pp"]),
          q.get("food_total", ""))
    right(QO["cost_total"], "TOTAL COST",
          '=IF($J$%d="","",SUM($E$%d:$E$%d)+$J$%d)'
          % (QO["food_total"], QR["labor"], QR["other"], QO["food_total"]),
          q.get("cost_total", ""))
    right(QO["price"], "RECOMMENDED PRICE",
          '=IF(OR($J$%d="",$E$%d="",1-$E$%d<=0),"",'
          'ROUND($J$%d/(1-$E$%d),2))'
          % (QO["cost_total"], QR["guests"], QR["margin"],
             QO["cost_total"], QR["margin"]),
          q.get("price", ""), hero=True, height=34)
    right(QO["per_guest"], "Price per guest",
          '=IF(OR($J$%d="",$E$%d=""),"",$J$%d/$E$%d)'
          % (QO["price"], QR["guests"], QO["price"], QR["guests"]),
          q.get("per_guest", ""))
    right(QO["profit"], "Your profit",
          '=IF($J$%d="","",$J$%d-$J$%d)'
          % (QO["price"], QO["price"], QO["cost_total"]),
          q.get("profit", ""))
    right(QO["margin"], "Effective margin",
          '=IF($J$%d="","",$J$%d/$J$%d)'
          % (QO["price"], QO["profit"], QO["price"]),
          (q["profit"] / q["price"]) if q.get("price") else "", fmt=out_pct)
    right(QO["deposit"], "Deposit to collect",
          '=IF($J$%d="","",ROUND($J$%d*$E$%d,2))'
          % (QO["price"], QO["price"], QR["deposit_pct"]),
          q.get("deposit", ""))
    right(QO["balance"], "Balance due later",
          '=IF($J$%d="","",$J$%d-$J$%d)'
          % (QO["price"], QO["price"], QO["deposit"]),
          q.get("balance", ""))
    right(QO["cost_pp"], "True cost per guest",
          '=IF(OR($J$%d="",$E$%d=""),"",$J$%d/$E$%d)'
          % (QO["cost_total"], QR["guests"], QO["cost_total"],
             QR["guests"]),
          q.get("cost_pp", ""))

    # ------------------------------------------------------------------
    # notes + nav
    # ------------------------------------------------------------------
    note_row = C.QUOTE_LAST_ROW - 1
    ws.set_row(r(note_row), 22)
    ws.set_row(r(note_row + 1), 22)
    ws.set_row(r(note_row + 2), 22)
    ws.merge_range(r(note_row), 1, r(note_row + 2), ci(LAST_COL),
                   "  HOW TO PRICE LIKE A PRO\n"
                   "  1.  Food cost per guest = what one plate costs you "
                   "(use the Recipe Calculator on \U0001F37D\uFE0F Menu "
                   "Costing).\n"
                   "  2.  Labor / equipment / transport / other = the "
                   "event's share of everything else.\n"
                   "  3.  Recommended price covers ALL of it and still "
                   "pays you your target margin.  Rule of thumb: food "
                   "should stay under ~30% of the final price.",
                   S.note)

    nav = note_row + 4
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)

    # ------------------------------------------------------------------
    # validation
    # ------------------------------------------------------------------
    bk.validate(KEY, QR["client"], ci("E"), QR["client"], ci("E"),
                "=ClientsList", title="Client",
                message="Who is this quote for?")
    ws.data_validation(r(QR["date"]), ci("E"), r(QR["date"]), ci("E"), {
        "validate": "date", "criteria": ">=", "value": "=DATE(2000,1,1)",
        "input_title": "Event date", "show_input": True, "ignore_blank": True,
        "input_message": "Type a date or pick from the calendar."})
    bk.stats["validations"] += 1
    ws.data_validation(r(QR["guests"]), ci("E"), r(QR["guests"]), ci("E"), {
        "validate": "integer", "criteria": "between", "minimum": 1,
        "maximum": 100000, "input_title": "Guests", "show_input": True,
        "input_message": "How many are you feeding?"})
    bk.stats["validations"] += 1
    for field in ("food_pp", "labor", "equipment", "transport", "other"):
        ws.data_validation(r(QR[field]), ci("E"), r(QR[field]), ci("E"), {
            "validate": "decimal", "criteria": ">=", "value": 0,
            "input_title": "Cost", "show_input": True, "ignore_blank": True})
        bk.stats["validations"] += 1
    for field, (lo, hi) in (("margin", (0, 0.95)), ("deposit_pct", (0, 1))):
        ws.data_validation(r(QR[field]), ci("E"), r(QR[field]), ci("E"), {
            "validate": "decimal", "criteria": "between", "minimum": lo,
            "maximum": hi, "show_input": True, "ignore_blank": True})
        bk.stats["validations"] += 1

    # CF: recommended price glows when computed
    bk.cond(KEY, QO["price"], ci("J"), QO["price"], ci("J"), {
        "type": "formula", "criteria": '=$J$%d<>""' % QO["price"],
        "format": S.cf(bg=th.ok, fg=th.white, bold=True, size=17)})

    bk.page(KEY, LAST_COL, nav + 1, landscape=False, zoom=100)
