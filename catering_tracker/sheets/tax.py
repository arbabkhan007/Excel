"""
\U0001F9FE Tax Tracker - sales tax collected, tax already paid, and what is
still due.  With the required "not professional advice" disclaimer.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "tax"
LAST_COL = "H"
TC = C.TAX_COLS


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY, {"A": 2.2, "B": 4.5, "C": 14, "D": 18, "E": 16, "F": 16,
                    "G": 15, "H": 16})
    log_last = C.TAX_LOG_FIRST + C.TAX_LOG_ROWS - 1
    bk.paint(KEY, 0, 0, log_last + 22, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F9FE  Tax Tracker",
        "  What you collected, what you already paid over, and what is "
        "still due \u2014 ready for filing day.",
        LAST_COL)

    lbl = S.f(**S.base(font_size=11, bold=True, font_color=th.ink,
                       bg_color=th.card, align="left", valign="vcenter",
                       border=1, border_color=th.border, indent=1))
    val = S.f(**S.base(font_size=13, bold=True, font_color=th.primary,
                       bg_color=th.primary_soft, align="center",
                       valign="vcenter", border=1, border_color=th.border,
                       num_format="#,##0.00"))
    val_pct = S.f(**S.base(font_size=13, bold=True, font_color=th.primary,
                           bg_color=th.primary_soft, align="center",
                           valign="vcenter", border=1,
                           border_color=th.border, num_format="0.0%"))
    hero = S.f(**S.base(font_name=th.title_font, font_size=16, bold=True,
                        font_color=th.white, bg_color=th.accent,
                        align="center", valign="vcenter", border=1,
                        border_color=th.accent, num_format="#,##0.00"))

    ws.set_row(r(8), 22)
    ws.merge_range(r(8), 1, r(8), ci(LAST_COL),
                   "  \U0001F4CA  SUMMARY", S.section_soft)

    inc_amt = bk.rng("income", "amount")
    taxable_cached = sum(i["amount"] for i in m.income)
    collected_cached = m.agg.get("tax_collected", 0)
    paid_cached = sum(t["amount"] for t in m.tax_paid)
    due_cached = m.agg.get("tax_due", 0)

    rows = [
        ("Taxable sales (all invoices)", "=SUM(%s)" % inc_amt,
         taxable_cached, val, "#,##0.00"),
        ("Tax rate (from \u2699\uFE0F Setup)", "=TaxRate",
         m.settings["tax"], val_pct, "0.0%"),
        ("Tax collected on sales", "=ROUND($G$%d*$G$%d,2)"
         % (C.TAX_SUM["sales"], C.TAX_SUM["rate"]), collected_cached,
         val, "#,##0.00"),
        ("Tax already paid (log below)", "=SUM($G$%d:$G$%d)"
         % (C.TAX_LOG_FIRST, log_last), paid_cached, val, "#,##0.00"),
        ("ESTIMATED TAX STILL DUE", "=ROUND($G$%d-$G$%d,2)"
         % (C.TAX_SUM["collected"], C.TAX_SUM["paid"]), due_cached,
         hero, "#,##0.00"),
    ]
    for j, (label, formula, cached, fmt, _nf) in enumerate(rows):
        row = C.TAX_SUM["sales"] + j
        ws.set_row(r(row), 30 if j == len(rows) - 1 else 24)
        ws.merge_range(r(row), 1, r(row), 5, label,
                       lbl if j < len(rows) - 1 else S.f(**S.base(
                           font_size=12, bold=True, font_color=th.white,
                           bg_color=th.accent, align="left",
                           valign="vcenter", border=1, border_color=th.accent,
                           indent=1)))
        ws.merge_range(r(row), 6, r(row), 7, "", fmt)
        ws.write_formula(r(row), 6, formula, fmt, cached)
        bk.stats["formulas"] += 1
    ws.set_row(r(14), 8)

    # ------------------------------------------------------------------
    # payments log
    # ------------------------------------------------------------------
    ws.set_row(r(C.TAX_LOG_HDR - 1), 22)
    ws.merge_range(r(C.TAX_LOG_HDR - 1), 1, r(C.TAX_LOG_HDR - 1),
                   ci(LAST_COL),
                   "  \U0001F4DD  TAX PAYMENTS LOG  \u2014  every payment "
                   "you make to the taxman", S.section_soft)
    ws.set_row(r(C.TAX_LOG_HDR), 24)
    ws.write(r(C.TAX_LOG_HDR), ci(TC["n"]), "#", S.header(th.primary))
    ws.write(r(C.TAX_LOG_HDR), ci(TC["date"]), "Date paid",
             S.header(th.primary))
    ws.merge_range(r(C.TAX_LOG_HDR), ci(TC["desc"]), r(C.TAX_LOG_HDR),
                   ci("F"), "What for", S.header(th.primary))
    ws.write(r(C.TAX_LOG_HDR), ci(TC["amount"]), "Amount",
             S.header(th.primary))
    ws.write(r(C.TAX_LOG_HDR), ci(TC["method"]), "Method",
             S.header(th.primary))

    for i in range(C.TAX_LOG_ROWS):
        row = C.TAX_LOG_FIRST + i
        a = i % 2
        ws.set_row(r(row), 20)
        ws.write(r(row), ci(TC["n"]), i + 1, S.idx(a))
        ws.write_blank(r(row), ci(TC["date"]), None, S.cell("date", a))
        ws.merge_range(r(row), ci(TC["desc"]), r(row), ci("F"), "",
                       S.cell("text", a))
        ws.write_blank(r(row), ci(TC["amount"]), None, S.cell("money", a))
        ws.write_blank(r(row), ci(TC["method"]), None, S.cell("center", a))
        if i < len(m.tax_paid):
            t = m.tax_paid[i]
            ws.write_datetime(r(row), ci(TC["date"]), t["date"],
                              S.cell("date", a))
            ws.write(r(row), ci(TC["desc"]), t["desc"], S.cell("text", a))
            ws.write(r(row), ci(TC["amount"]), t["amount"],
                     S.cell("money", a))

    ws.data_validation(r(C.TAX_LOG_FIRST), ci(TC["date"]), r(log_last),
                       ci(TC["date"]),
                       {"validate": "date", "criteria": ">=",
                        "value": "=DATE(2000,1,1)", "ignore_blank": True,
                        "input_title": "Date paid", "show_input": True,
                        "input_message": "When did you pay it?"})
    bk.stats["validations"] += 1
    ws.data_validation(r(C.TAX_LOG_FIRST), ci(TC["amount"]), r(log_last),
                       ci(TC["amount"]),
                       {"validate": "decimal", "criteria": ">=", "value": 0,
                        "ignore_blank": True})
    bk.stats["validations"] += 1
    bk.validate(KEY, C.TAX_LOG_FIRST, ci(TC["method"]), log_last,
                ci(TC["method"]), "=" + bk.listname("payment_methods"),
                title="Method")

    # ------------------------------------------------------------------
    # disclaimer + notes
    # ------------------------------------------------------------------
    drow = log_last + 2
    ws.set_row(r(drow), 20)
    ws.set_row(r(drow + 1), 20)
    ws.set_row(r(drow + 2), 20)
    ws.merge_range(r(drow), 1, r(drow + 2), ci(LAST_COL),
                   "  \u26A0\uFE0F  PLEASE NOTE: this tab is a simple "
                   "record-keeping helper only. It is NOT professional "
                   "tax, accounting or legal advice. Tax rules differ by "
                   "country, state and business structure \u2014 always "
                   "confirm your rates, registrations, deductions and "
                   "filing deadlines with a qualified accountant or tax "
                   "advisor.", S.note)

    nav = drow + 4
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)

    # CF: due > 0 glows
    bk.cond(KEY, C.TAX_SUM["due"], ci("G"), C.TAX_SUM["due"], ci("H"), {
        "type": "cell", "criteria": ">", "value": 0,
        "format": S.cf(bg=th.bad, fg=th.white, bold=True, size=16)})

    bk.page(KEY, LAST_COL, nav + 1, landscape=False, zoom=100)
