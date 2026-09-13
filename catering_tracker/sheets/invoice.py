"""
\U0001F5A8️ Invoice & Proposal - pick an Event ID and the whole document
fills itself: client, event details, money, deposit, balance and a PAID /
BALANCE DUE stamp.  Prints on one A4 page.  The bottom half doubles as a
simple proposal (scope + terms) you can send before the event.
"""

from .. import config as C
from ..book import r, ci

KEY = "invoice"
LAST_COL = "H"
SEL_ROW = 10                 # Event ID selector
ev_first, ev_last = C.ROW_FIRST, C.ROW_FIRST + C.CAP["events"] - 1


def _idx(colletter):
    """INDEX/MATCH pull from the events sheet by the selected Event ID."""
    ev = "'%s'" % C.SHEET_NAMES["events"]
    ids = "%s!$C$%d:$C$%d" % (ev, ev_first, ev_last)
    src = "%s!$%s$%d:$%s$%d" % (ev, colletter, ev_first, colletter, ev_last)
    return 'IFERROR(INDEX(%s,MATCH($D$%d,%s,0)),"")' % (src, SEL_ROW, ids)


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY, {"A": 2.5, "B": 16, "C": 16, "D": 16, "E": 14, "F": 14,
                    "G": 14, "H": 14})
    bk.paint(KEY, 0, 0, 66, ci(LAST_COL), S.canvas)

    # demo selection
    sel = ""
    ev = None
    if m.events:
        ev = next((e for e in m.events if e["id"] == "EV-2026-005"),
                  m.events[0])
        sel = ev["id"]

    # ------------------------------------------------------------------
    # letterhead
    # ------------------------------------------------------------------
    ws.set_row(r(1), 8)
    ws.set_row(r(2), 40)
    ws.merge_range(r(2), 1, r(2), ci(LAST_COL), "", S.hero_title)
    ws.write_formula(r(2), 1, '="\U0001F37D\uFE0F  "&BusinessName',
                     S.hero_title,
                     "\U0001F37D\uFE0F  %s" % m.settings["business"])
    bk.stats["formulas"] += 1
    ws.set_row(r(3), 22)
    ws.merge_range(r(3), 1, r(3), 4, "  INVOICE & PROPOSAL",
                   S.f(**S.base(font_name=th.title_font, font_size=15,
                                bold=True, font_color=th.accent,
                                bg_color=th.bg, align="left",
                                valign="vcenter")))
    issued_f = S.f(**S.base(font_size=10, italic=True, font_color=th.muted,
                            bg_color=th.bg, align="right", valign="vcenter"))
    ws.merge_range(r(3), 5, r(3), ci(LAST_COL), "", issued_f)
    ws.write_formula(r(3), 5, '="Issued: "&TEXT(TODAY(),"dd mmm yyyy")',
                     issued_f,
                     "Issued: 12 Sep 2026" if m.events else "Issued: ")
    bk.stats["formulas"] += 1
    ws.set_row(r(4), 8)

    # ------------------------------------------------------------------
    # selector
    # ------------------------------------------------------------------
    sel_lbl = S.f(**S.base(font_size=11, bold=True, font_color=th.white,
                           bg_color=th.primary, align="left",
                           valign="vcenter", border=1,
                           border_color=th.primary, indent=1))
    sel_inp = S.f(**S.base(font_size=12, bold=True, font_color=th.primary,
                           bg_color=th.gold_soft, align="center",
                           valign="vcenter", border=2, border_color=th.gold,
                           locked=False))
    ws.set_row(r(6), 8)
    ws.set_row(r(SEL_ROW - 1), 22)
    ws.merge_range(r(SEL_ROW - 1), 1, r(SEL_ROW - 1), ci(LAST_COL),
                   "  \U0001F3AF  Pick the event - everything below fills "
                   "itself in", S.section_soft)
    ws.set_row(r(SEL_ROW), 32)
    ws.merge_range(r(SEL_ROW), 1, r(SEL_ROW), 2, "Event ID", sel_lbl)
    ws.merge_range(r(SEL_ROW), 3, r(SEL_ROW), 4, "", sel_inp)
    ws.write(r(SEL_ROW), 3, sel, sel_inp)
    ws.merge_range(r(SEL_ROW), 5, r(SEL_ROW), ci(LAST_COL),
                   "   \u2190 dropdown lists every Event ID from the "
                   "\U0001F4C5 Events tab", S.note_plain)
    bk.validate(KEY, SEL_ROW, 3, SEL_ROW, 3, "=EventList",
                title="Event ID",
                message="Which event is this document for?")

    # ------------------------------------------------------------------
    # pulled details
    # ------------------------------------------------------------------
    lbl = S.f(**S.base(font_size=10, bold=True, font_color=th.muted,
                       bg_color=th.card, align="left", valign="vcenter",
                       border=1, border_color=th.border, indent=1))
    valc = S.f(**S.base(font_size=11, bold=True, font_color=th.ink,
                        bg_color=th.card, align="left", valign="vcenter",
                        border=1, border_color=th.border, indent=1))
    valc_c = S.f(**S.base(font_size=11, bold=True, font_color=th.ink,
                          bg_color=th.card, align="center", valign="vcenter",
                          border=1, border_color=th.border))
    valc_date = S.f(**S.base(font_size=11, bold=True, font_color=th.ink,
                             bg_color=th.card, align="center",
                             valign="vcenter", border=1,
                             border_color=th.border,
                             num_format="dd mmm yyyy"))

    def pull(row, label, colletter, fmt, cached, lcols, vcols):
        ws.set_row(r(row), 22)
        ws.merge_range(r(row), lcols[0], r(row), lcols[1], label, lbl)
        if vcols[0] == vcols[1]:
            ws.write_formula(r(row), vcols[0], "=" + _idx(colletter), fmt,
                             cached)
        else:
            ws.merge_range(r(row), vcols[0], r(row), vcols[1], "", fmt)
            ws.write_formula(r(row), vcols[0], "=" + _idx(colletter), fmt,
                             cached)
        bk.stats["formulas"] += 1

    d1 = 12
    pull(d1, "Client", "D", valc, ev["client"] if ev else "", (1, 2), (3, 4))
    pull(d1, "Event date", "E", valc_date, ev["date"] if ev else "",
         (5, 6), (7, 7))
    pull(d1 + 1, "Event type", "F", valc_c, ev["type"] if ev else "",
         (1, 2), (3, 4))
    pull(d1 + 1, "Guests", "G", valc_c, ev["guests"] if ev else "",
         (5, 6), (7, 7))
    pull(d1 + 2, "Menu", "H", valc, ev["menu"] if ev else "", (1, 2), (3, 7))
    pull(d1 + 3, "Notes", "R", valc, ev["notes"] if ev else "", (1, 2),
         (3, 7))

    # ------------------------------------------------------------------
    # money box
    # ------------------------------------------------------------------
    box = d1 + 5                     # row 17
    ws.set_row(r(box - 1), 8)
    ws.set_row(r(box), 22)
    ws.merge_range(r(box), 1, r(box), ci(LAST_COL),
                   "  \U0001F4B0  THE MONEY", S.section_soft)
    r_quote = box + 1
    r_tax = box + 2
    r_total = box + 3
    r_dep = box + 4
    r_bal = box + 5

    def money_line(row, label, formula, cached, bg, big=False):
        ws.set_row(r(row), 26 if big else 22)
        lab_f = S.f(**S.base(font_size=12 if big else 11, bold=True,
                             font_color=th.primary if big else th.ink,
                             bg_color=bg, align="left", valign="vcenter",
                             border=1, border_color=th.border, indent=1))
        val_f = S.f(**S.base(font_size=14 if big else 12, bold=True,
                             font_color=th.ink, bg_color=bg, align="right",
                             valign="vcenter", border=1,
                             border_color=th.border, indent=1,
                             num_format="#,##0.00"))
        ws.merge_range(r(row), 1, r(row), 4, label, lab_f)
        ws.merge_range(r(row), 5, r(row), ci(LAST_COL), "", val_f)
        ws.write_formula(r(row), 5, formula, val_f, cached)
        bk.stats["formulas"] += 1

    price = ev["price"] if ev else ""
    dep = ev["deposit"] if ev else ""
    tax = m.settings["tax"]
    total_c = round(ev["price"] * (1 + tax), 2) if ev else ""
    bal_c = round(total_c - ev["deposit"], 2) if ev else ""

    money_line(r_quote, "Quote / invoice total", "=" + _idx("L"), price,
               th.card)
    money_line(r_tax, "Tax at your Setup rate",
               '=IF($F$%d="","",ROUND($F$%d*TaxRate,2))' % (r_quote, r_quote),
               round(ev["price"] * tax, 2) if ev else "", th.card)
    money_line(r_total, "TOTAL",
               '=IF($F$%d="","",$F$%d+$F$%d)' % (r_quote, r_quote, r_tax),
               total_c, th.primary_soft, big=True)
    money_line(r_dep, "Deposit held", "=" + _idx("O"), dep, th.card)
    money_line(r_bal, "BALANCE DUE",
               '=IF($F$%d="","",$F$%d-$F$%d)' % (r_quote, r_total, r_dep),
               bal_c, th.gold_soft, big=True)

    # PAID / BALANCE DUE stamp
    stamp_row = r_bal + 2
    ws.set_row(r(stamp_row), 32)
    stamp_fmt = S.f(**S.base(font_name=th.title_font, font_size=15,
                             bold=True, font_color=th.white, bg_color=th.ok,
                             align="center", valign="vcenter", border=1,
                             border_color=th.ok))
    ws.merge_range(r(stamp_row), 1, r(stamp_row), ci(LAST_COL), "",
                   stamp_fmt)
    ws.write_formula(
        r(stamp_row), 1,
        '=IF($D$%d="","",IF($F$%d<=0,"\u2705 PAID IN FULL",'
        '"\u23F3 BALANCE DUE: "&Currency&TEXT($F$%d,"#,##0")))'
        % (SEL_ROW, r_bal, r_bal), stamp_fmt,
        ("\u23F3 BALANCE DUE: %s" % m.money(bal_c)) if ev else "")
    bk.stats["formulas"] += 1
    bk.cond(KEY, stamp_row, 1, stamp_row, ci(LAST_COL), {
        "type": "formula",
        "criteria": '=ISNUMBER(SEARCH("PAID",$B$%d))' % stamp_row,
        "format": S.cf(bg=th.ok, fg=th.white, bold=True, size=15)})
    bk.cond(KEY, stamp_row, 1, stamp_row, ci(LAST_COL), {
        "type": "formula",
        "criteria": '=ISNUMBER(SEARCH("BALANCE",$B$%d))' % stamp_row,
        "format": S.cf(bg=th.accent, fg=th.white, bold=True, size=15)})

    # ------------------------------------------------------------------
    # proposal / terms
    # ------------------------------------------------------------------
    terms_top = stamp_row + 2
    ws.set_row(r(terms_top), 22)
    ws.merge_range(r(terms_top), 1, r(terms_top), ci(LAST_COL),
                   "  \U0001F4DC  PROPOSAL TERMS  (edit these to match "
                   "your business)", S.section_soft)
    terms = [
        "1.  This quote is valid for 30 days from the issue date.",
        "2.  A deposit (see above) confirms your date; the balance is due "
        "on or before the event day.",
        "3.  Final guest count is required 72 hours before service; "
        "catering is prepared to that number.",
        "4.  Please advise allergies and dietary requirements with the "
        "final count.",
        "5.  Cancellations inside 7 days forfeit the deposit; inside 72 "
        "hours, 100% of the quote is chargeable.",
        "6.  Menu substitutions of equal or greater value may be made if "
        "ingredients are unavailable.",
    ]
    t_fmt = S.f(**S.base(font_size=10.5, font_color=th.ink, bg_color=th.card,
                         align="left", valign="vcenter", border=1,
                         border_color=th.border, indent=1, locked=False))
    for j, t in enumerate(terms):
        row = terms_top + 1 + j
        ws.set_row(r(row), 20)
        ws.merge_range(r(row), 1, r(row), ci(LAST_COL), t, t_fmt)

    sign = terms_top + len(terms) + 2
    ws.set_row(r(sign), 34)
    sig_f = S.f(**S.base(font_size=10, italic=True, font_color=th.muted,
                         bg_color=th.bg, align="center", valign="bottom",
                         top=1, top_color=th.border_strong))
    ws.merge_range(r(sign), 1, r(sign), 3, "Signed (client)", sig_f)
    ws.write_blank(r(sign), 4, None, S.canvas)
    ws.merge_range(r(sign), 5, r(sign), 7, "Date", sig_f)

    thanks = sign + 2
    ws.set_row(r(thanks), 24)
    thanks_f = S.f(**S.base(font_name=th.title_font, font_size=12,
                            bold=True, italic=True, font_color=th.accent,
                            bg_color=th.bg, align="center",
                            valign="vcenter"))
    ws.merge_range(r(thanks), 1, r(thanks), ci(LAST_COL), "", thanks_f)
    ws.write_formula(
        r(thanks), 1,
        '="Thank you for choosing "&BusinessName&" \u2014 we cannot wait '
        'to feed your guests!"', thanks_f,
        "Thank you for choosing %s \u2014 we cannot wait to feed your "
        "guests!" % m.settings["business"])
    bk.stats["formulas"] += 1

    nav = thanks + 2
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)

    bk.page(KEY, LAST_COL, nav + 1, landscape=False, zoom=100, fit=True)
    ws.print_area("A1:%s%d" % (LAST_COL, thanks))
