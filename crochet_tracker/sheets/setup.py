"""
The ⚙️ Lists & Settings tab: business profile, money defaults and every
editable dropdown list in the workbook.
"""

from .. import config as C
from ..book import r, ci

LAST_COL = "J"


def build(bk):
    key = "setup"
    ws = bk.ws(key)
    S, th, m = bk.S, bk.th, bk.demo
    st = m.settings if m else {}

    bk.widths(key, {"A": 2.2, "B": 34, "C": 30, "D": 16, "E": 16, "F": 16,
                    "G": 18, "H": 16, "I": 16, "J": 16})
    bk.paint(key, 0, 0, 70, ci(LAST_COL), S.canvas)

    # ------------------------------------------------------------------
    # title band
    # ------------------------------------------------------------------
    ws.set_row(r(1), 7)
    ws.set_row(r(2), 32)
    ws.set_row(r(3), 18)
    ws.merge_range(r(2), 1, r(2), ci(LAST_COL) - 3,
                   "\u2699\uFE0F  Lists & Settings \u2014 make it yours",
                   S.sheet_title)
    ws.merge_range(r(3), 1, r(3), ci(LAST_COL) - 3,
                   "  Set these once and every tab, calculator, alert and "
                   "dropdown follows automatically.", S.sheet_sub)
    ws.merge_range(r(3), ci(LAST_COL) - 2, r(3), ci(LAST_COL), "",
                   S.home_link)
    ws.write_url(r(3), ci(LAST_COL) - 2,
                 "internal:%s!A1" % bk.q("dashboard"), S.home_link,
                 "\U0001F3E0  Back to Dashboard")

    inp = S.f(**S.base(font_size=12, bold=True, font_color=th.primary,
                       bg_color=th.gold_soft, border=1,
                       border_color=th.border_strong, align="left",
                       valign="vcenter", indent=1))
    inp_pct = S.f(**S.base(font_size=12, bold=True, font_color=th.primary,
                           bg_color=th.gold_soft, border=1,
                           border_color=th.border_strong, align="left",
                           valign="vcenter", indent=1, num_format="0%"))
    inp_money = S.f(**S.base(font_size=12, bold=True, font_color=th.primary,
                             bg_color=th.gold_soft, border=1,
                             border_color=th.border_strong, align="left",
                             valign="vcenter", indent=1,
                             num_format="#,##0.00"))
    label = S.f(**S.base(font_size=11, font_color=th.ink, bg_color=th.card,
                         align="left", valign="vcenter", indent=1,
                         border=1, border_color=th.border))

    def setting(row, text, value, fmt=None, cached=None):
        ws.set_row(r(row), 22)
        ws.write(r(row), 1, text, label)
        if value is None:
            ws.write(r(row), 2, cached if cached is not None else "", fmt)
        else:
            ws.write(r(row), 2, value, fmt)

    def section(row, emoji, text):
        ws.set_row(r(row), 24)
        ws.merge_range(r(row), 1, r(row), ci(LAST_COL),
                       "  %s  %s" % (emoji, text), S.section)

    # ------------------------------------------------------------------
    section(5, "\U0001F3F7\uFE0F", "BUSINESS PROFILE")
    setting(C.SU_BUSINESS, "Business name (dashboard header)",
            st.get("business", ""), inp)
    setting(C.SU_MESSAGE, "Dashboard message of the day",
            st.get("message", ""), inp)

    section(9, "\U0001F4B5", "MONEY & SEASON DEFAULTS")
    setting(C.SU_CURRENCY, "Currency symbol", st.get("currency", "$"), inp)
    setting(C.SU_YEAR, "Report year (monthly summary)",
            st.get("year", 2026), inp_money)
    setting(C.SU_WAGE, "Your hourly wage (pricing calculator)",
            st.get("wage", 14.0), inp_money)
    setting(C.SU_OVERHEAD, "Overhead % on top of direct costs",
            st.get("overhead", 0.10), inp_pct)
    setting(C.SU_MARGIN, "Default target profit margin",
            st.get("margin", 0.45), inp_pct)

    # ------------------------------------------------------------------
    section(C.SU_LIST_HEADER - 1, "\U0001F4DD",
            "YOUR DROPDOWN LISTS (edit freely)")
    ws.set_row(r(C.SU_LIST_HEADER), 20)
    cols = sorted(C.LIST_COLS.items(), key=lambda kv: kv[1])
    for key2, col in cols:
        ws.write(r(C.SU_LIST_HEADER), ci(col), C.LIST_TITLES[key2], S.thead)
    values = C.LIST_VALUES
    for i in range(C.SU_LIST_ROWS):
        row = C.SU_LIST_FIRST + i
        ws.set_row(r(row), 16)
        for key2, col in cols:
            lst = values[key2]
            cell = S.f(**S.base(font_size=10, font_color=th.ink,
                                bg_color=th.card if i % 2 else th.alt,
                                border=1, border_color=th.border,
                                align="left", valign="vcenter", indent=1))
            if m and key2 in ("suppliers",):
                pass
            ws.write(r(row), ci(col),
                     lst[i] if i < len(lst) else "", cell)
    note_row = C.SU_LIST_FIRST + C.SU_LIST_ROWS
    ws.merge_range(r(note_row), 1, r(note_row + 1), ci(LAST_COL),
                   "  These five columns feed every dropdown in the "
                   "workbook.  Add, rename or delete entries and the "
                   "catalog, sales log and reorder list update instantly.",
                   S.note)

    # ------------------------------------------------------------------
    section(C.SU_FIXED_HEADER - 1, "\U0001F512", "FIXED LIST (do not edit)")
    ws.set_row(r(C.SU_FIXED_HEADER), 20)
    ws.write(r(C.SU_FIXED_HEADER), ci("C"), "Tick box", S.thead)
    ws.write(r(C.SU_FIXED_FIRST), ci("C"), C.TICK, S.note_plain)
    ws.write(r(C.SU_FIXED_FIRST + 1), ci("C"), "", S.note_plain)

    foot = C.SU_FIXED_FIRST + 5
    ws.set_row(r(foot), 30)
    ws.merge_range(r(foot), 1, r(foot), ci(LAST_COL),
                   "  \U0001F4A1  Cream cells are inputs everywhere in this "
                   "workbook; white cells are formulas.  Sheets are "
                   "protected (no password) so nothing breaks by accident.",
                   S.f(**S.base(font_size=10, italic=True, font_color=th.ink,
                                bg_color=th.gold_soft, align="left",
                                valign="vcenter", indent=1)))
    nav = foot + 2
    ws.set_row(r(nav), 24)
    bk.nav_row(key, nav, max_col=LAST_COL)
    bk.page(key, LAST_COL, nav + 1, landscape=False, zoom=100)
