"""
The ⚙️ Setup tab: business profile, money defaults, the calendar month and
every editable dropdown list in the workbook.

Everything typed here flows through defined names (BusinessName, Currency,
TaxRate, DefaultMargin, DepositPct, DueSoonDays, ReportYear, CalMonth,
CalYear + the six list names), so a change here updates every tab.
"""

from .. import config as C
from ..book import r, ci

LAST_COL = "J"


def build(bk):
    key = "setup"
    ws = bk.ws(key)
    S, th, m = bk.S, bk.th, bk.demo
    st = m.settings

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
                   "\u2699\uFE0F  Setup \u2014 make it your business",
                   S.sheet_title)
    ws.merge_range(r(3), 1, r(3), ci(LAST_COL) - 3,
                   "  Set these once and every tab, quote, alert and "
                   "dropdown follows automatically.", S.sheet_sub)
    ws.merge_range(r(3), ci(LAST_COL) - 2, r(3), ci(LAST_COL), "",
                   S.home_link)
    ws.write_url(r(3), ci(LAST_COL) - 2,
                 "internal:%s!A1" % bk.q("dashboard"), S.home_link,
                 "\U0001F3E0  Back to Dashboard")
    bk.stats["links"] += 1
    ws.set_row(r(4), 7)

    label = S.f(**S.base(font_size=11, bold=True, font_color=th.ink,
                         bg_color=th.card, align="left", valign="vcenter",
                         border=1, border_color=th.border, indent=1))
    note_f = S.f(**S.base(font_size=9.5, italic=True, font_color=th.muted,
                          bg_color=th.card, align="left", valign="top",
                          text_wrap=True,
                          border=1, border_color=th.border, indent=1))
    input_f = S.f(**S.base(font_size=11, bold=True, font_color=th.primary,
                           bg_color=th.gold_soft, align="center",
                           valign="vcenter", border=2, border_color=th.gold,
                           locked=False))

    def fmt(num_format=None):
        props = dict(font_size=11, bold=True, font_color=th.primary,
                     bg_color=th.gold_soft, align="center",
                     valign="vcenter", border=2, border_color=th.gold,
                     locked=False)
        if num_format:
            props["num_format"] = num_format
        return S.f(**S.base(**props))

    def setting(row, text, note, value, cell_fmt, height=26):
        ws.set_row(r(row), height)
        ws.write(r(row), ci("B"), text, label)
        ws.write(r(row), ci("C"), value, cell_fmt)
        ws.merge_range(r(row), ci("D"), r(row), ci(LAST_COL), note, note_f)

    def section(row, text, style=None):
        ws.set_row(r(row), 24)
        ws.merge_range(r(row), 1, r(row), ci(LAST_COL), text,
                       style or S.section)

    def decimal(row, lo, hi, title, message):
        ws.data_validation(r(row), ci("C"), r(row), ci("C"), {
            "validate": "decimal", "criteria": "between",
            "minimum": lo, "maximum": hi, "ignore_blank": True,
            "input_title": title, "input_message": message,
            "show_input": True, "show_error": True,
            "error_message": message})
        bk.stats["validations"] += 1

    # ------------------------------------------------------------------
    # 1. business profile
    # ------------------------------------------------------------------
    section(5, "  \U0001F468\u200D\U0001F373  YOUR BUSINESS")
    setting(C.SU_BUSINESS, "Business name",
            "Shown on the dashboard hero and on the printable invoice.",
            st["business"], input_f, height=28)
    setting(C.SU_MESSAGE, "Dashboard message / focus",
            "One line under the dashboard title: this week's priority, a "
            "reminder, or a little motivation.", st["message"], input_f)
    ws.set_row(r(8), 8)

    # ------------------------------------------------------------------
    # 2. money defaults
    # ------------------------------------------------------------------
    section(9, "  \U0001F4B0  MONEY DEFAULTS")
    setting(C.SU_CURRENCY, "Currency symbol",
            "Every summary readout prefixes amounts with this. To show it "
            "inside the tables too: select the money columns \u2192 Home "
            "\u2192 Number format \u2192 Currency.",
            st["currency"], input_f)
    bk.validate(key, C.SU_CURRENCY, ci("C"), C.SU_CURRENCY, ci("C"),
                '="%s"' % ",".join(C.CURRENCIES),
                title="Currency", message="Pick your currency symbol.",
                error="Choose one of the listed symbols.")
    setting(C.SU_TAX, "Tax rate (sales tax / GST)",
            "Applied on the \U0001F9FE Tax tab and the printable invoice. "
            "Use 0% if you are not tax registered.",
            st["tax"], fmt("0.0%"))
    decimal(C.SU_TAX, 0, 1, "Tax rate",
            "A decimal between 0 and 1 (5% = 0.05).")
    setting(C.SU_MARGIN, "Default target margin",
            "The \U0001F9EE Quote Calculator starts from this. A 35% margin "
            "means price = total cost \u00F7 0.65.",
            st["margin"], fmt("0%"))
    decimal(C.SU_MARGIN, 0, 0.95, "Margin",
            "Between 0% and 95% (0.35 = 35%).")
    setting(C.SU_DEPOSIT, "Default deposit %",
            "Suggested deposit on quotes and invoices.",
            st["deposit"], fmt("0%"))
    decimal(C.SU_DEPOSIT, 0, 1, "Deposit %",
            "Between 0% and 100% (0.4 = 40%).")
    setting(C.SU_DUESOON, "\u201CDue soon\u201D window (days)",
            "Payment due dates inside this many days turn orange; past-due "
            "turns red everywhere.", st["duesoon"], fmt("0"))
    decimal(C.SU_DUESOON, 1, 90, "Days", "A number of days, 1-90.")
    setting(C.SU_YEAR, "Reporting year",
            "The \U0001F4C8 P&L tab and the monthly charts report on this "
            "year.", st["year"], fmt("0"))
    ws.set_row(r(16), 8)

    # ------------------------------------------------------------------
    # 3. calendar month
    # ------------------------------------------------------------------
    section(17, "  \U0001F4C6  EVENT CALENDAR MONTH")
    setting(C.SU_CAL_MONTH, "Calendar month (1-12)",
            "Which month the \U0001F4C6 Event Calendar tab displays "
            "(9 = September).", st["cal_month"], fmt("0"))
    bk.validate(key, C.SU_CAL_MONTH, ci("C"), C.SU_CAL_MONTH, ci("C"),
                '"1,2,3,4,5,6,7,8,9,10,11,12"', title="Month",
                message="A number from 1 (January) to 12 (December).")
    setting(C.SU_CAL_YEAR, "Calendar year", "", st["cal_year"], fmt("0"))
    ws.set_row(r(20), 8)

    # ------------------------------------------------------------------
    # 4. editable lists
    # ------------------------------------------------------------------
    section(C.SU_LIST_HEADER - 1,
            "  \U0001F4CB  YOUR LISTS  \u2014  edit any time; every "
            "dropdown in the workbook updates itself")
    ws.set_row(r(C.SU_LIST_HEADER), 30)
    for lk, col in sorted(C.LIST_COLS.items(), key=lambda kv: kv[1]):
        ws.write(r(C.SU_LIST_HEADER), ci(col), C.LIST_TITLES[lk],
                 S.header(th.primary))
    values = {
        "event_types": C.EVENT_TYPES,
        "expense_categories": C.EXPENSE_CATEGORIES,
        "payment_methods": C.PAYMENT_METHODS,
        "staff_roles": C.STAFF_ROLES,
        "menu_categories": C.MENU_CATEGORIES,
        "ingredient_categories": C.INGREDIENT_CATEGORIES,
    }
    list_fmt = S.f(**S.base(font_size=10, font_color=th.ink, bg_color=th.card,
                            align="left", valign="vcenter", border=1,
                            border_color=th.border, indent=1, locked=False))
    list_fmt_alt = S.f(**S.base(font_size=10, font_color=th.ink,
                                bg_color=th.alt, align="left",
                                valign="vcenter", border=1,
                                border_color=th.border, indent=1,
                                locked=False))
    for i in range(C.SU_LIST_ROWS):
        row = C.SU_LIST_FIRST + i
        ws.set_row(r(row), 18)
        for lk, col in C.LIST_COLS.items():
            vals = values[lk]
            value = vals[i] if i < len(vals) else ""
            ws.write(r(row), ci(col), value,
                     list_fmt_alt if i % 2 else list_fmt)
    note_row = C.SU_LIST_FIRST + C.SU_LIST_ROWS
    ws.set_row(r(note_row), 26)
    ws.set_row(r(note_row + 1), 26)
    ws.merge_range(r(note_row), 1, r(note_row + 1), ci(LAST_COL),
                   "  \u2022  Type over the defaults with your own words "
                   "\u2014 add as many as you like (20 rows each).\n"
                   "  \u2022  Don't leave a blank row in the middle of a "
                   "list: the dropdown stops at the first gap.\n"
                   "  \u2022  Deleting a word here does NOT delete your "
                   "events or expenses \u2014 it only removes it from the "
                   "dropdown.", S.note)
    ws.set_row(r(note_row + 2), 8)

    # ------------------------------------------------------------------
    # 5. fixed lists (locked)
    # ------------------------------------------------------------------
    section(C.SU_FIXED_HEADER - 1,
            "  \U0001F512  FIXED LISTS  \u2014  colours, counts and "
            "dashboard maths depend on this exact wording, so they are "
            "locked", S.section_accent)
    ws.set_row(r(C.SU_FIXED_HEADER), 26)
    fixed_titles = {"event_statuses": "Event pipeline",
                    "payment_statuses": "Payment status",
                    "tick": "Tick box"}
    for fk, col in sorted(C.FIXED_COLS.items(), key=lambda kv: kv[1]):
        ws.write(r(C.SU_FIXED_HEADER), ci(col), fixed_titles[fk],
                 S.header(th.accent))
    fixed_values = {"event_statuses": C.EVENT_STATUSES,
                    "payment_statuses": C.PAYMENT_STATUSES,
                    "tick": [C.TICK, ""]}
    lock_fmt = S.f(**S.base(font_size=10, font_color=th.muted,
                            bg_color=th.primary_soft, align="left",
                            valign="vcenter", border=1,
                            border_color=th.border, indent=1, locked=True))
    for fk, col in C.FIXED_COLS.items():
        vals = fixed_values[fk]
        for i in range(6):
            row = C.SU_FIXED_FIRST + i
            ws.set_row(r(row), 18)
            ws.write(r(row), ci(col), vals[i] if i < len(vals) else "",
                     lock_fmt)

    # ------------------------------------------------------------------
    # footer
    # ------------------------------------------------------------------
    foot = C.SU_FIXED_FIRST + 7
    ws.set_row(r(foot), 30)
    ws.merge_range(r(foot), 1, r(foot), ci(LAST_COL),
                   "  \U0001F4A1  Everything else is formula-driven: you "
                   "never have to touch a formula, only these settings and "
                   "the white input cells on each tab.",
                   S.f(**S.base(font_size=10, italic=True, font_color=th.ink,
                                bg_color=th.gold_soft, align="left",
                                valign="vcenter", indent=1)))
    nav_row = foot + 2
    ws.set_row(r(nav_row), 24)
    bk.nav_row(key, nav_row, max_col=LAST_COL)

    bk.page(key, LAST_COL, nav_row + 1, landscape=False,
            freeze=(4, 0), fit=True, zoom=100)
