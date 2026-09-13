"""
\U0001F4C6 Event Calendar - a real monthly grid.

The month and year come from ⚙️ Setup (CalMonth / CalYear).  Each day cell
shows its day number plus the first event booked that day
(client • type), looked up live from the Events tab.  Days with events
glow in the accent colour.
"""

from datetime import date

from .. import config as C
from ..book import r, ci

KEY = "calendar"
LAST_COL = "H"
GRID = C.CAL_GRID_ROW
WEEKDAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]


def _cell_formula(k):
    """Full day-cell formula: day number + first event of that day."""
    start = "DATE(CalYear,CalMonth,1)"
    n = "%d+2-WEEKDAY(%s,1)" % (k, start)
    dim = "DAY(EOMONTH(%s,0))" % start
    d = "DATE(CalYear,CalMonth,%s)" % n
    ev = "'%s'" % C.SHEET_NAMES["events"]
    first, last = C.ROW_FIRST, C.CAP["events"] + C.ROW_FIRST - 1
    edates = "%s!$E$%d:$E$%d" % (ev, first, last)
    eclient = "%s!$D$%d:$D$%d" % (ev, first, last)
    etype = "%s!$F$%d:$F$%d" % (ev, first, last)
    estat = "%s!$Q$%d:$Q$%d" % (ev, first, last)
    hit = 'MATCH(1,INDEX((%s=%s)*(%s<>"%s"),0),1)' % (edates, d, estat,
                                                      C.ES_CANCEL)
    label = ('IFERROR(INDEX(%s,%s)&" \u2022 "&INDEX(%s,%s),"")'
             % (eclient, hit, etype, hit))
    return '=IF(OR(%s<1,%s>%s),"",%s&IF(%s="","",CHAR(10)&%s))' % (
        n, n, dim, n, label, label)


def _demo_cells(m):
    """Cached cell text for the sample month (September 2026)."""
    out = {}
    if not m.events:
        return out
    cm, cy = m.settings["cal_month"], m.settings["cal_year"]
    first = date(cy, cm, 1)
    offset = (first.weekday() + 1) % 7        # Sunday-first offset
    dim = (date(cy + (cm == 12), (cm % 12) + 1, 1) -
           date(cy, cm, 1)).days
    by_day = {}
    for e in m.events:
        if e["date"].year == cy and e["date"].month == cm and \
                e["status"] != C.ES_CANCEL:
            by_day.setdefault(e["date"].day, e)
    for k in range(42):
        day = k - offset + 1
        if 1 <= day <= dim:
            e = by_day.get(day)
            text = str(day)
            if e:
                text += "\n%s \u2022 %s" % (e["client"], e["type"])
            out[k] = text
        else:
            out[k] = ""
    return out


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY, {"A": 2.2, "B": 15.5, "C": 15.5, "D": 15.5, "E": 15.5,
                    "F": 15.5, "G": 15.5, "H": 15.5})
    foot = GRID + 6 + 12
    bk.paint(KEY, 0, 0, foot, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F4C6  Event Calendar",
        "  Change the month on \u2699\uFE0F Setup (Calendar month / year) "
        "and this grid redraws itself from the Events tab.",
        LAST_COL)

    # ------------------------------------------------------------------
    # stat chips: what the displayed month holds
    # ------------------------------------------------------------------
    start = "DATE(CalYear,CalMonth,1)"
    end = "EOMONTH(DATE(CalYear,CalMonth,1),0)"
    ev = bk.q("events")
    edates = bk.rng("events", "date")
    eprice = bk.rng("events", "price")
    eguests = bk.rng("events", "guests")
    estat = bk.rng("events", "status")
    demo_cells = _demo_cells(m)
    n_events = sum(1 for k, v in demo_cells.items() if "\n" in v)
    guests = 0
    booked = 0
    if m.events:
        cm, cy = m.settings["cal_month"], m.settings["cal_year"]
        for e in m.events:
            if e["date"].year == cy and e["date"].month == cm and \
                    e["status"] != C.ES_CANCEL:
                guests += e["guests"]
                booked += e["price"]
    chips = [
        ('\U0001F4C5 "&TEXT(%s,"mmmm")&": "&COUNTIFS(%s,">="&%s,%s,"<="&%s,'
         '%s,"<>%s")&" event(s)"'
         % (start, edates, start, edates, end, estat, C.ES_CANCEL),
         "primary",
         "%s: %d event(s)" % (C.MONTH_NAMES[m.settings["cal_month"] - 1],
                              n_events), 2),
        ('\U0001F465 Guests: "&SUMIFS(%s,%s,">="&%s,%s,"<="&%s,%s,"<>%s")'
         % (eguests, edates, start, edates, end, estat, C.ES_CANCEL),
         "info", "Guests: %d" % guests, 2),
        ('\U0001F4B0 Booked: "&Currency&TEXT(SUMIFS(%s,%s,">="&%s,%s,'
         '"<="&%s,%s,"<>%s"),"#,##0")'
         % (eprice, edates, start, edates, end, estat, C.ES_CANCEL),
         "ok", "Booked: %s" % m.money(booked), 2),
    ]
    col = 1
    for formula, color, cached, span in chips:
        fmt = S.pill(th.soft(color), getattr(th, color), size=10.5,
                     bold=True, align="left")
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), col + span - 1,
                       "", fmt)
        ws.write_formula(r(C.ROW_STATS), col, '="' + formula, fmt, cached)
        bk.stats["formulas"] += 1
        col += span
    if col < ci(LAST_COL):
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), ci(LAST_COL),
                       "", S.canvas)
    elif col == ci(LAST_COL):
        ws.write_blank(r(C.ROW_STATS), col, None, S.canvas)

    # ------------------------------------------------------------------
    # weekday header + grid
    # ------------------------------------------------------------------
    hdr = GRID - 1
    ws.set_row(r(hdr), 20)
    for d, name in enumerate(WEEKDAYS):
        weekend = d in (0, 6)
        ws.write(r(hdr), 1 + d, name, S.header(th.primary_2 if weekend
                                               else th.primary))
    day_fmt = S.f(**S.base(font_size=11, bold=True, font_color=th.ink,
                           bg_color=th.card, align="left", valign="top",
                           text_wrap=True, border=1, border_color=th.border,
                           indent=1))
    for k in range(42):
        row = GRID + k // 7
        col = 1 + k % 7
        ws.set_row(r(row), C.CAL_CELL_H)
        cached = demo_cells.get(k, "") if m.events else ""
        ws.write_formula(r(row), col, _cell_formula(k), day_fmt, cached)
        bk.stats["formulas"] += 1

    # event days glow
    first_cell_r, first_cell_c = GRID, 1
    last_cell_r, last_cell_c = GRID + 5, 7
    bk.cond(KEY, first_cell_r, first_cell_c, last_cell_r, last_cell_c, {
        "type": "formula",
        "criteria": '=ISNUMBER(SEARCH("\u2022",B%d))' % GRID,
        "format": S.cf(bg=th.gold_soft, fg=th.ink, bold=True,
                       border=th.gold)})

    # ------------------------------------------------------------------
    # legend + notes
    # ------------------------------------------------------------------
    leg = GRID + 6 + 1
    ws.set_row(r(leg), 22)
    ws.merge_range(r(leg), 1, r(leg), ci(LAST_COL),
                   "  \U0001F511  LEGEND", S.section_soft)
    legend = [
        ("Gold cell", "an event is booked that day (client \u2022 type)",
         th.gold_soft, th.ink),
        ("\U0001F4AC / \U0001F4DD / \U0001F4B0 / \u2705",
         "pipeline stage colours live on the \U0001F4C5 Events tab",
         th.primary_soft, th.primary),
        ("Cancelled events", "are left off the calendar on purpose",
         th.bad_soft, th.bad),
    ]
    for j, (chip, text, bg, fg) in enumerate(legend):
        row = leg + 1 + j
        ws.set_row(r(row), 20)
        ws.write(r(row), 1, chip, S.pill(bg, fg, size=10, align="center"))
        ws.merge_range(r(row), 2, r(row), ci(LAST_COL), text,
                       S.f(**S.base(font_size=10, font_color=th.ink,
                                    bg_color=th.card, align="left",
                                    valign="vcenter", indent=1)))

    note = leg + 5
    ws.set_row(r(note), 20)
    ws.set_row(r(note + 1), 20)
    ws.merge_range(r(note), 1, r(note + 1), ci(LAST_COL),
                   "  \U0001F4A1  Two events on one day? The cell shows the "
                   "first one - the \U0001F4C5 Events tab (sorted by date) "
                   "shows everything. Print this page and pin it to the "
                   "kitchen wall.", S.note)

    nav = note + 3
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)
    bk.page(KEY, LAST_COL, nav + 1, zoom=95)
