"""🎒 Packing Checklist - event-specific list with a % packed meter."""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "packing"
LAST_COL = "F"


def build(bk):
    th = bk.th
    ws = bk.ws(KEY)
    m = bk.demo
    K.sheet_head(bk, KEY, LAST_COL,
                 "\U0001F392  Packing Checklist",
                 "Pick the fair you are packing for, tick as you load - "
                 "the % packed meter fills in as you go.")
    bk.paint(KEY, 0, 0, 60, ci(LAST_COL), bk.S.canvas)

    lab = bk.S.f(**bk.S.base(font_size=11, font_color=th.ink,
                             bg_color=th.card, align="left",
                             valign="vcenter", indent=1, border=1,
                             border_color=th.border))
    pick = bk.S.f(**bk.S.base(font_size=12, bold=True,
                              font_color=th.primary,
                              bg_color=th.gold_soft, border=1,
                              border_color=th.border_strong, align="left",
                              valign="vcenter", indent=1, locked=False))
    item = bk.S.f(**bk.S.base(font_size=10.5, font_color=th.ink,
                              bg_color=th.card, align="left",
                              valign="vcenter", indent=1, border=1,
                              border_color=th.border))
    note = bk.S.f(**bk.S.base(font_size=10.5, font_color=th.muted,
                              bg_color=th.card, align="left",
                              valign="vcenter", indent=1, border=1,
                              border_color=th.border, locked=False))
    tick = bk.S.f(**bk.S.base(font_size=13, bold=True, font_color=th.ok,
                              bg_color=th.card, align="center",
                              valign="vcenter", border=1,
                              border_color=th.border, locked=False))

    ws.set_row(r(C.PACK_EVENT_ROW), 26)
    ws.merge_range(r(C.PACK_EVENT_ROW), 1, r(C.PACK_EVENT_ROW), 2,
                   "  \U0001F3EA  Packing for:", lab)
    ws.merge_range(r(C.PACK_EVENT_ROW), 3, r(C.PACK_EVENT_ROW),
                   ci(LAST_COL), "", pick)
    bk.validate(KEY, C.PACK_EVENT_ROW, 3, C.PACK_EVENT_ROW, 3,
                "=" + bk.listname("events"), title="Pick a fair")
    if m and m.events:
        from ..demo import today as _today
        nxt = [e for e in m.events if e["date"] >= _today()]
        ws.write(r(C.PACK_EVENT_ROW), 3,
                 (nxt[0]["name"] if nxt else m.events[-1]["name"]), pick)

    rows = []
    row = C.PACK_EVENT_ROW + 2
    for title, items in C.PACK_SECTIONS:
        ws.set_row(r(row), 22)
        ws.merge_range(r(row), 1, r(row), ci(LAST_COL), "",
                       bk.S.section_soft)
        ws.write_formula(r(row), 1,
                         '=CONCATENATE("  \U0001F9FA  %s   \u2014   "&'
                         'COUNTIF($B$%d:$B$%d,"\u2713")&" / %d packed")'
                         % (title.upper(), row + 1, row + len(items),
                            len(items)),
                         bk.S.section_soft,
                         "  \U0001F9FA  %s   \u2014   %d / %d packed"
                         % (title.upper(),
                            sum(1 for it in items
                                if m and m.packing.get(it)),
                            len(items)))
        bk.stats["formulas"] += 1
        row += 1
        for it in items:
            ws.set_row(r(row), 20)
            ws.write(r(row), 1, "", tick)
            ws.merge_range(r(row), 2, r(row), 3, "  " + it, item)
            ws.merge_range(r(row), 4, r(row), ci(LAST_COL), "", note)
            if m and m.packing.get(it):
                ws.write(r(row), 1, C.TICK, tick)
            rows.append(row)
            row += 1
        row += 1

    first, last = rows[0], rows[-1]
    bk.validate(KEY, first, 1, last, 1, "=Tick", title="Packed?")
    bk.cond(KEY, first, 1, last, 1, {
        "type": "formula", "criteria": '=$B%d=""' % first,
        "format": bk.S.cf(bg=th.card, fg=th.ok)})
    bk.cond(KEY, first, 1, last, 3, {
        "type": "formula", "criteria": '=$A%d="%s"' % (first, C.TICK),
        "format": bk.S.cf(bg=th.ok_soft, fg=th.ok, strike=True)})

    chips = [
        ('="\U0001F392 Packed: "&COUNTIF($B$%d:$B$%d,"\u2713")&" of %d"'
         % (first, last, len(rows)), "primary",
         "Packed: %d of %d" % (m.agg.get("packed", 0) if m else 0,
                               len(rows)), 4),
        ('="\U0001F4CA Progress: "&TEXT(COUNTIF($B$%d:$B$%d,"\u2713")/%d,'
         '"0%%")' % (first, last, len(rows)), "ok",
         "Progress: %.0f%%" % (100.0 * (m.agg.get("packed", 0) / len(rows)
                                        if m else 0)), 3),
        ('="\u23F3 Still to pack: "&%d-COUNTIF($B$%d:$B$%d,"\u2713")'
         % (len(rows), first, last), "warn",
         "Still to pack: %d" % (len(rows) - (m.agg.get("packed", 0)
                                             if m else 0)), 4),
    ]
    K.chips(bk, KEY, chips, LAST_COL)
    K.footer_nav(bk, KEY, last + 3, LAST_COL, landscape=False, tip=
                 "  \U0001F4A1  Print this tab (File \u2192 Print) and clip "
                 "it to the van door.  Ticks save mornings.")
