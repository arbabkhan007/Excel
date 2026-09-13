"""💵 Pricing Calculator - true cost in, suggested price out."""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "pricing"
LAST_COL = "J"


def build(bk):
    th = bk.th
    ws = bk.ws(KEY)
    m = bk.demo
    pr = m.agg.get("pricing", {}) if m else {}
    K.sheet_head(bk, KEY, LAST_COL,
                 "\U0001F4B5  Pricing Calculator",
                 "Never sell a $25 make for $15 again: materials + labour + "
                 "overhead in, suggested price out.")
    bk.paint(KEY, 0, 0, C.PR_LAST_ROW + 8, ci(LAST_COL), bk.S.canvas)

    lab = bk.S.f(**bk.S.base(font_size=11, font_color=th.ink,
                             bg_color=th.card, align="left",
                             valign="vcenter", indent=1, border=1,
                             border_color=th.border))
    inp = bk.S.f(**bk.S.base(font_size=12, bold=True, font_color=th.primary,
                             bg_color=th.gold_soft, border=1,
                             border_color=th.border_strong, align="right",
                             valign="vcenter", num_format="#,##0.00",
                             locked=False))
    inp_pct = bk.S.f(**bk.S.base(font_size=12, bold=True,
                                 font_color=th.primary,
                                 bg_color=th.gold_soft, border=1,
                                 border_color=th.border_strong,
                                 align="right", valign="vcenter",
                                 num_format="0%", locked=False))
    inp_num = bk.S.f(**bk.S.base(font_size=12, bold=True,
                                 font_color=th.primary,
                                 bg_color=th.gold_soft, border=1,
                                 border_color=th.border_strong,
                                 align="right", valign="vcenter",
                                 num_format="#,##0.0", locked=False))
    out = bk.S.f(**bk.S.base(font_size=12, bold=True, font_color=th.primary,
                             bg_color=th.card, align="right",
                             valign="vcenter", border=1,
                             border_color=th.border,
                             num_format="#,##0.00"))
    hero = bk.S.f(**bk.S.base(font_size=16, bold=True, font_color=th.white,
                              bg_color=th.accent, align="right",
                              valign="vcenter", border=1,
                              border_color=th.accent,
                              num_format="#,##0.00"))

    def left(row, label, value, fmt, cached=None):
        ws.set_row(r(row), 22)
        ws.merge_range(r(row), 1, r(row), 4, "  " + label, lab)
        ws.merge_range(r(row), 5, r(row), 6, "", fmt)
        if isinstance(value, str) and value.startswith("="):
            ws.write_formula(r(row), 5, value, fmt, cached)
            bk.stats["formulas"] += 1
        else:
            ws.write(r(row), 5, value, fmt)

    def right(row, label, formula, cached, fmt=None, hero_fmt=False):
        ws.set_row(r(row), 24 if hero_fmt else 20)
        ws.merge_range(r(row), 7, r(row), 8, "  " + label, lab)
        ws.merge_range(r(row), 9, r(row), ci(LAST_COL), "",
                       hero if hero_fmt else (fmt or out))
        ws.write_formula(r(row), 9, formula, hero if hero_fmt
                         else (fmt or out), cached)
        bk.stats["formulas"] += 1

    q = m.pricing if m else {}
    ws.merge_range(r(8), 1, r(8), 6, "  1 \u00B7 WHAT IT COSTS YOU",
                   bk.S.section_soft)
    ws.merge_range(r(8), 7, r(8), ci(LAST_COL), "  2 \u00B7 WHAT TO CHARGE",
                   bk.S.section_soft)
    ws.set_row(r(8), 22)

    left(C.PR_IN["material"], "Materials / yarn cost",
         q.get("material", ""), inp, q.get("material", 0))
    left(C.PR_IN["packaging"], "Packaging & labels",
         q.get("packaging", ""), inp, q.get("packaging", 0))
    left(C.PR_IN["hours"], "Hours to make", q.get("hours", ""), inp_num,
         q.get("hours", 0))
    left(C.PR_IN["wage"], "Your hourly wage", "=HourlyWage", inp,
         (m.settings["wage"] if m else 14.0))
    left(C.PR_IN["overhead"], "Overhead % (tools, fees, power)",
         "=OverheadPct", inp_pct, (m.settings["overhead"] if m else 0.1))
    left(C.PR_IN["margin"], "Target profit margin", "=TargetMargin",
         inp_pct, (m.settings["margin"] if m else 0.45))

    E = "$E$"
    right(C.PR_OUT["labor"], "Labour cost (hours \u00D7 wage)",
          '=IF(%s%d="","",%s%d*%s%d)' % (E, C.PR_IN["hours"], E,
                                         C.PR_IN["hours"], E,
                                         C.PR_IN["wage"]),
          pr.get("labor", 0))
    right(C.PR_OUT["overhead_amt"], "Overhead amount",
          '=IF(%s%d="","",(%s%d+%s%d+%s%d)*%s%d)'
          % (E, C.PR_IN["material"], E, C.PR_IN["material"], E,
             C.PR_IN["packaging"], E, C.PR_OUT["labor"], E,
             C.PR_IN["overhead"]),
          pr.get("overhead_amt", 0))
    right(C.PR_OUT["true_cost"], "TRUE COST",
          '=%s%d+%s%d+%s%d+%s%d' % (E, C.PR_IN["material"], E,
                                    C.PR_IN["packaging"], E,
                                    C.PR_OUT["labor"], E,
                                    C.PR_OUT["overhead_amt"]),
          pr.get("true_cost", 0))
    right(C.PR_OUT["price"], "SUGGESTED PRICE",
          '=IF(OR(%s%d="",1-%s%d<=0),"",ROUND(%s%d/(1-%s%d),2))'
          % (E, C.PR_OUT["true_cost"], E, C.PR_IN["margin"], E,
             C.PR_OUT["true_cost"], E, C.PR_IN["margin"]),
          pr.get("price", 0), hero_fmt=True)
    right(C.PR_OUT["charm"], "Charm price (\u2026.95)",
          '=IF(%s%d="","",ROUNDUP(%s%d,0)-0.05)'
          % (E, C.PR_OUT["price"], E, C.PR_OUT["price"]),
          pr.get("charm", 0))
    right(C.PR_OUT["profit"], "Profit at suggested price",
          '=IF(%s%d="","",%s%d-%s%d)' % (E, C.PR_OUT["price"], E,
                                         C.PR_OUT["price"], E,
                                         C.PR_OUT["true_cost"]),
          pr.get("profit", 0))
    right(C.PR_OUT["check"], "Sanity check",
          '=IF(%s%d="","",IF(%s%d<%s%d,"\u26A0\uFE0F below true cost!",'
          '"\u2705 healthy margin"))'
          % (E, C.PR_OUT["price"], E, C.PR_OUT["price"], E,
             C.PR_OUT["true_cost"]),
          "\u2705 healthy margin" if pr else "",
          fmt=bk.S.f(**bk.S.base(font_size=11, bold=True,
                                 font_color=th.ok, bg_color=th.ok_soft,
                                 align="center", valign="vcenter",
                                 border=1, border_color=th.border)))

    K.note_block(bk, KEY, C.PR_LAST_ROW - 4, "B", LAST_COL, [
        "Rules of thumb from market sellers:",
        "1.  Materials = yarn + labels + packaging for ONE item.",
        "2.  Overhead covers hooks, stall extras, power and marketplace "
        "fees - 10-15% is typical.",
        "3.  If the suggested price feels too high for your market, cut "
        "the hours (simpler pattern) before you cut your wage.",
    ], title="  \U0001F4A1  PRICING LIKE A PRO")
    K.footer_nav(bk, KEY, C.PR_LAST_ROW + 2, LAST_COL, landscape=False)
