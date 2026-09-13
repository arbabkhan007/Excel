"""📖 Start Here - quick start, tab tour, tips and the friendly rules."""

import os

from ..book import r, ci

LAST_COL = "J"


def build(bk):
    key = "guide"
    ws = bk.ws(key)
    S, th = bk.S, bk.th

    bk.widths(key, {"A": 2.2, "B": 4, "C": 15, "D": 15, "E": 15, "F": 15,
                    "G": 15, "H": 15, "I": 15, "J": 15})
    bk.paint(key, 0, 0, 90, ci(LAST_COL), S.canvas)

    body = S.f(**S.base(font_size=10.5, bg_color=th.bg, align="left",
                        valign="vcenter", text_wrap=True))
    body_soft = S.f(**S.base(font_size=10.5, bg_color=th.card, align="left",
                             valign="vcenter", text_wrap=True,
                             border=1, border_color=th.border))
    badge = S.f(**S.base(bold=True, font_size=12, font_color=th.white,
                         bg_color=th.accent, align="center", valign="vcenter"))
    badge2 = S.f(**S.base(bold=True, font_size=12, font_color=th.white,
                          bg_color=th.info, align="center", valign="vcenter"))
    step_title = S.f(**S.base(bold=True, font_size=11.5,
                              font_color=th.primary, bg_color=th.bg,
                              align="left", valign="vcenter", indent=1))
    mark = S.f(**S.base(font_size=10, font_color=th.muted, bg_color=th.bg,
                        align="center", valign="vcenter"))

    def sec(row, emoji, text):
        ws.set_row(r(row), 24)
        ws.merge_range(r(row), 1, r(row), ci(LAST_COL),
                       "  %s  %s" % (emoji, text), S.section)

    # ------------------------------------------------------------------
    # cover banner + overlaid title (plain title band when no artwork)
    # ------------------------------------------------------------------
    image = _banner_path(bk)
    row = 1
    ws.set_row(r(row), 7)
    row += 1
    if image:
        w, h = _png_size(image)
        scale = 860.0 / float(w)
        span = int(h * scale / 20.0) + 1
        cover = S.f(**S.base(bg_color=th.cover_bg))
        for i in range(span):
            ws.set_row(r(row + i), 20)
            for col in range(1, ci(LAST_COL) + 1):
                ws.write(r(row + i), col, "", cover)
        ws.insert_image(r(row), 1, image,
                        {"x_scale": scale, "y_scale": scale,
                         "object_position": 1})
        t_row = row + max(1, span // 2 - 1)
        ws.merge_range(r(t_row), ci("C"), r(t_row), ci("G"),
                       "\U0001F4D6  START HERE",
                       S.f(**S.base(font_name=th.title_font, font_size=26,
                                    bold=True, font_color=th.primary,
                                    bg_color=th.cover_bg, align="center",
                                    valign="vcenter")))
        ws.merge_range(r(t_row + 1), ci("C"), r(t_row + 1), ci("G"),
                       "  your 5-minute tour of the Crochet Craft Fair "
                       "Tracker",
                       S.f(**S.base(font_size=11.5, italic=True,
                                    font_color=th.muted,
                                    bg_color=th.cover_bg, align="center",
                                    valign="vcenter")))
        row += span + 1
        ws.set_row(r(row), 18)
        ws.merge_range(r(row), ci(LAST_COL) - 2, r(row), ci(LAST_COL), "",
                       S.home_link)
        ws.write_url(r(row), ci(LAST_COL) - 2,
                     "internal:%s!A1" % bk.q("dashboard"), S.home_link,
                     "\U0001F3E0  Back to Dashboard")
        row += 2
    else:
        ws.set_row(r(row), 32)
        ws.set_row(r(row + 1), 18)
        ws.merge_range(r(row), 1, r(row), ci(LAST_COL) - 3,
                       "\U0001F4D6  Start Here \u2014 your 5-minute setup",
                       S.sheet_title)
        ws.merge_range(r(row + 1), 1, r(row + 1), ci(LAST_COL) - 3,
                       "  Everything is wired together: log a sale once "
                       "and the fairs, reorder list, monthly summary and "
                       "dashboard follow.", S.sheet_sub)
        ws.merge_range(r(row + 1), ci(LAST_COL) - 2, r(row + 1), ci(LAST_COL),
                       "", S.home_link)
        ws.write_url(r(row + 1), ci(LAST_COL) - 2,
                     "internal:%s!A1" % bk.q("dashboard"), S.home_link,
                     "\U0001F3E0  Back to Dashboard")
        row += 3
    # ------------------------------------------------------------------
    sec(row, "\U0001F680", "QUICK START \u2014 FIVE STEPS TO LIVE")
    row += 1
    steps = [
        ("1", "Open \u2699\uFE0F Lists & Settings",
         "Type your studio name, currency, hourly wage, overhead % and "
         "target margin.  Every calculator and alert reads those numbers."),
        ("2", "Skim your dropdown lists",
         "Lower down on the same tab you can rename categories, payment "
         "methods, yarn weights, units and suppliers.  Every dropdown in "
         "the workbook updates instantly \u2014 nothing is hard-coded."),
        ("3", "List your makes on \U0001F9F6 Product Catalog",
         "One row per design: price, yarn cost, packaging, hours and a "
         "minimum stock level.  Profit and margin calculate themselves."),
        ("4", "Log yarn on \U0001F9F5 Yarn & Materials",
         "What you bought, what you have used, what is left \u2014 the "
         "reorder flag turns red before you run out mid-commission."),
        ("5", "Work your fairs",
         "\U0001F3EA Craft Fairs takes one row per market, \U0001F4B0 "
         "Sales Log one row per sale.  Profit, best-sellers, the monthly "
         "summary and the dashboard build themselves from those two."),
    ]
    for num, title, text in steps:
        ws.set_row(r(row), 16)
        ws.set_row(r(row + 1), 16)
        ws.merge_range(r(row), 1, r(row + 1), 1, num, badge)
        ws.merge_range(r(row), 2, r(row), 3, title, step_title)
        ws.merge_range(r(row), 4, r(row + 1), ci(LAST_COL), text, body)
        row += 2
    row += 1

    # ------------------------------------------------------------------
    sec(row, "\U0001F5FA\uFE0F", "THE TAB TOUR")
    row += 1
    ws.merge_range(r(row), 1, r(row), ci(LAST_COL),
                   "  Premium edition: every tab below.  Basic edition: the "
                   "tabs marked \u25CF only.",
                   S.f(**S.base(font_size=10, italic=True, font_color=th.muted,
                                bg_color=th.bg, align="left", valign="vcenter",
                                indent=1)))
    row += 1
    tour = [
        ("\u25CF", "\U0001F4CA Dashboard", "Sales, profit, margin, units, "
         "best-seller, six charts and live low-stock panels."),
        ("\u25CF", "\U0001F9F6 Product Catalog", "Every make with cost, "
         "profit, margin, stock and reorder status."),
        ("\u25CF", "\U0001F3EA Craft Fairs", "One row per market with all "
         "five cost lines, takings, net profit and best seller."),
        ("\u25CF", "\U0001F4B0 Sales Log", "Every sale: fair and product "
         "dropdowns, discount, payment method and live cost of sale."),
        ("\u25CF", "\U0001F4B5 Pricing Calculator", "Materials + labour + "
         "overhead in, suggested and charm price out."),
        ("\u25CF", "\U0001F4D6 Start Here", "This page."),
        ("\u25CB", "\U0001F9F5 Yarn & Materials", "Yarn and findings with "
         "remaining quantity and an automatic reorder flag."),
        ("\u25CB", "\U0001F4E6 Made & Stocked", "Production batches: made, "
         "reserved, sold, damaged, current and available."),
        ("\u25CB", "\U0001F9EE Event Profit", "Pick a fair - revenue, every "
         "cost, net profit, margin, break-even and ROI."),
        ("\u25CB", "\U0001F504 Reorder List", "Low products and yarn pulled "
         "in automatically, with priority and estimated budget."),
        ("\u25CB", "\U0001F392 Packing Checklist", "Four packing zones with "
         "ticks and a % packed meter, per fair."),
        ("\u25CB", "\U0001F4C5 Monthly Summary", "The whole year month by "
         "month: revenue, costs, profit, units and markets."),
    ]
    for mark_txt, tab, blurb in tour:
        ws.set_row(r(row), 18)
        ws.merge_range(r(row), 2, r(row), 3, mark_txt + "  " + tab, step_title)
        ws.merge_range(r(row), 4, r(row), ci(LAST_COL), blurb, body)
        row += 1
    row += 1

    # ------------------------------------------------------------------
    sec(row, "\U0001F4A1", "TEN TIPS FROM WORKING CATERERS")
    row += 1
    tips = [
        "Price from true cost, not from the stall next door: the Pricing "
        "Calculator adds your wage and overhead before margin.",
        "Log sales at the fair on your phone (Google Sheets) and the "
        "dashboard is already correct by teardown time.",
        "Set a minimum stock level for every best-seller \u2014 the "
        "Reorder List only protects what you have told it to protect.",
        "Weigh your yarn usage once per pattern (used column) and future "
        "costs stay honest without counting every metre.",
        "Charm prices (\u2026.95) test better at markets than round "
        "numbers; the calculator suggests one for you.",
        "After each fair, glance at Event Profit: any market under ~25% "
         "margin two years running is a hobby, not a channel.",
        "The Monthly Summary reads the report year from Settings \u2014 "
         "change it in January and reuse this file forever.",
        "Print the Packing Checklist per fair; ticks beat memory at 6am.",
        "Back up before big edits: a copy on your desktop costs nothing "
        "and saves entire weekends.",
        "The EXAMPLE file is a fictional studio so you can see it working; "
        "the blank file is the one you keep.",
    ]
    for i, tip in enumerate(tips, 1):
        ws.set_row(r(row), 16)
        ws.set_row(r(row + 1), 16)
        ws.merge_range(r(row), 1, r(row + 1), 1, str(i), badge2)
        ws.merge_range(r(row), 2, r(row + 1), ci(LAST_COL), tip, body)
        row += 2
    row += 1

    # ------------------------------------------------------------------
    sec(row, "\U0001F512", "THE GENTLE RULES")
    row += 1
    rules = [
        "Cream-filled cells are yours to type in.  White cells are formulas "
        "\u2014 leave them alone and they will keep working for you.",
        "Sheets are protected to stop accidental edits.  Review \u2192 "
        "Unprotect Sheet (no password) if you ever need to restructure "
        "something.",
        "Add rows by copying an existing data row and inserting below it "
        "\u2014 the formulas, dropdowns and colours travel with the copy.",
        "This file opens in Excel 2016+ and in Google Sheets (upload to "
        "Drive \u2192 open with Sheets).  A few chart styles look slightly "
        "different in Sheets; every number still calculates.",
        "The EXAMPLE edition is loaded with a fictional crochet studio so "
        "you can see it working.  The blank edition is the one you keep.",
    ]
    for rule in rules:
        ws.set_row(r(row), 16)
        ws.set_row(r(row + 1), 16)
        ws.merge_range(r(row), 1, r(row + 1), ci(LAST_COL),
                       "\u2022  " + rule, body_soft)
        row += 2
    row += 1

    # ------------------------------------------------------------------
    sec(row, "\U0001F91D", "SUPPORT & GOOD WISHES")
    row += 1
    ws.set_row(r(row), 16)
    ws.set_row(r(row + 1), 16)
    ws.merge_range(r(row), 1, r(row + 1), ci(LAST_COL),
                   "  Thank you for buying this template \u2014 it was "
                   "built by people who have frogged a blanket at 2am and "
                   "know what a spreadsheet owes you between markets.  If "
                   "something looks wrong or you would love an extra tab, "
                   "message the shop on Etsy and we will help quickly.",
                   body_soft)
    row += 3

    # ------------------------------------------------------------------
    # footer + nav
    # ------------------------------------------------------------------
    ws.set_row(r(row), 30)
    ws.merge_range(r(row), 1, r(row), ci(LAST_COL),
                   "  \U0001F4AC  Questions about this file?  Your Etsy "
                   "shop message reaches a human, usually the same day.",
                   S.f(**S.base(font_size=10, italic=True, font_color=th.ink,
                                bg_color=th.gold_soft, align="left",
                                valign="vcenter", indent=1)))
    nav = row + 2
    ws.set_row(r(nav), 24)
    bk.nav_row(key, nav, max_col=LAST_COL)
    bk.page(key, LAST_COL, nav + 1, landscape=False, zoom=90)


# ---------------------------------------------------------------------------
def _banner_path(bk):
    """Watercolour cover art for the Start Here tab, if it exists."""
    if not bk.images:
        return None
    here = os.path.dirname(os.path.dirname(os.path.dirname(
        os.path.abspath(__file__))))
    for ext in (".png", ".jpg", ".jpeg"):
        path = os.path.join(here, "assets", "banner_%s%s" % (bk.th.key, ext))
        if os.path.exists(path):
            return path
    return None


def _png_size(path):
    """Pixel size of a PNG or JPEG, without needing Pillow."""
    with open(path, "rb") as fh:
        data = fh.read(32)
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        import struct
        w, h = struct.unpack(">II", data[16:24])
        return w, h
    return 1600, 300
