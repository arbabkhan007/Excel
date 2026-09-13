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
                       "  your 5-minute tour of the Catering Business "
                       "Manager",
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
                       "  Everything is wired together: log an event once "
                       "and the quotes, shopping list, reports and "
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
        ("1", "Open \u2699\uFE0F Setup",
         "Type your business name, pick your currency symbol, and set your "
         "default margin, deposit % and tax rate.  Every tab in this "
         "workbook reads those numbers."),
        ("2", "Skim the dropdown lists",
         "Lower down on \u2699\uFE0F Setup you can rename your event types, "
         "expense categories, payment methods and staff roles.  Every "
         "dropdown in the workbook updates the moment you edit them \u2014 "
         "nothing is hard-coded."),
        ("3", "Add your clients and events",
         "\U0001F465 Clients holds one row per client; \U0001F4C5 Events "
         "holds one row per job.  Pick the client from the dropdown on the "
         "Events tab and the two tabs stay in step forever."),
        ("4", "Price the job",
         "Type the guest count and your costs \u2014 the recommended price, "
         "price per guest, deposit and balance appear instantly at your "
         "target margin.  Copy the recommended price into the event's Price "
         "column."),
        ("5", "Log money as it moves",
         "\U0001F4B8 Expenses for what you spend, \U0001F4B0 Payments for "
         "what clients pay.  The dashboard, P&L and tax tracker build "
         "themselves from those two logs."),
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
        ("\u25CF", "\U0001F4CA Dashboard", "Your command center \u2014 "
         "revenue, profit, margin, pipeline and the next money due, live."),
        ("\u25CF", "\U0001F465 Clients", "One row per client with contact "
         "details, package, quote, deposit and what they still owe."),
        ("\u25CF", "\U0001F4C5 Events", "The master planner: every job from "
         "enquiry to completed, with menu, crew and equipment notes."),
        ("\u25CF", "\U0001F9EE Quote Calculator", "Costs in, recommended "
         "price out \u2014 the fastest way to stop under-pricing a job."),
        ("\u25CF", "\U0001F4B8 Expenses", "Every purchase tagged by category "
         "and event \u2014 the fuel for your P&L and tax numbers."),
        ("\u25CF", "\U0001F4B0 Payments", "Invoices and deposits received, "
         "with an automatic OVERDUE flag when a due date slips."),
        ("\u25CF", "\U0001F4D6 Start Here", "This page."),
        ("\u25CB", "\U0001F37D\uFE0F Menu Costing", "Price every dish from "
         "its ingredients \u2014 the recipe calculator works out cost per "
         "portion for you."),
        ("\u25CB", "\U0001F4E6 Inventory", "What is in the store room, what "
         "it is worth and what to reorder before it runs out."),
        ("\u25CB", "\U0001F6D2 Shopping List", "Auto-built from your events: "
         "what you need, what you have, what to buy and what it costs."),
        ("\u25CB", "\U0001F477 Staff & Labor", "Shifts, hours, overtime and "
         "who has been paid \u2014 labour cost lands in reports for you."),
        ("\u25CB", "\U0001F373 Equipment", "Chafers, urns and glassware: "
         "owned, reserved, in service, and replacement cost."),
        ("\u25CB", "\U0001F69A Suppliers", "Your vendor book \u2014 contacts, "
         "terms and what each one supplies."),
        ("\u25CB", "\U0001F4C6 Event Calendar", "A month-at-a-glance wall "
         "calendar of every booked job, any month, any year."),
        ("\u25CB", "\U0001F4C8 P&L & Reports", "Monthly and annual profit & "
         "loss plus revenue by month, by type, best dishes and clients."),
        ("\u25CB", "\U0001F9FE Tax Tracker", "Sales tax collected vs paid "
         "with quarterly estimates and a log for your accountant."),
        ("\u25CB", "\u2705 Checklists", "Four printable checklists: prep, "
         "shopping, day-of and end-of-event."),
        ("\u25CB", "\U0001F5A8\uFE0F Invoice & Proposal", "A print-ready "
         "invoice and proposal \u2014 pick the event and it fills itself."),
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
        "Price from costs, not from competitors: if food is more than about "
        "30% of your quote, the margin column on \U0001F37D\uFE0F Menu "
        "Costing warns you before the client ever sees a number.",
        "Take the deposit before you book suppliers \u2014 the pipeline strip "
        "on the dashboard shows exactly which jobs are still only talk.",
        "Log expenses on the day you spend them; a Friday-night receipt "
        "typed in on Monday is how margins disappear.",
        "Keep \U0001F4E6 Inventory honest: the LOW STOCK flag only works if "
        "you tick items back in when a delivery lands.",
        "Use the \U0001F6D2 Shopping List per event \u2014 it subtracts what "
        "is already in the store room so you never buy saffron twice.",
        "The OVERDUE flag on \U0001F4B0 Payments is automatic \u2014 chase "
        "anything red before you start the next job for that client.",
        "Set your calendar month on \u2699\uFE0F Setup and the \U0001F4C6 "
        "Event Calendar re-draws itself for any month of any year \u2014 "
        "reuse this workbook season after season.",
        "Renaming an event type on \u2699\uFE0F Setup renames it everywhere, "
        "including the revenue-by-type chart on the dashboard.",
        "Print \U0001F5A8\uFE0F Invoice & Proposal to PDF (File \u2192 Print "
        "\u2192 Save as PDF) for a client-ready document without leaving the "
        "spreadsheet.",
        "Back up before big edits: a copy of the file on your desktop costs "
        "nothing and saves entire weekends.",
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
        "The EXAMPLE edition is loaded with a fictional catering business so "
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
                   "  Thank you for buying this template \u2014 it was built "
                   "by people who have plated four hundred covers at "
                   "midnight and know what a spreadsheet owes you on a busy "
                   "week.  If something looks wrong or you would love an "
                   "extra tab, message the shop on Etsy and we will help "
                   "you quickly.", body_soft)
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
