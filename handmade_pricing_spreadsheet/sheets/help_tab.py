"""`Help` - troubleshooting, a glossary and a map of every feature."""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Help"
LAST_COL = 9


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.ink)

    s.banner("Help & Troubleshooting",
             "What the colours mean, what to do when a number looks "
             "wrong, and where every feature lives.", LAST_COL)

    s.set_col(0, 0, 4)
    s.set_col(1, 1, 30)
    s.set_col(2, LAST_COL, 14)

    row = 4

    # ------------------------------------------------ the colour code --
    row = s.section(row, "The colour code", LAST_COL)
    row = s.table_head(row, 1, ["Looks like", "Means"], [30, 68])
    colours = [
        ("White cell with a coloured edge",
         "Type here. These are the only cells you need to fill in."),
        ("Shaded cell", "A formula. It is locked so you cannot "
                        "overwrite it by accident."),
        ("Green", "Good - on target, profitable, above your "
                  "minimum wage."),
        ("Amber", "Worth a look - thin margin, stock low, research "
                  "going stale."),
        ("Red", "Act today - losing money, below minimum wage, out "
                "of stock."),
        ("Blue", "For information - a position, a status, a note."),
        ("A coloured bar inside a cell",
         "A data bar: the longer the bar, the bigger the number "
         "compared with the rest of the column."),
        ("A tiny chart", "A sparkline - twelve months of that "
                         "number in one cell."),
    ]
    for look, means in colours:
        s.write(row, 1, look, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, means, fmts.chrome("label"))
        s.row_height(row, 22)
        row += 1
    row += 1

    # ----------------------------------------------- troubleshooting ---
    row = s.section(row, "When something looks wrong", LAST_COL)
    row = s.table_head(row, 1, ["What you see", "Why", "The fix"],
                       [26, 36, 36])
    problems = [
        ("'Not found' in a name column",
         "The ID you typed is not on the tab it points at",
         "Pick from the drop-down instead of typing"),
        ("#DIV/0!",
         "Something is being divided by zero - usually a price or "
         "a pack size that is still empty",
         "Fill in the missing number"),
        ("#VALUE!",
         "Text where a number is expected, often from a pasted "
         "value with a currency symbol in it",
         "Paste as values and strip the symbols"),
        ("A price of 'Check the settings'",
         "Your target margin plus the fee rate is 100% or more, so "
         "no price can work",
         "Lower the target margin or check the fee rates"),
        ("'No price works'",
         "The same thing on the Retail Calculator",
         "Settings, then target margin; or Channels, then fees"),
        ("Suggested price looks far too high",
         "Usually the labour minutes - a figure in hours typed into "
         "a minutes box",
         "Check the Labor tab for that product"),
        ("Suggested price looks far too low",
         "Usually a missing cost: no packaging, no overhead or no "
         "labour at all",
         "Open the Retail Calculator and read the cost block"),
        ("Cost per gram looks wrong",
         "The pack size and the unit do not match - 100 g entered "
         "as 1 pack, or ml against a gram unit",
         "Materials tab: check pack quantity and unit"),
        ("A product shows no cost",
         "Nothing on the Material Usage or Labor tabs refers to it",
         "Add the bill of materials and the minutes"),
        ("Two sets of costs on one product",
         "It has rows on a craft tab and on Material Usage",
         "Use one or the other - the craft tab's check column "
         "warns you"),
        ("Margin is lower than you set",
         "The target is applied before fees; the fees then come "
         "off the price",
         "That is correct - the price already covers them"),
        ("Numbers do not update",
         "Calculation has been switched to manual",
         "Formulas, Calculation Options, Automatic - or press F9"),
        ("A cell will not let you type",
         "It is a formula and the sheet is protected",
         f"Review, Unprotect Sheet. The password is "
         f"'{cfg.password}'"),
        ("Drop-down list is missing an option",
         "The lists live on one tab",
         "Lists tab - add your option and it appears everywhere"),
    ]
    for what, why, fix in problems:
        s.write(row, 1, what, fmts.chrome("label_bold"))
        s.merge(row, 2, 5, why, fmts.chrome("label"))
        s.merge(row, 6, LAST_COL, fix, fmts.chrome("note"))
        s.row_height(row, 26)
        row += 1
    row += 1

    # ---------------------------------------------------- the glossary -
    row = s.section(row, "The words, in plain English", LAST_COL)
    row = s.table_head(row, 1, ["Term", "What it means here"],
                       [26, 72])
    glossary = [
        ("Cost", "Everything one piece costs you: materials, "
                 "labour, packaging, the postage you absorb and a "
                 "share of your overhead."),
        ("Margin", "Profit as a share of the price. $10 profit on a "
                   "$40 sale is a 25% margin."),
        ("Markup", "Profit as a share of the cost. The same $10 on "
                   "$30 of cost is a 33% markup. Always a bigger "
                   "number than the margin."),
        ("Contribution", "What one sale puts towards your fixed "
                         "costs: price less fees less the costs that "
                         "only happen when you make one."),
        ("Break-even", "How many you have to sell to cover the "
                       "fixed costs. Fixed costs divided by "
                       "contribution, rounded up."),
        ("Overhead", "Costs that are not about any one product - "
                     "studio rent, insurance, software, your website."),
        ("Base unit", "The unit a material is really used in: grams, "
                      "millilitres or centimetres. Costs are worked "
                      "out per base unit so quantities are simple."),
        ("Waste allowance", "The share of a material that never "
                            "makes it into the product - offcuts, "
                            "frogging, trim loss."),
        ("Percentage fees", "Fees charged as a share of the sale: "
                            "transaction, payment processing, "
                            "offsite ads."),
        ("Fixed fees", "Fees charged per order or per listing, "
                       "whatever the price: the listing fee, the "
                       "fixed part of card processing."),
        ("Keystone", "The old trade rule: double your cost for "
                     "wholesale, double again for retail. A sanity "
                     "check, not a method."),
        ("Line sheet", "The one page catalogue you send a shop, "
                       "with retail and trade prices."),
        ("Seconds", "Pieces that cannot be sold at full price. "
                    "Their cost is carried by the good ones."),
        ("Difficulty factor", "A multiplier on the time a fiddly "
                              "product takes, so you do not have to "
                              "re-time every variation."),
    ]
    for term, meaning in glossary:
        s.write(row, 1, term, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, meaning, fmts.chrome("label"))
        s.row_height(row, 24)
        row += 1
    row += 1

    # ------------------------------------------------- common questions
    row = s.section(row, "Questions people actually ask", LAST_COL)
    row = s.table_head(row, 1, ["Question", "Answer"], [30, 68])
    questions = [
        ("Should I really pay myself for every hour?",
         "Yes. If the price only works because your time is free, "
         "it is not a price, it is a subsidy. Start at your local "
         "minimum wage if your proper rate feels impossible, then "
         "raise it."),
        ("My price is now higher than everyone else's.",
         "Then either your costs are higher, your margin target is "
         "ambitious, or they are not paying themselves. Check the "
         "Competitors tab, then decide which of those it is."),
        ("What margin should I aim for?",
         "55% retail and 35% wholesale are reasonable starting "
         "points for handmade goods. Below 30% retail there is no "
         "room for a bad month."),
        ("Do I include my own labour in wholesale?",
         "Always. A wholesale price that only covers materials is "
         "a donation to a shop."),
        ("How do I handle free shipping?",
         "Put what the label costs in the Shipping tab and charge "
         "the buyer nothing. The workbook moves it into your cost "
         "where you can see it."),
        ("Should the pattern or design time be in every item?",
         "Spread it over the pieces you expect to sell, on the "
         "Development tab or the Digital tab. Loading it all onto "
         "the first sale makes the first sale impossible."),
        ("Why is my hourly wage lower than my hourly rate?",
         "Because the price is not covering the full cost. The "
         "wage column is what you really earn once the materials "
         "and fees are paid."),
        ("Can I use this for a craft fair?",
         "Yes - the Events tab costs the stall and tells you what "
         "the day has to take before it breaks even."),
        ("How often should I re-price?",
         "Twice a year, and immediately when a main material moves "
         "more than 10%. The catalog flags both."),
    ]
    for question, answer in questions:
        s.write(row, 1, question, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, answer, fmts.chrome("label"))
        s.row_height(row, 32)
        row += 1
    row += 1

    # ----------------------------------------------------- feature map -
    row = s.section(row, "Where every feature lives", LAST_COL)
    row = s.table_head(row, 1, ["If you want to...", "Go to"],
                       [46, 52])
    where = [
        ("Work out what to charge for something new",
         "Retail Calculator"),
        ("Check a product is paying you properly", "Products"),
        ("Cost a new material", "Materials"),
        ("Change your hourly rate", "Settings"),
        ("Change the Etsy or Shopify fee rates", "Channels"),
        ("Decide whether to accept a wholesale order",
         "Wholesale Calculator"),
        ("Find out how many you must sell this month",
         "Profit & Break-Even"),
        ("Price a crochet piece by yarn weight and hook",
         "Crochet"),
        ("Price a ring by grams of silver and scrap recovery",
         "Jewelry"),
        ("Price a candle by vessel size and fragrance load",
         "Candles"),
        ("Price a garment by fabric, size and layout waste",
         "Clothing"),
        ("Price a board by timber length and offcut allowance",
         "Woodworking"),
        ("Work out whether a pattern will ever pay for itself",
         "Digital"),
        ("See what a 20% discount really costs", "Promotions"),
        ("Compare what a run cost with what you estimated",
         "Batches"),
        ("See whether a craft fair was worth the day", "Events"),
        ("Test a price rise across the whole range", "Scenarios"),
        ("Print a costing sheet for your records", "Print Costing"),
        ("Send a shop a line sheet", "Print Order"),
        ("Quote for a commission", "Print Quote"),
        ("Add an option to a drop-down", "Lists"),
    ]
    for want, go in where:
        s.merge(row, 1, 5, want, fmts.chrome("label"))
        s.merge(row, 6, LAST_COL, go, fmts.chrome("label_bold"))
        s.row_height(row, 20)
        row += 1
    row += 1

    # ------------------------------------------------------- the small print
    row = s.section(row, "The small print", LAST_COL)
    row = s.note(row,
                 "This workbook is a calculator, not advice. Fee "
                 "rates, tax rules, safety testing and labelling "
                 "requirements differ by country and change without "
                 "notice - check the current official sources for "
                 "your marketplace and your jurisdiction before you "
                 "rely on a number. The sheet password protects "
                 "formulas from accidental edits; it is not "
                 "encryption and it is not security.",
                 LAST_COL, first_col=1, height=56)
    row += 1

    s.note(row,
           f"{cfg.product_name} - version {cfg.version}, by "
           f"{cfg.author}. No macros, no add-ins, nothing to "
           f"install. Every formula is readable and every list is "
           f"editable.",
           LAST_COL, first_col=1, height=30)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
