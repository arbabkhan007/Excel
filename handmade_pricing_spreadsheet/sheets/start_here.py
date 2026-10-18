"""`Start Here` - the first thing anyone sees when they open the file.

An hour of setup, then this workbook answers the question that keeps
handmade businesses small: what should I actually charge? It explains
the order to do things in, what fills itself in, the pricing formula
in plain words, and - just as honestly - what a spreadsheet cannot do.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Start Here"
LAST_COL = 9


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.ink)
    ws = s.ws

    s.banner(cfg.product_name,
             "Materials, labour, packaging, fees and overhead - "
             "turned into a price that pays you.", LAST_COL)

    s.set_col(0, 0, 4)
    s.set_col(1, 1, 30)
    s.set_col(2, 2, 18)
    s.set_col(3, LAST_COL, 15)

    row = 4
    row = s.note(row,
                 "Most handmade pricing advice stops at 'materials "
                 "times two'. That is why so many makers work for "
                 "nothing. This workbook counts the yarn, the hours, "
                 "the box, the postage, the Etsy fees and the studio "
                 "rent, then solves for the price that leaves the "
                 "margin you actually want.",
                 LAST_COL, height=42)
    row += 1

    # ------------------------------------------------- the five minutes
    row = s.section(row, "Five minutes to your first honest price",
                    LAST_COL)
    row = s.table_head(row, 1, ["Step", "Where", "What to do"],
                       [10, 24, 62])
    quick = [
        ("1", "Settings", "Put your hourly rate in. That one number "
                          "changes more than anything else here."),
        ("2", "Materials", "Add the three or four materials the "
                           "product uses, with what you paid and how "
                           "much came in the pack."),
        ("3", "Material Usage", "List how much of each goes into one "
                                "piece."),
        ("4", "Labor", "Put the minutes in - all of them, including "
                       "finishing and packing."),
        ("5", "Products", "Add the product. Cost, suggested price, "
                          "margin and your hourly wage appear."),
        ("6", "Retail Calculator",
         "Check the price, try the ladder, and decide."),
    ]
    for step, where, what in quick:
        s.write(row, 1, step, fmts.chrome("label_bold"))
        s.write(row, 2, where, fmts.chrome("label_bold"))
        s.merge(row, 3, LAST_COL, what, fmts.chrome("label"))
        s.row_height(row, 22)
        row += 1
    row += 1

    # ----------------------------------------------------- the formula
    row = s.section(row, "The formula everything is built on", LAST_COL)
    row = s.note(row,
                 "Marketplace fees are charged on the price - the "
                 "price you have not worked out yet. That is why you "
                 "cannot just add 20% to your costs: the fee is "
                 "charged on the fee. Solving it properly looks like "
                 "this.",
                 LAST_COL, first_col=1, height=34)
    formulas = [
        ("Cost", "C = materials + labour + packaging + postage + "
                 "overhead", "What one piece costs you."),
        ("Labour", "L = minutes / 60 x your hourly rate",
         "Every minute, not just the fun ones."),
        ("Price", "Price = (C + fixed fees) / (1 - target margin - "
                  "fee rate)",
         "The price that leaves the margin you asked for, after the "
         "percentage fees."),
        ("Margin", "Margin = profit / price x 100",
         "The share of the price you keep."),
        ("Markup", "Markup = profit / cost x 100",
         "The share of the cost you added. Not the same thing."),
        ("Break-even", "Units = fixed costs / contribution a unit",
         "Rounded up, because you cannot sell two thirds of a "
         "candle."),
    ]
    for label, formula, meaning in formulas:
        s.write(row, 1, label, fmts.chrome("label_bold"))
        s.merge(row, 2, 5, formula, fmts.chrome("label_bold"))
        s.merge(row, 6, LAST_COL, meaning, fmts.chrome("note"))
        s.row_height(row, 22)
        row += 1
    row += 1
    row = s.note(row,
                 "Worked example. A beanie costs $8.40 in yarn and "
                 "labour, $1.85 to package and post, and carries "
                 "$1.20 of overhead - $11.45. Etsy takes about 9.5% "
                 "plus $0.45. You want a 55% margin. The price is "
                 "(11.45 + 0.45) / (1 - 0.55 - 0.095) = $33.52, "
                 "rounded to $33.99. Adding 55% to the cost would "
                 "have given $17.75 and a real margin of 25%.",
                 LAST_COL, first_col=1, height=56)
    row += 1

    # ------------------------------------------------------ the tabs --
    row = s.section(row, "What is in the file", LAST_COL)
    row = s.table_head(row, 1, ["Tab", "What it does", "Your job"],
                       [26, 54, 28])
    tabs = [
        ("Settings", "Your rates, targets, fee rates, rounding rule "
                     "and every default", "Fill in once, review "
                                          "twice a year"),
        ("Dashboard", "Your margins, your month and everything that "
                      "needs fixing", "Read it weekly"),
        ("Products", "The catalog: cost, price, margin, hourly wage "
                     "and what sold", "The heart of the workbook"),
        ("Materials", "Every yarn, bead, wax and board, with the "
                      "cost per gram worked out", "Add as you buy"),
        ("Material Usage", "How much of each material goes into "
                           "each product", "The bill of materials"),
        ("Labor", "Minutes per stage, per item or per batch",
         "Be honest about the time"),
        ("Packaging / Shipping", "Boxes, tissue, labels, postage "
                                 "services and what you absorb",
         "Set up once"),
        ("Overhead", "Studio, insurance, subscriptions - shared "
                     "across what you make", "Add your bills"),
        ("Channels", "Etsy, Shopify, markets, wholesale - every fee, "
                     "fully editable", "Check twice a year"),
        ("Retail Calculator", "The price, the ladder and the "
                              "sanity check", "Use before you list"),
        ("Wholesale Calculator", "What a stockist leaves you and "
                                 "what retail price supports it",
         "Use before you say yes"),
        ("Profit & Break-Even", "Contribution, break-even units, "
                                "margin against markup",
         "Use when you are planning"),
        ("Crochet / Jewelry / Candles", "Specialist calculators that "
                                        "work the way those crafts "
                                        "actually work", "Use yours"),
        ("Clothing / Woodworking / Digital", "Fabric and sizes, "
                                             "timber and offcuts, "
                                             "design time spread "
                                             "over sales",
         "Use yours"),
        ("Bundles / Variations", "Gift sets, sizes, personalisation "
                                 "and rush orders", "As needed"),
        ("Batches", "What a run really cost against the estimate",
         "Log your runs"),
        ("Sales", "Every order with the real fees and real postage",
         "Log them weekly"),
        ("Promotions / Price History / Competitors",
         "What discounts cost, what price changes did, what the "
         "market charges", "Keep it current"),
        ("Waste Log / Development / Events",
         "Material lost, design time, craft fair stalls",
         "Log as it happens"),
        ("Wholesale Customers / Orders", "Stockists, their terms and "
                                         "what each is worth",
         "If you wholesale"),
        ("Scenarios / Reports", "Reposition the range, compare "
                                "channels, survive a price rise",
         "Monthly"),
        ("Print Costing / Quote / Order", "Three printable forms, "
                                          "filled in from your data",
         "Print when needed"),
        ("Lists / Help", "Every drop-down, every factor table, and "
                         "the troubleshooting", "Edit the lists"),
    ]
    for tab, what, job in tabs:
        s.write(row, 1, tab, fmts.chrome("label_bold"))
        s.merge(row, 2, 6, what, fmts.chrome("label"))
        s.merge(row, 7, LAST_COL, job, fmts.chrome("note"))
        s.row_height(row, 24)
        row += 1
    row += 1

    # ---------------------------------------------------- the rhythm --
    row = s.section(row, "A rhythm that keeps it accurate", LAST_COL)
    rhythm = [
        ("As you buy", "Add new materials with the pack price and "
                       "pack size. Nothing else is needed.",
         "1 minute"),
        ("As you make", "Log batches with the real hours. Two runs "
                        "and your estimates become facts.",
         "2 minutes"),
        ("Every week", "Enter the week's sales with the real fees "
                       "and postage.", "15 minutes"),
        ("Every month", "Read the Dashboard. Fix anything below "
                        "minimum wage before you make more of it.",
         "20 minutes"),
        ("Every quarter", "Check your material prices and the fee "
                          "rates on the Channels tab. Both move.",
         "30 minutes"),
        ("Twice a year", "Re-price. Log every change on Price "
                         "History so you can see what worked.",
         "1 hour"),
    ]
    for when, what, how_long in rhythm:
        s.write(row, 1, when, fmts.chrome("label_bold"))
        s.merge(row, 2, 7, what, fmts.chrome("label"))
        s.merge(row, 8, LAST_COL, how_long, fmts.chrome("note"))
        s.row_height(row, 24)
        row += 1
    row += 1

    # ------------------------------------------------- what it does ---
    row = s.section(row, "What the workbook does by itself", LAST_COL)
    row = s.bullets(row, [
        "Turns a pack price into a cost per gram, millilitre or "
        "centimetre, and applies your waste allowance.",
        "Adds every material and every labour stage into a cost per "
        "piece, with a difficulty factor for the fiddly ones.",
        "Solves for the price that leaves your target margin after "
        "percentage fees and fixed fees - separately, because they "
        "behave differently.",
        "Applies your rounding rule so prices come out at $33.99 "
        "rather than $33.52.",
        "Warns you when a product pays under minimum wage, loses "
        "money, or has not been re-priced while its materials moved.",
        "Prices six crafts the way those crafts actually work - "
        "skeins and hooks, grams of silver and scrap recovery, wax "
        "and fragrance load, fabric and size factors, timber and "
        "offcuts, design time spread over expected sales.",
        "Works out wholesale both ways: what a discount leaves you, "
        "and what retail price would make it work.",
        "Shows break-even in units, in money and in hours.",
        "Compares what you estimated with what the batch really cost "
        "and what the marketplace really charged.",
        "Tells you what a discount costs in profit and how much more "
        "you would have to sell to stand still.",
    ], LAST_COL, first_col=1, height=16)
    row += 1

    # ------------------------------------------------ honest limits ---
    row = s.section(row, "What it honestly cannot do", LAST_COL)
    row = s.note(row,
                 "Every one of these needs something outside a "
                 "spreadsheet. Anyone who tells you otherwise is "
                 "selling you something.", LAST_COL, first_col=1,
                 height=22)
    row = s.bullets(row, [
        "Know the current Etsy, Shopify or Amazon Handmade fees. "
        "They change, they vary by country, and they are different "
        "for offsite ads. Put your real rates on the Channels tab "
        "and check them against the official fee schedule twice a "
        "year.",
        "Pull your orders in automatically. Export a CSV from your "
        "shop and paste it into the Sales tab.",
        "Tell you what your market will pay. It tells you what you "
        "need - the Competitors tab is where you record the other "
        "half.",
        "Make a product sell. A correct price on a bad photograph "
        "sells nothing.",
        "Handle sales tax, VAT or import duty properly. It has a "
        "simple tax rate and nothing more.",
        "Replace an accountant, and it does not know your local "
        "rules on labelling, safety testing or candle regulations.",
    ], LAST_COL, first_col=1, height=16)
    row += 1

    # ------------------------------------------------ the fee warning -
    row = s.section(row, "About marketplace fees", LAST_COL)
    s.merge(row, 1, LAST_COL, None, fmts.calc("text"))
    ws.write_formula(
        row, 1,
        '="Every fee in this workbook is a number you can edit on '
        'the Channels tab, because published fees change and vary by '
        'country. The rates in this file were last reviewed on "&'
        'TEXT(FeesReviewed,"d mmmm yyyy")&". Before you rely on a '
        'price, verify the current official fee schedule for your '
        'marketplace and your country, and update the Channels tab."',
        fmts.calc("text"))
    s.row_height(row, 44)
    row += 2

    # ------------------------------------------------ housekeeping ----
    row = s.section(row, "Housekeeping", LAST_COL)
    row = s.table_head(row, 1, ["Question", "Answer"], [30, 68])
    facts = [
        ("Is there example data in it?",
         "Both versions are in your download: one filled with a "
         "worked example shop, one empty and ready for yours."),
        ("How do I clear the example?",
         "Select the data rows on a tab and press Delete. The "
         "formula columns refill themselves - they are not deleted."),
        ("Why are some cells locked?",
         f"Every calculated cell is protected so a stray paste "
         f"cannot break it. The password is '{cfg.password}'. "
         f"Review, Unprotect Sheet if you want to change a formula."),
        ("Which cells do I type in?",
         "The white ones. Shaded cells work themselves out."),
        ("Can I change the currency?",
         "Yes - the Settings tab holds the symbol, and every money "
         "cell follows it."),
        ("Does it work in Google Sheets?",
         "Yes. File, Import, Upload. Drop-downs and colours come "
         "across; sheet protection behaves a little differently."),
        ("Which Excel versions?",
         "Excel 2010 and newer, Excel for Mac, Excel online, "
         "LibreOffice Calc and Google Sheets. No macros, so there "
         "is nothing to enable."),
        ("Can I add more rows?",
         "Yes - insert rows inside the existing block and copy the "
         "formulas down. Never insert above the header row."),
        ("My craft is not one of the six tabs",
         "Use the Products, Materials and Labor tabs - they work "
         "for anything. The six specialist tabs are shortcuts, not "
         "requirements."),
    ]
    for question, answer in facts:
        s.write(row, 1, question, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, answer, fmts.chrome("label"))
        s.row_height(row, 26)
        row += 1
    row += 1

    s.note(row,
           f"{cfg.product_name} - version {cfg.version}, by "
           f"{cfg.author}. Built with formulas only: no macros, no "
           f"add-ins, nothing to install. The Help tab has the "
           f"troubleshooting and a map of every feature.",
           LAST_COL, first_col=1, height=34)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
