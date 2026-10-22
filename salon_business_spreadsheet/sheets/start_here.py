"""`Start Here` - what this is, what to type, and what it will not do.

The first tab a buyer opens. It sets the file up in six steps, maps
all twenty tabs, and is honest about the limits: no texts, no
confirmations, no booking platform, no tax advice.
"""

from __future__ import annotations

from ..config import DISCLAIMER, Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter

NAME = "Start Here"
LAST_COL = 9

TABS = [
    ("Dashboard", "The month at a glance: takings against target, "
                  "average ticket, utilisation, rebooking, and the "
                  "jobs that need doing today.", "Read it"),
    ("Appointments", "The diary. Date, time, client, service, stylist "
                     "and status - with a double-booking alarm.",
     "Type here"),
    ("Clients", "The client book: contact details, formula, patch "
                "test, spend, last visit, and who is drifting away.",
     "Type here"),
    ("Services", "Your menu, with the minutes, the product cost and "
                 "what each treatment earns an hour.", "Type here"),
    ("Sales", "The till. One row per thing sold - gross, discount, "
              "refund, net, tax and profit.", "Type here"),
    ("Payments", "Money in and out: deposits, balances, tips, refunds "
                 "and the card fees.", "Type here"),
    ("Staff", "The team, their pay basis and what each chair brings "
              "in an hour.", "Type here"),
    ("Commissions", "One row per stylist per month. Eight methods "
                    "supported - you choose, per person.", "Type here"),
    ("Payroll", "Pay runs: hours, commission released, tips passed "
                "on and the employer costs on top.", "Type here"),
    ("Expenses", "Every cost, categorised the way a profit and loss "
                 "wants it, with the recurring ones spread.",
     "Type here"),
    ("Profit & Loss", "Twelve months of income, cost of sales, staff "
                      "cost, overheads and operating profit.",
     "Read it"),
    ("Inventory", "Retail and backbar stock: what is left, what it is "
                  "worth, what to order.", "Type here"),
    ("Purchases", "Orders placed and received, plus the supplier book "
                  "that feeds the drop-downs.", "Type here"),
    ("Marketing", "Campaigns and offers, measured against what they "
                  "cost - money out and discount given away.",
     "Type here"),
    ("Loyalty", "Points, rewards and referral credit, counted from "
                "the spend already logged.", "Type here"),
    ("Goals & Budget", "This year's targets, the month-by-month plan "
                       "and a budget for every category.",
     "Type here"),
    ("Analytics", "Categories, days, hours, sources and retention - "
                  "the patterns behind the takings.", "Read it"),
    ("Reports", "The week and the month on one printable page.",
     "Read it"),
    ("Settings", "Your salon, your tax rules, your pay model, your "
                 "lists. Set these first.", "Set it up"),
]

STEPS = [
    ("1. Settings first",
     "Open the Settings tab and work down it: salon name, currency, "
     "tax rate and what tax applies to, deposit rule, rounding, the "
     "commission defaults and your targets. Ten minutes here saves "
     "every other tab from guessing."),
    ("2. Put your menu in",
     "On the Services tab, type each treatment with its price, how "
     "many minutes it takes and what the product costs you. The "
     "minutes matter as much as the price - that is what turns a "
     "price list into a profit list."),
    ("3. Add the team",
     "On the Staff tab, add everyone and pick their commission "
     "method. Nothing is assumed: commission only, hourly, a flat "
     "amount per service, tiered, a chair renter who pays you rent, "
     "or none at all."),
    ("4. Bring the clients across",
     "Paste or type your client list - name, phone, email and the "
     "'found you by' column. Everything else fills itself in as you "
     "start logging visits."),
    ("5. Work the diary and the till",
     "Log appointments as you book them and sales as you ring them "
     "up. A deposit goes on the Payments tab against the booking; it "
     "is applied to the bill when the client comes in, so it can "
     "never be counted as revenue twice."),
    ("6. Once a month",
     "Add a commission row per stylist, run the payroll rows, log "
     "the bills on Expenses, count the stock, and read the Profit & "
     "Loss. Twenty minutes, and you know exactly how the month went."),
]

RULES = [
    "Blue cells are yours. Grey cells are formulas - they are locked "
    "so a stray paste cannot break the workbook.",
    "Every sheet is protected with the password 'premium'. Review > "
    "Unprotect Sheet if you want to change the structure.",
    "IDs number themselves from the row. Filter and sort freely, but "
    "do not cut and paste rows around - copy the values instead.",
    "Type dates as real dates. 14/03 or 3/14 depending on where you "
    "are, not '14 march'.",
    "Status is the engine: Completed, Cancelled and No show are the "
    "three final answers an appointment can have. Everything else is "
    "still open and is left out of the completion and cancellation "
    "rates.",
    "Never type over a formula to make a number look right. Find the "
    "row that is wrong instead - the alert columns are there to point "
    "at it.",
    "Add rows by typing in the first empty row of a tab. Every tab "
    "has hundreds already formatted and waiting.",
]

LIMITS = [
    "It does not send confirmations, reminders, texts or emails, and "
    "it does not sync with any booking platform, till or card reader. "
    "Everything in here is typed in or pasted in by you.",
    "It is not accounting, payroll or tax software, and nothing in it "
    "is accounting, payroll or tax advice. It organises your business "
    "so that the conversation with your accountant is a short one.",
    "It does not calculate income tax, national insurance or social "
    "security on wages. It shows what a pay run costs the salon.",
    "It does not take payments or store card details.",
    "Sales tax is whatever you set it to be. There is no country "
    "baked in, and the two switches let you tax retail, services, "
    "both or neither.",
]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary,
                    landscape=False)
    ws = s.ws

    s.banner(cfg.product_name,
             "A complete salon business system in one spreadsheet - "
             "diary, clients, prices, takings, staff pay, stock and "
             "profit.", LAST_COL)
    s.set_col(0, 0, 3)
    s.set_col(1, 1, 26)
    s.set_col(2, LAST_COL, 13)

    row = 4
    row = s.note(
        row,
        "Everything in this workbook is worked out from four tabs you "
        "fill in as you go: the diary, the till, the expenses and the "
        "pay runs. Fill those honestly and the dashboard, the profit "
        "and loss and the reports look after themselves.",
        LAST_COL, height=32)
    row += 1

    # --------------------------------------------------- six steps ----
    row = s.section(row, "Setting up, in six steps", LAST_COL)
    for title, text in STEPS:
        s.write(row, 1, title, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, text, fmts.chrome("note"))
        s.row_height(row, 32)
        row += 1
    row += 1

    # ------------------------------------------------------ the tabs ---
    row = s.section(row, "What is on each tab", LAST_COL)
    row = s.table_head(row, 1, ["Tab", "What it is for", "You"],
                       [22, 60, 12])
    for name, text, kind in TABS:
        if name in reg or name in ("Dashboard", "Profit & Loss",
                                   "Goals & Budget", "Analytics",
                                   "Reports", "Settings"):
            ws.write_url(row, 1, f"internal:{quote_sheet(name)}!A1",
                         fmts.chrome("link"), name)
        else:
            s.write(row, 1, name, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL - 1, text, fmts.chrome("note"))
        s.write(row, LAST_COL, kind, fmts.chrome("label"))
        s.row_height(row, 26)
        row += 1
    row += 1

    # ------------------------------------------------- the house rules --
    row = s.section(row, "Seven rules that keep it working", LAST_COL)
    for rule in RULES:
        s.merge(row, 1, LAST_COL, "-  " + rule, fmts.chrome("note"))
        s.row_height(row, 26)
        row += 1
    row += 1

    # -------------------------------------------------- what it is not --
    row = s.section(row, "What this workbook does not do", LAST_COL)
    for limit in LIMITS:
        s.merge(row, 1, LAST_COL, "-  " + limit, fmts.chrome("note"))
        s.row_height(row, 30)
        row += 1
    row += 1

    # ------------------------------------------------- the numbers used --
    row = s.section(row, "The sums behind the headline numbers",
                    LAST_COL)
    row = s.table_head(row, 1, ["Figure", "How it is worked out"],
                       [26, 70])
    sums = [
        ("Net sales", "Gross, less discounts, less refunds."),
        ("Operating profit",
         "Net sales, less the cost of what was sold, less the full "
         "cost of the pay runs, less the expenses logged."),
        ("Average ticket",
         "Net sales divided by completed appointments that were paid "
         "for."),
        ("Completion rate",
         "Completed appointments divided by the appointments that "
         "reached a final answer - completed, cancelled or no show."),
        ("Cancellation rate",
         "Cancelled appointments over the same final-answer total."),
        ("Client balance",
         "What was billed, less discounts, less every payment "
         "applied - deposits included."),
        ("Commission",
         "Eligible revenue times the rate, by whichever of the eight "
         "methods that stylist is on."),
        ("Service gross profit",
         "Price less the direct costs: the product used and the "
         "commission earned on it."),
        ("Stock on hand",
         "Opening, plus received, less used in services, less sold, "
         "less waste, plus adjustments."),
        ("Budget variance", "Actual spend less the budget. Positive "
                            "means over."),
        ("Repeat client rate",
         "Clients who came back at least twice, over the clients you "
         "have served."),
    ]
    for label, text in sums:
        s.write(row, 1, label, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, text, fmts.chrome("note"))
        s.row_height(row, 26)
        row += 1
    row += 1

    # ------------------------------------------------------- the small --
    row = s.section(row, "The small print", LAST_COL)
    row = s.note(row, DISCLAIMER, LAST_COL, height=50)
    row += 1
    s.write(row, 1, "Version", fmts.chrome("label_bold"))
    s.write(row, 2, cfg.version, fmts.chrome("label"))
    s.write(row, 4, "Made by", fmts.chrome("label_bold"))
    s.merge(row, 5, LAST_COL, cfg.author, fmts.chrome("label"))
    row += 1
    s.write(row, 1, "Password", fmts.chrome("label_bold"))
    s.write(row, 2, cfg.password, fmts.chrome("label"))
    s.write(row, 4, "Built for", fmts.chrome("label_bold"))
    s.merge(row, 5, LAST_COL, "Excel 2010 and later, Google Sheets, "
                              "LibreOffice, Numbers",
            fmts.chrome("label"))

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    s.protect()
    return s
