"""`Getting Started` - what this is, what to type, and what it will not do.

The first tab a buyer opens.  It sets the file up in seven steps,
maps all twenty-three tabs, and is straight about the limits: this is
a business management workbook, not a child development record, not a
medical system, and not a licensing compliance tool.
"""

from __future__ import annotations

from ..config import DISCLAIMER, PRIVACY_NOTE, Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter

NAME = "Getting Started"
LAST_COL = 9

TABS = [
    ("Dashboard", "The setting at a glance: occupancy, attendance, "
                  "money in and out, every room's ratio and the jobs "
                  "that need doing today.", "Read it"),
    ("Getting Started", "This page.", "Read it"),
    ("Settings & Lists", "Your setting, your fees, your age bands and "
                         "ratios, your tax rules, your lists. Set "
                         "these first.", "Set it up"),
    ("Child Enrollment", "The register. Every child, their family, "
                         "their room, their schedule and the rate "
                         "they are on.", "Type here"),
    ("Parent Database", "One row per family, not per child - so "
                        "siblings share a bill, an address and an "
                        "emergency contact.", "Type here"),
    ("Waitlist", "Enquiries, tours, offers and where each one sits in "
                 "the queue for the room they want.", "Type here"),
    ("Attendance", "Daily check in and check out, with late arrivals, "
                   "early collections and missing sign-outs flagged.",
     "Type here"),
    ("Tuition & Fees", "One row per charge. Tuition, extras, "
                       "discounts and subsidies - the lines that make "
                       "up an invoice.", "Type here"),
    ("Invoices & Receipts", "The money ledger: invoices, payments, "
                            "credit notes, refunds and write-offs, "
                            "one row per movement.", "Type here"),
    ("Daily Reports", "The printable note that goes home: mood, meals, "
                      "naps, nappies and what they did today.",
     "Type here"),
    ("Meal Tracker", "What was served, how much was eaten, and who "
                     "needed something different.", "Type here"),
    ("Allergy & Medical", "Allergies, conditions, care plans, consent "
                          "and the medication administration log.",
     "Type here"),
    ("Incident Reports", "Accidents, injuries and near misses, with "
                         "the parent notification timed to the "
                         "minute.", "Type here"),
    ("Staff Database", "The team: roles, pay, certificates and when "
                       "each one expires.", "Type here"),
    ("Staff Schedule", "The rota and the clock. Who is in which room, "
                       "for how long, and what it costs.",
     "Type here"),
    ("Payroll Tracker", "Pay runs built from the rota - basic, "
                        "overtime, employer costs and net pay.",
     "Type here"),
    ("Classroom Management", "Every room: licensed places, children "
                             "enrolled, who is in today and whether "
                             "you are within ratio.", "Type here"),
    ("Expense Tracker", "Every cost, categorised the way a profit and "
                        "loss wants it, against your budget.",
     "Type here"),
    ("Supply Inventory", "Nappies, food, cleaning and craft - what is "
                         "left, how fast it goes, what to order.",
     "Type here"),
    ("Parent Communication", "Every conversation, what was agreed and "
                             "who is still waiting on a reply.",
     "Type here"),
    ("Document Tracker", "Children's forms, staff certificates and "
                         "the setting's own licences, with expiry "
                         "dates.", "Type here"),
    ("Monthly Reports", "One printable page for the month - "
                        "enrollment, attendance, money and staffing.",
     "Read it"),
    ("Annual Reports", "Twelve months side by side, plus the year's "
                       "totals and what changed.", "Read it"),
]

STEPS = [
    ("1. Settings first",
     "Open Settings & Lists and work down it: the setting's name, "
     "your currency, your opening hours, your age bands and the "
     "staff-to-child ratio for each one, your fees, your tax "
     "switches and your targets. Twenty minutes here and no other "
     "tab has to guess."),
    ("2. Set the age bands to your own licence",
     "The ratio table on Settings is the most country-specific thing "
     "in this workbook, which is why it is a table you edit and not "
     "a formula you cannot. Infant, toddler, preschool - the "
     "boundaries in months and the children per adult are all yours "
     "to change. Every room warning, every ratio check and every "
     "'moves up on' date follows that table."),
    ("3. Open the rooms",
     "On Classroom Management, name each room, give it its licensed "
     "places and pick the age band it runs. Everything on the "
     "dashboard that talks about occupancy, free seats or ratio "
     "reads these rows."),
    ("4. Add the families, then the children",
     "Families first, on Parent Database - one row per household, "
     "with the emergency contacts and the authorised pickup list. "
     "Then the children on Child Enrollment, each pointing at their "
     "family ID. Two children in one family share a bill "
     "automatically, and the second one picks up the sibling "
     "discount without you doing anything."),
    ("5. Take the register every day",
     "Attendance is the tab that drives the rest. Date, child, "
     "status, in and out. From it come the attendance percentages, "
     "the late pickup fees, the daily reports, the room counts and "
     "the ratio checks."),
    ("6. Bill once a month",
     "Add the charge lines on Tuition & Fees, raise one invoice per "
     "family on Invoices & Receipts, and tag the lines with that "
     "invoice number. When the money comes in, add a Payment row "
     "pointing at the same invoice. Partial payments, instalments "
     "and credit notes all work with no extra columns."),
    ("7. Once a month, in twenty minutes",
     "Run the payroll rows from the rota, log the bills on Expense "
     "Tracker, count the supplies, check the Document Tracker for "
     "anything expiring, and read the Monthly Report. That is the "
     "whole routine."),
]

RULES = [
    "Blue cells are yours. Grey cells are formulas - they are locked "
    "so a stray paste cannot break the workbook.",
    "Every sheet is protected with the password 'premium'. Review > "
    "Unprotect Sheet if you want to change the structure.",
    "IDs number themselves from the row. Filter and sort freely, but "
    "do not cut and paste rows around - copy the values instead.",
    "Type dates as real dates and times as real times. 7:30 am, not "
    "'half seven'.",
    "The family is the billing unit and the child is the care unit. "
    "Invoices go to a family ID; attendance, meals and reports go to "
    "a child ID. Keeping those straight is what makes siblings work.",
    "Never type over a formula to make a number look right. Find the "
    "row that is wrong instead - the Action column on every tab is "
    "there to point at it.",
    "Add rows by typing in the first empty row of a tab. Every tab "
    "has hundreds already formatted and waiting.",
]

LIMITS = [
    "It is a business management workbook. It is not a child "
    "development record, a learning journal, or an assessment "
    "framework, and it makes no claim to meet EYFS, Head Start, "
    "NAEYC or any other curriculum or quality standard.",
    "It is not a medical system. The allergy, medication and "
    "incident tabs are a written record of what you did and when - "
    "they are not clinical advice and they do not replace a care "
    "plan written with a health professional.",
    "It does not make you compliant with anybody's licensing rules. "
    "The ratios and age bands are the ones you type in. Check them "
    "against your own regulator, and check them again when the rules "
    "change.",
    "It does not send anything. No emails, no texts, no app "
    "notifications, no parent portal. Every record in here is typed "
    "in or pasted in by you.",
    "It is not accounting, payroll or tax software, and nothing in "
    "it is accounting, payroll or tax advice. It does not calculate "
    "income tax or social security on wages - it shows what a pay "
    "run costs the setting.",
    "It does not take payments and it does not store card details.",
    "Tax is whatever you set it to be. There is no country baked in, "
    "and childcare is exempt in a great many places - which is why "
    "the default rate is zero.",
]

SUMS = [
    ("Occupancy", "Children holding a place, divided by the licensed "
                  "places across your rooms. On hold and notice "
                  "given still hold a place; withdrawn does not."),
    ("Attendance %", "Sessions actually attended, divided by sessions "
                     "scheduled. Public holidays, closures and snow "
                     "days are taken out of both sides, because they "
                     "are nobody's fault."),
    ("Staff-to-child ratio", "Children present in a room today, "
                             "against the adults rota'd into that "
                             "room today. A shift only counts when it "
                             "has a room against it and is not Off, "
                             "Leave, Training or Admin."),
    ("Weekly fee", "The agreed rate converted to a week by the "
                   "child's own billing basis, less the sibling or "
                   "manual discount, less any subsidy."),
    ("Invoice total", "The sum of the charge lines tagged with that "
                      "invoice number - or the typed amount, if there "
                      "are no lines behind it."),
    ("Billed", "Invoices, less credit notes, less write-offs. This is "
               "revenue."),
    ("Collected", "Payments, less refunds. This is cash. It is a "
                  "different number from billed, on purpose."),
    ("Family balance", "Everything that put the account up, less "
                       "everything that brought it down - deposits "
                       "included, whether or not they are matched to "
                       "a bill."),
    ("Late pickup fee", "Minutes past the child's own contracted "
                        "pickup time, less the grace period, charged "
                        "as a first block and then per minute."),
    ("Pay run cost", "Basic hours plus overtime at the multiplier, "
                     "plus bonus and adjustments, plus employer costs "
                     "- and contractors and volunteers carry no "
                     "employer costs."),
    ("Operating profit", "Billed for the month, less the full cost of "
                         "the pay runs, less the expenses logged. "
                         "Wages are counted once, from the Payroll "
                         "Tracker."),
    ("Weeks of cover", "Stock on hand divided by what you have used a "
                       "week since tracking started - which is how "
                       "you find out that the wipes run out on "
                       "Thursday."),
]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.accent)
    ws = s.ws

    s.banner(cfg.product_name,
             "A complete childcare business system in one spreadsheet "
             "- enrollment, attendance, tuition, staff, rooms and "
             "parents.", LAST_COL)
    s.set_col(0, 0, 3)
    s.set_col(1, 1, 26)
    s.set_col(2, LAST_COL, 13)

    row = 4
    row = s.note(
        row,
        "Everything in this workbook is worked out from five tabs you "
        "fill in as you go: the children, the register, the charges, "
        "the rota and the costs. Fill those honestly and the "
        "dashboard, the invoices and the reports look after "
        "themselves.",
        LAST_COL, height=32)
    row += 1

    # ------------------------------------------------- seven steps ----
    row = s.section(row, "Setting up, in seven steps", LAST_COL)
    for title, text in STEPS:
        s.write(row, 1, title, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, text, fmts.chrome("note"))
        s.row_height(row, 42)
        row += 1
    row += 1

    # ---------------------------------------------------- the tabs ----
    row = s.section(row, "What is on each tab", LAST_COL)
    row = s.table_head(row, 1, ["Tab", "What it is for", "You"],
                       [24, 58, 12])
    for name, text, kind in TABS:
        ws.write_url(row, 1, f"internal:{quote_sheet(name)}!A1",
                     fmts.chrome("link"), name)
        s.merge(row, 2, LAST_COL - 1, text, fmts.chrome("note"))
        s.write(row, LAST_COL, kind, fmts.chrome("label"))
        s.row_height(row, 28)
        row += 1
    row += 1

    # ---------------------------------------------- the house rules ---
    row = s.section(row, "Seven rules that keep it working", LAST_COL)
    for rule in RULES:
        s.merge(row, 1, LAST_COL, "-  " + rule, fmts.chrome("note"))
        s.row_height(row, 28)
        row += 1
    row += 1

    # -------------------------------------------- the printable ones --
    row = s.section(row, "The three things you will print", LAST_COL)
    printables = [
        ("Invoice or receipt",
         "Invoices & Receipts builds the document for you: reference, "
         "family, period, the charge lines behind it, what has been "
         "paid and what is left. Filter to the family, print the "
         "row's block, or copy it into the printable layout at the "
         "bottom of the Monthly Reports tab."),
        ("Daily child report",
         "Daily Reports writes the parent-facing summary sentence "
         "itself from the mood, the meals, the naps and the nappies "
         "you logged. Filter to today, filter to the child, print."),
        ("Incident report",
         "Incident Reports composes the whole paragraph - what "
         "happened, where, the injury, the first aid, who was told "
         "and how many minutes it took. Print it, have the parent "
         "sign it, tick Signed."),
    ]
    for label, text in printables:
        s.write(row, 1, label, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, text, fmts.chrome("note"))
        s.row_height(row, 44)
        row += 1
    row += 1

    # ------------------------------------------------ what it is not --
    row = s.section(row, "What this workbook does not do", LAST_COL)
    for limit in LIMITS:
        s.merge(row, 1, LAST_COL, "-  " + limit, fmts.chrome("note"))
        s.row_height(row, 34)
        row += 1
    row += 1

    # --------------------------------------------- the numbers used ---
    row = s.section(row, "The sums behind the headline numbers",
                    LAST_COL)
    row = s.table_head(row, 1, ["Figure", "How it is worked out"],
                       [26, 68])
    for label, text in SUMS:
        s.write(row, 1, label, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, text, fmts.chrome("note"))
        s.row_height(row, 32)
        row += 1
    row += 1

    # ------------------------------------------------- the small -------
    row = s.section(row, "The small print", LAST_COL)
    row = s.note(row, PRIVACY_NOTE, LAST_COL, height=34)
    row += 1
    row = s.note(row, DISCLAIMER, LAST_COL, height=76)
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
