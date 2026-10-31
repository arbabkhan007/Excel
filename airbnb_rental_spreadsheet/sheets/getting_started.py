"""`Getting Started` - twenty minutes from download to running.

Nobody reads a manual for a spreadsheet, so this one is short, in
order, and tells you the two or three things that are genuinely not
obvious: that every tab joins on the property ID and the booking ID,
that the cleaning schedule is driven by your check-outs, and that the
platform fee table on Settings is the one thing you must change before
the numbers mean anything.
"""

from __future__ import annotations

from ..config import Config, DISCLAIMER, Palette, PRIVACY_NOTE
from ..layout import Registry, col_letter
from ..writer import SheetWriter

NAME = "Getting Started"
LAST_COL = 9

STEPS = [
    ("1", "Set your business up",
     "Settings & Lists",
     "Name, currency and contact details at the top. Two minutes, "
     "and it puts your name on every report and printable checklist."),
    ("2", "Fix the platform fee table",
     "Settings & Lists",
     "The most important five minutes you will spend in this file. "
     "Airbnb's split fee, Vrbo's two models and your own negotiated "
     "Booking.com rate are all different, and every net revenue "
     "figure in the workbook depends on getting them right. The "
     "defaults are a starting point, not your contract."),
    ("3", "Add your properties",
     "Property Database",
     "One row each. The ID in the first column writes itself - that "
     "ID is what every other tab joins on, so you never type a "
     "property name twice. Give each one a nightly rate, a cleaning "
     "fee and what the cleaner actually costs you."),
    ("4", "Put your cleaners and contractors in",
     "Staff & Vendors",
     "Cleaners, laundry, the handyman. Each gets an ID too, and "
     "cleaning jobs and repairs point at it."),
    ("5", "Enter your bookings",
     "Booking Tracker",
     "Platform, dates, guests, rate. Nights, gross, commission, card "
     "fee and net revenue all work themselves out. Leave the rate "
     "blank and it uses the property's nightly rate. Past bookings "
     "are worth back-filling: three months of history makes every "
     "analytic on the dashboard real."),
    ("6", "Block the dates you are keeping",
     "Booking Tracker",
     "Your own weekend, a renovation, a gap you will not sell. Add "
     "it as a booking with the status 'Owner block' or 'Maintenance "
     "block'. It takes the night off the market, earns nothing, and "
     "keeps your occupancy honest."),
    ("7", "Check the calendar",
     "Booking Calendar",
     "Pick a property and a month. If you have typed a clash, the "
     "day goes red and says DOUBLE BOOKED. Worth doing after every "
     "batch of entry."),
    ("8", "Log the turnovers",
     "Cleaning Schedule",
     "One row per changeover: point it at the booking that is "
     "leaving and it finds the property, the guest, the check-out "
     "date, the next arrival and whether it is a same-day turn. Nine "
     "tick boxes for the rooms, and the true cost including laundry "
     "and consumables."),
    ("9", "Record what you spend",
     "Expense Tracker",
     "Everything that is not cleaning or a repair - the mortgage "
     "interest, the rent if you are rent-to-rent, insurance, "
     "utilities, subscriptions. Tag each one to a property where you "
     "can; leave the property blank for anything that covers the "
     "whole portfolio."),
    ("10", "Reconcile the payouts",
     "Platform Revenue & Fees",
     "When money lands, type what actually arrived. The tab compares "
     "it with what you assumed and tells you when your fee rates are "
     "wrong - which is how most hosts discover they have been out by "
     "three percent all year."),
    ("11", "Read the dashboard",
     "Dashboard",
     "Occupancy, ADR, RevPAR, profit by property, and a list of "
     "everything that needs doing. Nothing to refresh."),
]

CONCEPTS = [
    ("Everything joins on two IDs",
     "A property is PRP-01 and a booking is BKG-0007 everywhere in "
     "the workbook. Type the ID and the name, the rate, the cleaning "
     "fee and the guest come with it. This is why the file does not "
     "care whether you have one flat or twelve, or whether you add a "
     "fourth platform next spring."),
    ("Nothing here syncs with anything",
     "There is no connection to Airbnb, Vrbo or Booking.com, no "
     "iCal, no API, and nothing in this workbook will ever change on "
     "its own. Every figure traces back to something you typed. That "
     "is a limitation and it is also the reason it still works in "
     "five years when an API has been retired."),
    ("Cleaning is driven by your check-outs",
     "The Cleaning Schedule does not invent jobs. You add a row and "
     "point it at the booking that is leaving; the tab then works "
     "out the property, the date, the next arrival, the gap in "
     "hours and whether it is a same-day turn. The Dashboard counts "
     "check-outs with no cleaning row against you, so nothing slips."),
    ("ADR and RevPAR leave cleaning fees out",
     "A cleaning fee is a cost you are recovering, not a room rate. "
     "Counting it inflates your ADR and flatters your RevPAR, which "
     "is exactly the mistake that makes a one-night-stay strategy "
     "look clever. Both figures here use accommodation revenue only."),
    ("Costs are counted once",
     "Cleaning lives on the Cleaning Schedule and repairs live on "
     "the Maintenance Log, because both need a workflow. If you "
     "would rather type them as plain expenses, two switches on "
     "Settings move them - but move them, do not do both."),
    ("Occupancy excludes what you blocked",
     "A night you kept for yourself was never for sale, so it is not "
     "counted against you. Occupancy is booked nights over available "
     "nights, and blocking a fortnight in February does not make you "
     "look like a failing host."),
]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.deep)

    s.banner("Getting Started",
             "Eleven steps, in order. Twenty minutes and the "
             "dashboard is live.", LAST_COL)
    s.set_col(0, 0, 5)
    s.set_col(1, 1, 30)
    s.set_col(2, 2, 24)
    s.set_col(3, LAST_COL, 14)

    row = 4
    row = s.note(
        row,
        "White cells are yours. Everything grey is a formula and is "
        "locked so it cannot be typed over by accident - there is no "
        "password prompt to worry about in normal use, and if you "
        "ever need to change a formula the password is in the README "
        "that came with the file.",
        LAST_COL, height=32)
    row += 1

    row = s.section(row, "The order to do it in", LAST_COL)
    row = s.table_head(row, 0, ["", "Step", "Where", "What to do"],
                       [5, 30, 24, 90])
    for number, title, where, detail in STEPS:
        s.write(row, 0, number, fmts.chrome("label_bold"))
        s.write(row, 1, title, fmts.chrome("label_bold"))
        s.write(row, 2, where, fmts.chrome("label"))
        s.merge(row, 3, LAST_COL, detail, fmts.chrome("note"))
        s.row_height(row, 32)
        row += 1
    row += 1

    row = s.section(row, "Six things worth knowing", LAST_COL)
    for title, detail in CONCEPTS:
        s.merge(row, 0, 2, title, fmts.chrome("label_bold"))
        s.merge(row, 3, LAST_COL, detail, fmts.chrome("note"))
        s.row_height(row, 40)
        row += 1
    row += 1

    # ------------------------------------------------- where things are --
    row = s.section(row, "What each tab is for", LAST_COL)
    row = s.table_head(row, 0, ["", "Tab", "Type", "What it does"],
                       [5, 30, 24, 90])
    tabs = [
        ("Dashboard", "Report",
         "Everything, summarised. The page to open first."),
        ("Getting Started", "Guide", "This page."),
        ("Settings & Lists", "Setup",
         "Fees, targets, tax, costs and every drop-down."),
        ("Property Database", "Register",
         "One row per property, with its own profit and loss."),
        ("Booking Tracker", "Register",
         "Every stay, every block, every cancellation."),
        ("Booking Calendar", "Report",
         "A month per property, a fortnight across all of them."),
        ("Guest Database", "Register",
         "Who has stayed, how often, and what they are worth."),
        ("Check-in & Check-out", "Register",
         "Access codes, wifi, house rules, inspection."),
        ("Cleaning Schedule", "Register",
         "Every turnover, who did it and what it really cost."),
        ("Turnover Checklist", "Printable",
         "One job on one page, for the cleaner to carry."),
        ("Platform Revenue & Fees", "Register",
         "What each platform kept and what actually landed."),
        ("Expense Tracker", "Register",
         "Everything else you spend, tagged to a property."),
        ("Supplies Inventory", "Register",
         "Linen, consumables, what to reorder and when."),
        ("Maintenance Log", "Register",
         "Repairs, who did them, estimated against actual."),
        ("Reviews Tracker", "Register",
         "Seven scores per review and which one is the problem."),
        ("Staff & Vendors", "Register",
         "Cleaners and contractors, their rates and their record."),
        ("Tax Summary", "Report",
         "Income and deductible costs grouped for an accountant."),
        ("Occupancy & Performance", "Register",
         "Property by month: occupancy, ADR, RevPAR, profit."),
        ("Monthly Reports", "Report",
         "Any month you pick, in full."),
        ("Annual Reports", "Report",
         "Twelve months across, and the year against last year."),
    ]
    for index, (tab, kind, what) in enumerate(tabs, start=1):
        s.write(row, 0, index, fmts.chrome("label_bold"))
        s.write(row, 1, tab, fmts.chrome("label_bold"))
        s.write(row, 2, kind, fmts.chrome("label"))
        s.merge(row, 3, LAST_COL, what, fmts.chrome("note"))
        row += 1
    row += 1

    # --------------------------------------------------- live progress --
    row = s.section(row, "How far you have got", LAST_COL)
    row = s.note(
        row,
        "This counts what is actually in the file, so it moves as you "
        "fill it in.",
        LAST_COL, height=22)
    checks = [
        ("Properties added", '=COUNTIF([[Property Database.name]],"?*")',
         "Add at least one before anything else works."),
        ("Bookings entered", '=COUNTIF([[Booking Tracker.check_in]],'
                             '">0")',
         "Back-fill three months and the analytics become real."),
        ("Cleaners and vendors",
         '=COUNTIF([[Staff & Vendors.name]],"?*")',
         "Needed before a cleaning job can be assigned."),
        ("Turnovers logged", '=COUNTIF([[Cleaning Schedule.booking]],'
                             '"?*")',
         "Should roughly match your number of check-outs."),
        ("Expenses recorded", '=COUNTIF([[Expense Tracker.date]],">0")',
         "Without these, profit is only revenue."),
        ("Payouts reconciled",
         '=COUNTIF([[Platform Revenue & Fees.booking]],"?*")',
         "How you find out your fee rates are wrong."),
        ("Reviews captured", '=COUNTIF([[Reviews Tracker.booking]],'
                             '"?*")',
         "Seven scores each, so you can see which one slips."),
        ("Jobs needing attention",
         '=COUNTIF([[Property Database.alert]],"?*")-'
         'COUNTIF([[Property Database.alert]],"OK*")+'
         'COUNTIF([[Cleaning Schedule.alert]],"?*")-'
         'COUNTIF([[Cleaning Schedule.alert]],"OK*")',
         "The only number here you want to be zero."),
    ]
    row = s.table_head(row, 0, ["", "Check", "Count", "Why"],
                       [5, 30, 24, 90])
    for index, (label, formula, why) in enumerate(checks, start=1):
        s.write(row, 0, index, fmts.chrome("label_bold"))
        s.write(row, 1, label, fmts.chrome("label_bold"))
        s.formula(row, 2, s.reg.render(formula), fmts.calc("int"))
        s.merge(row, 3, LAST_COL, why, fmts.chrome("note"))
        row += 1
    row += 1

    row = s.section(row, "Housekeeping", LAST_COL)
    row = s.bullets(row, [
        "Add rows by typing on the first blank line - the formulas "
        "are already there, all the way down.",
        "Do not delete a column. Hide it if it is in your way; the "
        "hidden helper columns are doing real work.",
        "Sort a register on any column you like. The IDs are "
        "formulas based on the row, so sorting renumbers them - if "
        "you want the IDs to stay put, paste them as values first, "
        "or simply filter instead of sorting.",
        "Back it up. It is a file, not a service; nobody else has a "
        "copy.",
        "The Google Sheets version behaves the same way, with the "
        "same formulas, minus the cell protection.",
    ], LAST_COL)
    row += 1

    row = s.note(row, PRIVACY_NOTE, LAST_COL, height=44)
    row += 1
    row = s.note(row, DISCLAIMER, LAST_COL, height=104)

    s.finish(freeze_row=4, gridlines=False, fit_width=True,
             print_area=f"A1:{col_letter(LAST_COL)}{row + 2}")
    s.protect()
    return s
