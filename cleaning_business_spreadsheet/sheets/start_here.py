"""`Start Here` - the first thing anyone sees when they open the file.

Twenty minutes of setup, then the workbook runs the admin side of a
cleaning business for you. This page explains the order to do things
in, what fills itself in, and - just as important - what a spreadsheet
honestly cannot do.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Start Here"
LAST_COL = 9


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.ink)

    s.banner(cfg.product_name,
             "A complete back office for a cleaning business - "
             "clients, quotes, jobs, staff, invoices and profit.",
             LAST_COL)

    s.set_col(0, 0, 4)
    s.set_col(1, 1, 30)
    s.set_col(2, 2, 18)
    s.set_col(3, LAST_COL, 15)

    row = 4
    row = s.note(row,
                 "Welcome. This workbook replaces the diary, the "
                 "quote pad, the invoice book and the shoebox of "
                 "receipts. Everything is a formula you can read and "
                 "change - no macros, no add-ins, no subscription.",
                 LAST_COL, height=34)
    row += 1

    # -------------------------------------------------- what you get ---
    row = s.section(row, "What is in the file", LAST_COL)
    row = s.table_head(row, 1, ["Tab", "What it does", "Your job"],
                       [30, 52, 26])
    tabs = [
        ("Business Settings", "Your name, rates, tax, targets and "
                              "every reminder window",
         "Fill in once, revisit twice a year"),
        ("Dashboard", "Twelve live numbers, the jobs for the week and "
                      "everything that needs doing",
         "Read it every morning"),
        ("Pricing & Quote Calculator",
         "Prices a clean four different ways and shows the margin on "
         "each", "Use it before you quote"),
        ("Client CRM", "Every client, what they are worth and whether "
                       "they are drifting away", "Add clients"),
        ("Property & Site Database",
         "Addresses, access, keys, pets, parking and the hours each "
         "property needs", "Add properties"),
        ("Leads & Sales Pipeline",
         "Enquiries from first call to won or lost",
         "Log every enquiry"),
        ("Service Catalog", "What you offer, what it costs you and "
                            "what it earns an hour", "Set up once"),
        ("Quotes & Estimates", "Quotes, their lines, follow-ups and "
                               "acceptance", "Raise quotes"),
        ("Job Scheduler", "Every visit: who, where, when, how long, "
                          "how much", "The heart of the workbook"),
        ("Calendar", "The month at a glance and the next thirty "
                     "visits in order", "Read only"),
        ("Cleaning Checklists",
         "Room-by-room task templates, and the tasks ticked off on "
         "the day", "Build your own"),
        ("Employee Database", "Your team, their pay, skills, checks "
                              "and performance", "Add staff"),
        ("Employee Schedule & Attendance",
         "The week's rota, who is overloaded and who has room",
         "Read, then fix on Jobs"),
        ("Job Costing & Profitability",
         "Wages, supplies and travel against the price, job by job",
         "Automatic"),
        ("Invoices & Payments",
         "Invoices, lines, payments, ageing and chasing",
         "Raise and record"),
        ("Supplies & Equipment",
         "Stock, usage, reorder points, machines and their servicing",
         "Count monthly"),
        ("Expenses & Mileage", "Every cost, with tax and categories, "
                               "plus business travel", "Log weekly"),
        ("Client Retention", "Who is slipping away and what they are "
                             "worth", "Act on it"),
        ("Complaints & Quality",
         "Complaints, what they cost and whether you kept the client",
         "Log honestly"),
        ("Safety & Compliance",
         "Certificates, checks and assessments with expiry warnings",
         "Keep in date"),
        ("Financial Plan", "Break-even, targets and what if you grew",
         "Plan with it"),
        ("Reports & Analytics", "Where the money comes from and where "
                                "it goes", "Read monthly"),
        ("Printable Forms", "Quote, invoice and job sheet, ready for "
                            "A4", "Print as needed"),
        ("Settings & Dropdown Lists",
         "Every drop-down list in the workbook", "Edit to suit you"),
    ]
    for tab, what, job in tabs:
        s.write(row, 1, tab, fmts.chrome("label_bold"))
        s.merge(row, 2, 6, what, fmts.chrome("label"))
        s.merge(row, 7, LAST_COL, job, fmts.chrome("note"))
        row += 1
    row += 1

    # ------------------------------------------------- the setup list --
    row = s.section(row, "Twenty minutes to set it up", LAST_COL)
    row = s.table_head(row, 1, ["Step", "Where", "What to do"],
                       [8, 30, 60])
    steps = [
        ("1", "Business Settings",
         "Your business name, address, phone and email. They print on "
         "every quote and invoice."),
        ("2", "Business Settings",
         "Your hourly rate, minimum charge and what a cleaner costs "
         "you an hour. Nothing else works until these are right."),
        ("3", "Business Settings",
         "Tax rate and payment terms. Set the tax rate to 0 if you "
         "are not registered."),
        ("4", "Service Catalog",
         "Edit the services to match what you actually sell. Delete "
         "the ones you do not offer."),
        ("5", "Cleaning Checklists",
         "Adjust the task templates. These become the job sheets your "
         "cleaners carry."),
        ("6", "Employee Database",
         "Add your team - even if the team is just you. Put your own "
         "pay rate in; it is what makes the profit figures real."),
        ("7", "Client CRM and Property & Site Database",
         "Add your clients, then their properties. The property is "
         "where the hours and the access details live."),
        ("8", "Job Scheduler",
         "Put next week's visits in. The price, the finish time and "
         "the invoice link all work themselves out."),
        ("9", "Dashboard",
         "Open it every morning. If a number looks wrong, the tab it "
         "came from is named right beside it."),
    ]
    for num, where, what in steps:
        s.write(row, 1, num, fmts.chrome("kpi_label"))
        s.merge(row, 2, 3, where, fmts.chrome("label_bold"))
        s.merge(row, 4, LAST_COL, what, fmts.chrome("label"))
        s.row_height(row, 26)
        row += 1
    row += 1

    # ------------------------------------------------ the daily rhythm -
    row = s.section(row, "The rhythm that keeps it accurate", LAST_COL)
    row = s.table_head(row, 1, ["When", "What", "How long"],
                       [20, 62, 18])
    rhythm = [
        ("Every morning", "Open the Dashboard. Confirm tomorrow's "
                          "visits. Assign anything with no cleaner.",
         "2 minutes"),
        ("After each job", "Mark it Completed and put the actual "
                           "hours in.", "30 seconds"),
        ("Every Friday", "Approve timesheets, raise the week's "
                         "invoices, log receipts.", "20 minutes"),
        ("Every Monday", "Chase unpaid invoices and unanswered "
                         "quotes. This is the hour that pays for "
                         "itself.", "30 minutes"),
        ("Monthly", "Count supplies, check equipment servicing, read "
                    "the Reports tab.", "1 hour"),
        ("Quarterly", "Review prices against the Job Costing tab. Any "
                      "client under your margin floor needs a "
                      "conversation.", "1 hour"),
    ]
    for when, what, how_long in rhythm:
        s.write(row, 1, when, fmts.chrome("label_bold"))
        s.merge(row, 2, 7, what, fmts.chrome("label"))
        s.merge(row, 8, LAST_COL, how_long, fmts.chrome("note"))
        s.row_height(row, 24)
        row += 1
    row += 1

    # ------------------------------------------------ what it does ----
    row = s.section(row, "What the workbook does by itself", LAST_COL)
    row = s.bullets(row, [
        "Prices a clean from the floor area, the room count, the "
        "service or your target margin - and shows all four.",
        "Works out finish times, crew size and hours per cleaner.",
        "Applies the frequency discount, the pet surcharge, the "
        "minimum charge and your rounding rule.",
        "Flags double bookings, long days and jobs with nobody on "
        "them.",
        "Totals quotes from their lines, dates the expiry and tells "
        "you when to chase.",
        "Builds invoices from the lines, ages them, adds the late fee "
        "and tracks the balance after every payment.",
        "Costs every job with real wages from the timesheets, real "
        "supply usage and the travel, then shows the margin.",
        "Tracks stock down to days of cover and tells you what to "
        "reorder.",
        "Warns you before certificates expire and before machines are "
        "due a service.",
        "Spots clients going quiet long before they tell you they are "
        "leaving.",
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
        "Send reminders by itself. It flags what needs chasing - you "
        "or your phone still send the message.",
        "Sync with Google Calendar, Outlook or a phone diary. Export "
        "the Jobs tab to CSV if you need that.",
        "Take card payments or connect to your bank. Record the "
        "payment here once it lands.",
        "Collect a legally binding electronic signature. Print the "
        "quote, or use a proper e-signature service.",
        "Track your cleaners' location, or clock them in by GPS.",
        "File your tax return. It gets the numbers ready for someone "
        "who can.",
        "Replace advice from an accountant or an employment lawyer, "
        "and it does not know your local rules.",
    ], LAST_COL, first_col=1, height=16)
    row += 1

    # ------------------------------------------------ privacy ---------
    row = s.section(row, "A word about privacy", LAST_COL)
    row = s.note(row,
                 "This file will end up holding client addresses, "
                 "alarm codes, where the keys are kept, staff pay and "
                 "background checks. That is genuinely sensitive. "
                 "Keep it somewhere access-controlled, do not email "
                 "the whole workbook to anyone, give cleaners the "
                 "printed job sheet rather than the file, and delete "
                 "entry details when a client leaves. The sheet "
                 "password protects formulas from accidents - it is "
                 "not encryption.",
                 LAST_COL, first_col=1, height=56)
    row += 1

    # ------------------------------------------------ housekeeping ----
    row = s.section(row, "Housekeeping", LAST_COL)
    row = s.table_head(row, 1, ["Question", "Answer"], [30, 68])
    facts = [
        ("Is there example data in it?",
         "Both versions are in your download: one filled with a "
         "worked example, one empty and ready for your own."),
        ("How do I clear the example?",
         "Select the data rows on a tab and press Delete. The "
         "formula columns refill themselves - they are not deleted."),
        ("Why are some cells locked?",
         f"Every calculated cell is protected so a stray paste "
         f"cannot break it. The password is '{cfg.password}'. "
         f"Review, Unprotect Sheet if you want to change a formula."),
        ("Which cells do I type in?",
         "The white ones. Grey cells calculate themselves."),
        ("Does it work in Google Sheets?",
         "Yes. File, Import, Upload. Drop-downs and conditional "
         "colours come across; sheet protection behaves a little "
         "differently."),
        ("Which Excel versions?",
         "Excel 2010 and newer, Excel for Mac, Excel online, "
         "LibreOffice Calc and Google Sheets. No macros, so nothing "
         "to enable."),
        ("Can I add more rows?",
         "Yes - insert rows inside the existing block and copy the "
         "formulas down. Never insert above the header row."),
        ("Can I add my logo?",
         "Insert, Picture on the printable forms. Everything else is "
         "typed on the Settings tab."),
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
