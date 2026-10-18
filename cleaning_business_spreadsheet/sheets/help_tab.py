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
             "What the colours mean, what to do when a cell looks "
             "wrong, and where every feature lives.", LAST_COL)

    s.set_col(0, 0, 4)
    s.set_col(1, 1, 30)
    s.set_col(2, LAST_COL, 14)

    row = 4

    # ----------------------------------------------- the colour code ---
    row = s.section(row, "The colour code", LAST_COL)
    row = s.table_head(row, 1, ["Looks like", "Means"], [30, 68])
    colours = [
        ("White cell with a blue edge",
         "Type here. These are the only cells you need to fill in."),
        ("Pale grey cell", "A formula. It is locked so you cannot "
                           "overwrite it by accident."),
        ("Green text or fill", "Good - on target, paid, in date, "
                               "profitable."),
        ("Amber fill", "Needs attention soon - due, low, under "
                       "target."),
        ("Red fill", "Act today - overdue, out of stock, expired, "
                     "losing money."),
        ("Blue fill", "For information - booked, sent, in progress."),
        ("A coloured bar inside a cell",
         "A data bar: the longer the bar, the bigger the number "
         "compared with the rest of the column."),
        ("A tiny chart",
         "A sparkline - twelve months of that number in one cell."),
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
                       [24, 36, 38])
    problems = [
        ("#N/A in a name column",
         "The ID you typed is not on the tab it points at",
         "Pick the ID from the drop-down instead of typing it"),
        ("'Not found' instead of a client",
         "The client ID has been deleted or renumbered",
         "Re-pick the client on the row"),
        ("#VALUE! in a total",
         "Text has been typed into a number cell - '2 hours' "
         "instead of 2",
         "Clear the cell and type the number only"),
        ("#REF! anywhere",
         "A row or column the formula needed has been deleted",
         "Undo the deletion, or copy the formula down from the row "
         "above"),
        ("#DIV/0! in a margin or rate",
         "Dividing by a zero - usually a price or an hours figure "
         "that is still blank",
         "Fill in the missing number; most of these are already "
         "guarded"),
        ("A price of 0",
         "No rate, no service price and no override",
         "Check the Services tab and the hourly rate on Settings"),
        ("The next visit date is empty",
         "The client has no future job in the diary, or the "
         "frequency is One-off",
         "Book the next visit on the Job Scheduler"),
        ("Hours look far too low",
         "The property has no area and no room counts",
         "Fill in the property, or type an hours override on the "
         "job"),
        ("Profit is negative on every job",
         "The cost per hour on Settings is wrong - it should be pay "
         "plus on-costs, not what you charge",
         "Check Pay rate and Employer on-cost on the Settings tab"),
        ("A formula cell will not let me type",
         "The sheet is protected so formulas survive a stray paste",
         f"Review, Unprotect Sheet, password '{cfg.password}'"),
        ("A drop-down has the wrong options",
         "The list behind it lives on the Settings & Dropdown Lists "
         "tab",
         "Edit the list there and every drop-down updates"),
        ("Numbers have not changed after an edit",
         "Calculation has been switched to manual",
         "Formulas, Calculation Options, Automatic - or press F9"),
        ("New rows have no formulas",
         "Rows added below the block do not inherit anything",
         "Copy the last row of the block down into the new rows"),
        ("The dashboard shows nothing",
         "Jobs are still at Booked, not Completed",
         "Revenue counts Completed and Invoiced work only"),
    ]
    for what, why, fix in problems:
        s.write(row, 1, what, fmts.chrome("label_bold"))
        s.merge(row, 2, 5, why, fmts.chrome("label"))
        s.merge(row, 6, LAST_COL, fix, fmts.chrome("note"))
        s.row_height(row, 26)
        row += 1
    row += 1

    # ----------------------------------------------- the glossary ------
    row = s.section(row, "The words this workbook uses", LAST_COL)
    row = s.table_head(row, 1, ["Term", "What it means here"],
                       [26, 72])
    glossary = [
        ("Cost per hour", "What one cleaner-hour really costs you: "
                          "pay plus your on-costs. Not what you "
                          "charge."),
        ("Margin", "Profit as a percentage of the price. 40% margin "
                   "on a $100 clean means $40 left after wages, "
                   "supplies and travel."),
        ("Margin floor", "The lowest margin you will accept. Jobs "
                         "under it are flagged red all over the "
                         "workbook."),
        ("Contribution", "Price less the costs that only happen "
                         "because you did the job. What is left to "
                         "pay the fixed bills."),
        ("Break-even", "The number of jobs a month that covers your "
                       "fixed costs. Below it you are paying to "
                       "work."),
        ("Utilisation", "Billed hours as a share of paid hours. "
                        "Travel, waiting and admin are the gap."),
        ("Frequency factor", "The discount a regular client gets. "
                             "Weekly work takes less time per visit "
                             "because less dirt builds up."),
        ("Condition factor", "How much longer a dirty property takes. "
                             "1.0 is normal, 1.8 is a hoarder clean."),
        ("Days of cover", "How many days of stock you have at your "
                          "recent rate of use."),
        ("Ageing", "How long an invoice has been overdue, in "
                   "buckets. Anything past 60 days rarely gets paid "
                   "in full."),
        ("Recurring revenue", "What your regular clients are worth "
                              "over a year at today's frequencies. "
                              "The number a buyer would pay for."),
        ("Churn", "Clients who have stopped booking. Every cleaning "
                  "business has some; the question is how fast you "
                  "replace them."),
    ]
    for term, meaning in glossary:
        s.write(row, 1, term, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, meaning, fmts.chrome("label"))
        s.row_height(row, 24)
        row += 1
    row += 1

    # ----------------------------------------------- feature map -------
    row = s.section(row, "Where every feature lives", LAST_COL)
    row = s.table_head(row, 1, ["If you want to...", "Go to"],
                       [50, 48])
    where = [
        ("Price a clean before you quote",
         "Pricing & Quote Calculator"),
        ("See what a job will really cost you",
         "Pricing & Quote Calculator, section 5"),
        ("Compare weekly, fortnightly and monthly over a year",
         "Pricing & Quote Calculator, section 8"),
        ("Keep alarm codes, key numbers and parking notes",
         "Property & Site Database"),
        ("Know which clients are drifting away", "Client Retention"),
        ("Find out which advertising actually works",
         "Reports & Analytics"),
        ("See who is free on Thursday",
         "Employee Schedule & Attendance"),
        ("Check nobody is double booked",
         "Job Scheduler, Check column"),
        ("Build a room-by-room task list", "Cleaning Checklists"),
        ("Give the cleaner a sheet to take with them",
         "Print Job Sheet"),
        ("Work out what a job actually earned",
         "Job Costing & Profitability"),
        ("Chase an unpaid invoice", "Invoices, Action column"),
        ("Know what to reorder", "Supplies Inventory"),
        ("Remember to service the carpet machine",
         "Equipment & Assets"),
        ("Claim your mileage", "Mileage & Travel"),
        ("Keep insurance and background checks in date",
         "Safety & Compliance"),
        ("Work out how many jobs cover the bills", "Financial Plan"),
        ("Decide whether to hire",
         "Financial Plan, section 3"),
        ("Put your prices up", "Business Settings, then re-read "
                               "Job Costing"),
        ("Change a drop-down list", "Settings & Dropdown Lists"),
    ]
    for want, go in where:
        s.merge(row, 1, 5, want, fmts.chrome("label"))
        s.merge(row, 6, LAST_COL, go, fmts.chrome("label_bold"))
        s.row_height(row, 20)
        row += 1
    row += 1

    # ----------------------------------------------- questions ---------
    row = s.section(row, "Questions people ask", LAST_COL)
    row = s.table_head(row, 1, ["Question", "Answer"], [30, 68])
    faqs = [
        ("Can two people use it at once?",
         "Not safely in Excel on a shared drive. Upload it to Google "
         "Sheets or OneDrive if more than one person needs it open."),
        ("Will it work on a phone?",
         "It opens, and the dashboard reads well enough. Typing is "
         "painful on anything this wide - use the printed job "
         "sheet."),
        ("How many clients can it hold?",
         "The blank workbook is sized for a few hundred clients and "
         "several thousand jobs. Build it with more rows if you need "
         "them - the generator takes a row count."),
        ("Can I change the colours?",
         "Yes. They are ordinary cell formats - select and restyle."),
        ("Can I delete tabs I do not use?",
         "You can, but other tabs may reference them. Hiding is "
         "safer: right-click the tab, Hide."),
        ("Does it handle more than one currency?",
         "One at a time. The symbol is on the Settings tab."),
        ("Is the sheet password secure?",
         "No. Sheet protection stops accidents, not people. Protect "
         "the file itself if the data is sensitive - and it is."),
        ("Can I use it for a window cleaning or gardening round?",
         "Yes, with almost no changes. Rename the services and the "
         "checklist areas and the rest follows."),
    ]
    for question, answer in faqs:
        s.write(row, 1, question, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, answer, fmts.chrome("label"))
        s.row_height(row, 26)
        row += 1
    row += 1

    s.note(row,
           f"{cfg.product_name}, version {cfg.version}. Built by "
           f"{cfg.author} with formulas only - no macros, no add-ins, "
           f"nothing to install and nothing that phones home. Every "
           f"calculation on every tab can be read, checked and "
           f"changed by you.",
           LAST_COL, first_col=1, height=34)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    return s
