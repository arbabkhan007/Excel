"""Write every source file of a generator into one markdown document.

    python tools/export_source_md.py                   # -> SOURCE_CODE.md
    python tools/export_source_md.py --product social  # -> SOURCE_CODE_SOCIAL.md
    python tools/export_source_md.py -o /tmp/src.md

Handy when the code has to travel somewhere that is not a git checkout:
one file, every module in reading order, each in a fenced block with its
real path above it, plus a short description taken from the module's own
docstring so the two can never disagree.
"""

from __future__ import annotations

import argparse
import ast
import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

SHARED_TOOLS = [
    "tools/export_source_md.py",
    "tools/make_download_zip.py",
]

# ---------------------------------------------------------------------------
# Product A - pet sitting
# ---------------------------------------------------------------------------

PET_SECTIONS = [
    ("1. Run it",
     "Save the files below keeping these exact paths, then run the one "
     "liner. Nothing else is needed - no project file, no build step.\n\n"
     "```bash\n"
     "pip install XlsxWriter\n"
     "python build_pet_sitting_workbook.py\n"
     "```\n\n"
     "The workbook lands in `output/Pet Sitting Business Spreadsheet.xlsx`.",
     ["build_pet_sitting_workbook.py"]),
    ("2. The package",
     "Configuration, the formula template engine, cell formats, the sheet "
     "writer and the demo data. These are the parts every tab leans on.",
     ["pet_sitting_spreadsheet/__init__.py",
      "pet_sitting_spreadsheet/__main__.py",
      "pet_sitting_spreadsheet/cli.py",
      "pet_sitting_spreadsheet/config.py",
      "pet_sitting_spreadsheet/layout.py",
      "pet_sitting_spreadsheet/formats.py",
      "pet_sitting_spreadsheet/writer.py",
      "pet_sitting_spreadsheet/names.py",
      "pet_sitting_spreadsheet/build.py",
      "pet_sitting_spreadsheet/sample_data.py"]),
    ("3. The registers",
     "One module per data tab. Each returns a `SheetSpec` - a list of "
     "columns with their formulas - and the writer turns that into a "
     "formatted, validated, protected sheet.",
     ["pet_sitting_spreadsheet/sheets/__init__.py",
      "pet_sitting_spreadsheet/sheets/common.py",
      "pet_sitting_spreadsheet/sheets/lists.py",
      "pet_sitting_spreadsheet/sheets/settings.py",
      "pet_sitting_spreadsheet/sheets/services.py",
      "pet_sitting_spreadsheet/sheets/clients.py",
      "pet_sitting_spreadsheet/sheets/pets.py",
      "pet_sitting_spreadsheet/sheets/bookings.py",
      "pet_sitting_spreadsheet/sheets/visits.py",
      "pet_sitting_spreadsheet/sheets/health.py",
      "pet_sitting_spreadsheet/sheets/incidents.py",
      "pet_sitting_spreadsheet/sheets/messages.py",
      "pet_sitting_spreadsheet/sheets/invoices.py",
      "pet_sitting_spreadsheet/sheets/payments.py",
      "pet_sitting_spreadsheet/sheets/expenses.py",
      "pet_sitting_spreadsheet/sheets/mileage.py",
      "pet_sitting_spreadsheet/sheets/staff.py",
      "pet_sitting_spreadsheet/sheets/inventory.py",
      "pet_sitting_spreadsheet/sheets/reminders.py"]),
    ("4. The pages",
     "The tabs that read the registers: the hidden calculation engine, the "
     "dashboard, the reports, the planner, the welcome tab, the printable "
     "pages and the help tab.",
     ["pet_sitting_spreadsheet/sheets/calc.py",
      "pet_sitting_spreadsheet/sheets/dashboard.py",
      "pet_sitting_spreadsheet/sheets/reports.py",
      "pet_sitting_spreadsheet/sheets/budget.py",
      "pet_sitting_spreadsheet/sheets/start_here.py",
      "pet_sitting_spreadsheet/sheets/printables.py",
      "pet_sitting_spreadsheet/sheets/help_tab.py"]),
    ("5. Quality checks (optional)",
     "Not needed to build the workbook - these are the checks that keep it "
     "honest. The first two need `openpyxl`, the third needs `formulas`.",
     ["tools/validate_workbook.py",
      "tools/check_sample_data.py",
      "tools/recalc_check.py"] + SHARED_TOOLS),
]

PET = {
    "package": "pet_sitting_spreadsheet",
    "builder": "build_pet_sitting_workbook.py",
    "doc": "README.md",
    "tabs": 27,
    "title": "Pet Sitting & Dog Walking Business Spreadsheet",
    "subtitle": ("Complete source code of the generator that writes the "
                 "`.xlsx` product."),
    "output": "SOURCE_CODE.md",
    "intro": [
        "Writes a 27 tab `.xlsx` workbook with about 50,000 formulas: client "
        "and pet records, bookings, visit logs, pricing, invoices, payments, "
        "expenses, mileage, a dashboard, reports and automatic reminders. "
        "Document author is **Novality Store**; every formula cell is locked "
        "and every sheet protected with the password **`premium`**, while "
        "the cells you are meant to type in stay open.",
        "",
        "Pure Python, one dependency. Tested on XlsxWriter 3.2.0 and on the "
        "much older 0.3.0, so it does not rely on anything recent. The "
        "workbook itself is Excel 2010 safe — no dynamic arrays, no "
        "`XLOOKUP`, no `IFS`, nothing that needs array entry — and it opens "
        "in Google Sheets, LibreOffice and Numbers.",
        "",
        "```bash",
        "python build_pet_sitting_workbook.py --help",
        "",
        "  -o, --output PATH     where to write the file",
        "  --blank               no example data, just the empty workbook",
        "  --no-protect          leave the sheets unlocked",
        "  --password TEXT       protection password (default: premium)",
        "  --author TEXT         document author  (default: Novality Store)",
        "  --business TEXT       business name on the cover and invoices",
        "  --currency SYMBOL     currency symbol  (default: $)",
        "  --unit {km,miles}     mileage unit",
        "  --rows N              rows of capacity in the big registers",
        "  --today YYYY-MM-DD    pretend it is this date (for the demo data)",
        "```",
    ],
    "sections": PET_SECTIONS,
}

# ---------------------------------------------------------------------------
# Product B - social media content planner
# ---------------------------------------------------------------------------

_PKG = "social_media_spreadsheet"

SOCIAL_SECTIONS = [
    ("1. Run it",
     "Save the files below keeping these exact paths, then run the one "
     "liner. Nothing else is needed - no project file, no build step.\n\n"
     "```bash\n"
     "pip install XlsxWriter\n"
     "python build_social_media_workbook.py\n"
     "```\n\n"
     "The workbook lands in "
     "`output/Social Media Content Planner.xlsx`.",
     ["build_social_media_workbook.py"]),
    ("2. The package",
     "Configuration, the formula template engine, cell formats, the sheet "
     "writer, the cross sheet names and the demo data. These are the parts "
     "every tab leans on.",
     [f"{_PKG}/__init__.py",
      f"{_PKG}/__main__.py",
      f"{_PKG}/cli.py",
      f"{_PKG}/config.py",
      f"{_PKG}/layout.py",
      f"{_PKG}/formats.py",
      f"{_PKG}/writer.py",
      f"{_PKG}/names.py",
      f"{_PKG}/build.py",
      f"{_PKG}/sample_data.py"]),
    ("3. Setup & planning registers",
     "One module per data tab. Each returns a `SheetSpec` - a list of "
     "columns with their formulas - and the writer turns that into a "
     "formatted, validated, protected sheet. These are the tabs you fill "
     "in before you start posting.",
     [f"{_PKG}/sheets/__init__.py",
      f"{_PKG}/sheets/common.py",
      f"{_PKG}/sheets/lists.py",
      f"{_PKG}/sheets/settings.py",
      f"{_PKG}/sheets/accounts.py",
      f"{_PKG}/sheets/personas.py",
      f"{_PKG}/sheets/pillars.py",
      f"{_PKG}/sheets/calendar_sheet.py",
      f"{_PKG}/sheets/ideas.py",
      f"{_PKG}/sheets/captions.py",
      f"{_PKG}/sheets/hashtags.py",
      f"{_PKG}/sheets/platforms.py"]),
    ("4. Campaign & production registers",
     "The tabs that run the work: campaigns and launches, the seasonal "
     "plan, collaborations, repurposing, research and the team.",
     [f"{_PKG}/sheets/campaigns.py",
      f"{_PKG}/sheets/launches.py",
      f"{_PKG}/sheets/seasonal.py",
      f"{_PKG}/sheets/influencers.py",
      f"{_PKG}/sheets/ugc.py",
      f"{_PKG}/sheets/repurposing.py",
      f"{_PKG}/sheets/competitors.py",
      f"{_PKG}/sheets/trends.py",
      f"{_PKG}/sheets/assets.py",
      f"{_PKG}/sheets/team.py",
      f"{_PKG}/sheets/tasks.py",
      f"{_PKG}/sheets/compliance.py"]),
    ("5. Measurement registers",
     "Where the numbers go in: monthly account analytics, per post "
     "performance, tests, paid promotion, goals and the budget.",
     [f"{_PKG}/sheets/analytics.py",
      f"{_PKG}/sheets/performance.py",
      f"{_PKG}/sheets/ab_tests.py",
      f"{_PKG}/sheets/ads.py",
      f"{_PKG}/sheets/goals.py",
      f"{_PKG}/sheets/budget.py"]),
    ("6. The pages",
     "The tabs that read the registers: the hidden calculation engine, the "
     "two dashboards, the calendar grid, the production pipeline, the "
     "brand voice page, the reports, the welcome tab, the printables and "
     "the help tab.",
     [f"{_PKG}/sheets/calc.py",
      f"{_PKG}/sheets/dashboard.py",
      f"{_PKG}/sheets/annual.py",
      f"{_PKG}/sheets/month_view.py",
      f"{_PKG}/sheets/workflow.py",
      f"{_PKG}/sheets/brand_voice.py",
      f"{_PKG}/sheets/reports.py",
      f"{_PKG}/sheets/start_here.py",
      f"{_PKG}/sheets/printables.py",
      f"{_PKG}/sheets/help_tab.py"]),
    ("7. Quality checks (optional)",
     "Not needed to build the workbook - these are the checks that keep it "
     "honest. The first two need `openpyxl`, the third needs `formulas`.",
     ["tools/validate_social_workbook.py",
      "tools/check_social_sample.py",
      "tools/recalc_check.py"] + SHARED_TOOLS),
]

SOCIAL = {
    "package": "social_media_spreadsheet",
    "builder": "build_social_media_workbook.py",
    "doc": "README_SOCIAL.md",
    "tabs": 43,
    "title": "Social Media Content Calendar & Analytics Planner",
    "subtitle": ("Complete source code of the generator that writes the "
                 "`.xlsx` product."),
    "output": "SOURCE_CODE_SOCIAL.md",
    "intro": [
        "Writes a 43 tab `.xlsx` workbook with over 110,000 formulas: a "
        "12 month content calendar, an idea bank, caption and hashtag "
        "banks, five platform planners, campaigns, launches, influencer "
        "and UGC tracking, repurposing, competitor and trend research, "
        "monthly analytics, per post performance, two dashboards, printable "
        "briefs and a reports page. Document author is **Novality Store**; "
        "every formula cell is locked and every sheet protected with the "
        "password **`premium`**, while the cells you are meant to type in "
        "stay open.",
        "",
        "Pure Python, one dependency. Tested on XlsxWriter 3.2.0 and on the "
        "much older 0.3.0, so it does not rely on anything recent. The "
        "workbook itself is Excel 2010 safe — no dynamic arrays, no "
        "`XLOOKUP`, no `IFS`, nothing that needs array entry — and it opens "
        "in Google Sheets, LibreOffice and Numbers.",
        "",
        "```bash",
        "python build_social_media_workbook.py --help",
        "",
        "  -o, --output PATH     where to write the file",
        "  --blank               no example data, just the empty workbook",
        "  --no-protect          leave the sheets unlocked",
        "  --password TEXT       protection password (default: premium)",
        "  --author TEXT         document author  (default: Novality Store)",
        "  --brand TEXT          brand name on the cover and the dashboards",
        "  --currency SYMBOL     currency symbol  (default: $)",
        "  --year YYYY           the year the calendar is built for",
        "  --rows N              rows of capacity in the big registers",
        "  --today YYYY-MM-DD    pretend it is this date (for the demo data)",
        "```",
    ],
    "sections": SOCIAL_SECTIONS,
}

# ---------------------------------------------------------------------------
# Product C - automated meal planner
# ---------------------------------------------------------------------------

_MEAL = "meal_planner_spreadsheet"

MEAL_SECTIONS = [
    ("1. Run it",
     "Save the files below keeping these exact paths, then run the one "
     "liner. Nothing else is needed - no project file, no build step.\n\n"
     "```bash\n"
     "pip install XlsxWriter\n"
     "python build_meal_planner_workbook.py\n"
     "```\n\n"
     "The workbook lands in "
     "`output/Meal Planner & Grocery List.xlsx`.",
     ["build_meal_planner_workbook.py"]),
    ("2. The package",
     "Configuration, the formula template engine, cell formats, the sheet "
     "writer, the cross sheet names and the demo household. These are the "
     "parts every tab leans on.",
     [f"{_MEAL}/__init__.py",
      f"{_MEAL}/__main__.py",
      f"{_MEAL}/cli.py",
      f"{_MEAL}/config.py",
      f"{_MEAL}/layout.py",
      f"{_MEAL}/formats.py",
      f"{_MEAL}/writer.py",
      f"{_MEAL}/names.py",
      f"{_MEAL}/build.py",
      f"{_MEAL}/sample_data.py"]),
    ("3. The food registers",
     "One module per data tab. Each returns a `SheetSpec` - a list of "
     "columns with their formulas - and the writer turns that into a "
     "formatted, validated, protected sheet. These five are the chain "
     "that makes the shopping list write itself.",
     [f"{_MEAL}/sheets/__init__.py",
      f"{_MEAL}/sheets/common.py",
      f"{_MEAL}/sheets/lists.py",
      f"{_MEAL}/sheets/settings.py",
      f"{_MEAL}/sheets/members.py",
      f"{_MEAL}/sheets/recipes.py",
      f"{_MEAL}/sheets/recipe_lines.py",
      f"{_MEAL}/sheets/ingredients.py",
      f"{_MEAL}/sheets/meal_plan.py",
      f"{_MEAL}/sheets/grocery.py"]),
    ("4. The kitchen registers",
     "What is in the house and what happens to it: the pantry, the "
     "staples, leftovers, waste, batch cooking, the shops and what you "
     "spend in them, events and goals.",
     [f"{_MEAL}/sheets/pantry.py",
      f"{_MEAL}/sheets/staples.py",
      f"{_MEAL}/sheets/leftovers.py",
      f"{_MEAL}/sheets/waste.py",
      f"{_MEAL}/sheets/prep.py",
      f"{_MEAL}/sheets/spending.py",
      f"{_MEAL}/sheets/store_prices.py",
      f"{_MEAL}/sheets/stores.py",
      f"{_MEAL}/sheets/events.py",
      f"{_MEAL}/sheets/goals.py"]),
    ("5. The pages",
     "The tabs that read the registers: the welcome page, the dashboard, "
     "the two planners, the filters, the money and nutrition pages, the "
     "converter and scaler, rotation, food safety, cold storage, "
     "sustainability, the reports, the three printables and the help "
     "tab.",
     [f"{_MEAL}/sheets/start_here.py",
      f"{_MEAL}/sheets/dashboard.py",
      f"{_MEAL}/sheets/weekly.py",
      f"{_MEAL}/sheets/monthly.py",
      f"{_MEAL}/sheets/filters.py",
      f"{_MEAL}/sheets/nutrition.py",
      f"{_MEAL}/sheets/budget.py",
      f"{_MEAL}/sheets/recipe_costs.py",
      f"{_MEAL}/sheets/converter.py",
      f"{_MEAL}/sheets/rotation.py",
      f"{_MEAL}/sheets/allergens.py",
      f"{_MEAL}/sheets/fridge.py",
      f"{_MEAL}/sheets/sustainability.py",
      f"{_MEAL}/sheets/reports.py",
      f"{_MEAL}/sheets/printables.py",
      f"{_MEAL}/sheets/help_tab.py"]),
    ("6. Quality checks (optional)",
     "Not needed to build the workbook - these are the checks that keep "
     "it honest. The first two need `openpyxl`, the third needs "
     "`formulas`.",
     ["tools/validate_meal_workbook.py",
      "tools/check_meal_sample.py",
      "tools/recalc_check.py"] + SHARED_TOOLS),
]

MEAL = {
    "package": "meal_planner_spreadsheet",
    "builder": "build_meal_planner_workbook.py",
    "doc": "README_MEAL.md",
    "tabs": 36,
    "title": "Automated Meal Planner & Grocery List Spreadsheet",
    "subtitle": ("Complete source code of the generator that writes the "
                 "`.xlsx` product."),
    "output": "SOURCE_CODE_MEAL.md",
    "closing": ("**Activate the first tab** and close the file - there "
                "is no hidden sheet to hide in this one."),
    "engine_heading": "Ranking columns instead of a hidden engine",
    "intro": [
        "Writes a 36 tab `.xlsx` workbook with over 72,000 formulas: a "
        "weekly and a monthly meal planner, a recipe book with costed "
        "ingredient lines, a grocery list that builds itself from the "
        "plan and deducts the pantry, fridge and freezer tracking, "
        "leftovers and waste, a food budget, recipe costing and store "
        "price comparison, unit conversion and recipe scaling, dietary "
        "filters, calories and macros, batch cooking, meal rotation, "
        "food safety, a dashboard, reports and three printable pages. "
        "Document author is **Novality Store**; every formula cell is "
        "locked and every sheet protected with the password "
        "**`premium`**, while the cells you are meant to type in stay "
        "open.",
        "",
        "Pure Python, one dependency. Tested on XlsxWriter 3.2.0 and on "
        "the much older 0.3.0, so it does not rely on anything recent. "
        "The workbook itself is Excel 2010 safe \u2014 no dynamic arrays, "
        "no `XLOOKUP`, no `IFS`, nothing that needs array entry \u2014 and "
        "it opens in Google Sheets, LibreOffice and Numbers.",
        "",
        "```bash",
        "python build_meal_planner_workbook.py --help",
        "",
        "  -o, --output PATH     where to write the file",
        "  --blank               no example data, just the empty workbook",
        "  --no-protect          leave the sheets unlocked",
        "  --password TEXT       protection password (default: premium)",
        "  --author TEXT         document author  (default: Novality Store)",
        "  --household TEXT      household name on the cover and printables",
        "  --cook TEXT           the main cook / planner",
        "  --currency SYMBOL     currency symbol  (default: $)",
        "  --year YYYY           the year the planner is built for",
        "  --week-start DAY      Monday or Sunday",
        "  --servings N          default number of people a meal feeds",
        "  --budget AMOUNT       weekly food budget",
        "  --units {Metric,Imperial}   which unit set the drop downs lead with",
        "  --rows N              rows of capacity in the big registers",
        "  --today YYYY-MM-DD    pretend it is this date (for the demo data)",
        "```",
    ],
    "sections": MEAL_SECTIONS,
}

# ---------------------------------------------------------------------------
# Product D - cleaning business CRM
# ---------------------------------------------------------------------------

_CLEAN = "cleaning_business_spreadsheet"

CLEANING_SECTIONS = [
    ("1. Run it",
     "Save the files below keeping these exact paths, then run the one "
     "liner. Nothing else is needed - no project file, no build step.\n\n"
     "```bash\n"
     "pip install XlsxWriter\n"
     "python build_cleaning_workbook.py\n"
     "```\n\n"
     "The workbook lands in "
     "`output/Cleaning Business CRM & Pricing Calculator.xlsx`.",
     ["build_cleaning_workbook.py"]),
    ("2. The package",
     "Configuration, the formula template engine, cell formats, the sheet "
     "writer, the cross sheet names and the demo cleaning company. These "
     "are the parts every tab leans on.",
     [f"{_CLEAN}/__init__.py",
      f"{_CLEAN}/__main__.py",
      f"{_CLEAN}/cli.py",
      f"{_CLEAN}/config.py",
      f"{_CLEAN}/layout.py",
      f"{_CLEAN}/formats.py",
      f"{_CLEAN}/writer.py",
      f"{_CLEAN}/names.py",
      f"{_CLEAN}/build.py",
      f"{_CLEAN}/sample_data.py"]),
    ("3. The work registers",
     "One module per data tab. Each returns a `SheetSpec` - a list of "
     "columns with their formulas - and the writer turns that into a "
     "formatted, validated, protected sheet. These are the chain that "
     "turns an enquiry into a priced job and a paid invoice.",
     [f"{_CLEAN}/sheets/__init__.py",
      f"{_CLEAN}/sheets/common.py",
      f"{_CLEAN}/sheets/lists.py",
      f"{_CLEAN}/sheets/settings.py",
      f"{_CLEAN}/sheets/clients.py",
      f"{_CLEAN}/sheets/properties.py",
      f"{_CLEAN}/sheets/leads.py",
      f"{_CLEAN}/sheets/services.py",
      f"{_CLEAN}/sheets/quotes.py",
      f"{_CLEAN}/sheets/quote_lines.py",
      f"{_CLEAN}/sheets/jobs.py",
      f"{_CLEAN}/sheets/job_tasks.py",
      f"{_CLEAN}/sheets/checklists.py"]),
    ("4. The people and money registers",
     "Who did the work, what it cost and what came back in: staff, "
     "timesheets, job costing, invoices and their lines, payments, "
     "supplies, equipment, expenses, mileage, complaints, incidents and "
     "compliance documents.",
     [f"{_CLEAN}/sheets/employees.py",
      f"{_CLEAN}/sheets/timesheets.py",
      f"{_CLEAN}/sheets/job_costing.py",
      f"{_CLEAN}/sheets/invoices.py",
      f"{_CLEAN}/sheets/invoice_lines.py",
      f"{_CLEAN}/sheets/payments.py",
      f"{_CLEAN}/sheets/supplies.py",
      f"{_CLEAN}/sheets/supply_usage.py",
      f"{_CLEAN}/sheets/equipment.py",
      f"{_CLEAN}/sheets/expenses.py",
      f"{_CLEAN}/sheets/mileage.py",
      f"{_CLEAN}/sheets/complaints.py",
      f"{_CLEAN}/sheets/incidents.py",
      f"{_CLEAN}/sheets/safety.py"]),
    ("5. The pages",
     "The tabs that read the registers: the welcome page, the dashboard, "
     "the quote calculator, the month calendar, the staff rota, "
     "retention, the reports, the financial plan, the three printables "
     "and the help tab.",
     [f"{_CLEAN}/sheets/start_here.py",
      f"{_CLEAN}/sheets/dashboard.py",
      f"{_CLEAN}/sheets/quote_calc.py",
      f"{_CLEAN}/sheets/calendar_page.py",
      f"{_CLEAN}/sheets/staff_schedule.py",
      f"{_CLEAN}/sheets/retention.py",
      f"{_CLEAN}/sheets/reports.py",
      f"{_CLEAN}/sheets/financial.py",
      f"{_CLEAN}/sheets/printables.py",
      f"{_CLEAN}/sheets/help_tab.py"]),
    ("6. Quality checks (optional)",
     "Not needed to build the workbook - these are the checks that keep "
     "it honest. The first two need `openpyxl`, the third needs "
     "`formulas`.",
     ["tools/validate_cleaning_workbook.py",
      "tools/check_cleaning_sample.py",
      "tools/recalc_check.py"] + SHARED_TOOLS),
]

CLEANING = {
    "package": "cleaning_business_spreadsheet",
    "builder": "build_cleaning_workbook.py",
    "doc": "README_CLEANING.md",
    "tabs": 37,
    "title": "Cleaning Business CRM & Pricing Calculator",
    "subtitle": ("Complete source code of the generator that writes the "
                 "`.xlsx` product."),
    "output": "SOURCE_CODE_CLEANING.md",
    "closing": ("**Activate the first tab** and close the file - there "
                "is no hidden sheet to hide in this one."),
    "engine_heading": "Ranking columns instead of a hidden engine",
    "intro": [
        "Writes a 37 tab `.xlsx` workbook with over 110,000 formulas: a "
        "client CRM and property database, a lead pipeline, a priced "
        "service catalogue, a quote calculator that sizes a job from the "
        "property and prices it, quotes and estimates, a job scheduler "
        "with recurring visits, a month calendar, property specific "
        "cleaning checklists, staff records, a rota with attendance and "
        "timesheets, job costing and profit per job, invoices, payments, "
        "supplies, equipment, expenses, mileage, retention, complaints, "
        "safety documents, a dashboard, reports, a financial plan and "
        "three printable forms. Document author is **Novality Store**; "
        "every formula cell is locked and every sheet protected with the "
        "password **`premium`**, while the cells you are meant to type "
        "in stay open.",
        "",
        "Pure Python, one dependency. Tested on XlsxWriter 3.2.0 and on "
        "the much older 0.3.0, so it does not rely on anything recent. "
        "The workbook itself is Excel 2010 safe \u2014 no dynamic arrays, "
        "no `XLOOKUP`, no `IFS`, nothing that needs array entry \u2014 and "
        "it opens in Google Sheets, LibreOffice and Numbers.",
        "",
        "```bash",
        "python build_cleaning_workbook.py --help",
        "",
        "  -o, --output PATH     where to write the file",
        "  --blank               no example data, just the empty workbook",
        "  --no-protect          leave the sheets unlocked",
        "  --password TEXT       protection password (default: premium)",
        "  --author TEXT         document author  (default: Novality Store)",
        "  --business TEXT       business name on the cover and printables",
        "  --owner TEXT          owner / manager name",
        "  --currency SYMBOL     currency symbol  (default: $)",
        "  --currency-code CODE  three letter code (default: USD)",
        "  --year YYYY           the year the planner and reports open on",
        "  --week-start DAY      Monday or Sunday",
        "  --rate AMOUNT         hourly rate charged to the client",
        "  --pay AMOUNT          hourly pay rate of a cleaner",
        "  --minimum AMOUNT      minimum charge for any job",
        "  --tax PERCENT         sales tax / VAT rate, e.g. 20 for 20%",
        "  --terms DAYS          payment terms in days",
        "  --units {sq ft,sq m}  the area unit used for pricing",
        "  --distance {miles,km} the distance unit used for mileage",
        "  --rows N              rows of capacity in the big registers",
        "  --today YYYY-MM-DD    pretend it is this date (for the demo data)",
        "```",
    ],
    "sections": CLEANING_SECTIONS,
}

PRODUCTS = {"pet": PET, "social": SOCIAL, "meal": MEAL,
            "cleaning": CLEANING}


# ---------------------------------------------------------------------------
# Chapter 0 - how the generator is put together (same engine for both
# products, so the prose is written once and filled in per product)
# ---------------------------------------------------------------------------

ARCHITECTURE = """\
Everything here is plain Python writing a `.xlsx` file through XlsxWriter.
There is no template workbook, no copy of an existing file and no manual
step: the spreadsheet is described in code, and the code is the only
source of truth for what ends up in a cell.

Two kinds of tab exist.

* A **register** is a data tab — a header row and hundreds of rows under
  it. Its module exposes `NAME`, a `spec(cfg)` that returns a
  `SheetSpec` (a list of `Col` objects) and, optionally, a
  `decorate(sheet, fmts, cfg)` that adds totals or notes underneath the
  grid. Nothing in a register module writes cells directly.
* A **page** is a hand-built tab — the dashboards, the reports, the
  printables, the welcome and help pages. Its module exposes
  `build(book, fmts, cfg, reg)` and writes with `SheetWriter` helpers:
  `banner`, `section`, `kpi`, `table_head`, `calc_row`, `note`,
  `bullets`, `field`, `merge`, `finish`.

### The build pipeline

`{builder}` parses the command line into a `Config`, then
`{package}.build.build_workbook(path, cfg)` does six things in order:

1. **Open the workbook** and stamp the document properties — title,
   author, company, keywords and a comment naming the protection
   password. The author is the same on every generated file.
2. **Build the `FormatBook`.** Every visual style the product uses is
   created once, up front: a style (text, date, money, percent, status,
   link, …) crossed with a mode (`INPUT` cream and unlocked, `CALC` grey
   and locked, `PLAIN`), plus the page chrome (titles, section bars, KPI
   tiles, notes) and the conditional formatting palette.
3. **Register every `SheetSpec` first**, before a single tab is written.
   This is what lets any sheet refer to any other sheet's columns by
   name regardless of which tab happens to be written first.
4. **Write the tabs in `SHEET_ORDER`.** Registers go through
   `build_register()` — `SheetWriter.grid(spec)` writes widths, formats,
   headers, formulas, sample rows, drop downs, header comments,
   conditional formats, the autofilter, frozen panes and the print
   setup — then the module's `decorate()` runs and the sheet is
   protected. Pages call their own `build()`.
5. **Define the live lookup names** (`ContentIDs`, `TeamNames`, …) from
   the registry. Each is an `OFFSET` range that grows and shrinks with
   the number of filled rows, so a drop down never shows blanks.
6. {closing}

### Describing a column

A column is declared once, for one row, with placeholders:

```python
Col("er", "Engagement rate", 11, "percent",
    formula='=IFERROR({{engagements}}{{r}}/{{reach}}{{r}},"")',
    cf=[{{"type": "cell", "criteria": ">=", "value": "=TargetEngagement",
         "fmt": "good"}}])
```

`Col` carries the key, the header text, the width, the style, the
formula template, the data validation, the header comment, the
conditional formats, whether the column is hidden and which outline
group it belongs to. `Registry.render()` expands the template per row:

| Token | Becomes |
|-------|---------|
| `{{r}}` / `{{r_prev}}` | this row number / the row above |
| `{{key}}` | the column letter of another column on the same sheet |
| `{{first}}` / `{{last}}` | the first and last data row of the sheet |
| `[[Sheet.key]]` | an absolute range on another sheet, e.g. `'Content Calendar'!$G$5:$G$604` |
| `[[Sheet.key#12]]` | one cell of that column |
| `[[Sheet.key#]]` | the same row as the formula being written |

Helpers cover the patterns Excel 2010 makes awkward — `max_if`,
`min_if`, `blank_if`, `safe`, `lookup`, `dynamic_name`, `quote_sheet`.

### The rules every formula obeys

* **Excel 2010 safe.** `SUMIFS`, `COUNTIFS`, `INDEX/MATCH`, `SMALL`,
  `LARGE`, `RANK`, `EDATE`, `EOMONTH`, `SUMPRODUCT(MAX(...))`. No
  dynamic arrays, no `XLOOKUP`, no `IFS`, no `AGGREGATE`, nothing that
  needs Ctrl+Shift+Enter. The file opens in Excel 2010 and later, Google
  Sheets, LibreOffice and Numbers.
* **No charts.** The vendored XlsxWriter in this repository cannot write
  them, so trends are drawn with `add_sparkline`, data bars and
  three-colour scales — which also survive a Google Sheets import.
* **Nth-row lookups without array formulas.** Hidden composite key
  columns (`date + ROW()/100000`) plus `SMALL` and `INDEX/MATCH` answer
  "the third thing on 14 March" and "the next eight due" with ordinary
  formulas.
* **One number, one home.** Nothing is typed into two tabs. Every other
  appearance of a figure is a formula pointing back at the one place it
  was entered.
* **Locked where it matters.** Formula cells are locked, input cells are
  not, and every sheet is protected with the same password — so the
  maths cannot be typed over by accident but the cream cells stay open.

{engine}
"""

DEFAULT_CLOSING = ("**Hide the `Calc` engine**, activate the first tab "
                   "and close the file.")

CALC_ENGINE = """\
### The hidden `Calc` tab

Dashboards do not aggregate the registers directly. A hidden sheet
pre-computes one row per month, per platform/service, per category and
per owner, each with its own ranking key, and the visible pages read
those columns with plain `INDEX`/`MATCH`. It keeps the heavy `SUMIFS`
work in one place, keeps recalculation predictable, and makes every
dashboard number traceable to a single cell."""

RANK_ENGINE = """\
### Ranking columns instead of a hidden engine

There is no hidden `Calc` tab in this product. Anything a page needs to
rank - the next thing to buy, the recipe you have left alone longest,
the leftovers to eat first, the fridge item closest to its date - is a
**hidden rank column on the register itself**:

```python
Col("rotation_rank", "Rotation rank", 9, "int",
    formula='=IF({{name}}{{r}}="","",COUNTIFS({{sort_key_first}}:'
            '{{sort_key_last}},">"&{{sort_key}}{{r}})+'
            'COUNTIFS({{sort_key_first}}:{{sort_key}}{{r}},'
            '{{sort_key}}{{r}}))',
    hidden=True)
```

`COUNTIFS(whole column, ">" & this row) + COUNTIFS(up to this row, this
row)` is a dense rank that breaks ties by row order, so the ranks come
out 1, 2, 3 ... with no gaps and no duplicates. The page then asks for
the nth item with an ordinary `MATCH(n, rank_range, 0)` - a plain range
lookup, which is why there is **no array formula and nothing needing
Ctrl+Shift+Enter anywhere in the workbook**. The tempting alternative,
`LARGE(IF(...))` over a computed array, reads better in Python and
quietly returns the wrong answer in Excel 2010.

The ranks in use: `Ingredients.need_rank` (what to buy next),
`Grocery List.buy_rank` (position on the printed list),
`Ingredients.spend_rank` and `waste_rank`, `Recipes.cost_rank`,
`match_rank`, `rotation_rank`, `times_planned_rank`, `calories_rank`,
`protein_rank` and `fibre_rank`, `Pantry.fridge_rank` and
`freezer_rank`, `Leftovers.queue_rank`, `Meal Prep.todo_rank` and
`Stores.spend_rank`."""


def outline(path: str) -> list:
    """`[(kind, name, note)]` for the classes, functions and constants."""
    with open(os.path.join(ROOT, path), encoding="utf-8") as handle:
        source = handle.read()
    try:
        tree = ast.parse(source)
    except SyntaxError:                                  # pragma: no cover
        return []
    items = []
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            args = [a.arg for a in node.args.args]
            if node.args.vararg:
                args.append("*" + node.args.vararg.arg)
            if node.args.kwonlyargs:
                args.append("*")
                args += [a.arg for a in node.args.kwonlyargs]
            if node.args.kwarg:
                args.append("**" + node.args.kwarg.arg)
            note = first_sentence(ast.get_docstring(node) or "")
            items.append(("function", f"{node.name}({', '.join(args)})",
                          note))
        elif isinstance(node, ast.ClassDef):
            bases = []
            for base in node.bases:
                try:
                    bases.append(ast.unparse(base))
                except Exception:                        # pragma: no cover
                    pass
            name = node.name + (f"({', '.join(bases)})" if bases else "")
            note = first_sentence(ast.get_docstring(node) or "")
            fields = [n.target.id for n in node.body
                      if isinstance(n, ast.AnnAssign)
                      and isinstance(n.target, ast.Name)]
            if fields and not note:
                note = "fields: " + ", ".join(fields[:8])
            items.append(("class", name, note))
        elif isinstance(node, (ast.Assign, ast.AnnAssign)):
            targets = ([node.target] if isinstance(node, ast.AnnAssign)
                       else node.targets)
            for target in targets:
                if not isinstance(target, ast.Name):
                    continue
                if not re.match(r"^[A-Z][A-Z0-9_]*$", target.id):
                    continue
                items.append(("constant", target.id,
                              describe_value(node.value)))
    return items


def describe_value(node) -> str:
    """A short, human description of a module level constant."""
    if node is None:
        return ""
    if isinstance(node, ast.Constant):
        text = repr(node.value)
        return text if len(text) < 60 else f"{text[:57]}..."
    if isinstance(node, (ast.List, ast.Tuple)):
        return f"{len(node.elts)} items"
    if isinstance(node, ast.Dict):
        return f"{len(node.keys)} entries"
    if isinstance(node, ast.Call):
        try:
            name = ast.unparse(node.func)
        except Exception:                                # pragma: no cover
            name = "call"
        return f"{name}(...)"
    try:
        text = ast.unparse(node)
    except Exception:                                    # pragma: no cover
        return ""
    return text if len(text) < 60 else f"{text[:57]}..."


def first_sentence(text: str) -> str:
    text = " ".join(text.split())
    if not text:
        return ""
    cut = re.search(r"\.(\s|$)", text)
    if cut:
        text = text[:cut.start()]
    return text.strip().rstrip(".")


def blurb(path: str) -> str:
    """First sentence of the module docstring, as a one line description."""
    with open(os.path.join(ROOT, path), encoding="utf-8") as handle:
        source = handle.read()
    try:
        text = ast.get_docstring(ast.parse(source)) or ""
    except SyntaxError:                                  # pragma: no cover
        text = ""
    return first_sentence(text)


def full_docstring(path: str) -> str:
    with open(os.path.join(ROOT, path), encoding="utf-8") as handle:
        source = handle.read()
    try:
        return (ast.get_docstring(ast.parse(source)) or "").strip()
    except SyntaxError:                                  # pragma: no cover
        return ""


def anchor(heading: str) -> str:
    slug = heading.lower()
    slug = re.sub(r"[^\w\s-]", "", slug)
    return re.sub(r"\s+", "-", slug.strip())


def tab_table(product: dict) -> list:
    """A row per tab: order, name, kind, module, columns, rows.

    Imports the product package, so it degrades to nothing if the
    generator is not importable from where the tool is run.
    """
    if ROOT not in sys.path:
        sys.path.insert(0, ROOT)
    try:
        package = __import__(f"{product['package']}.build",
                             fromlist=["build"])
        config = __import__(f"{product['package']}.config",
                            fromlist=["config"])
    except Exception:                                    # pragma: no cover
        return []
    cfg = config.Config()
    lines = ["| # | Tab | Kind | Written by | Columns | Rows |",
             "|---|-----|------|------------|--------:|-----:|"]
    for number, name in enumerate(package.SHEET_ORDER, start=1):
        module = package.REGISTER_BY_NAME.get(name)
        if module is not None:
            spec = module.spec(cfg)
            # a register is usually a module, but the platform planners
            # are instances of one shared class
            where = (getattr(module, "__name__", None)
                     or type(module).__module__)
            path = where.replace(".", "/") + ".py"
            lines.append(f"| {number} | **{name}** | register | "
                         f"`{path}` | {len(spec.columns)} | {spec.rows} |")
        else:
            builder = package.PAGE_BUILDERS.get(name)
            path = (getattr(builder, "__module__", "")
                    .replace(".", "/") + ".py")
            lines.append(f"| {number} | **{name}** | page | `{path}` "
                         f"| - | - |")
    return lines


def data_dictionary(product: dict) -> list:
    """One table per register: every column, its type and where it comes
    from. Generated from the specs themselves, so it cannot drift."""
    if ROOT not in sys.path:
        sys.path.insert(0, ROOT)
    try:
        package = __import__(f"{product['package']}.build",
                             fromlist=["build"])
        config = __import__(f"{product['package']}.config",
                            fromlist=["config"])
    except Exception:                                    # pragma: no cover
        return []
    cfg = config.Config()
    lines = []
    for name in package.SHEET_ORDER:
        module = package.REGISTER_BY_NAME.get(name)
        if module is None:
            continue
        spec = module.spec(cfg)
        typed = sum(1 for c in spec.columns if not c.is_formula)
        lines += [f"#### {spec.name}", ""]
        if spec.subtitle:
            lines += [f"{spec.subtitle}", ""]
        lines += [f"{len(spec.columns)} columns, {spec.rows} rows of "
                  f"capacity, {typed} of them typed in by hand.", "",
                  "| Col | Header | Key | Format | Filled by |",
                  "|-----|--------|-----|--------|-----------|"]
        for col in spec.columns:
            source = "formula" if col.is_formula else "you type it"
            if col.validate:
                picked = str(col.validate.get("source", "")).lstrip("=")
                if picked and re.match(r"^[A-Za-z_][A-Za-z0-9_]*$", picked):
                    source = f"drop down `{picked}`"
                elif not col.is_formula:
                    source = "you type it (checked)"
            if col.hidden:
                source += ", hidden"
            lines.append(f"| {spec.letter(col.key)} | {col.header} | "
                         f"`{col.key}` | {col.style} | {source} |")
        lines.append("")
    return lines


def file_table(sections: list) -> list:
    lines = ["| File | Lines | What it is |", "|------|------:|------------|"]
    for _heading, _intro, paths in sections:
        for path in paths:
            with open(os.path.join(ROOT, path), encoding="utf-8") as handle:
                count = len(handle.read().splitlines())
            note = blurb(path) or ""
            if len(note) > 110:
                note = note[:107].rsplit(" ", 1)[0] + "..."
            lines.append(f"| [`{path}`](#{anchor(path)}) | {count:,} | "
                         f"{note} |")
    return lines


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--product", choices=sorted(PRODUCTS),
                        default="pet",
                        help="which generator to export (default: pet)")
    parser.add_argument("-o", "--output", default=None)
    parser.add_argument("--plain", action="store_true",
                        help="just the code, without the architecture "
                             "chapter and the per file outlines")
    args = parser.parse_args(argv)

    product = PRODUCTS[args.product]
    sections = product["sections"]
    detailed = not args.plain

    lines = [f"# {product['title']} — source code", "",
             product["subtitle"], ""]
    lines += product["intro"]
    lines += ["", "---", "", "## Contents", ""]

    if detailed:
        lines += ["* [0. How it is put together](#0-how-it-is-put-together)",
                  "  * [The build pipeline](#the-build-pipeline)",
                  "  * [Describing a column](#describing-a-column)",
                  "  * [The rules every formula obeys]"
                  "(#the-rules-every-formula-obeys)",
                  f"  * [{product.get('engine_heading', 'The hidden `Calc` tab')}]"
                  f"(#{anchor(product.get('engine_heading', 'The hidden Calc tab'))})",
                  "  * [Every tab at a glance](#every-tab-at-a-glance)",
                  "  * [Every file at a glance](#every-file-at-a-glance)"]
    for heading, _intro, paths in sections:
        lines.append(f"* [{heading}](#{anchor(heading)})")
        for path in paths:
            lines.append(f"  * [`{path}`](#{anchor(path)})")
    if detailed:
        lines.append("* [Appendix A. Data dictionary]"
                     "(#appendix-a-data-dictionary)")
    lines += ["", "---", ""]

    if detailed:
        lines += ["## 0. How it is put together", ""]
        lines += [ARCHITECTURE.format(
            builder=product["builder"], package=product["package"],
            engine=(RANK_ENGINE if product.get("engine_heading")
                    else CALC_ENGINE),
            closing=product.get("closing", DEFAULT_CLOSING)), ""]
        tabs = tab_table(product)
        if tabs:
            lines += ["### Every tab at a glance", "",
                      f"{product['tabs']} tabs in the order they appear in "
                      f"the finished workbook. \"Rows\" is the capacity of "
                      f"a register at the default `--rows` setting.", ""]
            lines += tabs + [""]
        lines += ["### Every file at a glance", ""]
        lines += file_table(sections) + ["", "---", ""]

    total = 0
    for heading, intro, paths in sections:
        lines += [f"## {heading}", "", intro, ""]
        for path in paths:
            full = os.path.join(ROOT, path)
            with open(full, encoding="utf-8") as handle:
                code = handle.read().rstrip("\n")
            count = len(code.splitlines())
            total += count
            lines.append(f"### `{path}`")
            lines.append("")
            if detailed:
                items = outline(path)
                classes = sum(1 for kind, _n, _d in items
                              if kind == "class")
                functions = sum(1 for kind, _n, _d in items
                                if kind == "function")
                stats = [f"{count:,} lines"]
                if classes:
                    stats.append(f"{classes} class"
                                 f"{'es' if classes > 1 else ''}")
                if functions:
                    stats.append(f"{functions} function"
                                 f"{'s' if functions > 1 else ''}")
                lines += [f"*{' · '.join(stats)}*", ""]
                doc = full_docstring(path)
                if doc:
                    lines += ["> " + line if line else ">"
                              for line in doc.splitlines()]
                    lines.append("")
                named = [item for item in items if item[0] != "constant"]
                constants = [item for item in items if item[0] == "constant"]
                if named:
                    lines += ["| Defines | What it does |",
                              "|---------|--------------|"]
                    for kind, name, note in named:
                        tag = "class " if kind == "class" else ""
                        lines.append(f"| `{tag}{name}` | {note} |")
                    lines.append("")
                if constants:
                    text = ", ".join(f"`{name}` ({note})" if note
                                     else f"`{name}`"
                                     for _k, name, note in constants[:14])
                    more = (f" and {len(constants) - 14} more"
                            if len(constants) > 14 else "")
                    lines += [f"**Module constants:** {text}{more}.", ""]
            else:
                note = blurb(path)
                if note:
                    lines.append(f"{note}. *({count} lines)*")
                else:
                    lines.append(f"*({count} lines)*")
                lines.append("")
            # four backticks, so a fence inside a docstring cannot
            # close the block early
            lines += ["````python", code, "````", ""]
        lines.append("---")
        lines.append("")

    if detailed:
        dictionary = data_dictionary(product)
        if dictionary:
            lines += ["## Appendix A. Data dictionary", "",
                      "Every column of every register tab, straight from "
                      "the specs that build them: the Excel column letter, "
                      "the header a buyer sees, the key formulas use, the "
                      "number format and whether the cell is typed in, "
                      "chosen from a drop down or calculated. Pages "
                      "(dashboards, reports, printables) are not listed - "
                      "they have no fixed grid.", ""]
            lines += dictionary + ["---", ""]

    lines += [f"*{total:,} lines of Python in "
              f"{sum(len(p) for _h, _i, p in sections)} files.*", ""]

    out = args.output or product["output"]
    with open(out, "w", encoding="utf-8") as handle:
        handle.write("\n".join(lines))
    size = os.path.getsize(out)
    print(f"{out}  ({total:,} lines of source, {size / 1024:.0f} KB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
