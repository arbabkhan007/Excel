"""Pack a generator into a single zip people can download and run.

    python tools/make_download_zip.py
    python tools/make_download_zip.py --product social
    python tools/make_download_zip.py --product meal

Writes `download/<product>-source.zip` containing the generator, the
quality tools, the documentation and a `requirements.txt`, all under one
folder so it unpacks tidily.

The vendored copy of XlsxWriter in this repository is deliberately left
out: the zip is meant to be unpacked anywhere and run against a normal
`pip install XlsxWriter`.  Timestamps are fixed so rebuilding the zip
without changing any source produces exactly the same bytes.
"""

from __future__ import annotations

import argparse
import os
import sys
import zipfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

#: Same date for every entry, so the zip only changes when the code does.
STAMP = (2026, 1, 1, 0, 0, 0)

SHARED_TOOLS = ["tools/recalc_check.py", "tools/export_source_md.py",
                "tools/make_download_zip.py"]

REQUIREMENTS = """\
# The generator needs one package.
XlsxWriter>=3.0

# Only for the optional checks in tools/
# openpyxl>=3.1     the validator and the sample data check
# formulas>=1.3     recalc_check.py
"""

PET_HOW_TO_RUN = """\
Pet Sitting & Dog Walking Business Spreadsheet
==============================================

    pip install XlsxWriter
    python build_pet_sitting_workbook.py

The workbook appears in:

    output/Pet Sitting Business Spreadsheet.xlsx

Useful options
--------------
    --blank                 empty workbook, no example data
    --business "My Name"    your business name on the cover and invoices
    --currency EUR-symbol   e.g. --currency "\u20ac" --currency-code EUR
    --unit miles            mileage in miles instead of km
    --password secret       change the sheet protection password
    --author "Your Name"    change the document author
    --rows 1000             more rows of capacity in the big registers

Every formula cell is locked and every sheet is protected with the
password `premium` so the maths cannot be typed over by accident; the
cream coloured cells are the ones meant for typing in.

README.md explains how it is built, what is on each of the 27 tabs and
how to change things.  SOURCE_CODE.md is the whole program in one
readable document.
"""

SOCIAL_HOW_TO_RUN = """\
Social Media Content Calendar & Analytics Planner
=================================================

    pip install XlsxWriter
    python build_social_media_workbook.py

The workbook appears in:

    output/Social Media Content Planner.xlsx

Useful options
--------------
    --blank                 empty workbook, no example data
    --brand "My Brand"      your brand name on the cover and dashboards
    --year 2027             the year the 12 month calendar is built for
    --currency EUR-symbol   e.g. --currency "\u20ac" --currency-code EUR
    --password secret       change the sheet protection password
    --author "Your Name"    change the document author
    --rows 600              more rows of capacity in the big registers

Every formula cell is locked and every sheet is protected with the
password `premium` so the maths cannot be typed over by accident; the
cream coloured cells are the ones meant for typing in.

Start on the "Start Here" tab inside the workbook.  README_SOCIAL.md
explains how it is built and what is on each of the 43 tabs;
SOURCE_CODE_SOCIAL.md is the whole program in one readable document.
"""

MEAL_HOW_TO_RUN = """\
Automated Meal Planner & Grocery List Spreadsheet
==================================================

    pip install XlsxWriter
    python build_meal_planner_workbook.py

The workbook appears in:

    output/Meal Planner & Grocery List.xlsx

Useful options
--------------
    --blank                 empty workbook, no example data
    --household "Our House" your household name on the cover and prints
    --cook "Your Name"      the main cook / planner
    --servings 6            how many people a meal is normally cooked for
    --budget 200            the weekly food budget
    --units Imperial        lead with cups, ounces and pounds
    --year 2027             the year the planner is built for
    --week-start Sunday     weeks run Sunday to Saturday
    --currency EUR-symbol   e.g. --currency "\u20ac" --currency-code EUR
    --password secret       change the sheet protection password
    --author "Your Name"    change the document author
    --rows 200              more rows of capacity in the big registers

Every formula cell is locked and every sheet is protected with the
password `premium` so the maths cannot be typed over by accident; the
cream coloured cells are the ones meant for typing in.

Fill in Settings and Household first, then Ingredients and Recipes, then
plan some meals - the Grocery List builds itself from the plan and takes
off whatever the Pantry says you already have.  README_MEAL.md explains
how it is built and what is on each of the 36 tabs; SOURCE_CODE_MEAL.md
is the whole program in one readable document.
"""


CLEANING_HOW_TO_RUN = """\
Cleaning Business CRM & Pricing Calculator
============================================

    pip install XlsxWriter
    python build_cleaning_workbook.py

The workbook appears in:

    output/Cleaning Business CRM & Pricing Calculator.xlsx

Useful options
--------------
    --blank                 empty workbook, no example data
    --business "Name Ltd"   your business name on the cover and prints
    --owner "Your Name"     the owner / manager
    --rate 45               hourly rate charged to the client
    --pay 18                hourly pay rate of a cleaner
    --minimum 90            minimum charge for any job
    --tax 20                sales tax / VAT rate, 0 if you are not registered
    --terms 14              payment terms in days
    --units "sq m"          price area work in square metres
    --distance km           log mileage in kilometres
    --year 2027             the year the planner and reports open on
    --week-start Sunday     weeks run Sunday to Saturday
    --currency EUR-symbol   e.g. --currency "\u20ac" --currency-code EUR
    --password secret       change the sheet protection password
    --author "Your Name"    change the document author
    --rows 1200             more rows of capacity in the big registers

Every formula cell is locked and every sheet is protected with the
password `premium` so the maths cannot be typed over by accident; the
cream coloured cells are the ones meant for typing in.

Fill in Settings, then Clients and their Properties, then the Service
Catalog - after that the Pricing & Quote Calculator will price a job
from the property, the Job Scheduler will cost it and the Invoices tab
will bill it.  README_CLEANING.md explains how it is built and what is
on each of the 37 tabs; SOURCE_CODE_CLEANING.md is the whole program in
one readable document.
"""


HANDMADE_HOW_TO_RUN = """\
Handmade Product Pricing & Profit Calculator
============================================

    pip install XlsxWriter
    python build_handmade_workbook.py

The workbook appears in:

    output/Handmade Product Pricing & Profit Calculator.xlsx

Useful options
--------------
    --blank                 empty workbook, no example shop
    --shop "My Shop"        shop name on the cover and the printables
    --owner "Your Name"     the maker
    --currency EUR-symbol   e.g. --currency "\u20ac" --currency-code EUR
    --rate 22               what you pay yourself an hour
    --margin 60             target retail margin, percent
    --wholesale-discount 50 the discount a stockist gets off retail
    --rounding "Nearest 0.95"   how a calculated price is rounded
    --tax 20                sales tax / VAT rate, percent
    --password secret       change the sheet protection password
    --author "Your Name"    change the document author
    --rows 400              more rows of capacity in the big registers

Every formula cell is locked and every sheet is protected with the
password `premium`; the cream coloured cells are the ones meant for
typing in.

Fill in Settings, then Materials, then Products - after that the Retail
Calculator prices anything in the catalog and the Dashboard tells you
which products are quietly losing money.  README_HANDMADE.md explains
how it is built and what is on each of the 40 tabs; SOURCE_CODE_HANDMADE.md
is the whole program in one readable document.
"""


SALON_HOW_TO_RUN = """\
Salon Business Management Spreadsheet
=====================================

    pip install XlsxWriter
    python build_salon_workbook.py

The workbook appears in:

    output/Salon Business Management Spreadsheet.xlsx

Useful options
--------------
    --blank                 empty workbook, no example salon
    --salon "My Salon"      your salon name on the cover and reports
    --owner "Your Name"     the owner
    --currency EUR-symbol   e.g. --currency "\u20ac" --currency-code EUR
    --tax 20 --tax-on both  sales tax rate, and what it applies to
                            (both | retail | services | neither)
    --commission 45         default service commission percentage
    --commission-method ... the method new stylists start on
    --hourly 15             default hourly rate
    --burden 12             employer costs on wages, percent
    --deposit 25            deposit as a percentage of the price
    --chairs 6 --open-days 6   capacity, for the utilisation figures
    --target 24000          monthly revenue target
    --password secret       change the sheet protection password
    --author "Your Name"    change the document author
    --rows 2000             more rows of capacity in the big registers

Every formula cell is locked and every sheet is protected with the
password `premium` so the maths cannot be typed over by accident; the
white cells are the ones meant for typing in.

Fill in Settings first - tax, deposits, commission defaults, targets -
then the Services menu, then the team on Staff.  After that the diary
feeds the till, the till feeds the commissions, the commissions feed
the payroll and the payroll lands in the Profit & Loss.
README_SALON.md explains how it is built and what is on each of the 20
tabs; SOURCE_CODE_SALON.md is the whole program in one readable
document.

It does not send confirmations, texts or emails, it does not sync with
any booking platform, and it is not accounting, payroll or tax advice.
"""


PRODUCTS = {
    "pet": {
        "folder": "pet-sitting-spreadsheet",
        "output": os.path.join("download",
                               "pet-sitting-spreadsheet-source.zip"),
        "files": ["build_pet_sitting_workbook.py", "README.md",
                  "SOURCE_CODE.md"],
        "trees": ["pet_sitting_spreadsheet"],
        "tools": ["tools/validate_workbook.py", "tools/check_sample_data.py"]
                 + SHARED_TOOLS,
        "how_to_run": PET_HOW_TO_RUN,
    },
    "social": {
        "folder": "social-media-content-planner",
        "output": os.path.join(
            "download", "social-media-content-planner-source.zip"),
        "files": ["build_social_media_workbook.py", "README_SOCIAL.md",
                  "SOURCE_CODE_SOCIAL.md"],
        "trees": ["social_media_spreadsheet"],
        "tools": ["tools/validate_social_workbook.py",
                  "tools/check_social_sample.py"] + SHARED_TOOLS,
        "how_to_run": SOCIAL_HOW_TO_RUN,
    },
    "meal": {
        "folder": "meal-planner-grocery-list",
        "output": os.path.join(
            "download", "meal-planner-grocery-list-source.zip"),
        "files": ["build_meal_planner_workbook.py", "README_MEAL.md",
                  "SOURCE_CODE_MEAL.md"],
        "trees": ["meal_planner_spreadsheet"],
        "tools": ["tools/validate_meal_workbook.py",
                  "tools/check_meal_sample.py"] + SHARED_TOOLS,
        "how_to_run": MEAL_HOW_TO_RUN,
    },
    "cleaning": {
        "folder": "cleaning-business-crm",
        "output": os.path.join(
            "download", "cleaning-business-crm-source.zip"),
        "files": ["build_cleaning_workbook.py", "README_CLEANING.md",
                  "SOURCE_CODE_CLEANING.md"],
        "trees": ["cleaning_business_spreadsheet"],
        "tools": ["tools/validate_cleaning_workbook.py",
                  "tools/check_cleaning_sample.py"] + SHARED_TOOLS,
        "how_to_run": CLEANING_HOW_TO_RUN,
    },
    "handmade": {
        "folder": "handmade-pricing-calculator",
        "output": os.path.join(
            "download", "handmade-pricing-calculator-source.zip"),
        "files": ["build_handmade_workbook.py", "README_HANDMADE.md",
                  "SOURCE_CODE_HANDMADE.md"],
        "trees": ["handmade_pricing_spreadsheet"],
        "tools": ["tools/validate_handmade_workbook.py",
                  "tools/check_handmade_sample.py"] + SHARED_TOOLS,
        "how_to_run": HANDMADE_HOW_TO_RUN,
    },
    "salon": {
        "folder": "salon-business-spreadsheet",
        "output": os.path.join(
            "download", "salon-business-spreadsheet-source.zip"),
        "files": ["build_salon_workbook.py", "README_SALON.md",
                  "SOURCE_CODE_SALON.md"],
        "trees": ["salon_business_spreadsheet"],
        "tools": ["tools/validate_salon_workbook.py",
                  "tools/check_salon_sample.py",
                  "tools/check_parens.py"] + SHARED_TOOLS,
        "how_to_run": SALON_HOW_TO_RUN,
    },
}


def gather(product: dict) -> list:
    """`[(archive name, text)]` for everything that goes in the zip."""
    items = []
    for name in product["files"]:
        with open(os.path.join(ROOT, name), encoding="utf-8") as handle:
            items.append((name, handle.read()))
    for tree in product["trees"]:
        for folder, dirs, names in os.walk(os.path.join(ROOT, tree)):
            dirs[:] = sorted(d for d in dirs if d != "__pycache__")
            for name in sorted(names):
                if not name.endswith(".py"):
                    continue
                full = os.path.join(folder, name)
                rel = os.path.relpath(full, ROOT).replace(os.sep, "/")
                with open(full, encoding="utf-8") as handle:
                    items.append((rel, handle.read()))
    for name in product["tools"]:
        with open(os.path.join(ROOT, name), encoding="utf-8") as handle:
            items.append((name, handle.read()))
    items.append(("requirements.txt", REQUIREMENTS))
    items.append(("HOW TO RUN.txt", product["how_to_run"]))
    return items


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--product", choices=sorted(PRODUCTS),
                        default="pet",
                        help="which generator to pack (default: pet)")
    parser.add_argument("-o", "--output", default=None)
    args = parser.parse_args(argv)

    product = PRODUCTS[args.product]
    out = args.output or product["output"]
    folder = os.path.dirname(os.path.abspath(out))
    if folder and not os.path.isdir(folder):
        os.makedirs(folder)

    items = gather(product)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, text in items:
            info = zipfile.ZipInfo(f"{product['folder']}/{name}",
                                   date_time=STAMP)
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o644 << 16
            zf.writestr(info, text)

    size = os.path.getsize(out)
    lines = sum(len(text.splitlines()) for name, text in items
                if name.endswith(".py"))
    print(f"{out}  ({len(items)} files, {lines:,} lines of Python, "
          f"{size / 1024:.0f} KB)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
