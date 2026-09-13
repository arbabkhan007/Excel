# 🍽️ Ultimate Catering Business Spreadsheet - Python Source Code

> **Catering Business Manager | Excel & Google Sheets Template**
> Version: 1.0 | No macros, nothing to install

This document contains the complete Python source code that generates the
**Ultimate Catering Business Spreadsheet** workbooks: up to 20 worksheets,
6 charts, 1,613 auto-calculating formulas, 82 drop-down validations and 76
conditional-formatting rules per premium workbook - in two editions
(Basic / Premium), two visual themes (Classic / Fresh) and two fill states
(blank template / filled EXAMPLE for listing screenshots).

---

## 📋 Table of Contents

1. Overview
2. Requirements
3. How to Run
4. Package Layout
5. The Sheets
6. Design System
7. How the Automation Works
8. Quality Checks
9. Complete Source Code
10. Appendix: QA Tooling Source
11. Changelog

---

## 📊 Overview

| Feature | Premium | Basic |
|---------|---------|-------|
| Worksheets | 20 (19 visible + hidden `_Data`) | 9 (8 visible + hidden `_Data`) |
| Chart visuals | 6 (8 chart objects) | 4 (5 chart objects) |
| Formulas | 1,613 | 896 |
| Data validations | 82 | 45 |
| Conditional formats | 76 | 42 |
| Named ranges | 22 | 20 |
| File size | ~0.7 MB (incl. cover art) | ~0.6 MB (incl. cover art) |

Both editions open straight into Excel 2016+ or Google Sheets (File > Import >
Upload). There are no macros and no add-ins.

### 🎨 Colour palettes

| Role | Classic | Fresh |
|------|---------|-------|
| Canvas / card | `#F8F4EC` / `#FFFFFF` | `#FAFBF8` / `#FFFFFF` |
| Primary (espresso / basil) | `#4B3626` | `#2F5D3A` |
| Accent (copper / tomato) | `#B4652F` | `#C4552F` |
| Gold (brass / lemon) | `#A9822B` | `#C9A227` |

---

## 🧰 Requirements

* Python 3.8+
* The vendored `xlsxwriter/` in this repository (no pip install needed)
* Optional, for the QA tooling in `tools/`: `openpyxl`, `formulas`, `Pillow`

---

## ▶️ How to Run

```bash
# the curated Etsy product set (6 files) into products/
python3 catering_business_tracker.py --all --outdir products/

# or one workbook at a time
python3 catering_business_tracker.py --edition premium --theme classic \
    --mode demo --out products/preview.xlsx

# text-only build (no watercolour cover art) for tiny files
python3 catering_business_tracker.py --all --no-images
```

---

## 📦 Package Layout

```
catering_business_tracker.py     CLI: editions, themes, modes, output paths
catering_tracker/
  config.py        sheet names, tab order, column maps, geometry constants
  theme.py         the two palettes (classic / fresh) as attribute bags
  styles.py        cached XlsxWriter format factory + text metrics
  demo.py          the fictional EXAMPLE business + cached aggregate values
  book.py          Workbook wrapper: sheets, names, ranges, CF, nav, pages
  workbook.py      BUILDERS table, build order, filenames, build_all()
  sheets/
    common.py      shared table/DV/CF/chip/note helpers
    data.py        hidden _Data: pools, KPIs, month table (chart fuel)
    setup.py .. guide.py   one module per visible tab
tools/
  verify_workbook.py  openpyxl audit (hidden writes, dangling refs, DV/CF)
  calc_check.py       full recalculation with the `formulas` engine
  layout_check.py     clipping / overflow / row-height audit
  render_preview.py   PNG approximation of a sheet for eyeballing
  make_banner_alpha.py transparent-window pass over the cover artwork
assets/
  banner_classic.png / banner_fresh.png   watercolour Start Here covers
```

---

## 📑 The Sheets

Premium order (Basic keeps the ● rows):

| Tab | Purpose |
|-----|---------|
| ● `_Data` (hidden) | chart fuel: month table, KPI cells, dropdown pools |
| ● ⚙️ Setup | business profile, money defaults, calendar month, all lists |
| ● 📊 Dashboard | KPIs, progress bars, pipeline strip, 4 charts, alert panels |
| ● 👥 Clients | client CRM with balances and payment status |
| ● 📅 Events | master order planner with the 6-stage pipeline |
| ● 🧮 Quote Calculator | costs in → recommended price out |
| ● 💸 Expenses | 11-category spend log |
| ● 💰 Payments | invoices & deposits with OVERDUE auto-flag |
| ○ 🍽️ Menu Costing | dish list + ingredient recipe calculator |
| ○ 📦 Inventory | stock levels with LOW STOCK / REORDER flags |
| ○ 🛒 Shopping List | auto required-vs-in-stock purchasing list |
| ○ 👷 Staff & Labor | shifts, overtime, unpaid flags |
| ○ 🍳 Equipment | fleet status, service dates, replacement value |
| ○ 🚚 Suppliers | vendor database |
| ○ 📆 Event Calendar | month wall built from Setup month/year |
| ○ 📈 P&L & Reports | monthly + annual P&L and four report tables |
| ○ 🧾 Tax Tracker | collected vs paid, quarterly estimates, disclaimer |
| ○ ✅ Checklists | prep / shopping / day-of / end-of-event |
| ○ 🖨️ Invoice & Proposal | print-ready documents from one dropdown |
| ● 📖 Start Here | banner cover, 5-step setup, tab tour, tips |

---

## 🤖 How the Automation Works

* **Defined names** (`EventList`, `ClientsList`, `MenuItems`, `SuppliersList`,
  `Currency`, `TaxRate`, `DefaultMargin`, …) point at OFFSET windows over the
  hidden `_Data` pools, so dropdowns grow and shrink with your data.
* **KPI cells** on `_Data` (column AE/AF) compute every dashboard number with
  SUMIFS/COUNTIFS over the tracker tabs; the dashboard only ever reads those
  names, so one edit ripples everywhere.
* **Cached values**: every formula is written with its expected result so the
  file shows correct numbers before Excel recalculates (and Google Sheets
  imports cleanly).
* **Conditional formatting** drives the colour language: status pills, LOW
  STOCK glow, OVERDUE red, due-soon amber, ticked-box green, margin warnings.
* **Protection**: input cells stay unlocked, formula cells locked; sheets are
  protected without a password.

---

## ✅ Quality Checks

Every build is audited before release:

```bash
/tmp/venv/bin/python tools/verify_workbook.py products/Catering*.xlsx
/tmp/venv/bin/python tools/calc_check.py      products/<file>.xlsx
/tmp/venv/bin/python tools/layout_check.py    products/Catering*.xlsx
```

* `verify_workbook` - no writes hidden under merges, no dangling sheet
  references, DV/CF/chart inventories.
* `calc_check` - the `formulas` engine recalculates all 1,600+ formulas and
  compares against the cached values (the OFFSET dropdown names are a known
  engine limitation and reported separately).
* `layout_check` - no clipped text, no overflow, wrapped rows tall enough.

---

## 📜 Complete Source Code

### `catering_business_tracker.py`

```python
#!/usr/bin/env python3
"""
Build the Ultimate Catering Business Spreadsheet workbooks.

Examples
--------
    python3 catering_business_tracker.py --all
    python3 catering_business_tracker.py --edition premium --theme classic \
        --mode demo --out products/hero.xlsx
    python3 catering_business_tracker.py --edition basic --theme fresh

The script needs nothing but the XlsxWriter library that lives in this
repository (it is imported from the checkout root), so it runs straight from
a fresh clone with a stock Python 3.
"""

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from catering_tracker import build_all, build_workbook  # noqa: E402
from catering_tracker.workbook import product_filename  # noqa: E402


def main(argv=None):
    p = argparse.ArgumentParser(
        description="Build Ultimate Catering Business Spreadsheet workbooks.")
    p.add_argument("--edition", choices=["basic", "premium", "both"],
                   default="premium")
    p.add_argument("--theme", choices=["classic", "fresh", "both"],
                   default="classic")
    p.add_argument("--mode", choices=["blank", "demo", "both"],
                   default="blank",
                   help="blank = clean template, demo = filled example")
    p.add_argument("--all", action="store_true",
                   help="build the curated Etsy product set into products/")
    p.add_argument("--out", default=None,
                   help="output path (single build only)")
    p.add_argument("--outdir", default="products",
                   help="output folder for --all / multi builds")
    p.add_argument("--protect", default=None, metavar="PASSWORD",
                   help="optionally lock every sheet with a password")
    p.add_argument("--no-images", action="store_true",
                   help="skip cover banner images (smaller files)")
    args = p.parse_args(argv)

    if args.all:
        stats = build_all(args.outdir, protect=args.protect,
                          images=not args.no_images)
        _report(stats)
        return 0

    editions = (["basic", "premium"] if args.edition == "both"
                else [args.edition])
    themes_ = (["classic", "fresh"] if args.theme == "both"
               else [args.theme])
    modes = ["blank", "demo"] if args.mode == "both" else [args.mode]

    stats = []
    for edition in editions:
        for theme_name in themes_:
            for mode in modes:
                if (args.out and len(editions) == 1 and len(themes_) == 1
                        and len(modes) == 1):
                    path = args.out
                else:
                    os.makedirs(args.outdir, exist_ok=True)
                    path = os.path.join(
                        args.outdir,
                        product_filename(edition, theme_name, mode))
                stats.append(build_workbook(path, edition, theme_name, mode,
                                            protect=args.protect,
                                            images=not args.no_images))
    _report(stats)
    return 0


def _report(stats):
    print()
    print("=" * 78)
    print("BUILD COMPLETE")
    print("=" * 78)
    for s in stats:
        print("  %-58s %7.1f KB" % (os.path.basename(s["path"]),
                                    s["size"] / 1024.0))
        print("      formulas %5d   validations %3d   cond. formats %3d   "
              "charts %2d   links %3d"
              % (s["formulas"], s["validations"], s["cond_formats"],
                 s["charts"], s["links"]))
    print()


if __name__ == "__main__":
    sys.exit(main())
```

### `catering_tracker/__init__.py`

```python
"""
Ultimate Catering Business Spreadsheet - Etsy-ready Excel workbook builder.

Built on the XlsxWriter library vendored in this repository.  The package is
deliberately split into small modules so that every tab of the product can be
read (and tweaked) on its own:

    theme.py       colour palettes (classic bistro + fresh market)
    config.py      sheet names, capacities, column maps, list contents
    styles.py      every cell format, derived from the active theme
    book.py        the build context: sheet registry, cross-sheet references
    demo.py        the optional "filled-in example" business + aggregates
    sheets/        one module per worksheet
    workbook.py    orchestrates a full build

Entry point:  python catering_business_tracker.py --all
"""

from .workbook import build_workbook, build_all  # noqa: F401

__version__ = "1.0.0"
__product__ = "Ultimate Catering Business Spreadsheet"
```

### `catering_tracker/config.py`

```python
"""
Every layout constant for the Ultimate Catering Business Spreadsheet.

Row numbers, column letters, sheet orders, dropdown values and list positions
all live here so the sheet builders stay readable and the whole workbook can be
re-skinned or re-rowed from one audited place.
"""

PRODUCT = "Ultimate Catering Business Spreadsheet"
PRODUCT_SHORT = "Catering Business Manager"
TAGLINE = "Events • Clients • Quotes • Food Cost • Profit"
AUTHOR = "Novality Store"
VERSION = "1.0"

# ===========================================================================
# SHEETS
# ===========================================================================
SHEET_NAMES = {
    "data": "_Data",
    "setup": "\u2699\uFE0F Setup",
    "dashboard": "\U0001F4CA Dashboard",
    "clients": "\U0001F465 Clients",
    "events": "\U0001F4C5 Events",
    "quote": "\U0001F9EE Quote Calculator",
    "menu": "\U0001F37D\uFE0F Menu Costing",
    "inventory": "\U0001F4E6 Inventory",
    "shopping": "\U0001F6D2 Shopping List",
    "expenses": "\U0001F4B8 Expenses",
    "income": "\U0001F4B0 Payments",
    "staff": "\U0001F477 Staff & Labor",
    "equipment": "\U0001F373 Equipment",
    "suppliers": "\U0001F69A Suppliers",
    "calendar": "\U0001F4C6 Event Calendar",
    "reports": "\U0001F4C8 P&L & Reports",
    "tax": "\U0001F9FE Tax Tracker",
    "checklists": "\u2705 Checklists",
    "invoice": "\U0001F5A8\uFE0F Invoice & Proposal",
    "guide": "\U0001F4D6 Start Here",
}

SHEET_SHORT = {
    "data": "_Data",
    "setup": "Setup",
    "dashboard": "\U0001F3E0 Home",
    "clients": "Clients",
    "events": "Events",
    "quote": "Quote",
    "menu": "Menu",
    "inventory": "Stock",
    "shopping": "Shop",
    "expenses": "Expenses",
    "income": "Payments",
    "staff": "Staff",
    "equipment": "Equip",
    "suppliers": "Suppliers",
    "calendar": "Calendar",
    "reports": "P&L",
    "tax": "Tax",
    "checklists": "Checks",
    "invoice": "Invoice",
    "guide": "Help",
}

EDITIONS = {
    "premium": ["data", "setup", "dashboard", "clients", "events", "quote",
                "menu", "inventory", "shopping", "expenses", "income",
                "staff", "equipment", "suppliers", "calendar", "reports",
                "tax", "checklists", "invoice", "guide"],
    "basic": ["data", "setup", "dashboard", "clients", "events", "quote",
              "expenses", "income", "guide"],
}

# ===========================================================================
# SHARED TABLE GEOMETRY
# ===========================================================================
ROW_SPACER_1 = 1
ROW_TITLE = 2
ROW_SUBTITLE = 3
ROW_SPACER_2 = 4
ROW_STATS = 5
ROW_SPACER_3 = 6
ROW_HEADER = 7
ROW_FIRST = 8

CAP = {                       # data rows per sheet
    "clients": 40,
    "events": 40,
    "menu": 30,
    "inventory": 40,
    "shopping": 40,
    "expenses": 60,
    "income": 60,
    "staff": 50,
    "equipment": 35,
    "suppliers": 30,
}


def last_row(key):
    return ROW_FIRST + CAP[key] - 1


TICK = "\u2713"

# event pipeline
ES_INQUIRY = "\U0001F4AC Inquiry"
ES_QUOTED = "\U0001F4DD Quote Sent"
ES_DEPOSIT = "\U0001F4B0 Deposit Paid"
ES_CONFIRMED = "\u2705 Confirmed"
ES_DONE = "\U0001F389 Completed"
ES_CANCEL = "\u274C Cancelled"
EVENT_STATUSES = [ES_INQUIRY, ES_QUOTED, ES_DEPOSIT, ES_CONFIRMED, ES_DONE,
                  ES_CANCEL]

PS_UNPAID = "\U0001F534 Unpaid"
PS_PART = "\U0001F7E0 Partial"
PS_PAID = "\U0001F7E2 Paid in full"
PAYMENT_STATUSES = [PS_UNPAID, PS_PART, PS_PAID]

EVENT_TYPES = ["Wedding", "Corporate", "Birthday", "Private Party", "Buffet",
               "Cocktail", "Conference", "Festival", "Other"]
EXPENSE_CATEGORIES = ["Ingredients", "Packaging", "Staff / Labor",
                      "Transportation", "Equipment", "Kitchen Rental",
                      "Marketing", "Insurance", "Utilities", "Software",
                      "Miscellaneous"]
PAYMENT_METHODS = ["Cash", "Bank Transfer", "Card", "Cheque", "Mobile Wallet",
                   "Other"]
STAFF_ROLES = ["Head Chef", "Chef", "Sous Chef", "Server", "Bartender",
               "Driver", "Setup Crew", "Cleaner", "Manager", "Other"]
MENU_CATEGORIES = ["Starter", "Main", "Dessert", "Sides", "Beverage",
                   "Canapé", "Other"]
INGREDIENT_CATEGORIES = ["Produce", "Meat & Fish", "Dairy", "Pantry",
                         "Bakery", "Beverages", "Packaging", "Other"]

# ===========================================================================
# SETUP SHEET
# ===========================================================================
SU_BUSINESS = 6
SU_MESSAGE = 7
SU_CURRENCY = 10
SU_TAX = 11
SU_MARGIN = 12
SU_DEPOSIT = 13
SU_DUESOON = 14
SU_YEAR = 15
SU_CAL_MONTH = 18
SU_CAL_YEAR = 19

SU_LIST_HEADER = 22
SU_LIST_FIRST = 23
SU_LIST_ROWS = 20
LIST_COLS = {"event_types": "B", "expense_categories": "C",
             "payment_methods": "D", "staff_roles": "E",
             "menu_categories": "F", "ingredient_categories": "G"}
LIST_TITLES = {"event_types": "Event types",
               "expense_categories": "Expense categories",
               "payment_methods": "Payment methods",
               "staff_roles": "Staff roles",
               "menu_categories": "Menu categories",
               "ingredient_categories": "Ingredient categories"}

SU_FIXED_HEADER = 46
SU_FIXED_FIRST = 47
FIXED_COLS = {"event_statuses": "C", "payment_statuses": "D", "tick": "E"}

CURRENCIES = ["$", "€", "£", "¥", "₨", "₹", "A$", "C$", "R", "kr", "CHF"]

# ===========================================================================
# COLUMN MAPS  (field -> column letter)
# ===========================================================================
COLS = {
    "clients": {"n": "B", "name": "C", "phone": "D", "email": "E",
                "event_date": "F", "event_type": "G", "guests": "H",
                "venue": "I", "package": "J", "quote": "K", "deposit": "L",
                "balance": "M", "status": "N", "notes": "O"},
    "events": {"n": "B", "id": "C", "client": "D", "date": "E", "type": "F",
               "guests": "G", "menu": "H", "staff_req": "I",
               "equipment_req": "J", "cost": "K", "price": "L",
               "profit": "M", "margin": "N", "deposit": "O",
               "balance": "P", "status": "Q", "notes": "R"},
    "menu": {"n": "B", "item": "C", "category": "D", "ingredients": "E",
             "portion": "F", "cost": "G", "price": "H", "profit": "I",
             "margin": "J", "notes": "K"},
    "inventory": {"n": "B", "ingredient": "C", "category": "D", "unit": "E",
                  "qty": "F", "min": "G", "unit_cost": "H", "value": "I",
                  "supplier": "J", "status": "K", "notes": "L"},
    "shopping": {"n": "B", "event": "C", "ingredient": "D", "required": "E",
                 "available": "F", "to_buy": "G", "supplier": "H",
                 "est_cost": "I", "purchased": "J", "date": "K",
                 "notes": "L"},
    "expenses": {"n": "B", "date": "C", "vendor": "D", "category": "E",
                 "desc": "F", "amount": "G", "method": "H", "event": "I",
                 "receipt": "J"},
    "income": {"n": "B", "invoice": "C", "client": "D", "event_date": "E",
               "amount": "F", "deposit": "G", "pay1": "H", "pay2": "I",
               "received": "J", "balance": "K", "due": "L", "status": "M"},
    "staff": {"n": "B", "name": "C", "role": "D", "event": "E", "hours": "F",
              "rate": "G", "ot": "H", "total": "I", "paid": "J",
              "notes": "K"},
    "equipment": {"n": "B", "item": "C", "category": "D", "owned": "E",
                  "reserved": "F", "damaged": "G", "available": "H",
                  "unit_value": "I", "value": "J", "maintenance": "K",
                  "status": "L"},
    "suppliers": {"n": "B", "supplier": "C", "contact": "D", "phone": "E",
                  "email": "F", "category": "G", "item": "H", "price": "I",
                  "min_order": "J", "delivery": "K", "terms": "L",
                  "notes": "M"},
}

WIDTHS = {
    "clients": {"A": 2.2, "B": 4.5, "C": 20, "D": 14, "E": 30, "F": 12,
                "G": 13, "H": 8, "I": 26, "J": 16, "K": 11, "L": 11,
                "M": 11, "N": 15, "O": 26},
    "events": {"A": 2.2, "B": 4.5, "C": 10, "D": 21, "E": 12, "F": 13,
               "G": 8, "H": 30, "I": 9, "J": 22, "K": 11, "L": 11, "M": 11,
               "N": 9, "O": 11, "P": 11, "Q": 16, "R": 24},
    "menu": {"A": 2.2, "B": 4.5, "C": 24, "D": 12, "E": 34, "F": 12,
             "G": 12, "H": 11, "I": 11, "J": 9, "K": 24},
    "inventory": {"A": 2.2, "B": 4.5, "C": 22, "D": 14, "E": 8, "F": 10,
                  "G": 10, "H": 10, "I": 12, "J": 18, "K": 14, "L": 22},
    "shopping": {"A": 2.2, "B": 4.5, "C": 12, "D": 22, "E": 10, "F": 10,
                 "G": 11, "H": 18, "I": 11, "J": 10, "K": 12, "L": 22},
    "expenses": {"A": 2.2, "B": 4.5, "C": 12, "D": 20, "E": 16, "F": 34,
                 "G": 11, "H": 14, "I": 12, "J": 14},
    "income": {"A": 2.2, "B": 4.5, "C": 12, "D": 20, "E": 12, "F": 12,
               "G": 11, "H": 11, "I": 11, "J": 12, "K": 12, "L": 12,
               "M": 15},
    "staff": {"A": 2.2, "B": 4.5, "C": 18, "D": 13, "E": 12, "F": 8,
              "G": 10, "H": 8, "I": 12, "J": 9, "K": 22},
    "equipment": {"A": 2.2, "B": 4.5, "C": 27, "D": 14, "E": 8, "F": 10,
                  "G": 9, "H": 10, "I": 11, "J": 12, "K": 13, "L": 16},
    "suppliers": {"A": 2.2, "B": 4.5, "C": 20, "D": 16, "E": 14, "F": 29,
                  "G": 16, "H": 27, "I": 10, "J": 10, "K": 10, "L": 14,
                  "M": 22},
}

# ===========================================================================
# _Data layout
# ===========================================================================
DATA_MONTH_FIRST = 2          # H..L : month, revenue, expenses, profit, events
DATA_EXP_FIRST = 2            # N..O : expense category, spend (pie chart)
DATA_TYPE_FIRST = 16          # N..O : event type, revenue
DATA_STATUS_FIRST = 28        # N..O : event status, count (doughnut chart)
DATA_CLIENT_FIRST = 26        # Q..R : client, revenue
DATA_MENU_FIRST = 26          # T..U : menu item, times booked
DATA_POOL_FIRST = 2           # AA..AC : upcoming events pool (date,label,kind)
DATA_DUE_FIRST = 42           # AA..AC second block: outstanding payments
DATA_POOL_ROWS = 40           # events sheet has 40 rows
DATA_DUE_ROWS = 60            # income sheet has 60 rows
KPI_COL_LABEL = "AE"
KPI_COL_VALUE = "AF"

# ===========================================================================
# quote calculator geometry (block sheet)
# ===========================================================================
QUOTE_ROWS = {"client": 10, "date": 11, "guests": 12,
              "food_pp": 14, "labor": 15, "equipment": 16, "transport": 17,
              "other": 18, "margin": 20, "deposit_pct": 21}
QUOTE_OUT = {"food_total": 14, "cost_total": 15, "price": 16,
             "per_guest": 17, "profit": 18, "margin": 19, "deposit": 20,
             "balance": 21, "cost_pp": 22}
QUOTE_LAST_ROW = 26

# ===========================================================================
# tax tracker geometry
# ===========================================================================
TAX_SUM = {"sales": 9, "rate": 10, "collected": 11, "paid": 12, "due": 13}
TAX_LOG_HDR = 19
TAX_LOG_FIRST = 20
TAX_LOG_ROWS = 20
TAX_COLS = {"n": "B", "date": "C", "desc": "D", "amount": "G", "method": "H"}


KPI_ROW = {}                  # filled below
KPI_FMT = {}

_KPI_SPECS = [
    ("revenue", "Cash received (deposits + payments)", "#,##0"),
    ("expenses", "Total expenses", "#,##0"),
    ("profit", "Net profit", "#,##0"),
    ("margin", "Net profit margin", "0%"),
    ("events_total", "Events on the books", "0"),
    ("events_done", "Events completed", "0"),
    ("events_confirmed", "Events confirmed (upcoming)", "0"),
    ("events_cancelled", "Events cancelled", "0"),
    ("avg_order", "Average order value", "#,##0"),
    ("avg_guests", "Average guests per event", "0"),
    ("avg_profit", "Average profit per event", "#,##0"),
    ("guests_total", "Guests catered (completed)", "0"),
    ("outstanding", "Outstanding receivables", "#,##0"),
    ("overdue", "Overdue invoices", "0"),
    ("invoices_open", "Invoices still open", "0"),
    ("food_cost", "Ingredient & packaging spend", "#,##0"),
    ("labor_cost", "Staff / labor spend", "#,##0"),
    ("food_pct", "Food cost % of revenue", "0%"),
    ("labor_pct", "Labor cost % of revenue", "0%"),
    ("repeat_pct", "Repeat customer %", "0%"),
    ("inv_value", "Inventory value", "#,##0.00"),
    ("low_stock", "Ingredients below minimum", "0"),
    ("shop_lines", "Shopping lines open", "0"),
    ("shop_cost", "Shopping list estimated cost", "#,##0.00"),
    ("staff_unpaid", "Unpaid staff shifts", "0"),
    ("equip_value", "Equipment value", "#,##0"),
    ("equip_service", "Equipment needing service", "0"),
    ("tax_collected", "Tax collected on sales", "#,##0.00"),
    ("tax_due", "Estimated tax still due", "#,##0.00"),
    ("quote_price", "Quote calculator: recommended price", "#,##0.00"),
    ("quote_per_guest", "Quote calculator: price per guest", "#,##0.00"),
    ("menu_items", "Menu items priced", "0"),
    ("menu_margin", "Average menu margin", "0%"),
    ("best_type", "Best earning event type", "@"),
    ("best_client", "Highest value client", "@"),
]
for _i, (_key, _label, _fmt) in enumerate(_KPI_SPECS):
    KPI_ROW[_key] = 2 + _i
    KPI_FMT[_key] = _fmt

KPI_LABEL = {k: label for k, label, _ in _KPI_SPECS}

# ===========================================================================
# calendar geometry
# ===========================================================================
CAL_GRID_ROW = 12             # first row of the 6x7 grid
CAL_CELL_H = 62
MONTH_NAMES = ["January", "February", "March", "April", "May", "June",
               "July", "August", "September", "October", "November",
               "December"]
```

### `catering_tracker/theme.py`

```python
"""
Colour palettes / design systems for the Catering Business Manager.

Two themes ship with the product:

  classic   warm cream canvas, espresso brown, copper, brass gold
  fresh     cool white canvas, basil green, tomato accent, lemon gold

Everything else in the workbook is derived from the active theme.
"""


class Theme(object):
    """Small attribute bag so sheet code can read ``th.primary`` etc."""

    def __init__(self, **kw):
        self.__dict__.update(kw)

    def soft(self, key):
        """Return the pale 'tint' partner of a colour key."""
        return self.__dict__[key + "_soft"]


# ===========================================================================
# CLASSIC - "bistro ledger"
# ===========================================================================
CLASSIC = Theme(
    key="classic",
    label="Classic",
    bg="#F8F4EC",
    cover_bg="#F8F4EC",
    card="#FFFFFF",
    alt="#FBF8F1",
    border="#E3D9C6",
    border_strong="#C8B795",
    primary="#4B3626",          # espresso
    primary_2="#7A5C43",        # latte
    primary_2_soft="#EBE1D4",
    primary_soft="#EDE6DA",
    accent="#B4652F",           # copper
    accent_soft="#F7E7DB",
    gold="#A9822B",             # brass
    gold_soft="#F6EED7",
    ink="#2E241A",
    muted="#8B7C6B",
    white="#FFFFFF",
    ok="#4C7A4C",
    ok_soft="#E2EEE0",
    warn="#B4761A",
    warn_soft="#FBF0D9",
    bad="#AC2F2F",
    bad_soft="#F7E2E0",
    info="#3F6E8C",
    info_soft="#E2ECF3",
    plum="#7A5273",
    plum_soft="#F0E6EE",
    title_font="Georgia",
    body_font="Calibri",
    mono_font="Consolas",
    tabs={
        "data": "#8B7C6B",
        "setup": "#7A5C43",
        "dashboard": "#4B3626",
        "clients": "#B4652F",
        "events": "#A9822B",
        "quote": "#3F6E8C",
        "menu": "#AC2F2F",
        "inventory": "#4C7A4C",
        "shopping": "#6C9A58",
        "expenses": "#AC2F2F",
        "income": "#4C7A4C",
        "staff": "#7A5273",
        "equipment": "#3F6E8C",
        "suppliers": "#B4761A",
        "calendar": "#A9822B",
        "reports": "#4B3626",
        "tax": "#7A5273",
        "checklists": "#4C7A4C",
        "invoice": "#B4652F",
        "guide": "#8B7C6B",
    },
)

# ===========================================================================
# FRESH - "market garden"
# ===========================================================================
FRESH = Theme(
    key="fresh",
    label="Fresh",
    bg="#FAFBF8",
    cover_bg="#FAFBF8",
    card="#FFFFFF",
    alt="#F4F7F3",
    border="#DCE4DA",
    border_strong="#B9C8B4",
    primary="#2F5D46",          # basil
    primary_2="#56806A",
    primary_2_soft="#E2EDE6",
    primary_soft="#E6F0EA",
    accent="#C4482F",           # tomato
    accent_soft="#F9E4DE",
    gold="#C99A18",             # lemon
    gold_soft="#FAF1D6",
    ink="#22302A",
    muted="#7C8A82",
    white="#FFFFFF",
    ok="#3E7C4F",
    ok_soft="#E0EFE4",
    warn="#C07C12",
    warn_soft="#FBF0D8",
    bad="#B23A3A",
    bad_soft="#F8E3E1",
    info="#3E6E8E",
    info_soft="#E2EDF3",
    plum="#6C5B94",
    plum_soft="#EBE7F3",
    title_font="Calibri Light",
    body_font="Calibri",
    mono_font="Consolas",
    tabs={
        "data": "#7C8A82",
        "setup": "#56806A",
        "dashboard": "#2F5D46",
        "clients": "#C4482F",
        "events": "#C99A18",
        "quote": "#3E6E8E",
        "menu": "#B23A3A",
        "inventory": "#3E7C4F",
        "shopping": "#6C9A58",
        "expenses": "#B23A3A",
        "income": "#3E7C4F",
        "staff": "#6C5B94",
        "equipment": "#3E6E8E",
        "suppliers": "#C07C12",
        "calendar": "#C99A18",
        "reports": "#2F5D46",
        "tax": "#6C5B94",
        "checklists": "#3E7C4F",
        "invoice": "#C4482F",
        "guide": "#7C8A82",
    },
)

THEMES = {"classic": CLASSIC, "fresh": FRESH}


def get(name):
    try:
        return THEMES[name]
    except KeyError:
        raise ValueError("unknown theme %r (have %s)"
                         % (name, ", ".join(sorted(THEMES))))
```

### `catering_tracker/styles.py`

```python
"""
Every cell format used by the workbook, generated from the active theme.

Sheet builders never build formats inline - they ask ``Styles`` for a
semantic style (``S.cell("money", alt)``), which keeps the two themes
visually consistent and keeps the format cache small.
"""

import math


class Styles(object):

    def __init__(self, wb, th):
        self.wb = wb
        self.th = th
        self._cache = {}
        self._build()

    # ------------------------------------------------------------------
    # core
    # ------------------------------------------------------------------
    def f(self, props=None, **kw):
        """Cached ``workbook.add_format`` (accepts a dict or keywords)."""
        p = dict(props or {})
        p.update(kw)
        key = tuple(sorted((k, str(v)) for k, v in p.items()))
        fmt = self._cache.get(key)
        if fmt is None:
            fmt = self.wb.add_format(p)
            self._cache[key] = fmt
        return fmt

    def base(self, **over):
        """Body-text defaults + overrides."""
        p = {"font_name": self.th.body_font, "font_size": 10.5,
             "font_color": self.th.ink, "valign": "vcenter"}
        p.update(over)
        return p

    # ------------------------------------------------------------------
    # canvases, banners, sections
    # ------------------------------------------------------------------
    def _build(self):
        th = self.th
        self.canvas = self.f({"bg_color": th.bg})
        self.canvas_card = self.f({"bg_color": th.card})

        self.hero_title = self.f(self.base(
            font_name=th.title_font, font_size=26, bold=True,
            font_color=th.white, bg_color=th.primary, align="left",
            valign="vcenter", indent=1))
        self.hero_count = self.f(self.base(
            font_name=th.title_font, font_size=17, bold=True,
            font_color=th.gold_soft, bg_color=th.accent, align="left",
            valign="vcenter", indent=1))
        self.hero_meta = self.f(self.base(
            font_size=10, italic=True, font_color=th.ink,
            bg_color=th.gold_soft, align="left", valign="vcenter", indent=1))

        self.sheet_title = self.f(self.base(
            font_name=th.title_font, font_size=20, bold=True,
            font_color=th.primary, bg_color=th.bg, align="left",
            valign="vcenter", indent=1))
        self.sheet_sub = self.f(self.base(
            font_size=10.5, italic=True, font_color=th.muted,
            bg_color=th.bg, align="left", valign="vcenter", indent=1))
        self.home_link = self.f(self.base(
            font_size=10, bold=True, font_color=th.white, bg_color=th.accent,
            align="center", valign="vcenter", border=1,
            border_color=th.accent))

        self.section = self.f(self.base(
            font_name=th.title_font, font_size=13, bold=True,
            font_color=th.white, bg_color=th.primary, align="left",
            valign="vcenter", indent=1))
        self.section_accent = self.f(self.base(
            font_name=th.title_font, font_size=13, bold=True,
            font_color=th.white, bg_color=th.accent, align="left",
            valign="vcenter", indent=1))
        self.section_gold = self.f(self.base(
            font_name=th.title_font, font_size=13, bold=True,
            font_color=th.ink, bg_color=th.gold_soft, align="left",
            valign="vcenter", indent=1, bottom=2, bottom_color=th.gold))
        self.section_soft = self.f(self.base(
            font_name=th.title_font, font_size=12, bold=True,
            font_color=th.primary, bg_color=th.primary_soft, align="left",
            valign="vcenter", indent=1, left=4, left_color=th.gold))

        # table header
        self.thead = self.f(self.base(
            font_size=10, bold=True, font_color=th.white, bg_color=th.primary,
            align="center", valign="vcenter", text_wrap=True, border=1,
            border_color=th.primary))

        self.note = self.f(self.base(
            font_size=10, italic=True, font_color=th.muted, bg_color=th.alt,
            align="left", valign="top", text_wrap=True, border=1,
            border_color=th.border, indent=1, locked=False))
        self.note_plain = self.f(self.base(
            font_size=10, italic=True, font_color=th.muted, bg_color=th.bg,
            align="left", valign="top", text_wrap=True))
        self.footer = self.f(self.base(
            font_size=9, italic=True, font_color=th.muted, bg_color=th.bg,
            align="left", valign="vcenter"))

        self.bar_text = self.f(self.base(
            font_name=th.mono_font, font_size=13, bold=True,
            font_color=th.primary_2, bg_color=th.card, align="left",
            valign="vcenter", border=1, border_color=th.border))
        self.bar_label = self.f(self.base(
            font_size=10, bold=True, font_color=th.ink, bg_color=th.card,
            align="left", valign="vcenter", border=1,
            border_color=th.border, indent=1))
        self.bar_pct = self.f(self.base(
            font_name=th.mono_font, font_size=13, bold=True,
            font_color=th.accent, bg_color=th.card, align="center",
            valign="vcenter", border=1, border_color=th.border))

    # ------------------------------------------------------------------
    # table headers in an accent colour
    # ------------------------------------------------------------------
    def header(self, color=None):
        color = color or self.th.primary
        return self.f(self.base(
            font_size=10, bold=True, font_color=self.th.white,
            bg_color=color, align="center", valign="vcenter", text_wrap=True,
            border=1, border_color=color))

    # ------------------------------------------------------------------
    # data cells
    # ------------------------------------------------------------------
    _KINDS = {
        "text":       {"align": "left", "locked": False},
        "center":     {"align": "center", "locked": False},
        "money":      {"align": "right", "num_format": "#,##0.00",
                       "locked": False},
        "money0":     {"align": "right", "num_format": "#,##0",
                       "locked": False},
        "num":        {"align": "right", "num_format": "#,##0",
                       "locked": False},
        "qty":        {"align": "center", "num_format": "#,##0",
                       "locked": False},
        "qty1":       {"align": "center", "num_format": "#,##0.0",
                       "locked": False},
        "pct":        {"align": "center", "num_format": "0%",
                       "locked": False},
        "date":       {"align": "center", "num_format": "dd mmm yyyy",
                       "locked": False},
        "tick":       {"align": "center", "font_size": 13, "bold": True,
                       "font_color": "#2F7D4F", "locked": False},
        "wrap":       {"align": "left", "text_wrap": True, "valign": "top",
                       "locked": False},
        "link":       {"align": "left", "font_color": "#2A5D8F",
                       "underline": 1, "locked": False},
        "url":        {"align": "center", "bold": True, "font_size": 10},
        # calculated columns (locked - they must not be typed over)
        "calc":       {"align": "left"},
        "calc_c":     {"align": "center"},
        "calc_wrap":  {"align": "left", "text_wrap": True},
        "calc_money": {"align": "right", "num_format": "#,##0.00"},
        "calc_num":   {"align": "right", "num_format": "#,##0"},
        "calc_qty1":  {"align": "center", "num_format": "#,##0.0"},
        "calc_pct":   {"align": "center", "num_format": "0%"},
        "calc_date":  {"align": "center", "num_format": "dd mmm yyyy"},
        "calc_tick":  {"align": "center", "font_size": 13, "bold": True},
    }

    def cell(self, kind, alt=False, **over):
        """Data-cell format.  ``alt`` selects the zebra-stripe background."""
        th = self.th
        props = self.base(**self._KINDS[kind])
        if kind.startswith("calc"):
            props["bg_color"] = th.alt
            props["font_color"] = th.ink if kind in ("calc", "calc_wrap",
                                                     "calc_c") else th.ink
            props["locked"] = True
            if kind in ("calc_money", "calc_num", "calc_pct", "calc_date"):
                props["font_color"] = th.primary
        else:
            props["bg_color"] = th.alt if alt else th.card
        props["border"] = 1
        props["border_color"] = th.border
        props.update(over)
        return self.f(**props)

    def idx(self, alt=False):
        """The little grey row-number column."""
        return self.f(self.base(
            font_size=9, font_color=self.th.muted, align="center",
            valign="vcenter", border=1, border_color=self.th.border,
            bg_color=self.th.alt if alt else self.th.card,
            num_format="0"))

    # ------------------------------------------------------------------
    # dashboard furniture
    # ------------------------------------------------------------------
    def kpi_label(self, color, size=9.5, align="left"):
        return self.f(self.base(
            font_size=size, bold=True, font_color=self.th.white,
            bg_color=color, align=align, valign="vcenter", indent=1,
            border=1, border_color=color))

    def kpi_value(self, color, num_format=None, size=21, align="center",
                  bg=None):
        p = self.base(font_name=self.th.title_font, font_size=size, bold=True,
                      font_color=color, bg_color=bg or self.th.card,
                      align=align, valign="vcenter", border=1,
                      border_color=self.th.border)
        if num_format:
            p["num_format"] = num_format
        return self.f(**p)

    def kpi_text(self, color, size=11, align="left", bg=None, bold=True,
                 wrap=False):
        return self.f(self.base(
            font_size=size, bold=bold, font_color=color,
            bg_color=bg or self.th.card, align=align, valign="vcenter",
            border=1, border_color=self.th.border, indent=1 if align == "left" else 0,
            text_wrap=wrap))

    def panel(self, color=None, bg=None, size=10.5, bold=False, align="left",
              wrap=True, font_color=None, valign="top", border=True):
        p = self.base(font_size=size, bold=bold, align=align, valign=valign,
                      text_wrap=wrap,
                      bg_color=bg if bg is not None else self.th.card,
                      font_color=font_color or self.th.ink)
        if border:
            p["border"] = 1
            p["border_color"] = color or self.th.border
            if color:
                p["left"] = 3
                p["left_color"] = color
        return self.f(**p)

    def pill(self, bg, fg, size=9.5, bold=True, align="center"):
        return self.f(self.base(font_size=size, bold=bold, font_color=fg,
                                bg_color=bg, align=align, valign="vcenter",
                                border=1, border_color=bg, text_wrap=False))

    def nav(self, color, size=10.5):
        return self.f(self.base(font_size=size, bold=True,
                                font_color=self.th.white, bg_color=color,
                                align="center", valign="vcenter", border=1,
                                border_color=color))

    def nav_soft(self, size=10.5):
        th = self.th
        return self.f(self.base(font_size=size, bold=True,
                                font_color=th.primary, bg_color=th.primary_soft,
                                align="center", valign="vcenter", border=1,
                                border_color=th.border))

    # ------------------------------------------------------------------
    # conditional-format formats (dxf sources)
    # ------------------------------------------------------------------
    def cf(self, bg=None, fg=None, bold=False, strike=False, italic=False,
           border=None, size=None, align=None, num_format=None):
        th = self.th
        p = {}
        if bg:
            p["bg_color"] = bg
        if fg:
            p["font_color"] = fg
        if bold:
            p["bold"] = True
        if strike:
            p["font_strikeout"] = True
        if italic:
            p["italic"] = True
        if size:
            p["font_size"] = size
        if align:
            p["align"] = align
        if num_format:
            p["num_format"] = num_format
        if border:
            p["border"] = 1
            p["border_color"] = border
        p["font_name"] = th.body_font
        return self.f(**p)

    def cf_state(self, state):
        """Semantic state colours: ok / warn / bad / info / plum / gold."""
        th = self.th
        return self.cf(bg=th.soft(state), fg=getattr(th, state), bold=True,
                       border=th.border)

    # ------------------------------------------------------------------
    # helpers
    # ------------------------------------------------------------------
    def guide_text(self, size=10.5, bold=False, italic=False, color=None,
                   bg=None, indent=0, align="left"):
        return self.f(self.base(font_size=size, bold=bold, italic=italic,
                                font_color=color or self.th.ink,
                                bg_color=bg or self.th.card, align=align,
                                valign="top", text_wrap=True, indent=indent,
                                border=1, border_color=self.th.border,
                                locked=False))

    def static_height(self, text, width_chars, line_px=14.6, pad=8,
                      minimum=18):
        """Row height needed to show wrapped ``text`` in ``width_chars``."""
        if not text:
            return minimum
        per_line = max(10, int(width_chars * 0.95))
        lines = 0
        for para in str(text).split("\n"):
            lines += max(1, int(math.ceil(len(para) / float(per_line))))
        return max(minimum, lines * line_px + pad)


def wrap_height(text, width_chars, line_px=14.6, pad=8, minimum=18):
    """Module-level twin of :meth:`Styles.static_height`."""
    if not text:
        return minimum
    per_line = max(10, int(width_chars * 0.95))
    lines = 0
    for para in str(text).split("\n"):
        lines += max(1, int(math.ceil(len(para) / float(per_line))))
    return max(minimum, lines * line_px + pad)
```

### `catering_tracker/demo.py`

```python
"""
Sample data for EXAMPLE builds.

``Model`` mirrors what a real user would type into the workbook for a busy
catering company (Saffron & Sage Catering Co.) in September-December 2026.
Every derived figure (profit, margin, stock value, KPI totals, the monthly
report pools...) is computed here in Python so the EXAMPLE workbooks ship
with correct *cached* values: they look right the instant they are opened,
before Excel recalculates anything.
"""

from datetime import date

from . import config as C

SETTINGS = {
    "business": "Saffron & Sage Catering Co.",
    "currency": "$",
    "tax": 0.05,
    "margin": 0.35,
    "deposit": 0.40,
    "duesoon": 7,
    "year": 2026,
    "cal_month": 9,
    "cal_year": 2026,
    "message": "Confirm October crews early!",
}

TODAY = date(2026, 9, 12)          # "as if opened on" date for the sample

MONTH_SHORT = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def _d(m, day):
    return date(2026, m, day)


# ---------------------------------------------------------------------------
# events  (id, client, date, type, guests, menu, staff_req, equipment_req,
#          cost, price, deposit, status, notes)
# ---------------------------------------------------------------------------
EVENTS = [
    dict(id="EV-2026-001", client="Aisha Malik", date=_d(9, 5),
         type="Wedding", guests=180,
         menu="Herb Roast Chicken Platter, Chocolate Fondant",
         staff_req=12, equipment_req="Chafers x8, Tables x20, Linen x20",
         cost=4920, price=7600, deposit=3040, status=C.ES_DONE,
         notes="Nikkah + reception, two seatings"),
    dict(id="EV-2026-002", client="TechNova Pvt Ltd", date=_d(9, 10),
         type="Corporate", guests=60,
         menu="Chicken Tikka Platter, Lemon Tart",
         staff_req=5, equipment_req="Chafers x4, Coffee urns x2",
         cost=1180, price=1850, deposit=740, status=C.ES_DONE,
         notes="Quarterly all-hands lunch"),
    dict(id="EV-2026-003", client="Daniel Fernandes", date=_d(9, 19),
         type="Birthday", guests=45,
         menu="Beef Biryani Feast, Chocolate Fondant",
         staff_req=4, equipment_req="Chafers x3, Beverage dispenser x1",
         cost=890, price=1400, deposit=560, status=C.ES_CONFIRMED,
         notes="40th birthday, garden venue"),
    dict(id="EV-2026-004", client="Meera Khan", date=_d(9, 26),
         type="Private Party", guests=80,
         menu="Lamb Kofta Curry, Seasonal Fruit Display",
         staff_req=6, equipment_req="Chafers x5, Tables x10",
         cost=1620, price=2500, deposit=1000, status=C.ES_CONFIRMED,
         notes="Engagement party - deposit chasing final payment"),
    dict(id="EV-2026-005", client="Zenith Bank", date=_d(10, 3),
         type="Conference", guests=120,
         menu="Grilled Salmon Plate, Lemon Tart",
         staff_req=9, equipment_req="Chafers x6, Coffee urns x3, Glassware x120",
         cost=2450, price=3900, deposit=1560, status=C.ES_DEPOSIT,
         notes="Annual leadership summit, breakfast + lunch"),
    dict(id="EV-2026-006", client="Aisha Malik", date=_d(10, 10),
         type="Cocktail", guests=90,
         menu="Chicken Canapes (4 pc), Grilled Salmon Plate",
         staff_req=7, equipment_req="Glassware x90, Dispensers x2",
         cost=1980, price=3100, deposit=1240, status=C.ES_CONFIRMED,
         notes="Repeat client - mehndi cocktail hour"),
    dict(id="EV-2026-007", client="Omar Sheikh", date=_d(10, 17),
         type="Wedding", guests=220,
         menu="Beef Biryani Feast, Herb Roast Chicken Platter",
         staff_req=14, equipment_req="Chafers x10, Tables x25, Chairs x220",
         cost=6100, price=9800, deposit=3920, status=C.ES_DEPOSIT,
         notes="Biggest booking of the season"),
    dict(id="EV-2026-008", client="Bright Futures School", date=_d(10, 24),
         type="Festival", guests=300,
         menu="Chicken Tikka Platter, Seasonal Fruit Display",
         staff_req=12, equipment_req="Chafers x8, Beverage dispensers x4",
         cost=3400, price=5200, deposit=0, status=C.ES_QUOTED,
         notes="Autumn fair - quote sent, awaiting board approval"),
    dict(id="EV-2026-009", client="Hina Raza", date=_d(11, 7),
         type="Birthday", guests=35,
         menu="Creamy Pasta Alfredo, Chocolate Fondant",
         staff_req=3, equipment_req="Chafers x2",
         cost=640, price=1050, deposit=0, status=C.ES_QUOTED,
         notes="Quote sent Sep 9 - follow up next week"),
    dict(id="EV-2026-010", client="TechNova Pvt Ltd", date=_d(11, 14),
         type="Corporate", guests=75,
         menu="Grilled Salmon Plate, Lemon Tart",
         staff_req=6, equipment_req="Chafers x4, Coffee urns x2",
         cost=1520, price=2400, deposit=960, status=C.ES_CONFIRMED,
         notes="Annual client appreciation dinner"),
    dict(id="EV-2026-011", client="Gulberg Sports Club", date=_d(11, 21),
         type="Buffet", guests=150,
         menu="Beef Biryani Feast, Lamb Kofta Curry",
         staff_req=10, equipment_req="Chafers x8, Tables x15",
         cost=3050, price=4700, deposit=1880, status=C.ES_DEPOSIT,
         notes="Awards night buffet"),
    dict(id="EV-2026-012", client="Zara Ahmed", date=_d(12, 5),
         type="Wedding", guests=200,
         menu="Herb Roast Chicken Platter, Grilled Salmon Plate",
         staff_req=13, equipment_req="Chafers x10, Tables x22, Linen x22",
         cost=5600, price=8900, deposit=3560, status=C.ES_CONFIRMED,
         notes="Winter wedding - confirm salmon supplier early"),
    dict(id="EV-2026-013", client="Liberty Mall Expo", date=_d(12, 19),
         type="Festival", guests=400,
         menu="Chicken Tikka Platter, Beef Biryani Feast",
         staff_req=16, equipment_req="Chafers x12, Stall setup x3",
         cost=4800, price=7200, deposit=0, status=C.ES_INQUIRY,
         notes="Food expo stall - first contact by phone"),
    dict(id="EV-2026-014", client="Bilal Traders", date=_d(12, 28),
         type="Corporate", guests=50,
         menu="Creamy Pasta Alfredo, Lemon Tart",
         staff_req=4, equipment_req="Chafers x3",
         cost=980, price=1500, deposit=300, status=C.ES_CANCEL,
         notes="Client postponed to next year - deposit refunded"),
]

ACTIVE = (C.ES_DEPOSIT, C.ES_CONFIRMED, C.ES_DONE)
BOOKED = (C.ES_QUOTED, C.ES_DEPOSIT, C.ES_CONFIRMED, C.ES_DONE)


# ---------------------------------------------------------------------------
# clients (12 - mirror of their next / most relevant event)
# ---------------------------------------------------------------------------
def client_rows():
    """Derive the client tracker from the event book (repeat clients show
    their *next* event)."""
    by_name = {}
    for ev in EVENTS:
        prev = by_name.get(ev["client"])
        if prev is None:
            by_name[ev["client"]] = ev
        else:
            # keep the most recent *upcoming* booking for repeat clients
            if ev["date"] >= TODAY and ev["status"] != C.ES_CANCEL:
                by_name[ev["client"]] = ev
    phones = {
        "Aisha Malik": "0300-8899112", "TechNova Pvt Ltd": "0311-5544332",
        "Daniel Fernandes": "0345-2233445", "Meera Khan": "0301-7788990",
        "Zenith Bank": "0321-9090808", "Omar Sheikh": "0308-4455667",
        "Bright Futures School": "0302-3344556", "Hina Raza": "0333-1122334",
        "Gulberg Sports Club": "0344-6677889", "Zara Ahmed": "0306-9988776",
        "Liberty Mall Expo": "0303-1212334", "Bilal Traders": "0315-5566778",
    }
    emails = {
        "Aisha Malik": "aisha.malik@mail.example",
        "TechNova Pvt Ltd": "events@technova.example",
        "Daniel Fernandes": "dan.fernandes@mail.example",
        "Meera Khan": "meera.khan@mail.example",
        "Zenith Bank": "hr.summits@zenithbank.example",
        "Omar Sheikh": "omar.sheikh@mail.example",
        "Bright Futures School": "admin@brightfutures.example",
        "Hina Raza": "hina.raza@mail.example",
        "Gulberg Sports Club": "clubhouse@gulbergtown.example",
        "Zara Ahmed": "zara.ahmed@mail.example",
        "Liberty Mall Expo": "expo@libertymall.example",
        "Bilal Traders": "bilal.traders@mail.example",
    }
    packages = {
        "Aisha Malik": "Gold", "TechNova Pvt Ltd": "Silver",
        "Daniel Fernandes": "Custom", "Meera Khan": "Silver",
        "Zenith Bank": "Gold", "Omar Sheikh": "Gold",
        "Bright Futures School": "Bronze", "Hina Raza": "Custom",
        "Gulberg Sports Club": "Silver", "Zara Ahmed": "Gold",
        "Liberty Mall Expo": "Custom", "Bilal Traders": "Bronze",
    }
    venues = {
        "EV-2026-001": "Royal Marquee, Model Town",
        "EV-2026-002": "TechNova HQ, Arfa Tower",
        "EV-2026-003": "Private home, Gulberg III",
        "EV-2026-004": "Lawn venue, DHA Phase 5",
        "EV-2026-005": "Conference centre, Gulberg",
        "EV-2026-006": "Marquee, Ferozepur Road",
        "EV-2026-007": "Grand Marquee, Raiwind Road",
        "EV-2026-008": "School ground, Johar Town",
        "EV-2026-009": "Private home, Cantonment",
        "EV-2026-010": "TechNova HQ, Arfa Tower",
        "EV-2026-011": "Clubhouse, Gulberg",
        "EV-2026-012": "Winter Marquee, Bedian Road",
        "EV-2026-013": "Liberty Market expo hall",
        "EV-2026-014": "Bilal Traders office",
    }
    notes = {
        "Aisha Malik": "VIP - 2nd booking this season, loves the chicken platter",
        "TechNova Pvt Ltd": "Invoice to accounts@technova.example",
        "Omar Sheikh": "Referred by Aisha Malik",
        "Zenith Bank": "Needs halal certification letter",
        "Bright Futures School": "School board must approve quote",
        "Liberty Mall Expo": "Big volume, thin margin - negotiate staffing",
    }
    rows = []
    for ev in EVENTS:                      # first-seen order == event order
        name = ev["client"]
        if any(r["name"] == name for r in rows):
            continue
        latest = by_name[name]
        paid = (latest["deposit"] if latest["status"] in
                (C.ES_DEPOSIT, C.ES_CONFIRMED, C.ES_DONE) else 0)
        if latest["status"] == C.ES_DONE:
            status = C.PS_PAID
        elif paid:
            status = C.PS_PART
        else:
            status = C.PS_UNPAID
        rows.append(dict(
            name=name, phone=phones[name], email=emails[name],
            event_date=latest["date"], event_type=latest["type"],
            guests=latest["guests"], venue=venues[latest["id"]],
            package=packages[name], quote=latest["price"], deposit=paid,
            status=status, notes=notes.get(name, "")))
    return rows


CLIENTS = client_rows()


# ---------------------------------------------------------------------------
# menu items  (cost / price are per portion)
# ---------------------------------------------------------------------------
MENU = [
    dict(item="Herb Roast Chicken Platter", category="Main",
         ingredients="Chicken breast, fresh herbs, butter, olive oil",
         portion="1 plate", cost=4.85, price=12.50,
         notes="Signature wedding main"),
    dict(item="Beef Biryani Feast", category="Main",
         ingredients="Beef mince, basmati rice, mixed spice, onions",
         portion="1 plate", cost=3.90, price=9.75, notes="Crowd favourite"),
    dict(item="Chicken Tikka Platter", category="Main",
         ingredients="Chicken breast, yogurt, tikka masala, herbs",
         portion="1 plate", cost=4.40, price=11.50, notes=""),
    dict(item="Grilled Salmon Plate", category="Main",
         ingredients="Salmon fillet, olive oil, lemon, fresh herbs",
         portion="1 plate", cost=8.20, price=18.50, notes="Premium option"),
    dict(item="Creamy Pasta Alfredo", category="Main",
         ingredients="Pasta, heavy cream, cheddar, butter",
         portion="1 plate", cost=2.60, price=7.90, notes="Kids love it"),
    dict(item="Lamb Kofta Curry", category="Main",
         ingredients="Beef mince, onions, tomatoes, whole spices",
         portion="1 bowl", cost=4.10, price=10.50, notes=""),
    dict(item="Garden Salad Bar", category="Sides",
         ingredients="Tomatoes, onions, leafy greens, olive oil",
         portion="1 bowl", cost=1.35, price=4.25, notes="Vegan"),
    dict(item="Garlic Bread Basket", category="Sides",
         ingredients="Flour, butter, garlic, parsley",
         portion="1 basket", cost=0.85, price=2.75, notes=""),
    dict(item="Roast Vegetable Tray", category="Sides",
         ingredients="Seasonal vegetables, olive oil, herbs",
         portion="1 tray", cost=1.20, price=3.90, notes="Vegan"),
    dict(item="Chocolate Fondant", category="Dessert",
         ingredients="Chocolate, butter, eggs, sugar, flour",
         portion="1 piece", cost=1.85, price=5.50, notes="Made fresh on site"),
    dict(item="Lemon Tart", category="Dessert",
         ingredients="Flour, butter, eggs, sugar, lemons",
         portion="1 slice", cost=1.55, price=4.75, notes=""),
    dict(item="Seasonal Fruit Display", category="Dessert",
         ingredients="Fresh seasonal fruit, mint",
         portion="1 plate", cost=1.10, price=3.25, notes=""),
    dict(item="Iced Tea & Lemonade Station", category="Beverage",
         ingredients="Tea, lemons, sugar, mint",
         portion="1 glass", cost=0.45, price=1.90, notes="Self-serve station"),
    dict(item="Chicken Canapes (4 pc)", category="Canape",
         ingredients="Chicken, bread, herbs, butter",
         portion="4 pieces", cost=1.25, price=3.60, notes="Cocktail hour"),
]


# ---------------------------------------------------------------------------
# ingredient inventory
# ---------------------------------------------------------------------------
INVENTORY = [
    dict(ingredient="Basmati Rice", category="Pantry", unit="kg", qty=42,
         min=20, unit_cost=3.80, supplier="Metro Wholesale", notes="Super kernel"),
    dict(ingredient="Chicken Breast", category="Meat & Fish", unit="kg",
         qty=15, min=25, unit_cost=12.00, supplier="City Foods",
         notes="Halal, boneless"),
    dict(ingredient="Beef Mince", category="Meat & Fish", unit="kg", qty=18,
         min=10, unit_cost=9.50, supplier="City Foods", notes=""),
    dict(ingredient="Salmon Fillet", category="Meat & Fish", unit="kg", qty=6,
         min=8, unit_cost=22.00, supplier="City Foods",
         notes="Frozen, order 1 week ahead"),
    dict(ingredient="Olive Oil", category="Pantry", unit="L", qty=9.2, min=6,
         unit_cost=8.20, supplier="Metro Wholesale", notes=""),
    dict(ingredient="Butter", category="Dairy", unit="kg", qty=7, min=5,
         unit_cost=6.40, supplier="Dairy Land", notes="Unsalted"),
    dict(ingredient="Heavy Cream", category="Dairy", unit="L", qty=14, min=8,
         unit_cost=4.10, supplier="Dairy Land", notes=""),
    dict(ingredient="Cheddar Block", category="Dairy", unit="kg", qty=11,
         min=6, unit_cost=7.80, supplier="Dairy Land", notes=""),
    dict(ingredient="Eggs", category="Dairy", unit="tray", qty=24, min=15,
         unit_cost=3.20, supplier="Dairy Land", notes="30 per tray"),
    dict(ingredient="Flour", category="Pantry", unit="kg", qty=30, min=15,
         unit_cost=1.10, supplier="Metro Wholesale", notes=""),
    dict(ingredient="Sugar", category="Pantry", unit="kg", qty=12, min=10,
         unit_cost=1.20, supplier="Metro Wholesale", notes=""),
    dict(ingredient="Mixed Spice", category="Pantry", unit="kg", qty=2.8,
         min=2, unit_cost=14.00, supplier="Spice World", notes="House blend"),
    dict(ingredient="Pasta", category="Pantry", unit="kg", qty=16, min=8,
         unit_cost=2.30, supplier="Metro Wholesale", notes="Penne"),
    dict(ingredient="Tomatoes", category="Produce", unit="kg", qty=8, min=12,
         unit_cost=1.90, supplier="Fresh Mart", notes=""),
    dict(ingredient="Onions", category="Produce", unit="kg", qty=22, min=12,
         unit_cost=0.95, supplier="Fresh Mart", notes=""),
    dict(ingredient="Fresh Herbs", category="Produce", unit="kg", qty=1.8,
         min=2, unit_cost=9.00, supplier="Fresh Mart",
         notes="Coriander, mint, parsley"),
]

INV_QTY = {row["ingredient"]: row["qty"] for row in INVENTORY}
INV_COST = {row["ingredient"]: row["unit_cost"] for row in INVENTORY}


# ---------------------------------------------------------------------------
# shopping list (drives off the next three events)
# ---------------------------------------------------------------------------
SHOPPING = [
    dict(event="EV-2026-003", ingredient="Chicken Breast", required=12,
         supplier="City Foods", purchased=C.TICK, date=_d(9, 8),
         notes="For birani + tikka"),
    dict(event="EV-2026-003", ingredient="Mixed Spice", required=2,
         supplier="Spice World", purchased=C.TICK, date=_d(9, 8), notes=""),
    dict(event="EV-2026-003", ingredient="Fresh Herbs", required=1.5,
         supplier="Fresh Mart", purchased="", date=None, notes="Buy 2 days before"),
    dict(event="EV-2026-004", ingredient="Beef Mince", required=10,
         supplier="City Foods", purchased="", date=None, notes=""),
    dict(event="EV-2026-004", ingredient="Tomatoes", required=8,
         supplier="Fresh Mart", purchased="", date=None, notes=""),
    dict(event="EV-2026-004", ingredient="Heavy Cream", required=6,
         supplier="Dairy Land", purchased=C.TICK, date=_d(9, 11), notes=""),
    dict(event="EV-2026-005", ingredient="Basmati Rice", required=15,
         supplier="Metro Wholesale", purchased="", date=None, notes=""),
    dict(event="EV-2026-005", ingredient="Chicken Breast", required=30,
         supplier="City Foods", purchased="", date=None,
         notes="Order by Sep 20"),
    dict(event="EV-2026-005", ingredient="Salmon Fillet", required=10,
         supplier="City Foods", purchased="", date=None,
         notes="Frozen - 1 week lead time"),
    dict(event="EV-2026-005", ingredient="Butter", required=4,
         supplier="Dairy Land", purchased="", date=None, notes=""),
]


# ---------------------------------------------------------------------------
# expenses  (actual money out, Aug 20 - Sep 12)
# ---------------------------------------------------------------------------
EXPENSES = [
    dict(date=_d(8, 20), vendor="Metro Wholesale", category="Ingredients",
         desc="Chicken breasts - 40 kg (EV-001)", amount=480.00,
         method="Bank Transfer", event="EV-2026-001", receipt=C.TICK),
    dict(date=_d(8, 22), vendor="Metro Wholesale", category="Ingredients",
         desc="Basmati rice - 25 kg sack", amount=95.00, method="Cash",
         event="EV-2026-001", receipt=C.TICK),
    dict(date=_d(8, 24), vendor="Fresh Mart", category="Ingredients",
         desc="Salad produce & vegetables", amount=380.00, method="Card",
         event="EV-2026-001", receipt=C.TICK),
    dict(date=_d(8, 26), vendor="Meta Ads", category="Marketing",
         desc="Instagram & Facebook campaign", amount=220.00, method="Card",
         event="", receipt=C.TICK),
    dict(date=_d(8, 28), vendor="Dairy Land", category="Ingredients",
         desc="Cream, butter & cheese restock", amount=165.00, method="Card",
         event="EV-2026-001", receipt=C.TICK),
    dict(date=_d(8, 30), vendor="ChefGear", category="Equipment",
         desc="Chafing fuel & spare pans", amount=120.00, method="Card",
         event="", receipt=C.TICK),
    dict(date=_d(8, 21), vendor="SecureNow", category="Insurance",
         desc="Public liability - quarterly", amount=450.00,
         method="Bank Transfer", event="", receipt=C.TICK),
    dict(date=_d(9, 1), vendor="Cloud Kitchen Co", category="Kitchen Rental",
         desc="September production kitchen", amount=900.00,
         method="Bank Transfer", event="", receipt=C.TICK),
    dict(date=_d(9, 1), vendor="Metro Wholesale", category="Ingredients",
         desc="Corporate lunch ingredients", amount=240.00,
         method="Bank Transfer", event="EV-2026-002", receipt=C.TICK),
    dict(date=_d(9, 2), vendor="City Utilities", category="Utilities",
         desc="Electricity - August", amount=180.00, method="Bank Transfer",
         event="", receipt=C.TICK),
    dict(date=_d(9, 3), vendor="Ledgerly", category="Software",
         desc="Accounting software - annual", amount=240.00, method="Card",
         event="", receipt=C.TICK),
    dict(date=_d(9, 4), vendor="PrintHub", category="Marketing",
         desc="Menus, flyers & business cards", amount=110.00, method="Card",
         event="", receipt=C.TICK),
    dict(date=_d(9, 5), vendor="Fuel Stop", category="Transportation",
         desc="Delivery van fuel - EV-001 week", amount=90.00, method="Card",
         event="EV-2026-001", receipt=C.TICK),
    dict(date=_d(9, 6), vendor="PackRight", category="Packaging",
         desc="Boxes & cutlery packs (EV-002)", amount=95.00, method="Card",
         event="EV-2026-002", receipt=C.TICK),
    dict(date=_d(9, 6), vendor="Payroll", category="Staff / Labor",
         desc="Event crew wages - EV-001", amount=1456.00,
         method="Bank Transfer", event="EV-2026-001", receipt=C.TICK),
    dict(date=_d(9, 7), vendor="General Store", category="Miscellaneous",
         desc="Cleaning supplies", amount=65.00, method="Cash", event="",
         receipt=C.TICK),
    dict(date=_d(9, 8), vendor="City Utilities", category="Utilities",
         desc="Gas - August", amount=95.00, method="Bank Transfer", event="",
         receipt=C.TICK),
    dict(date=_d(9, 8), vendor="Fresh Mart", category="Ingredients",
         desc="Ingredients - September parties", amount=950.00, method="Card",
         event="EV-2026-003", receipt=""),
    dict(date=_d(9, 9), vendor="ChefGear", category="Equipment",
         desc="Replacement serving utensils", amount=85.00, method="Card",
         event="", receipt=""),
    dict(date=_d(9, 10), vendor="MoveIt Rentals", category="Transportation",
         desc="Refrigerated van rental", amount=150.00, method="Card",
         event="EV-2026-003", receipt=""),
    dict(date=_d(9, 10), vendor="Metro Wholesale", category="Ingredients",
         desc="Rice & flour bulk top-up", amount=260.00,
         method="Bank Transfer", event="EV-2026-004", receipt=""),
    dict(date=_d(9, 11), vendor="PackRight", category="Packaging",
         desc="Platters, foil & serving trays", amount=180.00, method="Card",
         event="", receipt=""),
    dict(date=_d(9, 11), vendor="City Foods", category="Ingredients",
         desc="Meat & poultry bulk - Oct weddings", amount=1490.00,
         method="Bank Transfer", event="EV-2026-007", receipt=""),
    dict(date=_d(9, 11), vendor="Payroll", category="Staff / Labor",
         desc="Chef & servers - EV-002", amount=582.00,
         method="Bank Transfer", event="EV-2026-002", receipt=C.TICK),
    dict(date=_d(9, 12), vendor="Payroll", category="Staff / Labor",
         desc="Kitchen prep team - week 37", amount=560.00,
         method="Bank Transfer", event="", receipt=""),
    dict(date=_d(9, 12), vendor="City Foods", category="Ingredients",
         desc="Seafood & premium produce", amount=640.00,
         method="Bank Transfer", event="EV-2026-006", receipt=""),
]


# ---------------------------------------------------------------------------
# income / invoices  (one per booked event, cancelled excluded)
# ---------------------------------------------------------------------------
def income_rows():
    rows, n = [], 1000
    for ev in EVENTS:
        if ev["status"] == C.ES_CANCEL:
            continue
        n += 1
        price, dep = ev["price"], ev["deposit"]
        paid_full = ev["status"] == C.ES_DONE
        pay2 = price - dep if paid_full else 0
        # one deliberately late balance so the OVERDUE flag has something to do
        due = _d(9, 10) if ev["id"] == "EV-2026-004" else ev["date"]
        rows.append(dict(
            invoice="INV-%d" % n, client=ev["client"], event_date=ev["date"],
            amount=price, deposit=dep, pay1=0, pay2=pay2,
            due=due, event_id=ev["id"]))
    return rows


INCOME = income_rows()


# ---------------------------------------------------------------------------
# staff shifts
# ---------------------------------------------------------------------------
STAFF = [
    dict(name="Imran Qureshi", role="Head Chef", event="EV-2026-001", hours=14,
         rate=25, ot=2, paid=C.TICK, notes="Ran both seatings"),
    dict(name="Fatima Noor", role="Sous Chef", event="EV-2026-001", hours=12,
         rate=18, ot=0, paid=C.TICK, notes=""),
    dict(name="Ali Hassan", role="Chef", event="EV-2026-001", hours=12,
         rate=16, ot=0, paid=C.TICK, notes=""),
    dict(name="Zoya Sheikh", role="Server", event="EV-2026-001", hours=10,
         rate=11, ot=2, paid=C.TICK, notes=""),
    dict(name="Hassan Raza", role="Server", event="EV-2026-001", hours=10,
         rate=11, ot=0, paid=C.TICK, notes=""),
    dict(name="Maryam Javed", role="Bartender", event="EV-2026-001", hours=10,
         rate=13, ot=0, paid=C.TICK, notes=""),
    dict(name="Kashif Ali", role="Driver", event="EV-2026-001", hours=8,
         rate=12, ot=0, paid=C.TICK, notes="Two delivery runs"),
    dict(name="Nadia Aslam", role="Setup Crew", event="EV-2026-001", hours=9,
         rate=10, ot=0, paid=C.TICK, notes=""),
    dict(name="Usman Ghani", role="Cleaner", event="EV-2026-001", hours=6,
         rate=9, ot=0, paid=C.TICK, notes=""),
    dict(name="Imran Qureshi", role="Head Chef", event="EV-2026-002", hours=8,
         rate=25, ot=0, paid=C.TICK, notes=""),
    dict(name="Fatima Noor", role="Sous Chef", event="EV-2026-002", hours=8,
         rate=18, ot=0, paid=C.TICK, notes=""),
    dict(name="Ali Hassan", role="Chef", event="EV-2026-002", hours=7,
         rate=16, ot=0, paid=C.TICK, notes=""),
    dict(name="Zoya Sheikh", role="Server", event="EV-2026-002", hours=6,
         rate=11, ot=0, paid=C.TICK, notes=""),
    dict(name="Kashif Ali", role="Driver", event="EV-2026-002", hours=5,
         rate=12, ot=0, paid=C.TICK, notes=""),
    dict(name="Imran Qureshi", role="Head Chef", event="EV-2026-003", hours=10,
         rate=25, ot=0, paid="", notes="Scheduled"),
    dict(name="Sara Butt", role="Server", event="EV-2026-003", hours=8,
         rate=11, ot=0, paid="", notes="Scheduled"),
    dict(name="Hassan Raza", role="Server", event="EV-2026-004", hours=9,
         rate=11, ot=0, paid="", notes="Scheduled"),
    dict(name="Bilal Ahmad", role="Setup Crew", event="EV-2026-005", hours=8,
         rate=10, ot=0, paid="", notes="Scheduled"),
]


# ---------------------------------------------------------------------------
# equipment
# ---------------------------------------------------------------------------
EQUIPMENT = [
    dict(item="Chafing Dishes (full size)", category="Serving", owned=24,
         reserved=8, damaged=1, unit_value=45, maintenance=_d(10, 1),
         notes="1 dented lid"),
    dict(item="Round Banquet Tables", category="Furniture", owned=20,
         reserved=0, damaged=0, unit_value=60, maintenance=_d(12, 1),
         notes="Seats 10 each"),
    dict(item="Rectangular Tables", category="Furniture", owned=15,
         reserved=6, damaged=0, unit_value=45, maintenance=date(2027, 1, 15),
         notes=""),
    dict(item="Banquet Chairs", category="Furniture", owned=200, reserved=80,
         damaged=4, unit_value=12, maintenance=date(2026, 11, 1),
         notes="4 frames bent"),
    dict(item="Table Linen Sets", category="Linens", owned=40, reserved=12,
         damaged=2, unit_value=18, maintenance=date(2026, 10, 5),
         notes="2 sets stained - replace"),
    dict(item="Dinnerware Sets (10 pc)", category="Tableware", owned=30,
         reserved=10, damaged=0, unit_value=25, maintenance=date(2027, 2, 1),
         notes=""),
    dict(item="Glassware Sets", category="Tableware", owned=60, reserved=20,
         damaged=3, unit_value=6, maintenance=date(2026, 10, 10),
         notes="Chips on 3 goblets"),
    dict(item="Cutlery Sets", category="Tableware", owned=250, reserved=100,
         damaged=0, unit_value=3, maintenance=date(2027, 3, 1), notes=""),
    dict(item="Beverage Dispensers", category="Beverage", owned=8,
         reserved=2, damaged=0, unit_value=55, maintenance=date(2026, 9, 25),
         notes="Tap service check"),
    dict(item="Portable Gas Ranges", category="Kitchen", owned=4, reserved=1,
         damaged=0, unit_value=320, maintenance=date(2026, 10, 8),
         notes=""),
    dict(item="Refrigerated Trailer", category="Transport", owned=1,
         reserved=1, damaged=0, unit_value=4500, maintenance=date(2026, 9, 20),
         notes="Booked for EV-007"),
    dict(item="Coffee Urns (100 cup)", category="Beverage", owned=6,
         reserved=2, damaged=1, unit_value=70, maintenance=date(2026, 12, 15),
         notes="1 element burnt out"),
]


# ---------------------------------------------------------------------------
# suppliers
# ---------------------------------------------------------------------------
SUPPLIERS = [
    dict(supplier="Metro Wholesale", contact="Aslam Bhai",
         phone="0300-1234567", email="orders@metrowholesale.example",
         category="Pantry", item="Rice, flour, pulses, oil", price=3.80,
         min_order=100, delivery="Next day", terms="Net 15",
         notes="Bulk discount over $500"),
    dict(supplier="City Foods", contact="Rana Sahib", phone="0321-7654321",
         email="sales@cityfoods.example", category="Meat & Fish",
         item="Chicken, beef, lamb, fish", price=12.00, min_order=50,
         delivery="Same day", terms="Cash", notes="Halal certified"),
    dict(supplier="Fresh Mart", contact="Ayesha K.", phone="0333-2224444",
         email="hello@freshmart.example", category="Produce",
         item="Vegetables, fruit, herbs", price=1.90, min_order=30,
         delivery="Same day", terms="Cash", notes="Farm fresh, seasonal"),
    dict(supplier="Dairy Land", contact="Usman D.", phone="0345-8889999",
         email="info@dairyland.example", category="Dairy",
         item="Cream, butter, cheese, eggs", price=6.40, min_order=40,
         delivery="Next day", terms="Net 7", notes=""),
    dict(supplier="Spice World", contact="Sana M.", phone="0301-5557777",
         email="orders@spiceworld.example", category="Pantry",
         item="Whole & ground spices", price=14.00, min_order=25,
         delivery="2 days", terms="Prepaid", notes="Custom blends available"),
    dict(supplier="PackRight", contact="Kamran S.", phone="0302-1112222",
         email="sales@packright.example", category="Packaging",
         item="Platters, boxes, cutlery", price=0.55, min_order=100,
         delivery="3 days", terms="Net 15", notes="Eco range available"),
    dict(supplier="ChefGear", contact="Adnan R.", phone="0334-7778888",
         email="support@chefgear.example", category="Equipment",
         item="Chafers, utensils, spares", price=45.00, min_order=0,
         delivery="4 days", terms="Net 30", notes=""),
    dict(supplier="Linen & Co", contact="Hira T.", phone="0311-4445555",
         email="hello@linenco.example", category="Linens",
         item="Cloths, napkins, runners", price=18.00, min_order=20,
         delivery="5 days", terms="50% advance", notes=""),
    dict(supplier="Beverage House", contact="Zeeshan A.",
         phone="0322-6667777", email="orders@bevhouse.example",
         category="Beverage", item="Teas, juices, syrups, water", price=2.10,
         min_order=60, delivery="2 days", terms="Net 15", notes=""),
    dict(supplier="Event Ice Co", contact="Mariam N.", phone="0308-3334444",
         email="info@eventice.example", category="Other",
         item="Ice, dry ice, cooler rental", price=0.80, min_order=50,
         delivery="Same day", terms="Cash", notes="Delivers to venue"),
]


# ---------------------------------------------------------------------------
# quote calculator demo state
# ---------------------------------------------------------------------------
QUOTE = dict(client="Fatima & Co (new enquiry)", date=_d(10, 31),
             guests=120, food_pp=9.50, labor=1200.00, equipment=250.00,
             transport=120.00, other=150.00)


# ---------------------------------------------------------------------------
# tax payments logged so far (tax sheet)
# ---------------------------------------------------------------------------
TAX_PAID = [
    dict(date=_d(7, 15), desc="Q2 sales tax filing", amount=610.00),
    dict(date=_d(8, 14), desc="Advance income tax", amount=400.00),
]


# ===========================================================================
# derived numbers + aggregates
# ===========================================================================
class Model(object):
    """Computed sample business - mirrors the Christmas tracker Model API."""

    def __init__(self, mode="demo", edition="premium"):
        self.mode = mode
        self.edition = edition
        self.demo = (mode == "demo")
        self.settings = dict(SETTINGS) if self.demo else {
            "business": "", "currency": "$", "tax": 0.0, "margin": 0.35,
            "deposit": 0.40, "duesoon": 7, "year": date.today().year,
            "cal_month": date.today().month, "cal_year": date.today().year,
            "message": "",
        }
        self.events = EVENTS if self.demo else []
        self.clients = CLIENTS if self.demo else []
        self.menu = MENU if self.demo else []
        self.inventory = INVENTORY if self.demo else []
        self.shopping = SHOPPING if self.demo else []
        self.expenses = EXPENSES if self.demo else []
        self.income = INCOME if self.demo else []
        self.staff = STAFF if self.demo else []
        self.equipment = EQUIPMENT if self.demo else []
        self.suppliers = SUPPLIERS if self.demo else []
        self.quote = dict(QUOTE) if self.demo else {}
        self.tax_paid = TAX_PAID if self.demo else []
        self.checks = {}
        self.agg = {}
        if self.demo:
            self._derive()

    def money(self, value, decimals=0):
        fmt = ",.%df" % decimals
        return self.settings["currency"] + format(value or 0, fmt)

    # ------------------------------------------------------------------
    def _derive(self):
        ev = self.events
        live = [e for e in ev if e["status"] != C.ES_CANCEL]
        done = [e for e in ev if e["status"] == C.ES_DONE]
        upcoming = [e for e in ev if e["status"] in
                    (C.ES_DEPOSIT, C.ES_CONFIRMED) and e["date"] >= TODAY]

        # per-event money
        for e in ev:
            e["profit"] = e["price"] - e["cost"]
            e["margin"] = (e["profit"] / e["price"]) if e["price"] else 0
            e["balance"] = e["price"] - e["deposit"]

        # income derivations
        for inv in self.income:
            e = next(x for x in ev if x["id"] == inv["event_id"])
            inv["received"] = inv["deposit"] + inv["pay1"] + inv["pay2"]
            inv["balance"] = inv["amount"] - inv["received"]
            if inv["received"] >= inv["amount"]:
                inv["status"] = C.PS_PAID
            elif inv["received"] > 0:
                inv["status"] = C.PS_PART
            else:
                inv["status"] = C.PS_UNPAID
            inv["overdue"] = (inv["balance"] > 0 and inv["due"] < TODAY)

        # shopping derivations
        for s in self.shopping:
            s["available"] = INV_QTY.get(s["ingredient"], 0)
            s["to_buy"] = max(0.0, s["required"] - s["available"])
            s["est_cost"] = round(s["to_buy"] * INV_COST.get(s["ingredient"], 0), 2)

        # menu derivations
        for m in self.menu:
            m["profit"] = round(m["price"] - m["cost"], 2)
            m["margin"] = (m["profit"] / m["price"]) if m["price"] else 0

        # inventory derivations
        for it in self.inventory:
            it["value"] = round(it["qty"] * it["unit_cost"], 2)
            if it["qty"] < it["min"]:
                it["status"] = "\U0001F534 Reorder"
            elif it["qty"] < it["min"] * 1.5:
                it["status"] = "\U0001F7E1 Low"
            else:
                it["status"] = "\U0001F7E2 OK"

        # staff derivations
        for s in self.staff:
            s["total"] = round(s["hours"] * s["rate"]
                               + s["ot"] * s["rate"] * 1.5, 2)

        # equipment derivations
        for q in self.equipment:
            q["available"] = q["owned"] - q["reserved"] - q["damaged"]
            q["value"] = q["owned"] * q["unit_value"]
            if q["damaged"] > 0:
                q["status"] = "\U0001F527 Service"
            elif q["maintenance"] <= date(2026, 10, 12):
                q["status"] = "\U0001F7E1 Soon"
            else:
                q["status"] = "\U0001F7E2 Ready"

        # quote calculator outputs
        q = self.quote
        food_total = q["guests"] * q["food_pp"]
        cost_total = food_total + q["labor"] + q["equipment"] \
            + q["transport"] + q["other"]
        margin = self.settings["margin"]
        price = cost_total / (1 - margin) if margin < 1 else cost_total
        q.update(food_total=food_total, cost_total=cost_total,
                 profit=price - cost_total, price=price,
                 per_guest=price / q["guests"],
                 deposit=price * self.settings["deposit"],
                 balance=price * (1 - self.settings["deposit"]),
                 cost_pp=cost_total / q["guests"])

        # ------------------------------------------------------------
        # aggregates (cached KPI / report values)
        # ------------------------------------------------------------
        a = self.agg
        revenue = sum(i["received"] for i in self.income)
        expenses = sum(x["amount"] for x in self.expenses)
        food_cost = sum(x["amount"] for x in self.expenses
                        if x["category"] in ("Ingredients", "Packaging"))
        labor_cost = sum(x["amount"] for x in self.expenses
                         if x["category"] == "Staff / Labor")
        outstanding = sum(i["balance"] for i in self.income)
        overdue_n = sum(1 for i in self.income if i["overdue"])
        invoices_open = sum(1 for i in self.income if i["balance"] > 0)

        a["revenue"] = revenue
        a["expenses"] = expenses
        a["profit"] = revenue - expenses
        a["margin"] = (revenue - expenses) / revenue if revenue else 0
        a["events_total"] = len(ev)
        a["events_done"] = len(done)
        a["events_confirmed"] = len([e for e in ev if e["status"] in
                                     (C.ES_DEPOSIT, C.ES_CONFIRMED)])
        a["events_cancelled"] = sum(1 for e in ev
                                    if e["status"] == C.ES_CANCEL)
        a["avg_order"] = sum(e["price"] for e in live) / len(live)
        a["avg_guests"] = sum(e["guests"] for e in live) / len(live)
        a["avg_profit"] = sum(e["profit"] for e in live) / len(live)
        a["guests_total"] = sum(e["guests"] for e in done)
        a["outstanding"] = outstanding
        a["overdue"] = overdue_n
        a["invoices_open"] = invoices_open
        a["food_cost"] = food_cost
        a["labor_cost"] = labor_cost
        a["food_pct"] = food_cost / revenue if revenue else 0
        a["labor_pct"] = labor_cost / revenue if revenue else 0
        repeat = sum(1 for name in {e["client"] for e in ev}
                     if sum(1 for e in ev if e["client"] == name) > 1)
        a["repeat_pct"] = repeat / len(CLIENTS)
        a["inv_value"] = round(sum(i["value"] for i in self.inventory), 2)
        a["low_stock"] = sum(1 for i in self.inventory
                             if i["qty"] < i["min"])
        a["shop_lines"] = sum(1 for s in self.shopping
                              if s["to_buy"] > 0 and s["purchased"] != C.TICK)
        a["shop_cost"] = round(sum(s["est_cost"] for s in self.shopping
                                   if s["purchased"] != C.TICK), 2)
        a["staff_unpaid"] = sum(1 for s in self.staff if s["paid"] != C.TICK)
        a["equip_value"] = sum(q["value"] for q in self.equipment)
        a["equip_service"] = sum(1 for q in self.equipment
                                 if q["damaged"] > 0
                                 or q["maintenance"] <= date(2026, 10, 12))
        taxable = sum(i["amount"] for i in self.income)
        a["tax_collected"] = round(taxable * self.settings["tax"], 2)
        a["tax_due"] = round(a["tax_collected"]
                             - sum(t["amount"] for t in self.tax_paid), 2)
        a["quote_price"] = round(q["price"], 2)
        a["quote_per_guest"] = round(q["per_guest"], 2)
        a["menu_items"] = len(self.menu)
        a["menu_margin"] = sum(m["margin"] for m in self.menu) / len(self.menu)

        # monthly pools (cash in, by event month)
        months = []
        for mi in range(12):
            label = MONTH_SHORT[mi]
            mrev = sum(i["received"] for i in self.income
                       if i["event_date"].month == mi + 1)
            mexp = sum(x["amount"] for x in self.expenses
                       if x["date"].month == mi + 1)
            mcnt = sum(1 for e in ev if e["date"].month == mi + 1)
            months.append((label, mrev, mexp, mrev - mexp, mcnt))
        a["months"] = months

        # revenue by event type (booked price, cancelled excluded)
        a["types"] = [(t, sum(e["price"] for e in live if e["type"] == t))
                      for t in C.EVENT_TYPES]

        # revenue by client (cash received)
        a["clients_pool"] = [(cl["name"],
                              sum(i["received"] for i in self.income
                                  if i["client"] == cl["name"]))
                             for cl in self.clients]

        # menu popularity (events featuring each dish)
        a["menu_pool"] = [(m["item"],
                           sum(1 for e in live if m["item"] in e["menu"]))
                          for m in self.menu]

        # checklist tick state for the EXAMPLE build
        self.checks = {
            "prep": set(range(8)),
            "shop": {0, 1, 2, 3, 5, 6},
            "day": set(),
            "end": set(range(8)),
        }

        # expense-category + status pools for the charts
        a["exp_cats"] = {}
        for x in self.expenses:
            a["exp_cats"][x["category"]] = \
                a["exp_cats"].get(x["category"], 0) + x["amount"]
        a["status_counts"] = {}
        for e in ev:
            a["status_counts"][e["status"]] = \
                a["status_counts"].get(e["status"], 0) + 1

        best_type = max(a["types"], key=lambda t: t[1])
        best_client = max(a["clients_pool"], key=lambda t: t[1])
        a["best_type"] = best_type[0] if best_type[1] else "-"
        a["best_client"] = best_client[0] if best_client[1] else "-"

        # upcoming-events pool (future, not cancelled/done) + dues pool
        a["upcoming"] = [(e["date"], "%s - %s" % (e["client"], e["type"]),
                          e["status"])
                         for e in ev
                         if e["date"] >= TODAY and e["status"] not in
                         (C.ES_CANCEL, C.ES_DONE)]
        a["dues"] = [(i["due"], "%s (%s)" % (i["client"], i["invoice"]),
                      i["balance"])
                     for i in self.income if i["balance"] > 0]
```

### `catering_tracker/book.py`

```python
"""
The build context.

``Book`` owns the workbook, the worksheet registry, the defined names and all
of the little helpers that turn a cell address into a formula fragment.  Sheet
builders receive one ``Book`` and never talk to XlsxWriter's raw indices
directly, which is what keeps the cross-sheet formulas correct when a tab is
missing (Basic edition) or when a capacity changes.
"""

import datetime
import types

import xlsxwriter
from xlsxwriter.utility import datetime_to_excel_datetime

from . import config as C
from .styles import Styles, wrap_height

_XLWorksheet = xlsxwriter.worksheet.Worksheet


def _coerce(value):
    """Cached formula results must be a number, a bool or a string.

    XlsxWriter writes whatever it is given straight into ``<v>``; a
    ``datetime.date`` would produce invalid XML that Excel refuses to open.
    """
    if value is None:
        return ""
    if isinstance(value, bool):
        return value
    if isinstance(value, (datetime.datetime, datetime.date)):
        return datetime_to_excel_datetime(value, False, False)
    if isinstance(value, (str, int, float)):
        return value
    return str(value)


def _safe_write_formula(self, row, col, formula, cell_format=None, value=0):
    """``write_formula`` that also coerces the cached result."""
    return _XLWorksheet.write_formula(self, row, col, formula, cell_format,
                                      _coerce(value))


def r(n):
    """1-indexed spreadsheet row -> 0-indexed XlsxWriter row."""
    return n - 1


def cellref(col, row):
    return "%s%d" % (col, row)


class Book(object):

    def __init__(self, path, theme, edition="premium", mode="blank",
                 protect=None, images=True, demo=None):
        self.path = path
        self.th = theme
        self.edition = edition
        self.mode = mode            # "blank" | "demo"
        self.protect = protect
        self.images = images
        self.demo = demo

        self.order = C.EDITIONS[edition]
        self.wb = xlsxwriter.Workbook(path)
        self.S = Styles(self.wb, theme)
        self.sheets = {}
        self.stats = {"formulas": 0, "validations": 0, "cond_formats": 0,
                      "charts": 0, "links": 0, "cells": 0}

        self._create_sheets()
        self._define_names()

        self.wb.set_properties({
            "title": C.PRODUCT,
            "subject": C.TAGLINE,
            "author": C.AUTHOR,
            "manager": C.AUTHOR,
            "company": C.AUTHOR,
            "category": "Catering Business Management Spreadsheet",
            "keywords": ("catering business, catering planner, food cost "
                         "calculator, event tracker, client tracker, quote "
                         "calculator, recipe costing, inventory, profit "
                         "dashboard, excel template, google sheets, etsy "
                         "spreadsheet"),
            "comments": ("%s v%s - %s. Works in Excel 2016+ and Google "
                         "Sheets. No macros, nothing to install."
                         % (C.PRODUCT, C.VERSION, C.TAGLINE)),
        })
        self.wb.set_calc_mode("auto")

    # ------------------------------------------------------------------
    # sheet registry
    # ------------------------------------------------------------------
    def _create_sheets(self):
        for key in self.order:
            ws = self.wb.add_worksheet(C.SHEET_NAMES[key])
            ws.write_formula = types.MethodType(_safe_write_formula, ws)
            ws.set_tab_color(self.th.tabs[key])
            ws.hide_gridlines(2)
            self.sheets[key] = ws
        if "data" in self.sheets:
            self.sheets["data"].hide()
        self.sheets["dashboard"].set_first_sheet()
        self.sheets["dashboard"].activate()
        self.sheets["dashboard"].set_zoom(90)

    def ws(self, key):
        return self.sheets[key]

    def has(self, key):
        return key in self.sheets and key != "data"

    def name(self, key):
        return C.SHEET_NAMES[key]

    def q(self, key):
        """Quoted sheet name for use inside formulas."""
        return "'%s'" % C.SHEET_NAMES[key]

    # ------------------------------------------------------------------
    # defined names
    # ------------------------------------------------------------------
    def _define_names(self):
        su = self.q("setup")
        simple = {
            "BusinessName": "%s!$C$%d" % (su, C.SU_BUSINESS),
            "Currency": "%s!$C$%d" % (su, C.SU_CURRENCY),
            "TaxRate": "%s!$C$%d" % (su, C.SU_TAX),
            "DefaultMargin": "%s!$C$%d" % (su, C.SU_MARGIN),
            "DepositPct": "%s!$C$%d" % (su, C.SU_DEPOSIT),
            "DueSoonDays": "%s!$C$%d" % (su, C.SU_DUESOON),
            "CalMonth": "%s!$C$%d" % (su, C.SU_CAL_MONTH),
            "CalYear": "%s!$C$%d" % (su, C.SU_CAL_YEAR),
            "ReportYear": "%s!$C$%d" % (su, C.SU_YEAR),
        }
        for nm, ref in simple.items():
            self.wb.define_name(nm, "=" + ref)

        first = C.SU_LIST_FIRST
        lastc = C.SU_LIST_FIRST + C.SU_LIST_ROWS - 1
        for key, col in sorted(C.LIST_COLS.items()):
            rng = "%s!$%s$%d:$%s$%d" % (su, col, first, col, lastc)
            self.wb.define_name(
                _name_for_list(key),
                "=OFFSET(%s!$%s$%d,0,0,MAX(1,COUNTA(%s)),1)" % (su, col, first, rng))

        # Fixed (non-customisable) status lists, in the locked block.
        fixed = {
            "EventStatuses": ("event_statuses", len(C.EVENT_STATUSES)),
            "PaymentStatuses": ("payment_statuses", len(C.PAYMENT_STATUSES)),
            "Tick": ("tick", 2),
        }
        for nm, (key, n) in sorted(fixed.items()):
            col = C.FIXED_COLS[key]
            self.wb.define_name(
                nm, "=%s!$%s$%d:$%s$%d"
                % (su, col, C.SU_FIXED_FIRST, col, C.SU_FIXED_FIRST + n - 1))

        # Dynamic lists that live in the working sheets (grow with the data).
        dyn = {"EventList": ("events", "C"), "ClientsList": ("clients", "C")}
        if self.has("menu"):
            dyn["MenuItems"] = ("menu", "C")
        if self.has("suppliers"):
            dyn["SuppliersList"] = ("suppliers", "C")
        for nm, (sheet, colL) in sorted(dyn.items()):
            rng = "%s!$%s$%d:$%s$%d" % (self.q(sheet), colL, C.ROW_FIRST,
                                        colL, C.last_row(sheet))
            base = rng.rsplit("$", 1)[0].rsplit(":", 1)[0]
            self.wb.define_name(
                nm, "=OFFSET(%s,0,0,MAX(1,COUNTA(%s)),1)" % (base, rng))

    # ------------------------------------------------------------------
    # reference helpers
    # ------------------------------------------------------------------
    def col(self, key, field):
        return C.COLS[key][field]

    def rng(self, key, field, first=None, last=None, quoted=True):
        """Absolute range for a column of a tracker sheet."""
        col = self.col(key, field)
        first = C.ROW_FIRST if first is None else first
        last = C.last_row(key) if last is None else last
        ref = "%s!$%s$%d:$%s$%d" % (self.q(key) if quoted else self.name(key),
                                    col, first, col, last)
        return ref

    def cell(self, key, field, row):
        return "%s!$%s$%d" % (self.q(key), self.col(key, field), row)

    def kpi(self, key):
        return "%s!$AF$%d" % (self.q("data"), C.KPI_ROW[key])

    def kpi_fmt(self, key):
        return C.KPI_FMT[key]

    def data_rng(self, col, first, last):
        return "%s!$%s$%d:$%s$%d" % (self.q("data"), col, first, col, last)

    def data_cell(self, col, row):
        return "%s!$%s$%d" % (self.q("data"), col, row)

    def listname(self, key):
        return _name_for_list(key)

    # ------------------------------------------------------------------
    # cached values (so previews look right before Excel recalculates)
    # ------------------------------------------------------------------
    def cached(self, key, default=0):
        if self.demo and key in self.demo.agg:
            return self.demo.agg[key]
        return default

    # ------------------------------------------------------------------
    # money / text formula helpers
    # ------------------------------------------------------------------
    def money(self, value_formula, decimals="#,##0"):
        """Currency-prefixed text, driven by the Currency setting."""
        return 'Currency&TEXT(%s,"%s")' % (value_formula, decimals)

    def bar(self, numerator, denominator, blocks=18):
        """Text progress bar built with REPT()."""
        pct = "MIN(1,IFERROR((%s)/(%s),0))" % (numerator, denominator)
        filled = "ROUND(%s*%d,0)" % (pct, blocks)
        return ('=IFERROR(REPT("\u2588",%s)&REPT("\u2591",%d-%s),"%s")'
                % (filled, blocks, filled, "\u2591" * blocks))

    def bar_static(self, numerator, denominator, blocks=18):
        """Python twin of :meth:`bar` used for cached values."""
        try:
            pct = min(1.0, float(numerator) / float(denominator)) if denominator else 0.0
        except (TypeError, ValueError, ZeroDivisionError):
            pct = 0.0
        filled = int(round(pct * blocks))
        return "\u2588" * filled + "\u2591" * (blocks - filled)

    # ------------------------------------------------------------------
    # page furniture
    # ------------------------------------------------------------------
    def page(self, key, last_col, last_row, landscape=True, freeze=None,
             fit=True, zoom=90, title_rows=None, paper=9):
        ws = self.sheets[key]
        ws.set_zoom(zoom)
        if landscape:
            ws.set_landscape()
        if fit:
            ws.fit_to_pages(1, 0)
        ws.set_margins(0.4, 0.4, 0.5, 0.5)
        ws.set_paper(paper)
        ws.set_header("&L&\"%s,Bold\"&11%s&R&\"%s,Italic\"&9Page &P of &N"
                      % (self.th.body_font, C.PRODUCT_SHORT,
                         self.th.body_font))
        ws.set_footer("&C&\"%s,Italic\"&8&A  \u2022  %s v%s  \u2022  "
                      "personal licence"
                      % (self.th.body_font, C.PRODUCT, C.VERSION))
        ws.print_area("A1:%s%d" % (last_col, last_row))
        if title_rows:
            ws.repeat_rows(title_rows[0], title_rows[1])
        if freeze:
            ws.freeze_panes(*freeze)
        if self.protect:
            ws.protect(self.protect, {
                "objects": True, "scenarios": True, "format_cells": True,
                "format_columns": True, "format_rows": True,
                "insert_rows": True, "insert_columns": True,
                "insert_hyperlinks": True, "delete_rows": True,
                "delete_columns": True, "sort": True, "autofilter": True,
                "select_locked_cells": True, "select_unlocked_cells": True,
            })
        return ws

    def widths(self, key, extra=None):
        ws = self.sheets[key]
        for col, width in sorted((C.WIDTHS.get(key, {}) or {}).items()):
            ws.set_column("%s:%s" % (col, col), width)
        for col, width in sorted((extra or {}).items()):
            ws.set_column("%s:%s" % (col, col), width)
        return ws

    def title_block(self, key, title, subtitle, last_col, home=True):
        """Standard 6-row header band used by every tracker tab."""
        ws = self.sheets[key]
        S = self.S
        ws.set_row(r(C.ROW_SPACER_1), 7)
        ws.set_row(r(C.ROW_TITLE), 30)
        ws.set_row(r(C.ROW_SUBTITLE), 18)
        ws.set_row(r(C.ROW_SPACER_2), 7)
        ws.set_row(r(C.ROW_STATS), 26)
        ws.set_row(r(C.ROW_SPACER_3), 7)
        ws.merge_range(r(C.ROW_TITLE), 1, r(C.ROW_TITLE), _ci(last_col) - 3,
                       title, S.sheet_title)
        for c in range(_ci(last_col) - 2, _ci(last_col) + 1):
            ws.write_blank(r(C.ROW_TITLE), c, None, S.canvas)
        ws.merge_range(r(C.ROW_SUBTITLE), 1, r(C.ROW_SUBTITLE), _ci(last_col) - 3,
                       subtitle, S.sheet_sub)
        if home:
            ws.merge_range(r(C.ROW_SUBTITLE), _ci(last_col) - 2,
                           r(C.ROW_SUBTITLE), _ci(last_col),
                           "", S.home_link)
            ws.write_url(r(C.ROW_SUBTITLE), _ci(last_col) - 2,
                         "internal:%s!A1" % self.q("dashboard"), S.home_link,
                         "\U0001F3E0  Back to Dashboard")
            self.stats["links"] += 1
        else:
            ws.merge_range(r(C.ROW_SUBTITLE), _ci(last_col) - 2,
                           r(C.ROW_SUBTITLE), _ci(last_col), "", S.canvas)
        return ws

    def stats_strip(self, key, items, row=None, first_col=1, span=2):
        """Row of small formula chips under the sheet title.

        ``items`` is a list of ``(formula_or_text, colour_key, cached)``.
        """
        ws = self.sheets[key]
        row = C.ROW_STATS if row is None else row
        col = first_col
        for formula, color, cached in items:
            fmt = self.S.pill(self.th.soft(color), getattr(self.th, color),
                              size=10.5, bold=True, align="left")
            ws.merge_range(r(row), col, r(row), col + span - 1, "", fmt)
            if formula.startswith("="):
                ws.write_formula(r(row), col, formula, fmt,
                                 cached if cached is not None else 0)
                self.stats["formulas"] += 1
            else:
                ws.write(r(row), col, formula, fmt)
            col += span
        return ws

    def nav_row(self, key, row, first_col=1, span=2, max_col="N"):
        """Bottom navigation buttons (one per visible tab)."""
        ws = self.sheets[key]
        palette = ["primary", "accent", "gold", "info", "plum", "ok",
                   "primary_2", "warn", "accent", "info"]
        col = first_col
        n = 0
        for target in self.order:
            if target == "data":
                continue
            if col + span - 1 > _ci(max_col):
                row += 1
                col = first_col
                ws.set_row(r(row), 24)
            color = getattr(self.th, palette[n % len(palette)])
            fmt = self.S.nav(color)
            if span > 1:
                ws.merge_range(r(row), col, r(row), col + span - 1, "", fmt)
            label = C.SHEET_SHORT[target]
            if target == key:
                label = "\u25B6 " + label
            ws.write_url(r(row), col, "internal:%s!A1" % self.q(target),
                         fmt, label)
            self.stats["links"] += 1
            col += span
            n += 1
        return row

    # ------------------------------------------------------------------
    # painting / writing helpers
    # ------------------------------------------------------------------
    def paint(self, key, row1, col1, row2, col2, fmt):
        ws = self.sheets[key]
        for rr in range(row1, row2 + 1):
            for cc in range(col1, col2 + 1):
                ws.write_blank(rr, cc, None, fmt)

    def band(self, key, row, col1, col2, fmt, height=None):
        ws = self.sheets[key]
        ws.merge_range(r(row), col1, r(row), col2, "", fmt)
        if height:
            ws.set_row(r(row), height)

    def write(self, key, row, col, value, fmt=None, cached=None):
        ws = self.sheets[key]
        self.stats["cells"] += 1
        if isinstance(value, str) and value.startswith("="):
            ws.write_formula(r(row), col, value, fmt,
                             cached if cached is not None else 0)
            self.stats["formulas"] += 1
        elif value is None or value == "":
            ws.write_blank(r(row), col, None, fmt)
        else:
            ws.write(r(row), col, value, fmt)
        return ws

    def merge(self, key, row1, col1, row2, col2, value, fmt=None,
              cached=None):
        ws = self.sheets[key]
        self.stats["cells"] += 1
        if isinstance(value, str) and value.startswith("="):
            ws.merge_range(r(row1), col1, r(row2), col2, "", fmt)
            ws.write_formula(r(row1), col1, value, fmt,
                             cached if cached is not None else 0)
            self.stats["formulas"] += 1
        else:
            ws.merge_range(r(row1), col1, r(row2), col2, value, fmt)
        return ws

    def para(self, key, row, col1, col2, text, fmt, width_chars=None,
             minimum=18):
        """Merged, wrapped paragraph with a row height that actually fits."""
        if width_chars is None:
            width_chars = sum(
                (C.WIDTHS.get(key, {}) or {}).get(_cl(c), 10)
                for c in range(col1, col2 + 1))
        h = wrap_height(text, width_chars, minimum=minimum)
        self.merge(key, row, col1, row, col2, text, fmt)
        self.sheets[key].set_row(r(row), h)
        return row + 1

    def validate(self, key, row1, col1, row2, col2, source, title=None,
                 message=None, error=None, error_type="stop"):
        ws = self.sheets[key]
        opts = {"validate": "list", "source": source, "ignore_blank": True,
                "show_input": bool(title or message),
                "show_error": bool(error)}
        if title:
            opts["input_title"] = title[:32]
        if message:
            opts["input_message"] = message[:255]
        if error:
            opts["error_title"] = "Pick from the list"[:32]
            opts["error_message"] = error[:255]
            opts["error_type"] = error_type
        res = ws.data_validation(r(row1), col1, r(row2), col2, opts)
        if res == 0:
            self.stats["validations"] += 1
        return res

    def cond(self, key, row1, col1, row2, col2, opts):
        ws = self.sheets[key]
        res = ws.conditional_format(r(row1), col1, r(row2), col2, opts)
        if res == 0:
            self.stats["cond_formats"] += 1
        return res

    def chart(self, ctype, **opts):
        ch = self.wb.add_chart(dict({"type": ctype}, **opts))
        ch.show_hidden_data()          # data lives on the hidden _Data sheet
        ch.show_blanks_as("gap")
        self.stats["charts"] += 1
        return ch

    def close(self):
        self.wb.close()


# ----------------------------------------------------------------------
# small utilities
# ----------------------------------------------------------------------
def _ci(letter):
    """Column letter -> 0-indexed column number."""
    n = 0
    for ch in letter:
        n = n * 26 + (ord(ch.upper()) - 64)
    return n - 1


def _cl(index):
    """0-indexed column number -> column letter."""
    s = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        s = chr(65 + rem) + s
    return s


_LIST_NAMES = {
    "event_types": "EventTypes",
    "expense_categories": "ExpenseCategories",
    "payment_methods": "PaymentMethods",
    "staff_roles": "StaffRoles",
    "menu_categories": "MenuCategories",
    "ingredient_categories": "IngredientCategories",
}


def _name_for_list(key):
    return _LIST_NAMES[key]


def name_for_list(key):
    return _LIST_NAMES[key]


def ci(letter):
    return _ci(letter)


def cl(index):
    return _cl(index)
```

### `catering_tracker/workbook.py`

```python
"""
Orchestrator: turns (edition, theme, mode) into a finished .xlsx.

Editions
    basic    dashboard + clients + events + quote calculator + expenses +
             payments + setup + guide  (the essentials)
    premium  everything: menu costing, inventory, shopping list, staff,
             equipment, suppliers, calendar, P&L/reports, tax tracker,
             checklists and the printable invoice

Themes
    classic  cream / espresso / copper / brass
    fresh    white / basil / tomato / lemon

Modes
    blank    clean, ready-to-use template (lists + settings pre-seeded)
    demo     filled-in example business, for listing screenshots
"""

import os

from . import theme as themes
from .book import Book
from .demo import Model
from .sheets import (calendar, checklists, clients, dashboard, data,
                     equipment, events, expenses, guide, income, inventory,
                     invoice, menu, quote, reports, setup, shopping, staff,
                     suppliers, tax)

BUILDERS = {
    "data": data.build,
    "setup": setup.build,
    "dashboard": dashboard.build,
    "clients": clients.build,
    "events": events.build,
    "quote": quote.build,
    "menu": menu.build,
    "inventory": inventory.build,
    "shopping": shopping.build,
    "expenses": expenses.build,
    "income": income.build,
    "staff": staff.build,
    "equipment": equipment.build,
    "suppliers": suppliers.build,
    "calendar": calendar.build,
    "reports": reports.build,
    "tax": tax.build,
    "checklists": checklists.build,
    "invoice": invoice.build,
    "guide": guide.build,
}

# Build order: the hidden engine first, then the tabs left-to-right.
BUILD_ORDER = ["data", "setup", "dashboard", "clients", "events", "quote",
               "menu", "inventory", "shopping", "expenses", "income",
               "staff", "equipment", "suppliers", "calendar", "reports",
               "tax", "checklists", "invoice", "guide"]


def build_workbook(path, edition="premium", theme_name="classic",
                   mode="blank", protect=None, images=True):
    if edition not in ("basic", "premium"):
        raise ValueError("edition must be 'basic' or 'premium'")
    if mode not in ("blank", "demo"):
        raise ValueError("mode must be 'blank' or 'demo'")
    th = themes.get(theme_name)
    model = Model(mode, edition)
    bk = Book(path, th, edition=edition, mode=mode, protect=protect,
              images=images, demo=model)
    for key in BUILD_ORDER:
        if key in bk.order:
            BUILDERS[key](bk)
    bk.close()
    bk.stats["path"] = path
    bk.stats["size"] = os.path.getsize(path)
    return bk.stats


# ---------------------------------------------------------------------------
# the curated product set written to products/
# ---------------------------------------------------------------------------
def product_filename(edition, theme_name, mode):
    suffix = "_EXAMPLE" if mode == "demo" else ""
    return ("Catering_Business_Manager_%s_%s%s.xlsx"
            % (edition.upper(), theme_name.capitalize(), suffix))


def build_all(outdir="products", protect=None, images=True):
    """The set of files an Etsy listing actually ships / screenshots."""
    combos = [
        ("premium", "classic", "demo"),    # listing hero + screenshots
        ("premium", "classic", "blank"),   # the product
        ("premium", "fresh", "blank"),     # second style
        ("basic", "classic", "blank"),     # the cheaper tier
        ("basic", "fresh", "blank"),
        ("basic", "classic", "demo"),      # cheaper tier screenshot
    ]
    os.makedirs(outdir, exist_ok=True)
    stats = []
    for edition, theme_name, mode in combos:
        path = os.path.join(outdir,
                            product_filename(edition, theme_name, mode))
        stats.append(build_workbook(path, edition, theme_name, mode,
                                    protect=protect, images=images))
    return stats
```

### `catering_tracker/sheets/__init__.py`

```python
"""Worksheet builders - one module per tab."""
```

### `catering_tracker/sheets/common.py`

```python
"""Shared worksheet furniture: headers, data rows, totals, validations."""

from datetime import date, datetime

from .. import config as C
from ..book import r, ci, cl


# ---------------------------------------------------------------------------
# geometry helpers
# ---------------------------------------------------------------------------
def first_row():
    return C.ROW_FIRST


def last_row(key):
    return C.last_row(key)


def n_rows(key):
    return C.CAP[key]


def alt(rownum):
    """Zebra stripe: even spreadsheet rows get the tinted background."""
    return (rownum % 2) == 0


# ---------------------------------------------------------------------------
# headers
# ---------------------------------------------------------------------------
def header_row(bk, key, columns, row=None, height=36):
    """``columns`` = list of (field, label, kind, colour_key|None)."""
    ws = bk.ws(key)
    row = C.ROW_HEADER if row is None else row
    for field, label, kind, color in columns:
        col = ci(bk.col(key, field))
        ws.write(r(row), col, label,
                 bk.S.header(getattr(bk.th, color) if color else None))
    ws.set_row(r(row), height)
    return row


def data_rows(bk, key, height=20):
    ws = bk.ws(key)
    for i in range(n_rows(key)):
        ws.set_row(r(C.ROW_FIRST + i), height)


def table_frame(bk, key, columns, height=20):
    """Write the header and pre-format every empty data cell."""
    header_row(bk, key, columns)
    ws = bk.ws(key)
    data_rows(bk, key, height)
    for i in range(n_rows(key)):
        rownum = C.ROW_FIRST + i
        a = alt(rownum)
        for field, label, kind, color in columns:
            col = ci(bk.col(key, field))
            if kind == "idx":
                fmt = bk.S.idx(a)
            elif kind.startswith("calc"):
                fmt = bk.S.cell(kind)
            else:
                fmt = bk.S.cell(kind, a)
            ws.write_blank(r(rownum), col, None, fmt)
    return ws


# ---------------------------------------------------------------------------
# writing a single row of a tracker table
# ---------------------------------------------------------------------------
def write_row(bk, key, columns, rownum, values, cached=None):
    ws = bk.ws(key)
    a = alt(rownum)
    cached = cached or {}
    for field, label, kind, color in columns:
        col = ci(bk.col(key, field))
        value = values.get(field)
        if kind == "idx":
            fmt = bk.S.idx(a)
        elif kind.startswith("calc"):
            fmt = bk.S.cell(kind)
        else:
            fmt = bk.S.cell(kind, a)
        if value is None or value == "":
            ws.write_blank(r(rownum), col, None, fmt)
        elif isinstance(value, str) and value.startswith("="):
            ws.write_formula(r(rownum), col, value, fmt,
                             cached.get(field, 0))
            bk.stats["formulas"] += 1
            bk.stats["cells"] += 1
        elif isinstance(value, (date, datetime)):
            ws.write_datetime(r(rownum), col, value, fmt)
            bk.stats["cells"] += 1
        else:
            ws.write(r(rownum), col, value, fmt)
            bk.stats["cells"] += 1
    return ws


# ---------------------------------------------------------------------------
# validation presets
# ---------------------------------------------------------------------------
def list_dv(bk, key, field, list_key, title=None, message=None, error=None,
            first=None, last=None):
    return bk.validate(key, first or C.ROW_FIRST, ci(bk.col(key, field)),
                       last or C.last_row(key), ci(bk.col(key, field)),
                       "=" + bk.listname(list_key), title=title,
                       message=message, error=error)


def fixed_dv(bk, key, field, name, title=None, message=None, error=None):
    return bk.validate(key, C.ROW_FIRST, ci(bk.col(key, field)),
                       C.last_row(key), ci(bk.col(key, field)), "=" + name,
                       title=title, message=message, error=error)


def tick_dv(bk, key, fields, message="Pick \u2713 from the dropdown to tick "
                                     "it off (leave blank for not yet)."):
    if isinstance(fields, str):
        fields = (fields,)
    for field in fields:
        bk.validate(key, C.ROW_FIRST, ci(bk.col(key, field)),
                    C.last_row(key), ci(bk.col(key, field)), "=Tick",
                    title="Tick it off", message=message)


def money_dv(bk, key, fields, label="money"):
    if isinstance(fields, str):
        fields = (fields,)
    for field in fields:
        bk.ws(key).data_validation(
            r(C.ROW_FIRST), ci(bk.col(key, field)),
            r(C.last_row(key)), ci(bk.col(key, field)),
            {"validate": "decimal", "criteria": ">=", "value": 0,
             "ignore_blank": True, "show_error": True,
             "error_title": "Enter %s" % label,
             "error_message": "Please enter a positive number (no currency "
                              "symbol) - the symbol comes from Setup.",
             "error_type": "warning"})
        bk.stats["validations"] += 1


def date_dv(bk, key, fields, message="Type a date, or pick one from the "
                                     "calendar picker."):
    if isinstance(fields, str):
        fields = (fields,)
    for field in fields:
        bk.ws(key).data_validation(
            r(C.ROW_FIRST), ci(bk.col(key, field)),
            r(C.last_row(key)), ci(bk.col(key, field)),
            {"validate": "date", "criteria": "between",
             "minimum": date(2000, 1, 1), "maximum": date(2100, 12, 31),
             "ignore_blank": True, "show_input": True, "input_title": "Date",
             "input_message": message, "show_error": True,
             "error_title": "That's not a date",
             "error_message": "Enter a date between 2000 and 2100.",
             "error_type": "warning"})
        bk.stats["validations"] += 1


def whole_dv(bk, key, fields, minimum=0, maximum=9999):
    if isinstance(fields, str):
        fields = (fields,)
    for field in fields:
        bk.ws(key).data_validation(
            r(C.ROW_FIRST), ci(bk.col(key, field)),
            r(C.last_row(key)), ci(bk.col(key, field)),
            {"validate": "whole", "criteria": "between", "minimum": minimum,
             "maximum": maximum, "ignore_blank": True, "show_error": True,
             "error_title": "Whole numbers only",
             "error_message": "Please enter a whole number.",
             "error_type": "warning"})
        bk.stats["validations"] += 1


# ---------------------------------------------------------------------------
# conditional formatting presets
# ---------------------------------------------------------------------------
def tick_cf(bk, key, fields):
    """Green glow on any ticked box."""
    if isinstance(fields, str):
        fields = (fields,)
    for field in fields:
        col = ci(bk.col(key, field))
        L = bk.col(key, field)
        bk.cond(key, C.ROW_FIRST, col, C.last_row(key), col, {
            "type": "formula", "criteria": '=$%s%d="%s"' % (L, C.ROW_FIRST,
                                                            C.TICK),
            "format": bk.S.cf(bg=bk.th.ok_soft, fg=bk.th.ok, bold=True,
                              size=13, border=bk.th.border)})


def secret_cf(bk, key, field):
    """Secret Mode: make the hiding-spot text invisible (white on white)."""
    col = ci(bk.col(key, field))
    bk.cond(key, C.ROW_FIRST, col, C.last_row(key), col, {
        "type": "formula", "criteria": '=SecretMode="Yes"',
        "format": bk.S.cf(bg=bk.th.card, fg=bk.th.card)})


def status_cf(bk, key, field, mapping, first=None, last=None):
    """Colour a status column.  ``mapping`` = {cell text: (bg, fg)}."""
    L = bk.col(key, field)
    col = ci(L)
    first = first or C.ROW_FIRST
    last = last or C.last_row(key)
    for text, (bg, fg) in mapping.items():
        bk.cond(key, first, col, last, col, {
            "type": "formula",
            "criteria": '=$%s%d="%s"' % (L, first, text),
            "format": bk.S.cf(bg=bg, fg=fg, bold=True, border=bk.th.border)})


def deadline_cf(bk, key, date_field, tick_field=None):
    """Overdue = red, due inside the DueSoonDays window = amber."""
    L = bk.col(key, date_field)
    col = ci(L)
    first, last = C.ROW_FIRST, C.last_row(key)
    done = ""
    if tick_field:
        done = '*($%s%d<>"%s")' % (bk.col(key, tick_field), first, C.TICK)
    bk.cond(key, first, col, last, col, {
        "type": "formula",
        "criteria": '=AND($%s%d<>"",$%s%d<TODAY()%s)' % (L, first, L, first,
                                                         done),
        "format": bk.S.cf(bg=bk.th.bad_soft, fg=bk.th.bad, bold=True)})
    bk.cond(key, first, col, last, col, {
        "type": "formula",
        "criteria": '=AND($%s%d>=TODAY(),$%s%d-TODAY()<=DueSoonDays%s)'
                   % (L, first, L, first, done),
        "format": bk.S.cf(bg=bk.th.warn_soft, fg=bk.th.warn, bold=True)})


def databar(bk, key, field, color=None, first=None, last=None):
    col = ci(bk.col(key, field))
    bk.cond(key, first or C.ROW_FIRST, col, last or C.last_row(key), col, {
        "type": "data_bar", "bar_color": color or bk.th.primary_2,
        "bar_solid": True, "min_type": "num", "min_value": 0,
        "max_type": "num", "max_value": 1})


# ---------------------------------------------------------------------------
# totals row
# ---------------------------------------------------------------------------
def totals_row(bk, key, row, cells, first_col="B", last_col=None,
               label="TOTALS", label_span=None):
    """``cells`` = {field: (formula, kind, cached)}."""
    ws = bk.ws(key)
    th = bk.th
    lab_fmt = bk.S.f(**bk.S.base(font_size=11, bold=True, font_color=th.white,
                                 bg_color=th.primary, align="right",
                                 valign="vcenter", border=1,
                                 border_color=th.primary, indent=1))
    if label_span:
        c1, c2 = label_span
        if c1 == c2:
            ws.write(r(row), ci(c1), label, lab_fmt)
        else:
            ws.merge_range(r(row), ci(c1), r(row), ci(c2), label, lab_fmt)
    else:
        ws.write(r(row), ci(first_col), label, lab_fmt)
    for field, (formula, kind, cached) in cells.items():
        col = ci(bk.col(key, field))
        fmt = bk.S.f(**bk.S.base(
            font_size=11, bold=True, font_color=th.white, bg_color=th.primary,
            align="right" if kind != "center" else "center",
            valign="vcenter", border=1, border_color=th.primary,
            num_format=kind))
        if isinstance(formula, str) and formula.startswith("="):
            ws.write_formula(r(row), col, formula, fmt, cached or 0)
            bk.stats["formulas"] += 1
        else:
            ws.write(r(row), col, formula, fmt)
    if last_col:
        span_end = ci(label_span[1]) if label_span else ci(first_col)
        for c in range(ci(first_col), ci(last_col) + 1):
            letter = cl(c)
            if letter not in [bk.col(key, f) for f in cells] and \
                    c > span_end:
                ws.write_blank(r(row), c, None, bk.S.f(**bk.S.base(
                    bg_color=th.primary, border=1, border_color=th.primary)))
    ws.set_row(r(row), 24)
    return row


def blank_row(bk, key, row, first_col="A", last_col="N", height=8, fmt=None):
    ws = bk.ws(key)
    fmt = fmt or bk.S.canvas
    ws.set_row(r(row), height)
    for c in range(ci(first_col), ci(last_col) + 1):
        ws.write_blank(r(row), c, None, fmt)
    return row


def note_block(bk, key, row, first_col, last_col, lines, title=None):
    """A bordered, wrapped tip box."""
    ws = bk.ws(key)
    width = sum((C.WIDTHS.get(key, {}) or {}).get(cl(c), 10)
                for c in range(ci(first_col), ci(last_col) + 1))
    if title:
        ws.merge_range(r(row), ci(first_col), r(row), ci(last_col), title,
                       bk.S.section_soft)
        ws.set_row(r(row), 22)
        row += 1
    text = "\n".join(lines)
    from ..styles import wrap_height
    h = wrap_height(text, width, minimum=20)
    ws.merge_range(r(row), ci(first_col), r(row), ci(last_col), text,
                   bk.S.note)
    ws.set_row(r(row), h)
    return row + 1
```

### `catering_tracker/sheets/data.py`

```python
"""
The hidden ``_Data`` worksheet - the calculation engine.

Layout (see config.py):

  H..L   rows  2-13   monthly pool  (month, cash in, expenses, profit, events)
  N..O   rows  2-12   expense-category pool   (pie chart)
  N..O   rows 16-24   revenue by event type   (bar chart)
  N..O   rows 28-33   event status counts     (doughnut chart)
  Q..R   rows 26-45   revenue by client
  T..U   rows 26-45   menu popularity (times a dish appears on an event menu)
  AA..AC rows  2-41   upcoming-events pool    (dashboard strip)
  AA..AC rows 42-101  outstanding-payments pool (dashboard strip)
  AE..AF rows  2-36   the KPI table - every headline number in the workbook

Keeping the maths on one hidden sheet keeps the visible tabs clean and the
dashboard formulas readable (``=_Data!$AF$5`` instead of a nine-line
SUMPRODUCT), and gives the charts contiguous, pre-computed ranges.
"""

from .. import config as C
from ..book import r, ci

TODAY = "TODAY()"


def _month_start(m):
    return "DATE(ReportYear,%d,1)" % m


def _month_end(m):
    return "DATE(ReportYear,%d,1)" % (m + 1) if m < 12 \
        else "DATE(ReportYear+1,1,1)"


def build(bk):
    ws = bk.ws("data")
    S, th = bk.S, bk.th
    agg = (bk.demo.agg if bk.demo else {})
    prem = bk.has("inventory")

    hdr = S.f(**S.base(font_size=9, bold=True, font_color=th.white,
                       bg_color=th.muted, align="center", valign="vcenter",
                       border=1, border_color=th.muted))
    txt = S.f(**S.base(font_size=9, font_color=th.ink, align="left"))
    num = S.f(**S.base(font_size=9, num_format="#,##0.00"))
    integer = S.f(**S.base(font_size=9, num_format="0"))
    dt = S.f(**S.base(font_size=9, num_format="dd mmm"))
    pct = S.f(**S.base(font_size=9, num_format="0.0%"))

    ws.set_column("A:A", 14)
    ws.set_column("B:G", 2)
    ws.set_column("H:H", 10)
    ws.set_column("I:L", 12)
    ws.set_column("M:M", 2)
    ws.set_column("N:N", 20)
    ws.set_column("O:O", 12)
    ws.set_column("P:P", 2)
    ws.set_column("Q:Q", 22)
    ws.set_column("R:R", 12)
    ws.set_column("S:S", 2)
    ws.set_column("T:T", 26)
    ws.set_column("U:U", 9)
    ws.set_column("V:Z", 2)
    ws.set_column("AA:AA", 11)
    ws.set_column("AB:AB", 40)
    ws.set_column("AC:AC", 13)
    ws.set_column("AD:AD", 2)
    ws.set_column("AE:AE", 32)
    ws.set_column("AF:AF", 14)

    def fcount():
        bk.stats["formulas"] += 1

    # ------------------------------------------------------------------
    # monthly pool  (H..L rows 2-13) - cash basis, by event month
    # ------------------------------------------------------------------
    inc = bk.q("income")
    exp = bk.q("events")
    inc_recv = bk.rng("income", "received")
    inc_date = bk.rng("income", "event_date")
    inc_amt = bk.rng("income", "amount")
    exp_amt = bk.rng("expenses", "amount")
    exp_date = bk.rng("expenses", "date")
    ev_date = bk.rng("events", "date")
    ev_id = bk.rng("events", "id")

    ws.write(r(1), ci("H"), "Month", hdr)
    for col, label in (("H", "Month"), ("I", "Cash in"), ("J", "Expenses"),
                       ("K", "Profit"), ("L", "Events")):
        ws.write(r(1), ci(col), label, hdr)
    for i in range(12):
        row = C.DATA_MONTH_FIRST + i
        m = i + 1
        cached = agg.get("months", [("", 0, 0, 0, 0)] * 12)[i]
        ws.write(r(row), ci("H"), cached[0] or
                 ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug",
                  "Sep", "Oct", "Nov", "Dec"][i], txt)
        ws.write_formula(
            r(row), ci("I"),
            '=SUMIFS(%s,%s,">="&%s,%s,"<"&%s)'
            % (inc_recv, inc_date, _month_start(m), inc_date, _month_end(m)),
            num, cached[1]); fcount()
        ws.write_formula(
            r(row), ci("J"),
            '=SUMIFS(%s,%s,">="&%s,%s,"<"&%s)'
            % (exp_amt, exp_date, _month_start(m), exp_date, _month_end(m)),
            num, cached[2]); fcount()
        ws.write_formula(r(row), ci("K"), "=I%d-J%d" % (row, row),
                         num, cached[3]); fcount()
        ws.write_formula(
            r(row), ci("L"),
            '=COUNTIFS(%s,">="&%s,%s,"<"&%s,%s,"<>")'
            % (ev_date, _month_start(m), ev_date, _month_end(m), ev_id),
            integer, cached[4]); fcount()

    # ------------------------------------------------------------------
    # expense category pool (N..O rows 2-12) - pie chart
    # ------------------------------------------------------------------
    exp_cat = bk.rng("expenses", "category")
    ws.write(r(1), ci("N"), "Expense category", hdr)
    ws.write(r(1), ci("O"), "Spend", hdr)
    cats = C.EXPENSE_CATEGORIES
    cached_cat = agg.get("exp_cats")
    for i, cat in enumerate(cats):
        row = C.DATA_EXP_FIRST + i
        ws.write(r(row), ci("N"), cat, txt)
        cached = 0.0
        if cached_cat:
            cached = cached_cat.get(cat, 0.0)
        ws.write_formula(
            r(row), ci("O"), '=SUMIF(%s,$N%d,%s)' % (exp_cat, row, exp_amt),
            num, cached); fcount()

    # ------------------------------------------------------------------
    # revenue by event type (N..O rows 16-24)
    # ------------------------------------------------------------------
    ev_type = bk.rng("events", "type")
    ev_price = bk.rng("events", "price")
    ev_status = bk.rng("events", "status")
    ws.write(r(15), ci("N"), "Event type", hdr)
    ws.write(r(15), ci("O"), "Booked value", hdr)
    cached_types = dict(agg.get("types") or [])
    for i, t in enumerate(C.EVENT_TYPES):
        row = C.DATA_TYPE_FIRST + i
        ws.write(r(row), ci("N"), t, txt)
        ws.write_formula(
            r(row), ci("O"),
            '=SUMIFS(%s,%s,$N%d,%s,"<>%s")'
            % (ev_price, ev_type, row, ev_status, C.ES_CANCEL),
            num, cached_types.get(t, 0)); fcount()

    # ------------------------------------------------------------------
    # event status counts (N..O rows 28-33) - doughnut
    # ------------------------------------------------------------------
    ws.write(r(27), ci("N"), "Pipeline stage", hdr)
    ws.write(r(27), ci("O"), "Events", hdr)
    cached_status = agg.get("status_counts")
    for i, st in enumerate(C.EVENT_STATUSES):
        row = C.DATA_STATUS_FIRST + i
        ws.write(r(row), ci("N"), st, txt)
        cached = 0
        if cached_status:
            cached = cached_status.get(st, 0)
        ws.write_formula(
            r(row), ci("O"), '=COUNTIF(%s,$N%d)' % (ev_status, row),
            integer, cached); fcount()

    # ------------------------------------------------------------------
    # revenue by client (Q..R rows 26-45)
    # ------------------------------------------------------------------
    ws.write(r(25), ci("Q"), "Client", hdr)
    ws.write(r(25), ci("R"), "Cash received", hdr)
    ws.write(r(25), ci("T"), "Menu item", hdr)
    ws.write(r(25), ci("U"), "Booked", hdr)
    cl = bk.q("clients")
    cl_first = C.ROW_FIRST
    cached_clients = dict(agg.get("clients_pool") or [])
    cached_menu = dict(agg.get("menu_pool") or [])
    client_names = [c["name"] for c in (bk.demo.clients if bk.demo else [])]
    menu_names = [x["item"] for x in (bk.demo.menu if bk.demo else [])]
    for i in range(20):
        row = C.DATA_CLIENT_FIRST + i
        setup_row = cl_first + i
        cname = client_names[i] if i < len(client_names) else ""
        ws.write_formula(
            r(row), ci("Q"), "=IF(%s!$C$%d=\"\",\"\",%s!$C$%d)"
            % (cl, setup_row, cl, setup_row), txt, cname); fcount()
        ws.write_formula(
            r(row), ci("R"),
            '=IF($Q%d="",0,SUMIF(%s,$Q%d,%s))'
            % (row, bk.rng("income", "client"), row, inc_recv),
            num, cached_clients.get(cname, 0)); fcount()

    # menu popularity (T..U rows 26-45) - premium only
    if bk.has("menu"):
        mn = bk.q("menu")
        ev_menu = bk.rng("events", "menu")
        for i in range(20):
            row = C.DATA_MENU_FIRST + i
            menu_row = cl_first + i
            mname = menu_names[i] if i < len(menu_names) else ""
            ws.write_formula(
                r(row), ci("T"), "=IF(%s!$C$%d=\"\",\"\",%s!$C$%d)"
                % (mn, menu_row, mn, menu_row), txt, mname); fcount()
            ws.write_formula(
                r(row), ci("U"),
                '=IF($T%d="",0,COUNTIF(%s,"*"&$T%d&"*"))'
                % (row, ev_menu, row),
                integer, cached_menu.get(mname, 0)); fcount()
    else:
        for i in range(20):
            row = C.DATA_MENU_FIRST + i
            ws.write(r(row), ci("T"), "", txt)
            ws.write(r(row), ci("U"), "", integer)

    # ------------------------------------------------------------------
    # upcoming-events pool (AA..AC rows 2-41)
    # ------------------------------------------------------------------
    ws.write(r(1), ci("AA"), "Date", hdr)
    ws.write(r(1), ci("AB"), "Event", hdr)
    ws.write(r(1), ci("AC"), "Stage", hdr)
    ws.write(r(41), ci("AA"), "Due", hdr)
    ws.write(r(41), ci("AB"), "Invoice", hdr)
    ws.write(r(41), ci("AC"), "Balance", hdr)
    ev_client = bk.col("events", "client")
    ev_typec = bk.col("events", "type")
    ev_datec = bk.col("events", "date")
    ev_statusc = bk.col("events", "status")
    cached_up = agg.get("upcoming") or []
    for i in range(C.DATA_POOL_ROWS):
        row = C.DATA_POOL_FIRST + i
        er = C.ROW_FIRST + i                     # matching events row
        e = "%s!$%s$%d" % (exp, ev_datec, er)
        st = "%s!$%s$%d" % (exp, ev_statusc, er)
        cl_ref = "%s!$%s$%d" % (exp, ev_client, er)
        ty_ref = "%s!$%s$%d" % (exp, ev_typec, er)
        keep = ('AND(%s<>"",%s>=%s,%s<>"%s",%s<>"%s")'
                % (e, e, TODAY, st, C.ES_CANCEL, st, C.ES_DONE))
        ws.write_formula(r(row), ci("AA"),
                         "=IF(%s,%s,\"\")" % (keep, e), dt,
                         cached_up[i][0] if i < len(cached_up) else ""); fcount()
        ws.write_formula(r(row), ci("AB"),
                         '=IF($AA%d="","",%s&" \u2022 "&%s)' % (row, cl_ref, ty_ref),
                         txt, cached_up[i][1] if i < len(cached_up) else ""); fcount()
        ws.write_formula(r(row), ci("AC"),
                         '=IF($AA%d="","",%s)' % (row, st),
                         txt, cached_up[i][2] if i < len(cached_up) else ""); fcount()

    # outstanding payments pool (AA..AC rows 42-101)
    inc_client = bk.col("income", "client")
    inc_invoice = bk.col("income", "invoice")
    inc_due = bk.col("income", "due")
    inc_bal = bk.col("income", "balance")
    cached_dues = agg.get("dues") or []
    for i in range(C.DATA_DUE_ROWS):
        row = C.DATA_DUE_FIRST + i
        ir = C.ROW_FIRST + i                     # matching income row
        bal = "%s!$%s$%d" % (inc, inc_bal, ir)
        due = "%s!$%s$%d" % (inc, inc_due, ir)
        cli = "%s!$%s$%d" % (inc, inc_client, ir)
        invn = "%s!$%s$%d" % (inc, inc_invoice, ir)
        ws.write_formula(r(row), ci("AA"),
                         '=IF(AND(ISNUMBER(%s),IFERROR(%s*1,0)>0),%s,"")'
                         % (due, bal, due), dt,
                         cached_dues[i][0] if i < len(cached_dues) else ""); fcount()
        ws.write_formula(r(row), ci("AB"),
                         '=IF($AA%d="","",%s&" ("&%s&")")' % (row, cli, invn),
                         txt, cached_dues[i][1] if i < len(cached_dues) else ""); fcount()
        ws.write_formula(r(row), ci("AC"),
                         '=IF($AA%d="","",%s)' % (row, bal),
                         num, cached_dues[i][2] if i < len(cached_dues) else ""); fcount()

    # ------------------------------------------------------------------
    # the KPI table (AE..AF)
    # ------------------------------------------------------------------
    ws.write(r(1), ci(C.KPI_COL_LABEL), "KPI", hdr)
    ws.write(r(1), ci(C.KPI_COL_VALUE), "Value", hdr)
    K = _kpi_formulas(bk, prem)
    for key, label in C.KPI_LABEL.items():
        row = C.KPI_ROW[key]
        ws.write(r(row), ci(C.KPI_COL_LABEL), label, txt)
        formula, cached = K[key]
        fmt = {"#,##0": integer, "#,##0.00": num, "0": integer,
               "0%": pct, "@": txt}[C.KPI_FMT[key]]
        ws.write_formula(r(row), ci(C.KPI_COL_VALUE), formula, fmt,
                         agg.get(key, cached)); fcount()


# ---------------------------------------------------------------------------
# every KPI as (formula, fallback cached value)
# ---------------------------------------------------------------------------
def _kpi_formulas(bk, prem):
    A = lambda k: "$AF$%d" % C.KPI_ROW[k]          # noqa: E731
    inc = bk.q("income")
    recv = bk.rng("income", "received")
    amt = bk.rng("income", "amount")
    bal = bk.rng("income", "balance")
    due = bk.rng("income", "due")
    e_amt = bk.rng("expenses", "amount")
    e_cat = bk.rng("expenses", "category")
    ev = bk.q("events")
    ev_status = bk.rng("events", "status")
    ev_price = bk.rng("events", "price")
    ev_cost = bk.rng("events", "cost")
    ev_guests = bk.rng("events", "guests")
    ev_client = bk.rng("events", "client")
    cl_names = bk.rng("clients", "name")
    q = bk.q("quote")
    F = {}

    F["revenue"] = ("=SUM(%s)" % recv, 0)
    F["expenses"] = ("=SUM(%s)" % e_amt, 0)
    F["profit"] = ("=%s-%s" % (A("revenue"), A("expenses")), 0)
    F["margin"] = ("=IFERROR(%s/%s,0)" % (A("profit"), A("revenue")), 0)
    F["events_total"] = ("=COUNTA(%s)" % bk.rng("events", "id"), 0)
    F["events_done"] = ('=COUNTIF(%s,"%s")' % (ev_status, C.ES_DONE), 0)
    F["events_confirmed"] = (
        '=COUNTIF(%s,"%s")+COUNTIF(%s,"%s")'
        % (ev_status, C.ES_DEPOSIT, ev_status, C.ES_CONFIRMED), 0)
    F["events_cancelled"] = ('=COUNTIF(%s,"%s")' % (ev_status, C.ES_CANCEL), 0)
    F["avg_order"] = (
        '=IFERROR(AVERAGEIF(%s,"<>%s",%s),0)'
        % (ev_status, C.ES_CANCEL, ev_price), 0)
    F["avg_guests"] = (
        '=IFERROR(AVERAGEIF(%s,"<>%s",%s),0)'
        % (ev_status, C.ES_CANCEL, ev_guests), 0)
    F["avg_profit"] = (
        '=IFERROR(AVERAGEIFS(%s,%s,"<>%s"),0)'
        % (bk.rng("events", "profit"), ev_status, C.ES_CANCEL), 0)
    F["guests_total"] = (
        '=SUMIF(%s,"%s",%s)' % (ev_status, C.ES_DONE, ev_guests), 0)
    F["outstanding"] = ("=SUMIF(%s,\">0\")" % bal, 0)
    F["overdue"] = (
        '=SUMPRODUCT((%s>0)*(%s<>"")*(%s<TODAY()))' % (bal, due, due), 0)
    F["invoices_open"] = ('=COUNTIF(%s,">0")' % bal, 0)
    F["food_cost"] = (
        '=SUMIF(%s,"Ingredients",%s)+SUMIF(%s,"Packaging",%s)'
        % (e_cat, e_amt, e_cat, e_amt), 0)
    F["labor_cost"] = ('=SUMIF(%s,"Staff / Labor",%s)' % (e_cat, e_amt), 0)
    F["food_pct"] = ("=IFERROR(%s/%s,0)" % (A("food_cost"), A("revenue")), 0)
    F["labor_pct"] = ("=IFERROR(%s/%s,0)" % (A("labor_cost"), A("revenue")), 0)
    F["repeat_pct"] = (
        '=IFERROR(SUMPRODUCT((%s<>"")*(COUNTIF(%s,%s)>1))'
        '/COUNTIF(%s,"?*"),0)'
        % (cl_names, ev_client, cl_names, cl_names), 0)

    if prem:
        inv_val = bk.rng("inventory", "value")
        inv_status = bk.rng("inventory", "status")
        F["inv_value"] = ("=SUM(%s)" % inv_val, 0)
        F["low_stock"] = ('=COUNTIF(%s,"\U0001F534 Reorder")' % inv_status, 0)
        s_buy = bk.rng("shopping", "to_buy")
        s_tick = bk.rng("shopping", "purchased")
        s_cost = bk.rng("shopping", "est_cost")
        F["shop_lines"] = (
            '=COUNTIFS(%s,">0",%s,"<>%s")' % (s_buy, s_tick, C.TICK), 0)
        F["shop_cost"] = (
            '=SUMIF(%s,"<>%s",%s)' % (s_tick, C.TICK, s_cost), 0)
        st_name = bk.rng("staff", "name")
        st_paid = bk.rng("staff", "paid")
        F["staff_unpaid"] = (
            '=SUMPRODUCT((%s<>"")*(%s<>"%s"))' % (st_name, st_paid, C.TICK), 0)
        F["equip_value"] = ("=SUM(%s)" % bk.rng("equipment", "value"), 0)
        eq_status = bk.rng("equipment", "status")
        F["equip_service"] = (
            '=COUNTIF(%s,"\U0001F527 Service")+COUNTIF(%s,"\U0001F7E1 Soon")'
            % (eq_status, eq_status), 0)
        F["menu_items"] = ("=COUNTA(%s)" % bk.rng("menu", "item"), 0)
        F["menu_margin"] = (
            "=IFERROR(AVERAGE(%s),0)" % bk.rng("menu", "margin"), 0)
        F["quote_price"] = (
            "=%s!$J$%d" % (q, C.QUOTE_OUT["price"]), "")
        F["quote_per_guest"] = (
            "=%s!$J$%d" % (q, C.QUOTE_OUT["per_guest"]), "")
    else:
        for k in ("inv_value", "low_stock", "shop_lines", "shop_cost",
                  "staff_unpaid", "equip_value", "equip_service",
                  "menu_items", "menu_margin"):
            F[k] = ("=0", 0)
        F["quote_price"] = ("=0", 0)
        F["quote_per_guest"] = ("=0", 0)

    F["tax_collected"] = ("=ROUND(SUM(%s)*TaxRate,2)" % amt, 0)
    if bk.has("tax"):
        t = bk.q("tax")
        F["tax_due"] = ("=%s!$G$%d" % (t, C.TAX_SUM["due"]), 0)
    else:
        F["tax_due"] = ("=%s" % A("tax_collected"), 0)

    n_first, n_last = C.DATA_TYPE_FIRST, C.DATA_TYPE_FIRST \
        + len(C.EVENT_TYPES) - 1
    F["best_type"] = (
        '=IF(MAX($O$%d:$O$%d)<=0,"-",INDEX($N$%d:$N$%d,'
        'MATCH(MAX($O$%d:$O$%d),$O$%d:$O$%d,0)))'
        % (n_first, n_last, n_first, n_last, n_first, n_last,
           n_first, n_last), "-")
    c_first, c_last = C.DATA_CLIENT_FIRST, C.DATA_CLIENT_FIRST + 19
    F["best_client"] = (
        '=IF(MAX($R$%d:$R$%d)<=0,"-",INDEX($Q$%d:$Q$%d,'
        'MATCH(MAX($R$%d:$R$%d),$R$%d:$R$%d,0)))'
        % (c_first, c_last, c_first, c_last, c_first, c_last,
           c_first, c_last), "-")
    return F
```

### `catering_tracker/sheets/setup.py`

```python
"""
The ⚙️ Setup tab: business profile, money defaults, the calendar month and
every editable dropdown list in the workbook.

Everything typed here flows through defined names (BusinessName, Currency,
TaxRate, DefaultMargin, DepositPct, DueSoonDays, ReportYear, CalMonth,
CalYear + the six list names), so a change here updates every tab.
"""

from .. import config as C
from ..book import r, ci

LAST_COL = "J"


def build(bk):
    key = "setup"
    ws = bk.ws(key)
    S, th, m = bk.S, bk.th, bk.demo
    st = m.settings

    bk.widths(key, {"A": 2.2, "B": 34, "C": 30, "D": 16, "E": 16, "F": 16,
                    "G": 18, "H": 16, "I": 16, "J": 16})
    bk.paint(key, 0, 0, 70, ci(LAST_COL), S.canvas)

    # ------------------------------------------------------------------
    # title band
    # ------------------------------------------------------------------
    ws.set_row(r(1), 7)
    ws.set_row(r(2), 32)
    ws.set_row(r(3), 18)
    ws.merge_range(r(2), 1, r(2), ci(LAST_COL) - 3,
                   "\u2699\uFE0F  Setup \u2014 make it your business",
                   S.sheet_title)
    ws.merge_range(r(3), 1, r(3), ci(LAST_COL) - 3,
                   "  Set these once and every tab, quote, alert and "
                   "dropdown follows automatically.", S.sheet_sub)
    ws.merge_range(r(3), ci(LAST_COL) - 2, r(3), ci(LAST_COL), "",
                   S.home_link)
    ws.write_url(r(3), ci(LAST_COL) - 2,
                 "internal:%s!A1" % bk.q("dashboard"), S.home_link,
                 "\U0001F3E0  Back to Dashboard")
    bk.stats["links"] += 1
    ws.set_row(r(4), 7)

    label = S.f(**S.base(font_size=11, bold=True, font_color=th.ink,
                         bg_color=th.card, align="left", valign="vcenter",
                         border=1, border_color=th.border, indent=1))
    note_f = S.f(**S.base(font_size=9.5, italic=True, font_color=th.muted,
                          bg_color=th.card, align="left", valign="top",
                          text_wrap=True,
                          border=1, border_color=th.border, indent=1))
    input_f = S.f(**S.base(font_size=11, bold=True, font_color=th.primary,
                           bg_color=th.gold_soft, align="center",
                           valign="vcenter", border=2, border_color=th.gold,
                           locked=False))

    def fmt(num_format=None):
        props = dict(font_size=11, bold=True, font_color=th.primary,
                     bg_color=th.gold_soft, align="center",
                     valign="vcenter", border=2, border_color=th.gold,
                     locked=False)
        if num_format:
            props["num_format"] = num_format
        return S.f(**S.base(**props))

    def setting(row, text, note, value, cell_fmt, height=26):
        ws.set_row(r(row), height)
        ws.write(r(row), ci("B"), text, label)
        ws.write(r(row), ci("C"), value, cell_fmt)
        ws.merge_range(r(row), ci("D"), r(row), ci(LAST_COL), note, note_f)

    def section(row, text, style=None):
        ws.set_row(r(row), 24)
        ws.merge_range(r(row), 1, r(row), ci(LAST_COL), text,
                       style or S.section)

    def decimal(row, lo, hi, title, message):
        ws.data_validation(r(row), ci("C"), r(row), ci("C"), {
            "validate": "decimal", "criteria": "between",
            "minimum": lo, "maximum": hi, "ignore_blank": True,
            "input_title": title, "input_message": message,
            "show_input": True, "show_error": True,
            "error_message": message})
        bk.stats["validations"] += 1

    # ------------------------------------------------------------------
    # 1. business profile
    # ------------------------------------------------------------------
    section(5, "  \U0001F468\u200D\U0001F373  YOUR BUSINESS")
    setting(C.SU_BUSINESS, "Business name",
            "Shown on the dashboard hero and on the printable invoice.",
            st["business"], input_f, height=28)
    setting(C.SU_MESSAGE, "Dashboard message / focus",
            "One line under the dashboard title: this week's priority, a "
            "reminder, or a little motivation.", st["message"], input_f)
    ws.set_row(r(8), 8)

    # ------------------------------------------------------------------
    # 2. money defaults
    # ------------------------------------------------------------------
    section(9, "  \U0001F4B0  MONEY DEFAULTS")
    setting(C.SU_CURRENCY, "Currency symbol",
            "Every summary readout prefixes amounts with this. To show it "
            "inside the tables too: select the money columns \u2192 Home "
            "\u2192 Number format \u2192 Currency.",
            st["currency"], input_f)
    bk.validate(key, C.SU_CURRENCY, ci("C"), C.SU_CURRENCY, ci("C"),
                '="%s"' % ",".join(C.CURRENCIES),
                title="Currency", message="Pick your currency symbol.",
                error="Choose one of the listed symbols.")
    setting(C.SU_TAX, "Tax rate (sales tax / GST)",
            "Applied on the \U0001F9FE Tax tab and the printable invoice. "
            "Use 0% if you are not tax registered.",
            st["tax"], fmt("0.0%"))
    decimal(C.SU_TAX, 0, 1, "Tax rate",
            "A decimal between 0 and 1 (5% = 0.05).")
    setting(C.SU_MARGIN, "Default target margin",
            "The \U0001F9EE Quote Calculator starts from this. A 35% margin "
            "means price = total cost \u00F7 0.65.",
            st["margin"], fmt("0%"))
    decimal(C.SU_MARGIN, 0, 0.95, "Margin",
            "Between 0% and 95% (0.35 = 35%).")
    setting(C.SU_DEPOSIT, "Default deposit %",
            "Suggested deposit on quotes and invoices.",
            st["deposit"], fmt("0%"))
    decimal(C.SU_DEPOSIT, 0, 1, "Deposit %",
            "Between 0% and 100% (0.4 = 40%).")
    setting(C.SU_DUESOON, "\u201CDue soon\u201D window (days)",
            "Payment due dates inside this many days turn orange; past-due "
            "turns red everywhere.", st["duesoon"], fmt("0"))
    decimal(C.SU_DUESOON, 1, 90, "Days", "A number of days, 1-90.")
    setting(C.SU_YEAR, "Reporting year",
            "The \U0001F4C8 P&L tab and the monthly charts report on this "
            "year.", st["year"], fmt("0"))
    ws.set_row(r(16), 8)

    # ------------------------------------------------------------------
    # 3. calendar month
    # ------------------------------------------------------------------
    section(17, "  \U0001F4C6  EVENT CALENDAR MONTH")
    setting(C.SU_CAL_MONTH, "Calendar month (1-12)",
            "Which month the \U0001F4C6 Event Calendar tab displays "
            "(9 = September).", st["cal_month"], fmt("0"))
    bk.validate(key, C.SU_CAL_MONTH, ci("C"), C.SU_CAL_MONTH, ci("C"),
                '"1,2,3,4,5,6,7,8,9,10,11,12"', title="Month",
                message="A number from 1 (January) to 12 (December).")
    setting(C.SU_CAL_YEAR, "Calendar year", "", st["cal_year"], fmt("0"))
    ws.set_row(r(20), 8)

    # ------------------------------------------------------------------
    # 4. editable lists
    # ------------------------------------------------------------------
    section(C.SU_LIST_HEADER - 1,
            "  \U0001F4CB  YOUR LISTS  \u2014  edit any time; every "
            "dropdown in the workbook updates itself")
    ws.set_row(r(C.SU_LIST_HEADER), 30)
    for lk, col in sorted(C.LIST_COLS.items(), key=lambda kv: kv[1]):
        ws.write(r(C.SU_LIST_HEADER), ci(col), C.LIST_TITLES[lk],
                 S.header(th.primary))
    values = {
        "event_types": C.EVENT_TYPES,
        "expense_categories": C.EXPENSE_CATEGORIES,
        "payment_methods": C.PAYMENT_METHODS,
        "staff_roles": C.STAFF_ROLES,
        "menu_categories": C.MENU_CATEGORIES,
        "ingredient_categories": C.INGREDIENT_CATEGORIES,
    }
    list_fmt = S.f(**S.base(font_size=10, font_color=th.ink, bg_color=th.card,
                            align="left", valign="vcenter", border=1,
                            border_color=th.border, indent=1, locked=False))
    list_fmt_alt = S.f(**S.base(font_size=10, font_color=th.ink,
                                bg_color=th.alt, align="left",
                                valign="vcenter", border=1,
                                border_color=th.border, indent=1,
                                locked=False))
    for i in range(C.SU_LIST_ROWS):
        row = C.SU_LIST_FIRST + i
        ws.set_row(r(row), 18)
        for lk, col in C.LIST_COLS.items():
            vals = values[lk]
            value = vals[i] if i < len(vals) else ""
            ws.write(r(row), ci(col), value,
                     list_fmt_alt if i % 2 else list_fmt)
    note_row = C.SU_LIST_FIRST + C.SU_LIST_ROWS
    ws.set_row(r(note_row), 26)
    ws.set_row(r(note_row + 1), 26)
    ws.merge_range(r(note_row), 1, r(note_row + 1), ci(LAST_COL),
                   "  \u2022  Type over the defaults with your own words "
                   "\u2014 add as many as you like (20 rows each).\n"
                   "  \u2022  Don't leave a blank row in the middle of a "
                   "list: the dropdown stops at the first gap.\n"
                   "  \u2022  Deleting a word here does NOT delete your "
                   "events or expenses \u2014 it only removes it from the "
                   "dropdown.", S.note)
    ws.set_row(r(note_row + 2), 8)

    # ------------------------------------------------------------------
    # 5. fixed lists (locked)
    # ------------------------------------------------------------------
    section(C.SU_FIXED_HEADER - 1,
            "  \U0001F512  FIXED LISTS  \u2014  colours, counts and "
            "dashboard maths depend on this exact wording, so they are "
            "locked", S.section_accent)
    ws.set_row(r(C.SU_FIXED_HEADER), 26)
    fixed_titles = {"event_statuses": "Event pipeline",
                    "payment_statuses": "Payment status",
                    "tick": "Tick box"}
    for fk, col in sorted(C.FIXED_COLS.items(), key=lambda kv: kv[1]):
        ws.write(r(C.SU_FIXED_HEADER), ci(col), fixed_titles[fk],
                 S.header(th.accent))
    fixed_values = {"event_statuses": C.EVENT_STATUSES,
                    "payment_statuses": C.PAYMENT_STATUSES,
                    "tick": [C.TICK, ""]}
    lock_fmt = S.f(**S.base(font_size=10, font_color=th.muted,
                            bg_color=th.primary_soft, align="left",
                            valign="vcenter", border=1,
                            border_color=th.border, indent=1, locked=True))
    for fk, col in C.FIXED_COLS.items():
        vals = fixed_values[fk]
        for i in range(6):
            row = C.SU_FIXED_FIRST + i
            ws.set_row(r(row), 18)
            ws.write(r(row), ci(col), vals[i] if i < len(vals) else "",
                     lock_fmt)

    # ------------------------------------------------------------------
    # footer
    # ------------------------------------------------------------------
    foot = C.SU_FIXED_FIRST + 7
    ws.set_row(r(foot), 30)
    ws.merge_range(r(foot), 1, r(foot), ci(LAST_COL),
                   "  \U0001F4A1  Everything else is formula-driven: you "
                   "never have to touch a formula, only these settings and "
                   "the white input cells on each tab.",
                   S.f(**S.base(font_size=10, italic=True, font_color=th.ink,
                                bg_color=th.gold_soft, align="left",
                                valign="vcenter", indent=1)))
    nav_row = foot + 2
    ws.set_row(r(nav_row), 24)
    bk.nav_row(key, nav_row, max_col=LAST_COL)

    bk.page(key, LAST_COL, nav_row + 1, landscape=False,
            freeze=(4, 0), fit=True, zoom=100)
```

### `catering_tracker/sheets/dashboard.py`

```python
"""
\U0001F4CA Dashboard - the command centre.

Hero band (business name, next-event countdown, today's summary), two rows
of KPI cards, two progress bars, the six-stage pipeline strip, four charts
(monthly cash-in vs expenses vs profit, expense breakdown, pipeline
doughnut, revenue by event type) and the live "next events" / "money
coming in" panels.

Every number is a formula pointing at the hidden _Data sheet, so the
dashboard is always in sync with the tabs - and every formula ships with a
cached value so the EXAMPLE workbooks look perfect before Excel even
recalculates.
"""

from .. import config as C
from ..book import r, ci
from .events import STATUS_COLORS

KEY = "dashboard"
LAST_COL = "M"
CARDS = [(1, 2), (3, 4), (5, 6), (7, 8), (9, 10), (11, 12)]

ROW_HERO_1 = 2
ROW_HERO_2 = 3
ROW_HERO_3 = 4
ROW_MSG = 5
ROW_SEC_1 = 7
ROW_LBL_1 = 8
ROW_VAL_1 = 9
ROW_LBL_2 = 10
ROW_VAL_2 = 11
ROW_BAR_1 = 12
ROW_BAR_2 = 13
ROW_PIPE_SEC = 15
ROW_PIPE = 16
ROW_CHART_1 = 18
ROW_CHART_2 = 34
ROW_PANEL_SEC = 50
ROW_PANEL = 51
PANEL_ROWS = 6
ROW_FOOT = ROW_PANEL + PANEL_ROWS + 1


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo
    agg = m.agg
    d = bk.q("data")
    prem = bk.has("inventory")

    ws.set_column("A:A", 2.2)
    ws.set_column("B:M", 13)
    bk.paint(KEY, 0, 0, ROW_FOOT + 4, ci(LAST_COL), S.canvas)

    # ------------------------------------------------------------------
    # hero
    # ------------------------------------------------------------------
    ws.set_row(r(1), 8)
    ws.set_row(r(ROW_HERO_1), 42)
    ws.set_row(r(ROW_HERO_2), 32)
    ws.set_row(r(ROW_HERO_3), 20)
    ws.set_row(r(ROW_MSG), 22)

    ws.merge_range(r(ROW_HERO_1), 1, r(ROW_HERO_1), ci(LAST_COL), "",
                   S.hero_title)
    ws.write_formula(
        r(ROW_HERO_1), 1,
        '=IF(BusinessName="","\U0001F37D\uFE0F  Catering Command Center",'
        '"\U0001F37D\uFE0F  "&BusinessName&"   \u2022   Catering Command '
        'Center")', S.hero_title,
        "\U0001F37D\uFE0F  Catering Command Center"
        if not m.settings["business"] else
        "\U0001F37D\uFE0F  %s   \u2022   Catering Command Center"
        % m.settings["business"])
    bk.stats["formulas"] += 1

    ws.merge_range(r(ROW_HERO_2), 1, r(ROW_HERO_2), ci(LAST_COL), "",
                   S.hero_count)
    nxt_formula = _next_event(bk)
    ws.write_formula(r(ROW_HERO_2), 1, "=" + nxt_formula, S.hero_count,
                     _next_event_cached(m))
    bk.stats["formulas"] += 1

    ws.merge_range(r(ROW_HERO_3), 1, r(ROW_HERO_3), ci(LAST_COL), "",
                   S.hero_meta)
    meta = ('=TEXT(TODAY(),"dddd, dd mmmm yyyy")&"   \u2022   "&%s&'
            '" events on the books   \u2022   "&Currency&TEXT(%s,"#,##0")&'
            '" cash received   \u2022   "&Currency&TEXT(%s,"#,##0")&'
            '" still outstanding"'
            % (bk.kpi("events_total"), bk.kpi("revenue"),
               bk.kpi("outstanding")))
    ws.write_formula(
        r(ROW_HERO_3), 1, meta, S.hero_meta,
        "%s   \u2022   %d events on the books   \u2022   %s cash received"
        "   \u2022   %s still outstanding"
        % (_today_text(), agg.get("events_total", 0),
           m.money(agg.get("revenue", 0)), m.money(agg.get("outstanding", 0))))
    bk.stats["formulas"] += 1

    # message strip
    msg_fmt = S.f(**S.base(font_size=11, bold=True, italic=True,
                           font_color=th.white, bg_color=th.gold,
                           align="left", valign="vcenter", indent=1))
    ws.merge_range(r(ROW_MSG), 1, r(ROW_MSG), ci(LAST_COL), "", msg_fmt)
    ws.write_formula(
        r(ROW_MSG), 1,
        '=IF(%s!$C$%d="","","\U0001F4E3  "&%s!$C$%d)'
        % (bk.q("setup"), C.SU_MESSAGE, bk.q("setup"), C.SU_MESSAGE),
        msg_fmt,
        ("\U0001F4E3  " + m.settings["message"]) if m.settings["message"]
        else "")
    bk.stats["formulas"] += 1
    ws.set_row(r(6), 8)

    # ------------------------------------------------------------------
    # KPI cards
    # ------------------------------------------------------------------
    _section(ws, S, ROW_SEC_1, "  \U0001F4B0  BUSINESS AT A GLANCE")
    cards1 = [
        ("CASH RECEIVED", "ok",
         '=Currency&TEXT(%s,"#,##0")' % bk.kpi("revenue"),
         m.money(agg.get("revenue", 0))),
        ("TOTAL EXPENSES", "accent",
         '=Currency&TEXT(%s,"#,##0")' % bk.kpi("expenses"),
         m.money(agg.get("expenses", 0))),
        ("NET PROFIT", "primary",
         '=Currency&TEXT(%s,"#,##0")' % bk.kpi("profit"),
         m.money(agg.get("profit", 0))),
        ("NET MARGIN", "gold", "=%s" % bk.kpi("margin"),
         agg.get("margin", 0)),
        ("EVENTS ON BOOKS", "info", "=%s" % bk.kpi("events_total"),
         agg.get("events_total", 0)),
        ("AVG ORDER VALUE", "primary_2",
         '=Currency&TEXT(%s,"#,##0")' % bk.kpi("avg_order"),
         m.money(agg.get("avg_order", 0))),
    ]
    if prem:
        cards2 = [
            ("UPCOMING BOOKED", "primary",
             "=%s" % bk.kpi("events_confirmed"), agg.get("events_confirmed", 0)),
            ("OUTSTANDING", "warn",
             '=Currency&TEXT(%s,"#,##0")' % bk.kpi("outstanding"),
             m.money(agg.get("outstanding", 0))),
            ("OVERDUE INVOICES", "bad", "=%s" % bk.kpi("overdue"),
             agg.get("overdue", 0)),
            ("LOW-STOCK ITEMS", "bad", "=%s" % bk.kpi("low_stock"),
             agg.get("low_stock", 0)),
            ("UNPAID SHIFTS", "plum", "=%s" % bk.kpi("staff_unpaid"),
             agg.get("staff_unpaid", 0)),
            ("INVENTORY VALUE", "info",
             '=Currency&TEXT(%s,"#,##0")' % bk.kpi("inv_value"),
             m.money(agg.get("inv_value", 0))),
        ]
    else:
        cards2 = [
            ("UPCOMING BOOKED", "primary",
             "=%s" % bk.kpi("events_confirmed"), agg.get("events_confirmed", 0)),
            ("OUTSTANDING", "warn",
             '=Currency&TEXT(%s,"#,##0")' % bk.kpi("outstanding"),
             m.money(agg.get("outstanding", 0))),
            ("OVERDUE INVOICES", "bad", "=%s" % bk.kpi("overdue"),
             agg.get("overdue", 0)),
            ("OPEN INVOICES", "info", "=%s" % bk.kpi("invoices_open"),
             agg.get("invoices_open", 0)),
            ("REPEAT CUSTOMERS", "gold",
             '=TEXT(%s,"0%%")' % bk.kpi("repeat_pct"),
             agg.get("repeat_pct", 0)),
            ("GUESTS CATERED", "primary_2", "=%s" % bk.kpi("guests_total"),
             agg.get("guests_total", 0)),
        ]
    _cards(bk, ROW_LBL_1, ROW_VAL_1, cards1,
           num_formats=[None, None, None, "0%", "0", None])
    _cards(bk, ROW_LBL_2, ROW_VAL_2, cards2,
           num_formats=["0", None, "0", "0", "0", None])

    # ------------------------------------------------------------------
    # progress bars
    # ------------------------------------------------------------------
    bars = [
        (ROW_BAR_1, "  Net margin", "profit", "revenue", "margin",
         agg.get("margin", 0)),
        (ROW_BAR_2, "  Food cost % of cash in", "food_cost", "revenue",
         "food_pct", agg.get("food_pct", 0)),
    ]
    for row, label, num_k, den_k, pct_k, cached in bars:
        ws.set_row(r(row), 22)
        ws.merge_range(r(row), 1, r(row), 2, label, S.bar_label)
        ws.merge_range(r(row), 3, r(row), 9, "", S.bar_text)
        ws.write_formula(r(row), 3,
                         bk.bar(bk.kpi(num_k), bk.kpi(den_k), 34),
                         S.bar_text, bk.bar_static(cached, 1, 34))
        ws.merge_range(r(row), 10, r(row), ci(LAST_COL), "", S.bar_pct)
        ws.write_formula(r(row), 10,
                         '=TEXT(%s,"0.0%%")' % bk.kpi(pct_k),
                         S.bar_pct, "%.1f%%" % (cached * 100))
        bk.stats["formulas"] += 2
    ws.set_row(r(14), 8)

    # ------------------------------------------------------------------
    # pipeline strip
    # ------------------------------------------------------------------
    _section(ws, S, ROW_PIPE_SEC, "  \U0001F4C5  THE PIPELINE",
             style=S.section_accent)
    ws.set_row(r(ROW_PIPE), 30)
    ev_status = bk.rng("events", "status")
    counts = (agg.get("status_counts") or {})
    col = 1
    for st in C.EVENT_STATUSES:
        bg, fg = STATUS_COLORS[st]
        fmt = S.pill(getattr(th, bg), getattr(th, fg), size=11, bold=True,
                     align="center")
        ws.merge_range(r(ROW_PIPE), col, r(ROW_PIPE), col + 1, "", fmt)
        ws.write_formula(
            r(ROW_PIPE), col,
            '="%s: "&COUNTIF(%s,"%s")' % (st, ev_status, st), fmt,
            "%s: %d" % (st, counts.get(st, 0)))
        bk.stats["formulas"] += 1
        col += 2

    # ------------------------------------------------------------------
    # charts
    # ------------------------------------------------------------------
    _charts(bk)

    # ------------------------------------------------------------------
    # live panels: next events + money coming in
    # ------------------------------------------------------------------
    _panels(bk)

    # ------------------------------------------------------------------
    # footer
    # ------------------------------------------------------------------
    foot = ROW_FOOT + 1
    ws.set_row(r(foot), 30)
    ws.merge_range(r(foot), 1, r(foot), ci(LAST_COL),
                   "  \U0001F4A1  Everything on this page is calculated "
                   "\u2014 type your business into the tabs and watch it "
                   "update.  Start with the \U0001F4D6 Start Here guide, "
                   "set up \u2699\uFE0F Setup, then log \U0001F4C5 Events "
                   "and \U0001F4B0 Payments.",
                   S.f(**S.base(font_size=10, italic=True, font_color=th.ink,
                                bg_color=th.gold_soft, align="left",
                                valign="vcenter", indent=1)))
    nav = foot + 2
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)
    bk.page(KEY, LAST_COL, nav + 1, landscape=False, zoom=85)


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def _section(ws, S, row, text, style=None):
    ws.set_row(r(row), 22)
    ws.merge_range(r(row), 1, r(row), ci(LAST_COL), text, style or S.section)


def _cards(bk, label_row, value_row, cards, num_formats=None):
    ws, S, th = bk.ws(KEY), bk.S, bk.th
    num_formats = num_formats or [None] * len(cards)
    ws.set_row(r(label_row), 18)
    ws.set_row(r(value_row), 36)
    for i, (label, color, formula, cached) in enumerate(cards):
        c1, c2 = CARDS[i]
        lbl_fmt = S.kpi_label(getattr(th, color), size=9.5, align="center")
        ws.merge_range(r(label_row), c1, r(label_row), c2, label, lbl_fmt)
        val_fmt = S.kpi_value(getattr(th, color),
                              num_format=num_formats[i], size=20)
        ws.merge_range(r(value_row), c1, r(value_row), c2, "", val_fmt)
        ws.write_formula(r(value_row), c1, formula, val_fmt, cached)
        bk.stats["formulas"] += 1


def _next_event(bk):
    """Hero line: next upcoming event + countdown."""
    d = bk.q("data")
    pool_d = "%s!$AA$%d:$AA$%d" % (d, C.DATA_POOL_FIRST,
                                   C.DATA_POOL_FIRST + C.DATA_POOL_ROWS - 1)
    pool_l = "%s!$AB$%d:$AB$%d" % (d, C.DATA_POOL_FIRST,
                                   C.DATA_POOL_FIRST + C.DATA_POOL_ROWS - 1)
    nxt = "IFERROR(SMALL(%s,1),\"\")" % pool_d
    label = 'IFERROR(INDEX(%s,MATCH(%s,%s,0)),"")' % (pool_l, nxt, pool_d)
    return ('IF(%s="","\U0001F334  No upcoming events \u2014 add one on the '
            '\U0001F4C5 Events tab!","\U0001F525  Next event: "&%s&"  '
            '\u2014  in "&(%s-TODAY())&" day(s)")' % (nxt, label, nxt))


def _next_event_cached(m):
    up = m.agg.get("upcoming") or []
    if not up:
        return ("\U0001F334  No upcoming events \u2014 add one on the "
                "\U0001F4C5 Events tab!")
    from ..demo import TODAY
    when, label, _st = up[0]
    return "\U0001F525  Next event: %s  \u2014  in %d day(s)" % (
        label, (when - TODAY).days)


def _today_text():
    from ..demo import TODAY
    return TODAY.strftime("%A, %d %B %Y")


def _charts(bk):
    ws, th = bk.ws(KEY), bk.th
    d = bk.q("data")
    mf, ml = C.DATA_MONTH_FIRST, C.DATA_MONTH_FIRST + 11
    cats = "=%s!$H$%d:$H$%d" % (d, mf, ml)

    # 1 - monthly combo: cash in / expenses columns + profit line
    ch = bk.chart("column")
    ch.add_series({
        "name": "Cash in", "categories": cats,
        "values": "=%s!$I$%d:$I$%d" % (d, mf, ml),
        "fill": {"color": th.primary}, "border": {"color": th.primary},
        "gap": 60})
    ch.add_series({
        "name": "Expenses", "categories": cats,
        "values": "=%s!$J$%d:$J$%d" % (d, mf, ml),
        "fill": {"color": th.accent}, "border": {"color": th.accent}})
    line = bk.chart("line")
    line.add_series({
        "name": "Profit", "categories": cats,
        "values": "=%s!$K$%d:$K$%d" % (d, mf, ml),
        "line": {"color": th.ok, "width": 2.5},
        "marker": {"type": "circle", "size": 5, "fill": {"color": th.ok}}})
    ch.combine(line)
    ch.set_title({"name": "Money by month (reporting year on Setup)",
                  "name_font": {"size": 12, "bold": True,
                                "color": th.primary, "name": th.title_font}})
    ch.set_y_axis({"num_format": "#,##0", "num_font": {"size": 9},
                   "major_gridlines": {"visible": True,
                                       "line": {"color": th.border}}})
    ch.set_x_axis({"num_font": {"size": 9}})
    ch.set_size({"width": 555, "height": 285})
    ch.set_legend({"position": "bottom", "font": {"size": 9}})
    ws.insert_chart(r(ROW_CHART_1), ci("B"), ch, {"x_offset": 8,
                                                  "y_offset": 6})

    # 2 - expense breakdown doughnut
    ef, el = C.DATA_EXP_FIRST, C.DATA_EXP_FIRST + len(C.EXPENSE_CATEGORIES) - 1
    ch = bk.chart("doughnut")
    palette = [th.primary, th.accent, th.gold, th.info, th.plum, th.ok,
               th.primary_2, th.warn, th.bad, "#8C7B6B", "#6B8C7B"]
    ch.add_series({
        "name": "Expenses",
        "categories": "=%s!$N$%d:$N$%d" % (d, ef, el),
        "values": "=%s!$O$%d:$O$%d" % (d, ef, el),
        "points": [{"fill": {"color": palette[i % len(palette)]}}
                   for i in range(el - ef + 1)],
        "data_labels": {"percentage": True,
                        "font": {"size": 8, "color": th.white}},
    })
    ch.set_title({"name": "Where the money goes",
                  "name_font": {"size": 12, "bold": True,
                                "color": th.primary, "name": th.title_font}})
    ch.set_hole_size(58)
    ch.set_size({"width": 555, "height": 285})
    ch.set_legend({"position": "right", "font": {"size": 8}})
    ws.insert_chart(r(ROW_CHART_1), ci("H"), ch, {"x_offset": 8,
                                                  "y_offset": 6})

    # 3 - pipeline doughnut
    sf, sl = C.DATA_STATUS_FIRST, C.DATA_STATUS_FIRST + 5
    ch = bk.chart("doughnut")
    ch.add_series({
        "name": "Pipeline",
        "categories": "=%s!$N$%d:$N$%d" % (d, sf, sl),
        "values": "=%s!$O$%d:$O$%d" % (d, sf, sl),
        "points": [{"fill": {"color": c}} for c in
                   [th.info, th.plum, th.warn, th.primary, th.ok, th.bad]],
        "data_labels": {"value": True,
                        "font": {"size": 9, "color": th.white}},
    })
    ch.set_title({"name": "Event pipeline",
                  "name_font": {"size": 12, "bold": True,
                                "color": th.primary, "name": th.title_font}})
    ch.set_hole_size(58)
    ch.set_size({"width": 555, "height": 285})
    ch.set_legend({"position": "right", "font": {"size": 8}})
    ws.insert_chart(r(ROW_CHART_2), ci("B"), ch, {"x_offset": 8,
                                                  "y_offset": 6})

    # 4 - revenue by event type
    tf, tl = C.DATA_TYPE_FIRST, C.DATA_TYPE_FIRST + len(C.EVENT_TYPES) - 1
    ch = bk.chart("bar")
    ch.add_series({
        "name": "Booked value",
        "categories": "=%s!$N$%d:$N$%d" % (d, tf, tl),
        "values": "=%s!$O$%d:$O$%d" % (d, tf, tl),
        "fill": {"color": th.gold}, "border": {"color": th.gold},
        "gap": 45})
    ch.set_title({"name": "Revenue by event type",
                  "name_font": {"size": 12, "bold": True,
                                "color": th.primary, "name": th.title_font}})
    ch.set_x_axis({"num_format": "#,##0", "num_font": {"size": 9},
                   "major_gridlines": {"visible": True,
                                       "line": {"color": th.border}}})
    ch.set_y_axis({"num_font": {"size": 8}, "label_position": "low"})
    ch.set_size({"width": 555, "height": 285})
    ch.set_legend({"none": True})
    ws.insert_chart(r(ROW_CHART_2), ci("H"), ch, {"x_offset": 8,
                                                  "y_offset": 6})


def _panels(bk):
    ws, S, th, m = bk.ws(KEY), bk.S, bk.th, bk.demo
    d = bk.q("data")
    row = ROW_PANEL_SEC
    ws.set_row(r(row), 22)
    ws.merge_range(r(row), 1, r(row), 5,
                   "  \U0001F4C5  NEXT EVENTS", S.section_soft)
    ws.merge_range(r(row), 6, r(row), ci(LAST_COL),
                   "  \U0001F4B5  MONEY COMING IN", S.section_soft)

    def mk(**props):
        base = dict(font_size=10, font_color=th.ink, bg_color=th.card,
                    align="left", valign="vcenter", border=1,
                    border_color=th.border, indent=1)
        base.update(props)
        alt_props = dict(base)
        alt_props["bg_color"] = th.alt
        return (S.f(**S.base(**base)), S.f(**S.base(**alt_props)))

    t_fmts = mk()
    c_fmts = mk(align="center", bold=True)
    m_fmts = mk(align="right", bold=True, font_color=th.primary,
                num_format="#,##0.00")

    upcoming = (m.agg.get("upcoming") or [])[:PANEL_ROWS]
    dues = (m.agg.get("dues") or [])[:PANEL_ROWS]

    for i in range(PANEL_ROWS):
        prow = ROW_PANEL + i
        ws.set_row(r(prow), 20)
        a = 1 if i % 2 else 0
        drow = C.DATA_POOL_FIRST + i
        due_row = C.DATA_DUE_FIRST + i

        # left panel: date | event | stage
        ws.write_formula(
            r(prow), 1,
            '=IF(%s!$AA$%d="","\u2014",TEXT(%s!$AA$%d,"ddd dd mmm"))'
            % (d, drow, d, drow), c_fmts[a],
            upcoming[i][0].strftime("%a %d %b") if i < len(upcoming)
            else "\u2014")
        bk.stats["formulas"] += 1
        ws.merge_range(r(prow), 2, r(prow), 3, "", t_fmts[a])
        ws.write_formula(
            r(prow), 2,
            '=IF(%s!$AB$%d="","(nothing booked yet)",%s!$AB$%d)'
            % (d, drow, d, drow), t_fmts[a],
            upcoming[i][1] if i < len(upcoming) else "(nothing booked yet)")
        bk.stats["formulas"] += 1
        ws.merge_range(r(prow), 4, r(prow), 5, "", c_fmts[a])
        ws.write_formula(
            r(prow), 4, '=IF(%s!$AC$%d="","",%s!$AC$%d)'
            % (d, drow, d, drow), c_fmts[a],
            upcoming[i][2] if i < len(upcoming) else "")
        bk.stats["formulas"] += 1

        # right panel: due | invoice | balance
        ws.write_formula(
            r(prow), 6,
            '=IF(%s!$AA$%d="","\u2014",TEXT(%s!$AA$%d,"dd mmm"))'
            % (d, due_row, d, due_row), c_fmts[a],
            dues[i][0].strftime("%d %b") if i < len(dues) else "\u2014")
        bk.stats["formulas"] += 1
        ws.merge_range(r(prow), 7, r(prow), 9, "", t_fmts[a])
        ws.write_formula(
            r(prow), 7,
            '=IF(%s!$AB$%d="","(all square \u2014 nothing owed)",%s!$AB$%d)'
            % (d, due_row, d, due_row), t_fmts[a],
            dues[i][1] if i < len(dues)
            else "(all square \u2014 nothing owed)")
        bk.stats["formulas"] += 1
        ws.merge_range(r(prow), 10, r(prow), ci(LAST_COL), "", m_fmts[a])
        ws.write_formula(
            r(prow), 10,
            '=IF(%s!$AA$%d="","",%s!$AC$%d)' % (d, due_row, d, due_row),
            m_fmts[a], dues[i][2] if i < len(dues) else "")
        bk.stats["formulas"] += 1

        # overdue rows glow red
        bk.cond(KEY, prow, 6, prow, ci(LAST_COL), {
            "type": "formula",
            "criteria": '=AND(%s!$AA$%d<>"",%s!$AA$%d<TODAY())'
                       % (d, due_row, d, due_row),
            "format": S.cf(bg=th.bad_soft, fg=th.bad, bold=True,
                           border=th.border)})
        # events inside the due-soon window glow amber
        bk.cond(KEY, prow, 1, prow, 5, {
            "type": "formula",
            "criteria": '=AND(%s!$AA$%d<>"",%s!$AA$%d-TODAY()<=DueSoonDays)'
                       % (d, drow, d, drow),
            "format": S.cf(bg=th.warn_soft, fg=th.warn, bold=True,
                           border=th.border)})
```

### `catering_tracker/sheets/clients.py`

```python
"""
\U0001F465 Client Tracker - the address book + latest booking per client.

One row per client with contact details and the money picture of their
current/next event (quote, deposit, balance, payment status).
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "clients"
LAST_COL = "O"

COLUMNS = [
    ("n", "#", "idx", None),
    ("name", "Client name", "text", None),
    ("phone", "Phone", "center", None),
    ("email", "Email", "text", None),
    ("event_date", "Next / last event", "date", None),
    ("event_type", "Event type", "center", None),
    ("guests", "Guests", "qty", None),
    ("venue", "Venue", "text", None),
    ("package", "Package", "center", None),
    ("quote", "Quote value", "money", None),
    ("deposit", "Deposit held", "money", None),
    ("balance", "Balance", "calc_money", "primary_2"),
    ("status", "Payment status", "center", None),
    ("notes", "Notes", "text", None),
]

PACKAGES = "Custom,Bronze,Silver,Gold,Platinum"


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY)
    bk.paint(KEY, 0, 0, C.last_row(KEY) + 24, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F465  Client Tracker",
        "  One row per client: contacts, their next (or latest) event and "
        "the money picture.  Repeat clients are gold for your business "
        "\u2014 the dashboard counts them.",
        LAST_COL)

    # ------------------------------------------------------------------
    # quick-stat strip
    # ------------------------------------------------------------------
    name_rng = bk.rng(KEY, "name")
    quote_rng = bk.rng(KEY, "quote")
    dep_rng = bk.rng(KEY, "deposit")
    ev_client = bk.rng("events", "client")
    chips = [
        ('\U0001F465 Clients: "&COUNTIF(%s,"?*")' % name_rng,
         "primary", "Clients: %d" % len(m.clients)),
        ('\U0001F501 Repeat customers: "&TEXT(%s,"0%%")'
         % bk.kpi("repeat_pct"), "gold",
         "Repeat customers: %d%%" % round(m.agg.get("repeat_pct", 0) * 100)),
        ('\U0001F4B0 Total quoted: "&Currency&TEXT(SUM(%s),"#,##0")'
         % quote_rng, "primary_2",
         "Total quoted: %s" % m.money(sum(c["quote"] for c in m.clients))),
        ('\U0001F4B5 Deposits held: "&Currency&TEXT(SUM(%s),"#,##0")'
         % dep_rng, "ok",
         "Deposits held: %s" % m.money(sum(c["deposit"] for c in m.clients))),
    ]
    col = 1
    for formula, color, cached in chips:
        fmt = S.pill(th.soft(color), getattr(th, color), size=10.5,
                     bold=True, align="left")
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), col + 2, "", fmt)
        ws.write_formula(r(C.ROW_STATS), col, '="' + formula, fmt, cached)
        bk.stats["formulas"] += 1
        col += 3
    ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), ci(LAST_COL),
                   "", S.canvas)

    # ------------------------------------------------------------------
    # table
    # ------------------------------------------------------------------
    K.table_frame(bk, KEY, COLUMNS, height=22)
    for i in range(K.n_rows(KEY)):
        rownum = C.ROW_FIRST + i
        values = {
            "n": i + 1,
            "balance": '=IF($K%d="","",$K%d-$L%d)' % (rownum, rownum, rownum),
        }
        cached = {"balance": ""}
        if i < len(m.clients):
            cl = m.clients[i]
            values.update({k: cl[k] for k in
                           ("name", "phone", "email", "event_date",
                            "event_type", "guests", "venue", "package",
                            "quote", "deposit", "status", "notes")})
            cached["balance"] = cl["quote"] - cl["deposit"]
        K.write_row(bk, KEY, COLUMNS, rownum, values, cached)

    ws.autofilter(r(C.ROW_HEADER), ci("B"), r(C.last_row(KEY)), ci(LAST_COL))
    ws.freeze_panes(r(C.ROW_FIRST), ci("D"))

    # ------------------------------------------------------------------
    # validation
    # ------------------------------------------------------------------
    K.date_dv(bk, KEY, ("event_date",))
    K.list_dv(bk, KEY, "event_type", "event_types", title="Event type")
    K.whole_dv(bk, KEY, "guests", minimum=0, maximum=100000)
    bk.validate(KEY, C.ROW_FIRST, ci(bk.col(KEY, "package")),
                C.last_row(KEY), ci(bk.col(KEY, "package")),
                '"%s"' % PACKAGES, title="Package",
                message="Your service tiers - edit the list on \u2699\uFE0F "
                        "Setup if you named them differently.")
    K.money_dv(bk, KEY, ("quote", "deposit"))
    K.fixed_dv(bk, KEY, "status", "PaymentStatuses", title="Payment status",
               message="\U0001F534 Unpaid \u2192 \U0001F7E0 Partial \u2192 "
                       "\U0001F7E2 Paid in full.",
               error="Pick one of the three payment statuses.")

    # ------------------------------------------------------------------
    # conditional formatting
    # ------------------------------------------------------------------
    K.status_cf(bk, KEY, "status", {
        C.PS_UNPAID: (th.bad_soft, th.bad),
        C.PS_PART: (th.warn_soft, th.warn),
        C.PS_PAID: (th.ok_soft, th.ok),
    })
    # repeat customers get a golden name
    bk.cond(KEY, C.ROW_FIRST, ci("C"), C.last_row(KEY), ci("C"), {
        "type": "formula",
        "criteria": '=AND($C%d<>"",COUNTIF(%s,$C%d)>1)'
                   % (C.ROW_FIRST, ev_client, C.ROW_FIRST),
        "format": S.cf(bg=th.gold_soft, fg=th.gold, bold=True)})
    K.deadline_cf(bk, KEY, "event_date")

    # ------------------------------------------------------------------
    # totals + notes
    # ------------------------------------------------------------------
    trow = C.last_row(KEY) + 2
    K.totals_row(bk, KEY, trow, {
        "guests": ("=SUM($H$%d:$H$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                   "#,##0", sum(c["guests"] for c in m.clients)),
        "quote": ("=SUM($K$%d:$K$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                  "#,##0.00", sum(c["quote"] for c in m.clients)),
        "deposit": ("=SUM($L$%d:$L$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                    "#,##0.00", sum(c["deposit"] for c in m.clients)),
        "balance": ("=SUM($M$%d:$M$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                    "#,##0.00",
                    sum(c["quote"] - c["deposit"] for c in m.clients)),
    }, label="TOTALS  \u2192", label_span=("B", "G"), last_col=LAST_COL)

    K.note_block(
        bk, KEY, trow + 2, "B", "O",
        ["Names typed here feed the Client dropdown on the Events, "
         "Payments and Invoice tabs.",
         "A golden name means the client appears on more than one event "
         "- repeat business!",
         "Quote / Deposit are inputs; Balance calculates itself. For the "
         "full payment schedule use the \U0001F4B0 Payments tab."],
        title="\U0001F4A1 How to use this tab")

    nav = trow + 8
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)
    bk.page(KEY, LAST_COL, nav + 1, freeze=(C.ROW_FIRST, 3),
            title_rows=(C.ROW_HEADER, C.ROW_HEADER))
```

### `catering_tracker/sheets/events.py`

```python
"""
\U0001F4C5 Event / Order Tracker - the pipeline.

💬 Inquiry → 📝 Quote Sent → 💰 Deposit Paid → ✅ Confirmed → 🎉 Completed
(❌ Cancelled).  White cells are inputs; Profit, Margin and Balance are
formulas and locked.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "events"
LAST_COL = "R"

COLUMNS = [
    ("n", "#", "idx", None),
    ("id", "Event ID", "text", None),
    ("client", "Client", "text", None),
    ("date", "Event date", "date", None),
    ("type", "Type", "center", None),
    ("guests", "Guests", "qty", None),
    ("menu", "Menu / package", "wrap", None),
    ("staff_req", "Staff req.", "qty", None),
    ("equipment_req", "Equipment req.", "wrap", None),
    ("cost", "Total cost", "money", None),
    ("price", "Price quoted", "money", None),
    ("profit", "Profit", "calc_money", "primary_2"),
    ("margin", "Margin", "calc_pct", "primary_2"),
    ("deposit", "Deposit held", "money", None),
    ("balance", "Balance", "calc_money", "primary_2"),
    ("status", "Pipeline status", "center", None),
    ("notes", "Notes", "text", None),
]

STATUS_COLORS = {
    C.ES_INQUIRY: ("info_soft", "info"),
    C.ES_QUOTED: ("plum_soft", "plum"),
    C.ES_DEPOSIT: ("warn_soft", "warn"),
    C.ES_CONFIRMED: ("primary_soft", "primary"),
    C.ES_DONE: ("ok_soft", "ok"),
    C.ES_CANCEL: ("bad_soft", "bad"),
}


def _formulas(row):
    return {
        "n": row - C.ROW_FIRST + 1,
        "profit": '=IF(OR($K%d="",$L%d=""),"",$L%d-$K%d)' % (row, row, row, row),
        "margin": '=IF(OR($L%d="",$L%d=0),"",$M%d/$L%d)' % (row, row, row, row),
        "balance": '=IF($L%d="","",$L%d-$O%d)' % (row, row, row),
    }


def _cached(m, i):
    if i < len(m.events):
        e = m.events[i]
        return {"profit": e["profit"], "margin": e["margin"],
                "balance": e["balance"]}
    return {}


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY)
    bk.paint(KEY, 0, 0, C.last_row(KEY) + 26, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F4C5  Event & Order Tracker",
        "  The pipeline: \U0001F4AC Inquiry \u2192 \U0001F4DD Quote Sent "
        "\u2192 \U0001F4B0 Deposit Paid \u2192 \u2705 Confirmed \u2192 "
        "\U0001F389 Completed.  White cells are yours; tinted cells "
        "calculate.",
        LAST_COL)

    # ------------------------------------------------------------------
    # quick-stat strip
    # ------------------------------------------------------------------
    price = bk.rng(KEY, "price")
    profit = bk.rng(KEY, "profit")
    status = bk.rng(KEY, "status")
    not_cancel = '"<>%s"' % C.ES_CANCEL
    booked = sum(e["price"] for e in m.events if e["status"] != C.ES_CANCEL)
    projected = sum(e["profit"] for e in m.events
                    if e["status"] != C.ES_CANCEL)
    chips = [
        ('\U0001F4C5 Events: "&%s' % bk.kpi("events_total"),
         "primary", "Events: %d" % m.agg.get("events_total", 0)),
        ('\u2705 Confirmed + deposit: "&%s' % bk.kpi("events_confirmed"),
         "ok", "Confirmed + deposit: %d" % m.agg.get("events_confirmed", 0)),
        ('\U0001F389 Completed: "&%s&"  \u274C Cancelled: "&%s'
         % (bk.kpi("events_done"), bk.kpi("events_cancelled")),
         "gold", "Completed: %d  Cancelled: %d"
         % (m.agg.get("events_done", 0), m.agg.get("events_cancelled", 0))),
        ('\U0001F4B0 Booked value: "&Currency&TEXT(SUMIF(%s,%s,%s),"#,##0")'
         % (status, not_cancel, price),
         "primary_2", "Booked value: %s" % m.money(booked)),
        ('\U0001F4C8 Projected profit: "&Currency&TEXT(SUMIF(%s,%s,%s),"#,##0")'
         % (status, not_cancel, profit),
         "accent", "Projected profit: %s" % m.money(projected)),
    ]
    col = 1
    for formula, color, cached in chips:
        fmt = S.pill(th.soft(color), getattr(th, color), size=10.5,
                     bold=True, align="left")
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), col + 2, "", fmt)
        ws.write_formula(r(C.ROW_STATS), col, '="' + formula, fmt, cached)
        bk.stats["formulas"] += 1
        col += 3
    if col <= ci(LAST_COL):
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), ci(LAST_COL),
                       "", S.canvas)

    # ------------------------------------------------------------------
    # table
    # ------------------------------------------------------------------
    K.table_frame(bk, KEY, COLUMNS, height=34)
    for i in range(K.n_rows(KEY)):
        rownum = C.ROW_FIRST + i
        values = _formulas(rownum)
        cached = {"profit": "", "margin": "", "balance": ""}
        if i < len(m.events):
            e = m.events[i]
            values.update({k: e[k] for k in
                           ("id", "client", "date", "type", "guests", "menu",
                            "staff_req", "equipment_req", "cost", "price",
                            "deposit", "status", "notes")})
            cached = _cached(m, i)
        K.write_row(bk, KEY, COLUMNS, rownum, values, cached)

    ws.autofilter(r(C.ROW_HEADER), ci("B"), r(C.last_row(KEY)), ci(LAST_COL))
    ws.freeze_panes(r(C.ROW_FIRST), ci("D"))

    # ------------------------------------------------------------------
    # validation
    # ------------------------------------------------------------------
    bk.validate(KEY, C.ROW_FIRST, ci(bk.col(KEY, "client")),
                C.last_row(KEY), ci(bk.col(KEY, "client")),
                "=ClientsList", title="Client",
                message="Pick a client from the \U0001F465 Clients tab, or "
                        "type a new name.")
    K.date_dv(bk, KEY, ("date",),
              message="When is the event? Type a date or pick from the "
                      "calendar.")
    K.list_dv(bk, KEY, "type", "event_types", title="Event type",
              message="Wedding, corporate, birthday\u2026 editable on "
                      "\u2699\uFE0F Setup.")
    K.whole_dv(bk, KEY, "guests", minimum=0, maximum=100000)
    K.whole_dv(bk, KEY, "staff_req", minimum=0, maximum=500)
    K.money_dv(bk, KEY, ("cost", "price", "deposit"))
    K.fixed_dv(bk, KEY, "status", "EventStatuses", title="Pipeline status",
               message="\U0001F4AC Inquiry \u2192 \U0001F4DD Quote Sent "
                       "\u2192 \U0001F4B0 Deposit Paid \u2192 \u2705 "
                       "Confirmed \u2192 \U0001F389 Completed.",
               error="Pick one of the six pipeline statuses \u2014 the "
                     "colours and the dashboard depend on them.")

    # ------------------------------------------------------------------
    # conditional formatting
    # ------------------------------------------------------------------
    K.status_cf(bk, KEY, "status",
                {text: (getattr(th, bg), getattr(th, fg))
                 for text, (bg, fg) in STATUS_COLORS.items()})
    K.deadline_cf(bk, KEY, "date")
    L = bk.col(KEY, "price")
    bk.cond(KEY, C.ROW_FIRST, ci(L), C.last_row(KEY), ci(L), {
        "type": "formula", "criteria": '=$%s%d<>""' % (bk.col(KEY, "id"),
                                                       C.ROW_FIRST),
        "format": S.cf(bg=th.gold_soft, fg=th.ink, bold=True)})
    # cancelled rows: strike the client + price
    for field in ("client", "price"):
        col = bk.col(KEY, field)
        bk.cond(KEY, C.ROW_FIRST, ci(col), C.last_row(KEY), ci(col), {
            "type": "formula",
            "criteria": '=$%s%d="%s"' % (bk.col(KEY, "status"), C.ROW_FIRST,
                                         C.ES_CANCEL),
            "format": S.cf(strike=True, fg=th.muted)})

    # ------------------------------------------------------------------
    # totals
    # ------------------------------------------------------------------
    trow = C.last_row(KEY) + 2
    K.totals_row(bk, KEY, trow, {
        "guests": ("=SUM($G$%d:$G$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                   "#,##0", sum(e["guests"] for e in m.events)),
        "cost": ("=SUM($K$%d:$K$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                 "#,##0.00", sum(e["cost"] for e in m.events)),
        "price": ("=SUM($L$%d:$L$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                  "#,##0.00", sum(e["price"] for e in m.events)),
        "profit": ("=SUM($M$%d:$M$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                   "#,##0.00", sum(e["profit"] for e in m.events)),
        "deposit": ("=SUM($O$%d:$O$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                    "#,##0.00", sum(e["deposit"] for e in m.events)),
    }, label="TOTALS  (all events incl. cancelled)  \u2192",
        label_span=("B", "F"), last_col=LAST_COL)

    K.note_block(
        bk, KEY, trow + 2, "B", "R",
        ["Fill the white cells - the tinted Profit / Margin / Balance "
         "columns do the maths.",
         "Need a price? Build it on the \U0001F9EE Quote Calculator first, "
         "then copy the recommended price into 'Price quoted'.",
         "Cancelled events stay in the list (greyed + struck through) so "
         "your history is complete - they are excluded from every total "
         "and average."],
        title="\U0001F4A1 How to use this tab")

    nav = trow + 8
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)
    bk.page(KEY, LAST_COL, nav + 1, freeze=(C.ROW_FIRST, 3),
            title_rows=(C.ROW_HEADER, C.ROW_HEADER))
```

### `catering_tracker/sheets/quote.py`

```python
"""
\U0001F9EE Quote / Pricing Calculator - the headline selling point.

Type the guest count and your five cost buckets; the right-hand card turns
them into total cost, recommended price (at your target margin), price per
guest, profit, deposit and balance due.  Copy the recommended price onto
the Events tab or the printable Invoice.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "quote"
LAST_COL = "J"
QR = C.QUOTE_ROWS
QO = C.QUOTE_OUT


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo
    q = m.quote

    bk.widths(KEY, {"A": 2.2, "B": 17, "C": 17, "D": 17, "E": 16, "F": 3,
                    "G": 17, "H": 15, "I": 15, "J": 17})
    bk.paint(KEY, 0, 0, C.QUOTE_LAST_ROW + 8, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F9EE  Quote & Pricing Calculator",
        "  Never guess a price again: costs in \u2192 recommended quote "
        "out, at the margin YOU decide.",
        LAST_COL)
    ws.set_row(r(5), 26)
    ws.merge_range(r(5), 1, r(5), ci(LAST_COL), "", S.canvas)
    ws.set_row(r(8), 8)

    lbl = S.f(**S.base(font_size=11, bold=True, font_color=th.ink,
                       bg_color=th.card, align="left", valign="vcenter",
                       border=1, border_color=th.border, indent=1))
    note_f = S.f(**S.base(font_size=9.5, italic=True, font_color=th.muted,
                          bg_color=th.card, align="left", valign="vcenter",
                          text_wrap=True, border=1, border_color=th.border,
                          indent=1))
    inp = S.f(**S.base(font_size=12, bold=True, font_color=th.primary,
                       bg_color=th.gold_soft, align="center",
                       valign="vcenter", border=2, border_color=th.gold,
                       locked=False))
    inp_n = S.f(**S.base(font_size=12, bold=True, font_color=th.primary,
                         bg_color=th.gold_soft, align="center",
                         valign="vcenter", border=2, border_color=th.gold,
                         num_format="#,##0.00", locked=False))
    inp_i = S.f(**S.base(font_size=12, bold=True, font_color=th.primary,
                         bg_color=th.gold_soft, align="center",
                         valign="vcenter", border=2, border_color=th.gold,
                         num_format="#,##0", locked=False))
    inp_p = S.f(**S.base(font_size=12, bold=True, font_color=th.primary,
                         bg_color=th.gold_soft, align="center",
                         valign="vcenter", border=2, border_color=th.gold,
                         num_format="0%", locked=False))
    inp_d = S.f(**S.base(font_size=12, bold=True, font_color=th.primary,
                         bg_color=th.gold_soft, align="center",
                         valign="vcenter", border=2, border_color=th.gold,
                         num_format="dd mmm yyyy", locked=False))
    out = S.f(**S.base(font_size=12, bold=True, font_color=th.primary,
                       bg_color=th.primary_soft, align="center",
                       valign="vcenter", border=1, border_color=th.border,
                       num_format="#,##0.00"))
    out_hero = S.f(**S.base(font_name=th.title_font, font_size=17, bold=True,
                            font_color=th.white, bg_color=th.ok,
                            align="center", valign="vcenter", border=1,
                            border_color=th.ok, num_format="#,##0.00"))
    out_pct = S.f(**S.base(font_size=12, bold=True, font_color=th.primary,
                           bg_color=th.primary_soft, align="center",
                           valign="vcenter", border=1,
                           border_color=th.border, num_format="0.0%"))

    def left(row, label, value, fmt, cached=0, height=26):
        from datetime import date as _date
        ws.set_row(r(row), height)
        ws.merge_range(r(row), 1, r(row), 3, label, lbl)
        if isinstance(value, str) and value.startswith("="):
            ws.write_formula(r(row), 4, value, fmt, cached)
            bk.stats["formulas"] += 1
        elif isinstance(value, _date):
            ws.write_datetime(r(row), 4, value, fmt)
        else:
            ws.write(r(row), 4, value, fmt)
        ws.write_blank(r(row), 5, None, S.canvas)
        return

    def right(row, label, formula, cached, fmt=None, height=26, hero=False):
        ws.set_row(r(row), height)
        ws.merge_range(r(row), 6, r(row), 8, label,
                       lbl if not hero else S.f(**S.base(
                           font_size=11, bold=True, font_color=th.white,
                           bg_color=th.ok, align="left", valign="vcenter",
                           border=1, border_color=th.ok, indent=1)))
        ws.write_formula(r(row), 9, formula, fmt or (out_hero if hero else out),
                         cached)
        bk.stats["formulas"] += 1

    # ------------------------------------------------------------------
    # left card: inputs
    # ------------------------------------------------------------------
    ws.set_row(r(9), 24)
    ws.merge_range(r(9), 1, r(9), 4, "  1 \u00B7 THE EVENT", S.section_soft)
    ws.merge_range(r(9), 6, r(9), ci(LAST_COL), "", S.canvas)
    left(QR["client"], "Client / enquiry", q.get("client", ""), inp)
    left(QR["date"], "Event date", q.get("date", ""), inp_d)
    left(QR["guests"], "Guest count", q.get("guests", ""), inp_i)

    ws.set_row(r(13), 24)
    ws.merge_range(r(13), 1, r(13), 4, "  2 \u00B7 YOUR COSTS",
                   S.section_soft)
    ws.merge_range(r(13), 6, r(13), ci(LAST_COL), "", S.canvas)
    left(QR["food_pp"], "Food cost per guest", q.get("food_pp", ""), inp_n)
    left(QR["labor"], "Labor total", q.get("labor", ""), inp_n)
    left(QR["equipment"], "Equipment / rental", q.get("equipment", ""),
         inp_n)
    left(QR["transport"], "Transport", q.get("transport", ""), inp_n)
    left(QR["other"], "Other costs", q.get("other", ""), inp_n)

    ws.set_row(r(19), 24)
    ws.merge_range(r(19), 1, r(19), 4, "  3 \u00B7 PRICING POLICY",
                   S.section_soft)
    left(QR["margin"], "Target margin %", "=DefaultMargin", inp_p,
         cached=m.settings["margin"])
    left(QR["deposit_pct"], "Deposit %", "=DepositPct", inp_p,
         cached=m.settings["deposit"])
    ws.set_row(r(22), 8)

    # ------------------------------------------------------------------
    # right card: outputs
    # ------------------------------------------------------------------
    ws.merge_range(r(9), 6, r(12), ci(LAST_COL), "", S.canvas)
    right(QO["food_total"], "Food cost (guests x per guest)",
          '=IF(OR($E$%d="",$E$%d=""),"",$E$%d*$E$%d)'
          % (QR["guests"], QR["food_pp"], QR["guests"], QR["food_pp"]),
          q.get("food_total", ""))
    right(QO["cost_total"], "TOTAL COST",
          '=IF($J$%d="","",SUM($E$%d:$E$%d)+$J$%d)'
          % (QO["food_total"], QR["labor"], QR["other"], QO["food_total"]),
          q.get("cost_total", ""))
    right(QO["price"], "RECOMMENDED PRICE",
          '=IF(OR($J$%d="",$E$%d="",1-$E$%d<=0),"",'
          'ROUND($J$%d/(1-$E$%d),2))'
          % (QO["cost_total"], QR["guests"], QR["margin"],
             QO["cost_total"], QR["margin"]),
          q.get("price", ""), hero=True, height=34)
    right(QO["per_guest"], "Price per guest",
          '=IF(OR($J$%d="",$E$%d=""),"",$J$%d/$E$%d)'
          % (QO["price"], QR["guests"], QO["price"], QR["guests"]),
          q.get("per_guest", ""))
    right(QO["profit"], "Your profit",
          '=IF($J$%d="","",$J$%d-$J$%d)'
          % (QO["price"], QO["price"], QO["cost_total"]),
          q.get("profit", ""))
    right(QO["margin"], "Effective margin",
          '=IF($J$%d="","",$J$%d/$J$%d)'
          % (QO["price"], QO["profit"], QO["price"]),
          (q["profit"] / q["price"]) if q.get("price") else "", fmt=out_pct)
    right(QO["deposit"], "Deposit to collect",
          '=IF($J$%d="","",ROUND($J$%d*$E$%d,2))'
          % (QO["price"], QO["price"], QR["deposit_pct"]),
          q.get("deposit", ""))
    right(QO["balance"], "Balance due later",
          '=IF($J$%d="","",$J$%d-$J$%d)'
          % (QO["price"], QO["price"], QO["deposit"]),
          q.get("balance", ""))
    right(QO["cost_pp"], "True cost per guest",
          '=IF(OR($J$%d="",$E$%d=""),"",$J$%d/$E$%d)'
          % (QO["cost_total"], QR["guests"], QO["cost_total"],
             QR["guests"]),
          q.get("cost_pp", ""))

    # ------------------------------------------------------------------
    # notes + nav
    # ------------------------------------------------------------------
    note_row = C.QUOTE_LAST_ROW - 1
    ws.set_row(r(note_row), 22)
    ws.set_row(r(note_row + 1), 22)
    ws.set_row(r(note_row + 2), 22)
    ws.merge_range(r(note_row), 1, r(note_row + 2), ci(LAST_COL),
                   "  HOW TO PRICE LIKE A PRO\n"
                   "  1.  Food cost per guest = what one plate costs you "
                   "(use the Recipe Calculator on \U0001F37D\uFE0F Menu "
                   "Costing).\n"
                   "  2.  Labor / equipment / transport / other = the "
                   "event's share of everything else.\n"
                   "  3.  Recommended price covers ALL of it and still "
                   "pays you your target margin.  Rule of thumb: food "
                   "should stay under ~30% of the final price.",
                   S.note)

    nav = note_row + 4
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)

    # ------------------------------------------------------------------
    # validation
    # ------------------------------------------------------------------
    bk.validate(KEY, QR["client"], ci("E"), QR["client"], ci("E"),
                "=ClientsList", title="Client",
                message="Who is this quote for?")
    ws.data_validation(r(QR["date"]), ci("E"), r(QR["date"]), ci("E"), {
        "validate": "date", "criteria": ">=", "value": "=DATE(2000,1,1)",
        "input_title": "Event date", "show_input": True, "ignore_blank": True,
        "input_message": "Type a date or pick from the calendar."})
    bk.stats["validations"] += 1
    ws.data_validation(r(QR["guests"]), ci("E"), r(QR["guests"]), ci("E"), {
        "validate": "integer", "criteria": "between", "minimum": 1,
        "maximum": 100000, "input_title": "Guests", "show_input": True,
        "input_message": "How many are you feeding?"})
    bk.stats["validations"] += 1
    for field in ("food_pp", "labor", "equipment", "transport", "other"):
        ws.data_validation(r(QR[field]), ci("E"), r(QR[field]), ci("E"), {
            "validate": "decimal", "criteria": ">=", "value": 0,
            "input_title": "Cost", "show_input": True, "ignore_blank": True})
        bk.stats["validations"] += 1
    for field, (lo, hi) in (("margin", (0, 0.95)), ("deposit_pct", (0, 1))):
        ws.data_validation(r(QR[field]), ci("E"), r(QR[field]), ci("E"), {
            "validate": "decimal", "criteria": "between", "minimum": lo,
            "maximum": hi, "show_input": True, "ignore_blank": True})
        bk.stats["validations"] += 1

    # CF: recommended price glows when computed
    bk.cond(KEY, QO["price"], ci("J"), QO["price"], ci("J"), {
        "type": "formula", "criteria": '=$J$%d<>""' % QO["price"],
        "format": S.cf(bg=th.ok, fg=th.white, bold=True, size=17)})

    bk.page(KEY, LAST_COL, nav + 1, landscape=False, zoom=100)
```

### `catering_tracker/sheets/menu.py`

```python
"""
\U0001F37D️ Menu & Recipe Costing.

Top: the priced menu (cost per portion, price per portion, profit and
margin - with a red flag under 30%).
Bottom: the RECIPE COST CALCULATOR - list up to six ingredients with the
quantity per portion; unit prices are looked up from Inventory, the cost
per portion and a suggested selling price (at your default margin) build
themselves.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "menu"
LAST_COL = "K"

COLUMNS = [
    ("n", "#", "idx", None),
    ("item", "Menu item", "text", None),
    ("category", "Category", "center", None),
    ("ingredients", "Main ingredients", "wrap", None),
    ("portion", "Portion", "center", None),
    ("cost", "Cost / portion", "money", None),
    ("price", "Price / portion", "money", None),
    ("profit", "Profit", "calc_money", "primary_2"),
    ("margin", "Margin", "calc_pct", "primary_2"),
    ("notes", "Notes", "text", None),
]




def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY)
    trow = C.last_row(KEY) + 2
    rc_top = trow + 9
    last = rc_top + 16
    bk.paint(KEY, 0, 0, last + 12, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F37D\uFE0F  Menu & Recipe Costing",
        "  Know the true cost of every dish.  Margins under 30% light up "
        "red \u2014 reprice them or retire them.",
        LAST_COL)

    # ------------------------------------------------------------------
    # quick-stat strip
    # ------------------------------------------------------------------
    item_c = bk.rng(KEY, "item")
    marg_c = bk.rng(KEY, "margin")
    chips = [
        ('\U0001F37D\uFE0F Items priced: "&%s' % bk.kpi("menu_items"),
         "primary", "Items priced: %d" % m.agg.get("menu_items", 0)),
        ('\U0001F4C8 Average margin: "&TEXT(%s,"0%%")' % bk.kpi("menu_margin"),
         "ok", "Average margin: %d%%"
         % round(m.agg.get("menu_margin", 0) * 100)),
        ('\u26A0\uFE0F Under 30%% margin: "&TEXT(COUNTIF(%s,">0")-'
         'COUNTIF(%s,">=0.3"),"0")' % (marg_c, marg_c), "bad",
         "Under 30%% margin: %d"
         % sum(1 for x in m.menu if x["margin"] < 0.30)),
    ]
    col = 1
    for formula, color, cached in chips:
        fmt = S.pill(th.soft(color), getattr(th, color), size=10.5,
                     bold=True, align="left")
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), col + 2, "", fmt)
        ws.write_formula(r(C.ROW_STATS), col, '="' + formula, fmt, cached)
        bk.stats["formulas"] += 1
        col += 3
    if col < ci(LAST_COL):
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), ci(LAST_COL),
                       "", S.canvas)
    elif col == ci(LAST_COL):
        ws.write_blank(r(C.ROW_STATS), col, None, S.canvas)

    # ------------------------------------------------------------------
    # menu table
    # ------------------------------------------------------------------
    K.table_frame(bk, KEY, COLUMNS, height=34)
    for i in range(K.n_rows(KEY)):
        rownum = C.ROW_FIRST + i
        values = {
            "n": i + 1,
            "profit": '=IF(OR($C%d="",$H%d=""),"",$H%d-$G%d)'
                      % (rownum, rownum, rownum, rownum),
            "margin": '=IF(OR($C%d="",$H%d="",$H%d=0),"",$I%d/$H%d)'
                      % (rownum, rownum, rownum, rownum, rownum),
        }
        cached = {"profit": "", "margin": ""}
        if i < len(m.menu):
            it = m.menu[i]
            values.update({k: it[k] for k in
                           ("item", "category", "ingredients", "portion",
                            "cost", "price", "notes")})
            cached.update({"profit": it["profit"], "margin": it["margin"]})
        K.write_row(bk, KEY, COLUMNS, rownum, values, cached)

    ws.autofilter(r(C.ROW_HEADER), ci("B"), r(C.last_row(KEY)), ci(LAST_COL))
    ws.freeze_panes(r(C.ROW_FIRST), ci("C"))

    K.list_dv(bk, KEY, "category", "menu_categories", title="Menu category")
    K.money_dv(bk, KEY, ("cost", "price"))
    bk.cond(KEY, C.ROW_FIRST, ci("J"), C.last_row(KEY), ci("J"), {
        "type": "formula",
        "criteria": '=AND($C%d<>"",$J%d<0.3)' % (C.ROW_FIRST, C.ROW_FIRST),
        "format": S.cf(bg=th.bad_soft, fg=th.bad, bold=True)})
    bk.cond(KEY, C.ROW_FIRST, ci("J"), C.last_row(KEY), ci("J"), {
        "type": "formula",
        "criteria": '=AND($C%d<>"",$J%d>=0.6)' % (C.ROW_FIRST, C.ROW_FIRST),
        "format": S.cf(bg=th.ok_soft, fg=th.ok, bold=True)})
    K.databar(bk, KEY, "price", color=th.gold)

    K.totals_row(bk, KEY, trow, {
        "cost": ("=IFERROR(AVERAGE($G$%d:$G$%d),0)"
                 % (C.ROW_FIRST, C.last_row(KEY)), "#,##0.00",
                 sum(x["cost"] for x in m.menu) / max(1, len(m.menu))),
        "price": ("=IFERROR(AVERAGE($H$%d:$H$%d),0)"
                  % (C.ROW_FIRST, C.last_row(KEY)), "#,##0.00",
                  sum(x["price"] for x in m.menu) / max(1, len(m.menu))),
        "margin": ("=IFERROR(AVERAGE($J$%d:$J$%d),0)"
                   % (C.ROW_FIRST, C.last_row(KEY)), "0%",
                   m.agg.get("menu_margin", 0)),
    }, label="AVERAGES  \u2192", label_span=("B", "F"), last_col=LAST_COL)

    # ------------------------------------------------------------------
    # recipe cost calculator
    # ------------------------------------------------------------------
    sec = rc_top
    ws.set_row(r(sec), 26)
    ws.merge_range(r(sec), 1, r(sec), ci(LAST_COL),
                   "  \U0001F9EE  RECIPE COST CALCULATOR  \u2014  price a "
                   "new dish before it goes on the menu", S.section_accent)

    dish_row = sec + 1
    ws.set_row(r(dish_row), 32)
    lbl = S.f(**S.base(font_size=11, bold=True, font_color=th.ink,
                       bg_color=th.card, align="left", valign="vcenter",
                       border=1, border_color=th.border, indent=1))
    inp = S.f(**S.base(font_size=11, bold=True, font_color=th.primary,
                       bg_color=th.gold_soft, align="center",
                       valign="vcenter", border=2, border_color=th.gold,
                       locked=False))
    inp_l = S.f(**S.base(font_size=11, bold=True, font_color=th.primary,
                         bg_color=th.gold_soft, align="left",
                         valign="vcenter", border=2, border_color=th.gold,
                         indent=1, locked=False))
    calc = S.f(**S.base(font_size=11, bold=True, font_color=th.primary,
                        bg_color=th.primary_soft, align="center",
                        valign="vcenter", border=1, border_color=th.border,
                        num_format="#,##0.00"))
    ws.merge_range(r(dish_row), 1, r(dish_row), 2, "Dish being costed", lbl)
    ws.merge_range(r(dish_row), 3, r(dish_row), 6, "", inp_l)
    dish_name = m.menu[0]["item"] if m.menu else ""
    ws.write(r(dish_row), 3, dish_name, inp_l)
    ws.merge_range(r(dish_row), 7, r(dish_row), ci(LAST_COL),
                   "Unit prices below are looked up from \U0001F4E6 "
                   "Inventory - add the ingredient there first.",
                   S.note_plain)

    hdr_row = dish_row + 1
    ws.set_row(r(hdr_row), 20)
    for c1, c2, text in ((1, 1, "#"), (2, 3, "Ingredient"), (4, 5, "Qty per portion"),
                         (6, 7, "Unit"), (8, 9, "Unit cost (auto)"),
                         (10, 10, "Line cost")):
        if c1 == c2:
            ws.write(r(hdr_row), c1, text, S.header(th.primary_2))
        else:
            ws.merge_range(r(hdr_row), c1, r(hdr_row), c2, text,
                           S.header(th.primary_2))

    demo_recipe = []
    if m.menu:
        demo_recipe = [("Chicken Breast", 0.22, "kg", 12.00),
                       ("Butter", 0.03, "kg", 6.40),
                       ("Olive Oil", 0.02, "L", 8.20),
                       ("Fresh Herbs", 0.01, "kg", 9.00)]

    ing_first = hdr_row + 1
    for i in range(6):
        row = ing_first + i
        ws.set_row(r(row), 20)
        a = i % 2
        ws.write(r(row), 1, i + 1, S.idx(a))
        ws.merge_range(r(row), 2, r(row), 3, "", S.cell("text", a))
        ws.merge_range(r(row), 4, r(row), 5, "", S.cell("qty1", a))
        ws.merge_range(r(row), 6, r(row), 7, "", S.cell("center", a))
        ws.merge_range(r(row), 8, r(row), 9, "", S.cell("calc_money", a))
        ws.write_formula(
            r(row), 8,
            '=IF($C%d="","",IFERROR(SUMIF(%s,$C%d,%s),0))'
            % (row, bk.rng("inventory", "ingredient"), row,
               bk.rng("inventory", "unit_cost")),
            S.cell("calc_money", a),
            demo_recipe[i][3] if i < len(demo_recipe) else "")
        bk.stats["formulas"] += 1
        ws.write_formula(
            r(row), 10, '=IF(OR($C%d="",$E%d=""),"",$E%d*$I%d)'
            % (row, row, row, row), S.cell("calc_money", a),
            round(demo_recipe[i][1] * demo_recipe[i][3], 4)
            if i < len(demo_recipe) else "")
        bk.stats["formulas"] += 1

    ing_last = ing_first + 5
    # demo ingredients for the signature dish
    if m.menu:
        for i, (ing, qty, unit, _uc) in enumerate(demo_recipe):
            row = ing_first + i
            ws.write(r(row), 2, ing, S.cell("text", i % 2))
            ws.write(r(row), 4, qty, S.cell("qty1", i % 2))
            ws.write(r(row), 6, unit, S.cell("center", i % 2))
    demo_cost = round(sum(q * u for _i, q, _un, u in demo_recipe), 4)

    out_row = ing_last + 1
    margin_d = m.settings["margin"] or 0.35
    demo_price = round(demo_cost / (1 - margin_d), 4) if demo_cost else 0
    outs = [
        ("Cost per portion",
         "=SUM($K$%d:$K$%d)" % (ing_first, ing_last), demo_cost),
        ("Suggested price at your default margin",
         "=IFERROR($D$%d/(1-DefaultMargin),\"\")" % out_row, demo_price),
        ("Profit per portion at that price",
         "=IFERROR($D$%d-$D$%d,\"\")" % (out_row + 1, out_row),
         round(demo_price - demo_cost, 4)),
    ]
    hints = ["Add the finished dish to the menu table above when you're "
             "happy with the numbers.",
             "DefaultMargin comes from \u2699\uFE0F Setup - change it once, "
             "every suggested price follows.",
             "Ingredient names must match the \U0001F4E6 Inventory spelling "
             "for the unit-cost lookup to work."]
    for j, (label, formula, cached) in enumerate(outs):
        row = out_row + j
        ws.set_row(r(row), 24)
        ws.merge_range(r(row), 1, r(row), 2, label, lbl)
        ws.merge_range(r(row), 3, r(row), 4, "", calc)
        ws.write_formula(r(row), 3, formula, calc, cached)
        bk.stats["formulas"] += 1
        ws.merge_range(r(row), 5, r(row), ci(LAST_COL), hints[j],
                       S.note_plain)

    nav = out_row + 4
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)
    bk.page(KEY, LAST_COL, nav + 1, freeze=(C.ROW_FIRST, 2),
            title_rows=(C.ROW_HEADER, C.ROW_HEADER))
```

### `catering_tracker/sheets/inventory.py`

```python
"""
\U0001F4E6 Ingredient Inventory - with LOW STOCK conditional formatting.

Qty, minimum level and unit cost are inputs; Stock value and Reorder status
are formulas.  🔴 Reorder / 🟡 Low / 🟢 OK colour themselves.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "inventory"
LAST_COL = "L"

COLUMNS = [
    ("n", "#", "idx", None),
    ("ingredient", "Ingredient", "text", None),
    ("category", "Category", "center", None),
    ("unit", "Unit", "center", None),
    ("qty", "Qty on hand", "qty1", None),
    ("min", "Minimum level", "qty1", None),
    ("unit_cost", "Unit cost", "money", None),
    ("value", "Stock value", "calc_money", "primary_2"),
    ("supplier", "Supplier", "text", None),
    ("status", "Reorder status", "calc_c", "primary_2"),
    ("notes", "Notes", "text", None),
]

UNITS = "kg,g,L,mL,pc,tray,box,bag,bottle"


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY)
    bk.paint(KEY, 0, 0, C.last_row(KEY) + 24, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F4E6  Ingredient Inventory",
        "  What's on the shelf right now.  Stock below its minimum lights "
        "up \U0001F534 red - the shopping list and recipe calculator read "
        "these prices.",
        LAST_COL)

    # ------------------------------------------------------------------
    # quick-stat strip
    # ------------------------------------------------------------------
    ing_c = bk.rng(KEY, "ingredient")
    status_c = bk.rng(KEY, "status")
    chips = [
        ('\U0001F4E6 Ingredients: "&COUNTIF(%s,"?*")' % ing_c,
         "primary", "Ingredients: %d" % len(m.inventory)),
        ('\U0001F4B0 Stock value: "&Currency&TEXT(%s,"#,##0.00")'
         % bk.kpi("inv_value"), "primary_2",
         "Stock value: %s" % m.money(m.agg.get("inv_value", 0), 2)),
        ('\U0001F534 Reorder now: "&COUNTIF(%s,"\U0001F534 Reorder")'
         % status_c, "bad",
         "Reorder now: %d" % sum(1 for i in m.inventory
                                 if i["status"] == "\U0001F534 Reorder")),
        ('\U0001F7E1 Running low: "&COUNTIF(%s,"\U0001F7E1 Low")' % status_c,
         "warn", "Running low: %d"
         % sum(1 for i in m.inventory if i["status"] == "\U0001F7E1 Low")),
        ('\U0001F7E2 Well stocked: "&COUNTIF(%s,"\U0001F7E2 OK")' % status_c,
         "ok", "Well stocked: %d"
         % sum(1 for i in m.inventory if i["status"] == "\U0001F7E2 OK")),
    ]
    col = 1
    for formula, color, cached in chips:
        fmt = S.pill(th.soft(color), getattr(th, color), size=10.5,
                     bold=True, align="left")
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), col + 1, "", fmt)
        ws.write_formula(r(C.ROW_STATS), col, '="' + formula, fmt, cached)
        bk.stats["formulas"] += 1
        col += 2
    if col < ci(LAST_COL):
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), ci(LAST_COL),
                       "", S.canvas)
    elif col == ci(LAST_COL):
        ws.write_blank(r(C.ROW_STATS), col, None, S.canvas)

    # ------------------------------------------------------------------
    # table
    # ------------------------------------------------------------------
    K.table_frame(bk, KEY, COLUMNS, height=20)
    for i in range(K.n_rows(KEY)):
        rownum = C.ROW_FIRST + i
        values = {
            "n": i + 1,
            "value": '=IF(OR($C%d="",$F%d=""),"",$F%d*$H%d)'
                     % (rownum, rownum, rownum, rownum),
            "status": ('=IF($C%d="","",IF($F%d<$G%d,"\U0001F534 Reorder",'
                       'IF($F%d<$G%d*1.5,"\U0001F7E1 Low","\U0001F7E2 OK")))'
                       % (rownum, rownum, rownum, rownum, rownum)),
        }
        cached = {"value": "", "status": ""}
        if i < len(m.inventory):
            it = m.inventory[i]
            values.update({k: it[k] for k in
                           ("ingredient", "category", "unit", "qty", "min",
                            "unit_cost", "supplier", "notes")})
            cached.update({"value": it["value"], "status": it["status"]})
        K.write_row(bk, KEY, COLUMNS, rownum, values, cached)

    ws.autofilter(r(C.ROW_HEADER), ci("B"), r(C.last_row(KEY)), ci(LAST_COL))
    ws.freeze_panes(r(C.ROW_FIRST), ci("C"))

    # ------------------------------------------------------------------
    # validation
    # ------------------------------------------------------------------
    K.list_dv(bk, KEY, "category", "ingredient_categories",
              title="Ingredient category")
    bk.validate(KEY, C.ROW_FIRST, ci(bk.col(KEY, "unit")),
                C.last_row(KEY), ci(bk.col(KEY, "unit")),
                '"%s"' % UNITS, title="Unit",
                message="kg, L, tray\u2026 or type your own.")
    bk.validate(KEY, C.ROW_FIRST, ci(bk.col(KEY, "supplier")),
                C.last_row(KEY), ci(bk.col(KEY, "supplier")),
                "=SuppliersList", title="Supplier",
                message="Pick from the \U0001F69A Suppliers tab, or type a "
                        "new one.")
    ws.data_validation(r(C.ROW_FIRST), ci(bk.col(KEY, "qty")),
                       r(C.last_row(KEY)), ci(bk.col(KEY, "min")),
                       {"validate": "decimal", "criteria": ">=", "value": 0,
                        "input_title": "Quantity",
                        "input_message": "Zero or more (decimals are fine).",
                        "show_input": True, "ignore_blank": True})
    bk.stats["validations"] += 1
    K.money_dv(bk, KEY, ("unit_cost",))

    # ------------------------------------------------------------------
    # conditional formatting
    # ------------------------------------------------------------------
    K.status_cf(bk, KEY, "status", {
        "\U0001F534 Reorder": (th.bad_soft, th.bad),
        "\U0001F7E1 Low": (th.warn_soft, th.warn),
        "\U0001F7E2 OK": (th.ok_soft, th.ok),
    })
    # the whole qty cell goes red when below minimum
    bk.cond(KEY, C.ROW_FIRST, ci("F"), C.last_row(KEY), ci("F"), {
        "type": "formula",
        "criteria": '=AND($C%d<>"",$F%d<$G%d)' % (C.ROW_FIRST, C.ROW_FIRST,
                                                  C.ROW_FIRST),
        "format": S.cf(bg=th.bad_soft, fg=th.bad, bold=True)})
    K.databar(bk, KEY, "value", color=th.primary_2)

    # ------------------------------------------------------------------
    # totals + notes
    # ------------------------------------------------------------------
    trow = C.last_row(KEY) + 2
    K.totals_row(bk, KEY, trow, {
        "value": ("=SUM($I$%d:$I$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                  "#,##0.00", m.agg.get("inv_value", 0)),
    }, label="TOTAL STOCK VALUE  \u2192", label_span=("B", "H"),
        last_col=LAST_COL)

    K.note_block(
        bk, KEY, trow + 2, "B", "L",
        ["\U0001F534 Reorder = below minimum. \U0001F7E1 Low = below 1.5x "
         "the minimum. \U0001F7E2 OK = comfortable.",
         "Unit costs typed here power the \U0001F37D\uFE0F Recipe "
         "Calculator and the \U0001F6D2 Shopping List estimates.",
         "Do a quick shelf count before every big event - ten minutes now "
         "saves an emergency run later."],
        title="\U0001F4A1 How to use this tab")

    nav = trow + 8
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)
    bk.page(KEY, LAST_COL, nav + 1, freeze=(C.ROW_FIRST, 2),
            title_rows=(C.ROW_HEADER, C.ROW_HEADER))
```

### `catering_tracker/sheets/shopping.py`

```python
"""
\U0001F6D2 Shopping List - auto-built from what events need vs what the
shelf already holds.

Required qty is the only real input besides the ingredient name:
Available = SUMIF over the inventory, To buy = MAX(0, required-available),
Est. cost = to-buy x the inventory unit cost.  Tick 🛒 Purchased when it's
in the van.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "shopping"
LAST_COL = "L"

COLUMNS = [
    ("n", "#", "idx", None),
    ("event", "Event ID", "center", None),
    ("ingredient", "Ingredient", "text", None),
    ("required", "Qty needed", "qty1", None),
    ("available", "In stock", "calc_qty1", "primary_2"),
    ("to_buy", "To buy", "calc_qty1", "primary_2"),
    ("supplier", "Supplier", "text", None),
    ("est_cost", "Est. cost", "calc_money", "primary_2"),
    ("purchased", "\U0001F6D2 Got it", "tick", None),
    ("date", "Bought on", "date", None),
    ("notes", "Notes", "text", None),
]


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY)
    bk.paint(KEY, 0, 0, C.last_row(KEY) + 24, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F6D2  Shopping List",
        "  What each event needs vs what the shelf holds \u2014 To buy and "
        "Est. cost fill themselves in from the \U0001F4E6 Inventory.",
        LAST_COL)

    # ------------------------------------------------------------------
    # quick-stat strip
    # ------------------------------------------------------------------
    ing_c = bk.rng(KEY, "ingredient")
    tick_c = bk.rng(KEY, "purchased")
    est_c = bk.rng(KEY, "est_cost")
    chips = [
        ('\U0001F6D2 Lines: "&COUNTIF(%s,"?*")' % ing_c,
         "primary", "Lines: %d" % len(m.shopping)),
        ('\U0001F4B0 Still to buy: "&%s&" line(s)"' % bk.kpi("shop_lines"),
         "warn", "Still to buy: %d line(s)" % m.agg.get("shop_lines", 0)),
        ('\U0001F4B5 Estimated cost to buy: "&Currency&TEXT(%s,"#,##0.00")'
         % bk.kpi("shop_cost"), "accent",
         "Estimated cost to buy: %s" % m.money(m.agg.get("shop_cost", 0), 2)),
        ('\u2705 Purchased: "&COUNTIF(%s,"%s")&" line(s)"'
         % (tick_c, C.TICK), "ok",
         "Purchased: %d line(s)"
         % sum(1 for s in m.shopping if s["purchased"] == C.TICK)),
        ('\U0001F9FE Spent on this list: "&Currency&TEXT(SUMIF(%s,"%s",%s),'
         '"#,##0.00")' % (tick_c, C.TICK, est_c), "primary_2",
         "Spent on this list: %s"
         % m.money(sum(s["est_cost"] for s in m.shopping
                       if s["purchased"] == C.TICK), 2)),
    ]
    col = 1
    for formula, color, cached in chips:
        fmt = S.pill(th.soft(color), getattr(th, color), size=10.5,
                     bold=True, align="left")
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), col + 1, "", fmt)
        ws.write_formula(r(C.ROW_STATS), col, '="' + formula, fmt, cached)
        bk.stats["formulas"] += 1
        col += 2
    if col < ci(LAST_COL):
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), ci(LAST_COL),
                       "", S.canvas)
    elif col == ci(LAST_COL):
        ws.write_blank(r(C.ROW_STATS), col, None, S.canvas)

    # ------------------------------------------------------------------
    # table
    # ------------------------------------------------------------------
    inv_ing = bk.rng("inventory", "ingredient")
    inv_qty = bk.rng("inventory", "qty")
    inv_cost = bk.rng("inventory", "unit_cost")
    K.table_frame(bk, KEY, COLUMNS, height=20)
    for i in range(K.n_rows(KEY)):
        rownum = C.ROW_FIRST + i
        values = {
            "n": i + 1,
            "available": '=IF($D%d="",0,SUMIF(%s,$D%d,%s))'
                         % (rownum, inv_ing, rownum, inv_qty),
            "to_buy": '=IF($D%d="","",MAX(0,$E%d-$F%d))'
                      % (rownum, rownum, rownum),
            "est_cost": '=IF($G%d="","",ROUND($G%d*SUMIF(%s,$D%d,%s),2))'
                        % (rownum, rownum, inv_ing, rownum, inv_cost),
        }
        cached = {"available": 0, "to_buy": "", "est_cost": ""}
        if i < len(m.shopping):
            s = m.shopping[i]
            values.update({k: s[k] for k in
                           ("event", "ingredient", "required", "supplier",
                            "purchased", "date", "notes")})
            cached.update({"available": s["available"],
                           "to_buy": s["to_buy"], "est_cost": s["est_cost"]})
        K.write_row(bk, KEY, COLUMNS, rownum, values, cached)

    ws.autofilter(r(C.ROW_HEADER), ci("B"), r(C.last_row(KEY)), ci(LAST_COL))
    ws.freeze_panes(r(C.ROW_FIRST), ci("C"))

    # ------------------------------------------------------------------
    # validation
    # ------------------------------------------------------------------
    bk.validate(KEY, C.ROW_FIRST, ci(bk.col(KEY, "event")),
                C.last_row(KEY), ci(bk.col(KEY, "event")),
                "=EventList", title="Event ID",
                message="Which event needs this?")
    bk.validate(KEY, C.ROW_FIRST, ci(bk.col(KEY, "ingredient")),
                C.last_row(KEY), ci(bk.col(KEY, "ingredient")),
                "=OFFSET(%s!$C$%d,0,0,MAX(1,COUNTA(%s)),1)"
                % (bk.q("inventory"), C.ROW_FIRST, inv_ing),
                title="Ingredient",
                message="Match the spelling on the \U0001F4E6 Inventory tab "
                        "so In stock and Est. cost can look themselves up.")
    bk.validate(KEY, C.ROW_FIRST, ci(bk.col(KEY, "supplier")),
                C.last_row(KEY), ci(bk.col(KEY, "supplier")),
                "=SuppliersList", title="Supplier")
    ws.data_validation(r(C.ROW_FIRST), ci(bk.col(KEY, "required")),
                       r(C.last_row(KEY)), ci(bk.col(KEY, "required")),
                       {"validate": "decimal", "criteria": ">=", "value": 0,
                        "input_title": "Qty needed",
                        "input_message": "How much does the event need?",
                        "show_input": True, "ignore_blank": True})
    bk.stats["validations"] += 1
    K.tick_dv(bk, KEY, ("purchased",))
    K.date_dv(bk, KEY, ("date",))

    # ------------------------------------------------------------------
    # conditional formatting
    # ------------------------------------------------------------------
    K.tick_cf(bk, KEY, ("purchased",))
    # To buy > 0 and not purchased yet → red; purchased → whole row calms down
    bk.cond(KEY, C.ROW_FIRST, ci("G"), C.last_row(KEY), ci("G"), {
        "type": "formula",
        "criteria": '=AND($G%d>0,$J%d<>"%s")' % (C.ROW_FIRST, C.ROW_FIRST,
                                                 C.TICK),
        "format": S.cf(bg=th.bad_soft, fg=th.bad, bold=True)})
    bk.cond(KEY, C.ROW_FIRST, ci("G"), C.last_row(KEY), ci("G"), {
        "type": "formula",
        "criteria": '=AND($G%d=0,$D%d<>"")' % (C.ROW_FIRST, C.ROW_FIRST),
        "format": S.cf(bg=th.ok_soft, fg=th.ok, bold=True)})
    bk.cond(KEY, C.ROW_FIRST, ci("C"), C.last_row(KEY), ci("L"), {
        "type": "formula",
        "criteria": '=$J%d="%s"' % (C.ROW_FIRST, C.TICK),
        "format": S.cf(fg=th.muted, strike=True)})

    # ------------------------------------------------------------------
    # totals + notes
    # ------------------------------------------------------------------
    trow = C.last_row(KEY) + 2
    K.totals_row(bk, KEY, trow, {
        "required": ("=SUM($E$%d:$E$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                     "#,##0.0", sum(s["required"] for s in m.shopping)),
        "to_buy": ("=SUM($G$%d:$G$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                   "#,##0.0", sum(s["to_buy"] for s in m.shopping)),
        "est_cost": ("=SUM($I$%d:$I$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                     "#,##0.00",
                     sum(s["est_cost"] for s in m.shopping)),
    }, label="TOTALS  \u2192", label_span=("B", "D"), last_col=LAST_COL)

    K.note_block(
        bk, KEY, trow + 2, "B", "L",
        ["Add one line per ingredient per event - In stock, To buy and Est. "
         "cost fill themselves in.",
         "Red 'To buy' = still needs purchasing.  Green = the shelf already "
         "covers it.  Struck-through = bought and done.",
         "When you buy, log the real spend on \U0001F4B8 Expenses (category "
         "Ingredients) so the P&L stays true."],
        title="\U0001F4A1 How to use this tab")

    nav = trow + 8
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)
    bk.page(KEY, LAST_COL, nav + 1, freeze=(C.ROW_FIRST, 2),
            title_rows=(C.ROW_HEADER, C.ROW_HEADER))
```

### `catering_tracker/sheets/expenses.py`

```python
"""
\U0001F4B8 Expense Tracker - every rand/dollar/rupee that leaves the till.

Categories are the eleven from the spec (Ingredients ... Miscellaneous),
editable on Setup.  Feeds the P&L, the tax tab and the dashboard pie.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "expenses"
LAST_COL = "J"

COLUMNS = [
    ("n", "#", "idx", None),
    ("date", "Date", "date", None),
    ("vendor", "Vendor / payee", "text", None),
    ("category", "Category", "center", None),
    ("desc", "Description", "text", None),
    ("amount", "Amount", "money", None),
    ("method", "Paid by", "center", None),
    ("event", "Event ID", "center", None),
    ("receipt", "\U0001F9FE Receipt filed", "tick", None),
]


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY)
    bk.paint(KEY, 0, 0, C.last_row(KEY) + 24, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F4B8  Expense Tracker",
        "  Log every business cost as it happens \u2014 the P&L, tax "
        "summary and dashboard pie all read this table.",
        LAST_COL)

    # ------------------------------------------------------------------
    # quick-stat strip
    # ------------------------------------------------------------------
    amt = bk.rng(KEY, "amount")
    cat = bk.rng(KEY, "category")
    dt = bk.rng(KEY, "date")
    this_month = ('SUMIFS(%s,%s,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1),'
                  '%s,"<"&EDATE(DATE(YEAR(TODAY()),MONTH(TODAY()),1),1))'
                  % (amt, dt, dt))
    month_cached = 0
    if bk.demo:
        from ..demo import TODAY as _now
        month_cached = sum(x["amount"] for x in m.expenses
                           if x["date"].month == _now.month
                           and x["date"].year == _now.year)
    chips = [
        ('\U0001F4B8 Total expenses: "&Currency&TEXT(%s,"#,##0")'
         % bk.kpi("expenses"), "bad",
         "Total expenses: %s" % m.money(m.agg.get("expenses", 0))),
        ('\U0001F4C5 This month: "&Currency&TEXT(%s,"#,##0")' % this_month,
         "warn", "This month: %s" % m.money(month_cached)),
        ('\U0001F373 Food + packaging: "&Currency&TEXT(%s,"#,##0")'
         % bk.kpi("food_cost"), "primary_2",
         "Food + packaging: %s" % m.money(m.agg.get("food_cost", 0))),
        ('\U0001F477 Staff / labor: "&Currency&TEXT(%s,"#,##0")'
         % bk.kpi("labor_cost"), "info",
         "Staff / labor: %s" % m.money(m.agg.get("labor_cost", 0))),
    ]
    col = 1
    for formula, color, cached in chips:
        fmt = S.pill(th.soft(color), getattr(th, color), size=10.5,
                     bold=True, align="left")
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), col + 1, "", fmt)
        ws.write_formula(r(C.ROW_STATS), col, '="' + formula, fmt, cached)
        bk.stats["formulas"] += 1
        col += 2
    if col < ci(LAST_COL):
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), ci(LAST_COL),
                       "", S.canvas)
    elif col == ci(LAST_COL):
        ws.write_blank(r(C.ROW_STATS), col, None, S.canvas)

    # ------------------------------------------------------------------
    # table
    # ------------------------------------------------------------------
    K.table_frame(bk, KEY, COLUMNS, height=20)
    for i in range(K.n_rows(KEY)):
        rownum = C.ROW_FIRST + i
        values = {"n": i + 1}
        if i < len(m.expenses):
            values.update(m.expenses[i])
        K.write_row(bk, KEY, COLUMNS, rownum, values)

    ws.autofilter(r(C.ROW_HEADER), ci("B"), r(C.last_row(KEY)), ci(LAST_COL))
    ws.freeze_panes(r(C.ROW_FIRST), ci("B"))

    # ------------------------------------------------------------------
    # validation
    # ------------------------------------------------------------------
    K.date_dv(bk, KEY, "date")
    K.list_dv(bk, KEY, "category", "expense_categories", title="Category",
              message="The eleven built-in categories - editable on "
                      "\u2699\uFE0F Setup.",
              error="Pick a category (or add your own on Setup first).")
    K.list_dv(bk, KEY, "method", "payment_methods", title="Paid by")
    K.money_dv(bk, KEY, ("amount",))
    bk.validate(KEY, C.ROW_FIRST, ci(bk.col(KEY, "event")),
                C.last_row(KEY), ci(bk.col(KEY, "event")),
                "=EventList", title="Event ID",
                message="Optional: tie this expense to an event.")
    K.tick_dv(bk, KEY, ("receipt",))

    # ------------------------------------------------------------------
    # conditional formatting
    # ------------------------------------------------------------------
    K.tick_cf(bk, KEY, ("receipt",))
    K.databar(bk, KEY, "amount", color=th.accent)
    # missing receipts glow red once the row has an amount
    bk.cond(KEY, C.ROW_FIRST, ci(bk.col(KEY, "receipt")),
            C.last_row(KEY), ci(bk.col(KEY, "receipt")), {
        "type": "formula",
        "criteria": '=AND($G%d<>"",$J%d<>"%s")' % (C.ROW_FIRST, C.ROW_FIRST,
                                                   C.TICK),
        "format": S.cf(bg=th.bad_soft, fg=th.bad, bold=True)})

    # ------------------------------------------------------------------
    # totals + notes
    # ------------------------------------------------------------------
    trow = C.last_row(KEY) + 2
    K.totals_row(bk, KEY, trow, {
        "amount": ("=SUM($G$%d:$G$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                   "#,##0.00", sum(x["amount"] for x in m.expenses)),
    }, label="TOTAL SPENT  \u2192", label_span=("B", "F"), last_col=LAST_COL)

    K.note_block(
        bk, KEY, trow + 2, "B", "J",
        ["Log the expense the day it happens - future-you doing tax season "
         "will say thanks.",
         "The Event ID column is optional, but tagging event purchases "
         "makes each event's true cost visible.",
         "Red tick column = money spent with no receipt filed yet."],
        title="\U0001F4A1 How to use this tab")

    nav = trow + 8
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)
    bk.page(KEY, LAST_COL, nav + 1, freeze=(C.ROW_FIRST, 0),
            title_rows=(C.ROW_HEADER, C.ROW_HEADER))
```

### `catering_tracker/sheets/income.py`

```python
"""
\U0001F4B0 Income & Payments - one row per invoice.

Inputs: invoice #, client, event date, invoice amount, deposit, payment 1,
payment 2, due date.  Automatic: received, balance, payment status.
Overdue balances flag themselves red - the OVERDUE auto-flag from the spec.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "income"
LAST_COL = "M"

COLUMNS = [
    ("n", "#", "idx", None),
    ("invoice", "Invoice #", "center", None),
    ("client", "Client", "text", None),
    ("event_date", "Event date", "date", None),
    ("amount", "Invoice total", "money", None),
    ("deposit", "Deposit", "money", None),
    ("pay1", "Payment 1", "money", None),
    ("pay2", "Payment 2", "money", None),
    ("received", "Received", "calc_money", "primary_2"),
    ("balance", "Balance due", "calc_money", "primary_2"),
    ("due", "Due date", "date", None),
    ("status", "Status", "calc_c", "primary_2"),
]


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY)
    bk.paint(KEY, 0, 0, C.last_row(KEY) + 24, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F4B0  Income & Payments",
        "  One row per invoice.  Log deposits and payments as they arrive "
        "\u2014 Received, Balance and Status calculate themselves, and "
        "overdue money turns red.",
        LAST_COL)

    # ------------------------------------------------------------------
    # quick-stat strip
    # ------------------------------------------------------------------
    amt = bk.rng(KEY, "amount")
    recv = bk.rng(KEY, "received")
    bal = bk.rng(KEY, "balance")
    due = bk.rng(KEY, "due")
    invoiced_cached = sum(i["amount"] for i in m.income)
    chips = [
        ('\U0001F4B0 Cash received: "&Currency&TEXT(%s,"#,##0")'
         % bk.kpi("revenue"), "ok",
         "Cash received: %s" % m.money(m.agg.get("revenue", 0))),
        ('\U0001F9FE Invoiced: "&Currency&TEXT(SUM(%s),"#,##0")' % amt,
         "primary", "Invoiced: %s" % m.money(invoiced_cached)),
        ('\u23F3 Outstanding: "&Currency&TEXT(SUM(%s),"#,##0")' % bal,
         "warn", "Outstanding: %s" % m.money(m.agg.get("outstanding", 0))),
        ('\U0001F534 Overdue: "&%s&" invoice(s)"' % bk.kpi("overdue"),
         "bad", "Overdue: %d invoice(s)" % m.agg.get("overdue", 0)),
        ('\U0001F4C2 Open invoices: "&%s' % bk.kpi("invoices_open"),
         "info", "Open invoices: %d" % m.agg.get("invoices_open", 0)),
    ]
    col = 1
    for formula, color, cached in chips:
        fmt = S.pill(th.soft(color), getattr(th, color), size=10.5,
                     bold=True, align="left")
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), col + 1, "", fmt)
        ws.write_formula(r(C.ROW_STATS), col, '="' + formula, fmt, cached)
        bk.stats["formulas"] += 1
        col += 2
    if col <= ci(LAST_COL):
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), ci(LAST_COL),
                       "", S.canvas)

    # ------------------------------------------------------------------
    # table
    # ------------------------------------------------------------------
    K.table_frame(bk, KEY, COLUMNS, height=20)
    for i in range(K.n_rows(KEY)):
        rownum = C.ROW_FIRST + i
        values = {
            "n": i + 1,
            "received": '=IF($F%d="","",SUM($G%d:$I%d))'
                        % (rownum, rownum, rownum),
            "balance": '=IF($F%d="","",$F%d-$J%d)' % (rownum, rownum, rownum),
            "status": ('=IF($F%d="","",IF($J%d>=$F%d,"%s",IF($J%d>0,"%s","%s")))'
                       % (rownum, rownum, rownum, C.PS_PAID,
                          rownum, C.PS_PART, C.PS_UNPAID)),
        }
        cached = {"received": "", "balance": "", "status": ""}
        if i < len(m.income):
            inv = m.income[i]
            values.update({k: inv[k] for k in
                           ("invoice", "client", "event_date", "amount",
                            "deposit", "pay1", "pay2", "due")})
            cached.update({"received": inv["received"],
                           "balance": inv["balance"], "status": inv["status"]})
        K.write_row(bk, KEY, COLUMNS, rownum, values, cached)

    ws.autofilter(r(C.ROW_HEADER), ci("B"), r(C.last_row(KEY)), ci(LAST_COL))
    ws.freeze_panes(r(C.ROW_FIRST), ci("D"))

    # ------------------------------------------------------------------
    # validation
    # ------------------------------------------------------------------
    bk.validate(KEY, C.ROW_FIRST, ci(bk.col(KEY, "client")),
                C.last_row(KEY), ci(bk.col(KEY, "client")),
                "=ClientsList", title="Client",
                message="Pick the client (from the \U0001F465 Clients tab).")
    K.date_dv(bk, KEY, ("event_date", "due"))
    K.money_dv(bk, KEY, ("amount", "deposit", "pay1", "pay2"))

    # ------------------------------------------------------------------
    # conditional formatting
    # ------------------------------------------------------------------
    K.status_cf(bk, KEY, "status", {
        C.PS_UNPAID: (th.bad_soft, th.bad),
        C.PS_PART: (th.warn_soft, th.warn),
        C.PS_PAID: (th.ok_soft, th.ok),
    })
    # OVERDUE auto-flag: balance > 0 and due date passed
    L = bk.col(KEY, "due")
    Kb = bk.col(KEY, "balance")
    for field in ("due", "balance", "status"):
        colL = bk.col(KEY, field)
        bk.cond(KEY, C.ROW_FIRST, ci(colL), C.last_row(KEY), ci(colL), {
            "type": "formula",
            "criteria": '=AND($%s%d<>"",$%s%d>0,$%s%d<TODAY())'
                       % (L, C.ROW_FIRST, Kb, C.ROW_FIRST,
                          L, C.ROW_FIRST),
            "format": S.cf(bg=th.bad_soft, fg=th.bad, bold=True)})
        bk.cond(KEY, C.ROW_FIRST, ci(colL), C.last_row(KEY), ci(colL), {
            "type": "formula",
            "criteria": '=AND($%s%d<>"",$%s%d>0,$%s%d>=TODAY(),'
                       '$%s%d-TODAY()<=DueSoonDays)'
                       % (L, C.ROW_FIRST, Kb, C.ROW_FIRST,
                          L, C.ROW_FIRST, L, C.ROW_FIRST),
            "format": S.cf(bg=th.warn_soft, fg=th.warn, bold=True)})

    # ------------------------------------------------------------------
    # totals + notes
    # ------------------------------------------------------------------
    trow = C.last_row(KEY) + 2
    K.totals_row(bk, KEY, trow, {
        "amount": ("=SUM($F$%d:$F$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                   "#,##0.00", invoiced_cached),
        "deposit": ("=SUM($G$%d:$G$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                    "#,##0.00", sum(i["deposit"] for i in m.income)),
        "pay1": ("=SUM($H$%d:$H$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                 "#,##0.00", sum(i["pay1"] for i in m.income)),
        "pay2": ("=SUM($I$%d:$I$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                 "#,##0.00", sum(i["pay2"] for i in m.income)),
        "received": ("=SUM($J$%d:$J$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                     "#,##0.00", m.agg.get("revenue", 0)),
        "balance": ("=SUM($K$%d:$K$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                    "#,##0.00", m.agg.get("outstanding", 0)),
    }, label="TOTALS  \u2192", label_span=("B", "E"), last_col=LAST_COL)

    K.note_block(
        bk, KEY, trow + 2, "B", "M",
        ["Invoice total, deposit and the two payment columns are the only "
         "money you type - Received, Balance and Status are formulas.",
         "RED row ending = the due date has passed and money is still "
         "owed. Chase it. AMBER = due inside your \u201Cdue soon\u201D "
         "window (Setup).",
         "Need a pretty document? The \U0001F5A8\uFE0F Invoice tab turns "
         "any row into a printable invoice."],
        title="\U0001F4A1 How to use this tab")

    nav = trow + 8
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)
    bk.page(KEY, LAST_COL, nav + 1, freeze=(C.ROW_FIRST, 3),
            title_rows=(C.ROW_HEADER, C.ROW_HEADER))
```

### `catering_tracker/sheets/staff.py`

```python
"""
\U0001F477 Staff & Labor Calculator - one row per person per shift.

Total pay = hours x rate + overtime hours x rate x 1.5.  The Paid tick
feeds the "unpaid shifts" alert on the dashboard.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "staff"
LAST_COL = "K"

COLUMNS = [
    ("n", "#", "idx", None),
    ("name", "Team member", "text", None),
    ("role", "Role", "center", None),
    ("event", "Event ID", "center", None),
    ("hours", "Hours", "qty", None),
    ("rate", "Rate / hr", "money", None),
    ("ot", "OT hrs", "qty", None),
    ("total", "Total pay", "calc_money", "primary_2"),
    ("paid", "\u2705 Paid", "tick", None),
    ("notes", "Notes", "text", None),
]


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY)
    bk.paint(KEY, 0, 0, C.last_row(KEY) + 24, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F477  Staff & Labor Calculator",
        "  One row per person per shift.  Overtime is paid at 1.5x "
        "automatically; the total feeds your labor cost picture.",
        LAST_COL)

    # ------------------------------------------------------------------
    # quick-stat strip
    # ------------------------------------------------------------------
    tot = bk.rng(KEY, "total")
    paid_c = bk.rng(KEY, "paid")
    name_c = bk.rng(KEY, "name")
    shifts_cached = len(m.staff)
    hours_cached = sum(s["hours"] + s["ot"] for s in m.staff)
    cost_cached = sum(s["total"] for s in m.staff)
    paid_cached = sum(s["total"] for s in m.staff if s["paid"] == C.TICK)
    chips = [
        ('\U0001F477 Shifts logged: "&COUNTIF(%s,"?*")' % name_c,
         "primary", "Shifts logged: %d" % shifts_cached),
        ('\u23F1 Hours worked: "&TEXT(SUM($F$%d:$F$%d)+SUM($H$%d:$H$%d),"0")'
         % (C.ROW_FIRST, C.last_row(KEY), C.ROW_FIRST, C.last_row(KEY)),
         "info", "Hours worked: %d" % round(hours_cached)),
        ('\U0001F4B0 Labor booked: "&Currency&TEXT(SUM(%s),"#,##0")' % tot,
         "primary_2", "Labor booked: %s" % m.money(cost_cached)),
        ('\u2705 Paid out: "&Currency&TEXT(SUMIF(%s,"%s",%s),"#,##0")'
         % (paid_c, C.TICK, tot), "ok",
         "Paid out: %s" % m.money(paid_cached)),
        ('\U0001F534 Unpaid shifts: "&%s' % bk.kpi("staff_unpaid"),
         "bad", "Unpaid shifts: %d" % m.agg.get("staff_unpaid", 0)),
    ]
    col = 1
    for formula, color, cached in chips:
        fmt = S.pill(th.soft(color), getattr(th, color), size=10.5,
                     bold=True, align="left")
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), col + 1, "", fmt)
        ws.write_formula(r(C.ROW_STATS), col, '="' + formula, fmt, cached)
        bk.stats["formulas"] += 1
        col += 2
    if col < ci(LAST_COL):
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), ci(LAST_COL),
                       "", S.canvas)
    elif col == ci(LAST_COL):
        ws.write_blank(r(C.ROW_STATS), col, None, S.canvas)

    # ------------------------------------------------------------------
    # table
    # ------------------------------------------------------------------
    K.table_frame(bk, KEY, COLUMNS, height=20)
    for i in range(K.n_rows(KEY)):
        rownum = C.ROW_FIRST + i
        values = {
            "n": i + 1,
            "total": ('=IF(OR($C%d="",$G%d=""),"",'
                      '$F%d*$G%d+$H%d*$G%d*1.5)'
                      % (rownum, rownum, rownum, rownum, rownum, rownum)),
        }
        cached = {"total": ""}
        if i < len(m.staff):
            s = m.staff[i]
            values.update({k: s[k] for k in
                           ("name", "role", "event", "hours", "rate", "ot",
                            "paid", "notes")})
            cached["total"] = s["total"]
        K.write_row(bk, KEY, COLUMNS, rownum, values, cached)

    ws.autofilter(r(C.ROW_HEADER), ci("B"), r(C.last_row(KEY)), ci(LAST_COL))
    ws.freeze_panes(r(C.ROW_FIRST), ci("C"))

    # ------------------------------------------------------------------
    # validation
    # ------------------------------------------------------------------
    K.list_dv(bk, KEY, "role", "staff_roles", title="Role",
              message="Head Chef, Server, Driver\u2026 editable on "
                      "\u2699\uFE0F Setup.")
    bk.validate(KEY, C.ROW_FIRST, ci(bk.col(KEY, "event")),
                C.last_row(KEY), ci(bk.col(KEY, "event")),
                "=EventList", title="Event ID",
                message="Which event is this shift for?")
    K.whole_dv(bk, KEY, "hours", minimum=0, maximum=24)
    K.whole_dv(bk, KEY, "ot", minimum=0, maximum=24)
    K.money_dv(bk, KEY, ("rate",))
    K.tick_dv(bk, KEY, ("paid",))

    # ------------------------------------------------------------------
    # conditional formatting
    # ------------------------------------------------------------------
    K.tick_cf(bk, KEY, ("paid",))
    # unpaid highlight once a total exists
    bk.cond(KEY, C.ROW_FIRST, ci(bk.col(KEY, "total")),
            C.last_row(KEY), ci(bk.col(KEY, "total")), {
        "type": "formula",
        "criteria": '=AND($I%d<>"",$J%d<>"%s")' % (C.ROW_FIRST, C.ROW_FIRST,
                                                   C.TICK),
        "format": S.cf(bg=th.bad_soft, fg=th.bad, bold=True)})

    # ------------------------------------------------------------------
    # totals + notes
    # ------------------------------------------------------------------
    trow = C.last_row(KEY) + 2
    K.totals_row(bk, KEY, trow, {
        "hours": ("=SUM($F$%d:$F$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                  "#,##0", sum(s["hours"] for s in m.staff)),
        "ot": ("=SUM($H$%d:$H$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
               "#,##0", sum(s["ot"] for s in m.staff)),
        "total": ("=SUM($I$%d:$I$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                  "#,##0.00", cost_cached),
    }, label="TOTALS  \u2192", label_span=("B", "E"), last_col=LAST_COL)

    K.note_block(
        bk, KEY, trow + 2, "B", "K",
        ["Overtime hours are paid at 1.5x the hourly rate automatically.",
         "Red Total = the shift has not been ticked Paid yet.",
         "When you actually pay wages, also log them on the \U0001F4B8 "
         "Expenses tab (Staff / Labor) so the P&L stays honest."],
        title="\U0001F4A1 How to use this tab")

    nav = trow + 8
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)
    bk.page(KEY, LAST_COL, nav + 1, freeze=(C.ROW_FIRST, 2),
            title_rows=(C.ROW_HEADER, C.ROW_HEADER))
```

### `catering_tracker/sheets/equipment.py`

```python
"""
\U0001F373 Equipment Tracker - what you own, what's out, what's broken.

Available = owned - reserved - damaged.  Status flags anything damaged or
with maintenance due inside 30 days.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "equipment"
LAST_COL = "L"

COLUMNS = [
    ("n", "#", "idx", None),
    ("item", "Equipment", "text", None),
    ("category", "Category", "center", None),
    ("owned", "Owned", "qty", None),
    ("reserved", "On events", "qty", None),
    ("damaged", "Damaged", "qty", None),
    ("available", "Available", "calc_num", "primary_2"),
    ("unit_value", "Unit value", "money", None),
    ("value", "Total value", "calc_money", "primary_2"),
    ("maintenance", "Next service", "date", None),
    ("status", "Status", "calc_c", "primary_2"),
]

CATEGORIES = "Serving,Furniture,Linens,Tableware,Beverage,Kitchen,Transport,Other"


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY)
    bk.paint(KEY, 0, 0, C.last_row(KEY) + 24, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F373  Equipment Tracker",
        "  Know exactly what you own, what is committed to events, what is "
        "damaged and what needs servicing \u2014 before you promise it to "
        "a client.",
        LAST_COL)

    # ------------------------------------------------------------------
    # quick-stat strip
    # ------------------------------------------------------------------
    item_c = bk.rng(KEY, "item")
    damaged_c = bk.rng(KEY, "damaged")
    chips = [
        ('\U0001F373 Equipment lines: "&COUNTIF(%s,"?*")' % item_c,
         "primary", "Equipment lines: %d" % len(m.equipment)),
        ('\U0001F4B0 Fleet value: "&Currency&TEXT(%s,"#,##0")'
         % bk.kpi("equip_value"), "primary_2",
         "Fleet value: %s" % m.money(m.agg.get("equip_value", 0))),
        ('\U0001F527 Needs attention: "&%s' % bk.kpi("equip_service"),
         "warn", "Needs attention: %d" % m.agg.get("equip_service", 0)),
        ('\U0001F4A5 Damaged units: "&SUM(%s)' % damaged_c, "bad",
         "Damaged units: %d" % sum(q["damaged"] for q in m.equipment)),
    ]
    col = 1
    for formula, color, cached in chips:
        fmt = S.pill(th.soft(color), getattr(th, color), size=10.5,
                     bold=True, align="left")
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), col + 1, "", fmt)
        ws.write_formula(r(C.ROW_STATS), col, '="' + formula, fmt, cached)
        bk.stats["formulas"] += 1
        col += 2
    if col < ci(LAST_COL):
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), ci(LAST_COL),
                       "", S.canvas)
    elif col == ci(LAST_COL):
        ws.write_blank(r(C.ROW_STATS), col, None, S.canvas)

    # ------------------------------------------------------------------
    # table
    # ------------------------------------------------------------------
    K.table_frame(bk, KEY, COLUMNS, height=20)
    for i in range(K.n_rows(KEY)):
        rownum = C.ROW_FIRST + i
        values = {
            "n": i + 1,
            "available": '=IF($C%d="","",$E%d-$F%d-$G%d)'
                         % (rownum, rownum, rownum, rownum),
            "value": '=IF(OR($C%d="",$I%d=""),"",$E%d*$I%d)'
                     % (rownum, rownum, rownum, rownum),
            "status": ('=IF($C%d="","",IF($G%d>0,"\U0001F527 Service",'
                       'IF(AND($K%d<>"",$K%d<=TODAY()+30),"\U0001F7E1 Soon",'
                       '"\U0001F7E2 Ready")))'
                       % (rownum, rownum, rownum, rownum)),
        }
        cached = {"available": "", "value": "", "status": ""}
        if i < len(m.equipment):
            q = m.equipment[i]
            values.update({k: q[k] for k in
                           ("item", "category", "owned", "reserved",
                            "damaged", "unit_value", "maintenance", "notes")
                           if k in q})
            cached.update({"available": q["available"], "value": q["value"],
                           "status": q["status"]})
        K.write_row(bk, KEY, COLUMNS, rownum, values, cached)

    ws.autofilter(r(C.ROW_HEADER), ci("B"), r(C.last_row(KEY)), ci(LAST_COL))
    ws.freeze_panes(r(C.ROW_FIRST), ci("C"))

    # ------------------------------------------------------------------
    # validation
    # ------------------------------------------------------------------
    bk.validate(KEY, C.ROW_FIRST, ci(bk.col(KEY, "category")),
                C.last_row(KEY), ci(bk.col(KEY, "category")),
                '"%s"' % CATEGORIES, title="Category")
    K.whole_dv(bk, KEY, ("owned", "reserved", "damaged"), minimum=0,
               maximum=10000)
    K.money_dv(bk, KEY, ("unit_value",))
    K.date_dv(bk, KEY, ("maintenance",),
              message="Next service / inspection date.")

    # ------------------------------------------------------------------
    # conditional formatting
    # ------------------------------------------------------------------
    K.status_cf(bk, KEY, "status", {
        "\U0001F527 Service": (th.bad_soft, th.bad),
        "\U0001F7E1 Soon": (th.warn_soft, th.warn),
        "\U0001F7E2 Ready": (th.ok_soft, th.ok),
    })
    K.deadline_cf(bk, KEY, "maintenance")
    # available == 0 → red number
    bk.cond(KEY, C.ROW_FIRST, ci("H"), C.last_row(KEY), ci("H"), {
        "type": "cell", "criteria": "<=", "value": 0,
        "format": S.cf(bg=th.bad_soft, fg=th.bad, bold=True)})

    # ------------------------------------------------------------------
    # totals + notes
    # ------------------------------------------------------------------
    trow = C.last_row(KEY) + 2
    K.totals_row(bk, KEY, trow, {
        "owned": ("=SUM($E$%d:$E$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                  "#,##0", sum(q["owned"] for q in m.equipment)),
        "reserved": ("=SUM($F$%d:$F$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                     "#,##0", sum(q["reserved"] for q in m.equipment)),
        "damaged": ("=SUM($G$%d:$G$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                    "#,##0", sum(q["damaged"] for q in m.equipment)),
        "available": ("=SUM($H$%d:$H$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                      "#,##0", sum(q["available"] for q in m.equipment)),
        "value": ("=SUM($J$%d:$J$%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                  "#,##0.00", m.agg.get("equip_value", 0)),
    }, label="TOTALS  \u2192", label_span=("B", "D"), last_col=LAST_COL)

    K.note_block(
        bk, KEY, trow + 2, "B", "L",
        ["\U0001F527 Service = something is damaged. \U0001F7E1 Soon = "
         "maintenance due inside 30 days. \U0001F7E2 Ready = good to go.",
         "\u201COn events\u201D is how many units are committed right now "
         "- update it as bookings come and go.",
         "Fleet value feeds your balance sheet (and your insurance "
         "conversation)."],
        title="\U0001F4A1 How to use this tab")

    nav = trow + 8
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)
    bk.page(KEY, LAST_COL, nav + 1, freeze=(C.ROW_FIRST, 2),
            title_rows=(C.ROW_HEADER, C.ROW_HEADER))
```

### `catering_tracker/sheets/suppliers.py`

```python
"""
\U0001F69A Vendor / Supplier Database.

A simple, searchable rolodex: who supplies what, at what price, with what
lead time and terms.  The Supplier dropdowns elsewhere read this list.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "suppliers"
LAST_COL = "M"

COLUMNS = [
    ("n", "#", "idx", None),
    ("supplier", "Supplier", "text", None),
    ("contact", "Contact person", "text", None),
    ("phone", "Phone", "center", None),
    ("email", "Email", "text", None),
    ("category", "Supplies", "center", None),
    ("item", "Key items", "text", None),
    ("price", "Typical unit price", "money", None),
    ("min_order", "Min order", "money0", None),
    ("delivery", "Delivery time", "center", None),
    ("terms", "Payment terms", "center", None),
    ("notes", "Notes", "text", None),
]


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY)
    bk.paint(KEY, 0, 0, C.last_row(KEY) + 22, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F69A  Supplier Database",
        "  Your wholesale rolodex.  Names typed here feed the Supplier "
        "dropdowns on the Inventory and Shopping tabs.",
        LAST_COL)

    # ------------------------------------------------------------------
    # quick-stat strip
    # ------------------------------------------------------------------
    sup_c = bk.rng(KEY, "supplier")
    chips = [
        ('\U0001F69A Suppliers: "&COUNTIF(%s,"?*")' % sup_c,
         "primary", "Suppliers: %d" % len(m.suppliers)),
        ('\U0001F4E6 Ingredients tracked: "&COUNTA(%s)'
         % bk.rng("inventory", "ingredient"),
         "primary_2", "Ingredients tracked: %d" % len(m.inventory)),
        ('\U0001F534 Below minimum: "&%s&" item(s)"' % bk.kpi("low_stock"),
         "bad", "Below minimum: %d item(s)" % m.agg.get("low_stock", 0)),
    ]
    col = 1
    for formula, color, cached in chips:
        fmt = S.pill(th.soft(color), getattr(th, color), size=10.5,
                     bold=True, align="left")
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), col + 2, "", fmt)
        ws.write_formula(r(C.ROW_STATS), col, '="' + formula, fmt, cached)
        bk.stats["formulas"] += 1
        col += 3
    if col < ci(LAST_COL):
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), ci(LAST_COL),
                       "", S.canvas)
    elif col == ci(LAST_COL):
        ws.write_blank(r(C.ROW_STATS), col, None, S.canvas)

    # ------------------------------------------------------------------
    # table
    # ------------------------------------------------------------------
    K.table_frame(bk, KEY, COLUMNS, height=20)
    for i in range(K.n_rows(KEY)):
        rownum = C.ROW_FIRST + i
        values = {"n": i + 1}
        if i < len(m.suppliers):
            values.update(m.suppliers[i])
        K.write_row(bk, KEY, COLUMNS, rownum, values)

    ws.autofilter(r(C.ROW_HEADER), ci("B"), r(C.last_row(KEY)), ci(LAST_COL))
    ws.freeze_panes(r(C.ROW_FIRST), ci("C"))

    # ------------------------------------------------------------------
    # validation
    # ------------------------------------------------------------------
    K.list_dv(bk, KEY, "category", "ingredient_categories",
              title="Supplies what?",
              message="Broad category of what this supplier sells.")
    K.money_dv(bk, KEY, ("price", "min_order"))

    # ------------------------------------------------------------------
    # conditional formatting
    # ------------------------------------------------------------------
    K.databar(bk, KEY, "price", color=th.info)

    # ------------------------------------------------------------------
    # notes
    # ------------------------------------------------------------------
    trow = C.last_row(KEY) + 2
    K.note_block(
        bk, KEY, trow, "B", "M",
        ["Keep at least two suppliers for your critical ingredients - "
         "wedding weekends do not forgive stock-outs.",
         "\u201CDelivery time\u201D is the lead time from order to door: "
         "plan shopping runs around it.",
         "Terms like Net 15 mean you have 15 days to pay - useful for "
         "cash-flow timing before big events."],
        title="\U0001F4A1 How to use this tab")

    nav = trow + 6
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)
    bk.page(KEY, LAST_COL, nav + 1, freeze=(C.ROW_FIRST, 2),
            title_rows=(C.ROW_HEADER, C.ROW_HEADER))
```

### `catering_tracker/sheets/calendar.py`

```python
"""
\U0001F4C6 Event Calendar - a real monthly grid.

The month and year come from ⚙️ Setup (CalMonth / CalYear).  Each day cell
shows its day number plus the first event booked that day
(client • type), looked up live from the Events tab.  Days with events
glow in the accent colour.
"""

from datetime import date

from .. import config as C
from ..book import r, ci

KEY = "calendar"
LAST_COL = "H"
GRID = C.CAL_GRID_ROW
WEEKDAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"]


def _cell_formula(k):
    """Full day-cell formula: day number + first event of that day."""
    start = "DATE(CalYear,CalMonth,1)"
    n = "%d+2-WEEKDAY(%s,1)" % (k, start)
    dim = "DAY(EOMONTH(%s,0))" % start
    d = "DATE(CalYear,CalMonth,%s)" % n
    ev = "'%s'" % C.SHEET_NAMES["events"]
    first, last = C.ROW_FIRST, C.CAP["events"] + C.ROW_FIRST - 1
    edates = "%s!$E$%d:$E$%d" % (ev, first, last)
    eclient = "%s!$D$%d:$D$%d" % (ev, first, last)
    etype = "%s!$F$%d:$F$%d" % (ev, first, last)
    estat = "%s!$Q$%d:$Q$%d" % (ev, first, last)
    hit = 'MATCH(1,INDEX((%s=%s)*(%s<>"%s"),0),1)' % (edates, d, estat,
                                                      C.ES_CANCEL)
    label = ('IFERROR(INDEX(%s,%s)&" \u2022 "&INDEX(%s,%s),"")'
             % (eclient, hit, etype, hit))
    return '=IF(OR(%s<1,%s>%s),"",%s&IF(%s="","",CHAR(10)&%s))' % (
        n, n, dim, n, label, label)


def _demo_cells(m):
    """Cached cell text for the sample month (September 2026)."""
    out = {}
    if not m.events:
        return out
    cm, cy = m.settings["cal_month"], m.settings["cal_year"]
    first = date(cy, cm, 1)
    offset = (first.weekday() + 1) % 7        # Sunday-first offset
    dim = (date(cy + (cm == 12), (cm % 12) + 1, 1) -
           date(cy, cm, 1)).days
    by_day = {}
    for e in m.events:
        if e["date"].year == cy and e["date"].month == cm and \
                e["status"] != C.ES_CANCEL:
            by_day.setdefault(e["date"].day, e)
    for k in range(42):
        day = k - offset + 1
        if 1 <= day <= dim:
            e = by_day.get(day)
            text = str(day)
            if e:
                text += "\n%s \u2022 %s" % (e["client"], e["type"])
            out[k] = text
        else:
            out[k] = ""
    return out


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY, {"A": 2.2, "B": 15.5, "C": 15.5, "D": 15.5, "E": 15.5,
                    "F": 15.5, "G": 15.5, "H": 15.5})
    foot = GRID + 6 + 12
    bk.paint(KEY, 0, 0, foot, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F4C6  Event Calendar",
        "  Change the month on \u2699\uFE0F Setup (Calendar month / year) "
        "and this grid redraws itself from the Events tab.",
        LAST_COL)

    # ------------------------------------------------------------------
    # stat chips: what the displayed month holds
    # ------------------------------------------------------------------
    start = "DATE(CalYear,CalMonth,1)"
    end = "EOMONTH(DATE(CalYear,CalMonth,1),0)"
    ev = bk.q("events")
    edates = bk.rng("events", "date")
    eprice = bk.rng("events", "price")
    eguests = bk.rng("events", "guests")
    estat = bk.rng("events", "status")
    demo_cells = _demo_cells(m)
    n_events = sum(1 for k, v in demo_cells.items() if "\n" in v)
    guests = 0
    booked = 0
    if m.events:
        cm, cy = m.settings["cal_month"], m.settings["cal_year"]
        for e in m.events:
            if e["date"].year == cy and e["date"].month == cm and \
                    e["status"] != C.ES_CANCEL:
                guests += e["guests"]
                booked += e["price"]
    chips = [
        ('\U0001F4C5 "&TEXT(%s,"mmmm")&": "&COUNTIFS(%s,">="&%s,%s,"<="&%s,'
         '%s,"<>%s")&" event(s)"'
         % (start, edates, start, edates, end, estat, C.ES_CANCEL),
         "primary",
         "%s: %d event(s)" % (C.MONTH_NAMES[m.settings["cal_month"] - 1],
                              n_events), 2),
        ('\U0001F465 Guests: "&SUMIFS(%s,%s,">="&%s,%s,"<="&%s,%s,"<>%s")'
         % (eguests, edates, start, edates, end, estat, C.ES_CANCEL),
         "info", "Guests: %d" % guests, 2),
        ('\U0001F4B0 Booked: "&Currency&TEXT(SUMIFS(%s,%s,">="&%s,%s,'
         '"<="&%s,%s,"<>%s"),"#,##0")'
         % (eprice, edates, start, edates, end, estat, C.ES_CANCEL),
         "ok", "Booked: %s" % m.money(booked), 2),
    ]
    col = 1
    for formula, color, cached, span in chips:
        fmt = S.pill(th.soft(color), getattr(th, color), size=10.5,
                     bold=True, align="left")
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), col + span - 1,
                       "", fmt)
        ws.write_formula(r(C.ROW_STATS), col, '="' + formula, fmt, cached)
        bk.stats["formulas"] += 1
        col += span
    if col < ci(LAST_COL):
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), ci(LAST_COL),
                       "", S.canvas)
    elif col == ci(LAST_COL):
        ws.write_blank(r(C.ROW_STATS), col, None, S.canvas)

    # ------------------------------------------------------------------
    # weekday header + grid
    # ------------------------------------------------------------------
    hdr = GRID - 1
    ws.set_row(r(hdr), 20)
    for d, name in enumerate(WEEKDAYS):
        weekend = d in (0, 6)
        ws.write(r(hdr), 1 + d, name, S.header(th.primary_2 if weekend
                                               else th.primary))
    day_fmt = S.f(**S.base(font_size=11, bold=True, font_color=th.ink,
                           bg_color=th.card, align="left", valign="top",
                           text_wrap=True, border=1, border_color=th.border,
                           indent=1))
    for k in range(42):
        row = GRID + k // 7
        col = 1 + k % 7
        ws.set_row(r(row), C.CAL_CELL_H)
        cached = demo_cells.get(k, "") if m.events else ""
        ws.write_formula(r(row), col, _cell_formula(k), day_fmt, cached)
        bk.stats["formulas"] += 1

    # event days glow
    first_cell_r, first_cell_c = GRID, 1
    last_cell_r, last_cell_c = GRID + 5, 7
    bk.cond(KEY, first_cell_r, first_cell_c, last_cell_r, last_cell_c, {
        "type": "formula",
        "criteria": '=ISNUMBER(SEARCH("\u2022",B%d))' % GRID,
        "format": S.cf(bg=th.gold_soft, fg=th.ink, bold=True,
                       border=th.gold)})

    # ------------------------------------------------------------------
    # legend + notes
    # ------------------------------------------------------------------
    leg = GRID + 6 + 1
    ws.set_row(r(leg), 22)
    ws.merge_range(r(leg), 1, r(leg), ci(LAST_COL),
                   "  \U0001F511  LEGEND", S.section_soft)
    legend = [
        ("Gold cell", "an event is booked that day (client \u2022 type)",
         th.gold_soft, th.ink),
        ("\U0001F4AC / \U0001F4DD / \U0001F4B0 / \u2705",
         "pipeline stage colours live on the \U0001F4C5 Events tab",
         th.primary_soft, th.primary),
        ("Cancelled events", "are left off the calendar on purpose",
         th.bad_soft, th.bad),
    ]
    for j, (chip, text, bg, fg) in enumerate(legend):
        row = leg + 1 + j
        ws.set_row(r(row), 20)
        ws.write(r(row), 1, chip, S.pill(bg, fg, size=10, align="center"))
        ws.merge_range(r(row), 2, r(row), ci(LAST_COL), text,
                       S.f(**S.base(font_size=10, font_color=th.ink,
                                    bg_color=th.card, align="left",
                                    valign="vcenter", indent=1)))

    note = leg + 5
    ws.set_row(r(note), 20)
    ws.set_row(r(note + 1), 20)
    ws.merge_range(r(note), 1, r(note + 1), ci(LAST_COL),
                   "  \U0001F4A1  Two events on one day? The cell shows the "
                   "first one - the \U0001F4C5 Events tab (sorted by date) "
                   "shows everything. Print this page and pin it to the "
                   "kitchen wall.", S.note)

    nav = note + 3
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)
    bk.page(KEY, LAST_COL, nav + 1, zoom=95)
```

### `catering_tracker/sheets/reports.py`

```python
"""
\U0001F4C8 P&L & Reports.

  * Annual P&L card: revenue, COGS, gross profit, operating expenses,
    net profit, margins.
  * Month-by-month table (all formulas read the hidden _Data pools, which
    read the Payments / Expenses / Events tabs) + combo chart.
  * Revenue by event type + bar chart.
  * Top clients and best-selling menu items tables.
"""

from .. import config as C
from ..book import r, ci

KEY = "reports"
LAST_COL = "N"
PL_FIRST = 10               # monthly table first row (12 rows)


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo
    d = bk.q("data")
    months = m.agg.get("months") or [("", 0, 0, 0, 0)] * 12

    bk.widths(KEY, {"A": 2.2, "B": 13, "C": 13, "D": 13, "E": 13, "F": 11,
                    "G": 3, "H": 20, "I": 13, "J": 3, "K": 22, "L": 12,
                    "M": 3, "N": 4})
    bk.paint(KEY, 0, 0, 74, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F4C8  Profit & Loss \u00B7 Reports",
        "  The year at a glance, month by month \u2014 everything below is "
        "calculated from your tabs (reporting year: \u2699\uFE0F Setup).",
        LAST_COL)

    lbl = S.f(**S.base(font_size=10.5, bold=True, font_color=th.ink,
                       bg_color=th.card, align="left", valign="vcenter",
                       border=1, border_color=th.border, indent=1))
    val = S.f(**S.base(font_size=12, bold=True, font_color=th.primary,
                       bg_color=th.primary_soft, align="center",
                       valign="vcenter", border=1, border_color=th.border,
                       num_format="#,##0.00"))
    val_pct = S.f(**S.base(font_size=12, bold=True, font_color=th.primary,
                           bg_color=th.primary_soft, align="center",
                           valign="vcenter", border=1,
                           border_color=th.border, num_format="0.0%"))
    hero = S.f(**S.base(font_name=th.title_font, font_size=15, bold=True,
                        font_color=th.white, bg_color=th.ok, align="center",
                        valign="vcenter", border=1, border_color=th.ok,
                        num_format="#,##0.00"))

    # ------------------------------------------------------------------
    # annual P&L card
    # ------------------------------------------------------------------
    mf, ml = C.DATA_MONTH_FIRST, C.DATA_MONTH_FIRST + 11
    mrev = "%s!$I$%d:$I$%d" % (d, mf, ml)
    mexp = "%s!$J$%d:$J$%d" % (d, mf, ml)
    rev_c = m.agg.get("revenue", 0)
    cogs_c = m.agg.get("food_cost", 0)
    opex_c = m.agg.get("expenses", 0) - cogs_c
    net_c = rev_c - m.agg.get("expenses", 0)
    ws.set_row(r(8), 22)
    ws.merge_range(r(8), 1, r(8), ci("F"), "", S.section_soft)
    ws.write_formula(r(8), 1,
                     '="  \U0001F4CA  ANNUAL P&L  \u00B7  "&ReportYear',
                     S.section_soft,
                     "  \U0001F4CA  ANNUAL P&L  \u00B7  %d"
                     % m.settings["year"])
    bk.stats["formulas"] += 1
    pl_rows = [
        ("Revenue (cash received)", "=SUM(%s)" % mrev, rev_c, val),
        ("Cost of goods sold (ingredients + packaging)", "=%s"
         % bk.kpi("food_cost"), cogs_c, val),
        ("GROSS PROFIT", "=IF($D$%d=\"\",\"\",$D$%d-$D$%d)"
         % (PL_FIRST - 1, PL_FIRST - 1, PL_FIRST), None, hero),
        ("Operating expenses (everything else)", "=%s-%s"
         % (bk.kpi("expenses"), bk.kpi("food_cost")), opex_c, val),
        ("NET PROFIT", "=IF($D$%d=\"\",\"\",$D$%d-$D$%d)"
         % (PL_FIRST + 2, PL_FIRST, PL_FIRST + 2), None, hero),
        ("Gross margin %", "=IFERROR($D$%d/$D$%d,0)"
         % (PL_FIRST + 1, PL_FIRST - 1), None, val_pct),
        ("Net margin %", "=IFERROR($D$%d/$D$%d,0)"
         % (PL_FIRST + 3, PL_FIRST - 1), None, val_pct),
    ]
    gross_c = rev_c - cogs_c
    cached_pl = [rev_c, cogs_c, gross_c, opex_c, net_c,
                 gross_c / rev_c if rev_c else 0,
                 net_c / rev_c if rev_c else 0]
    for j, ((label, formula, _c, fmt), cached) in enumerate(zip(pl_rows,
                                                                cached_pl)):
        row = PL_FIRST - 1 + j        # starts at row 9
        big = fmt is hero
        ws.set_row(r(row), 28 if big else 22)
        ws.merge_range(r(row), 1, r(row), 2, label,
                       lbl if not big else S.f(**S.base(
                           font_size=11, bold=True, font_color=th.white,
                           bg_color=th.ok, align="left", valign="vcenter",
                           border=1, border_color=th.ok, indent=1)))
        ws.merge_range(r(row), 3, r(row), 5, "", fmt)
        ws.write_formula(r(row), 3, formula, fmt, cached)
        bk.stats["formulas"] += 1
        ws.write_blank(r(row), 6, None, S.canvas)
    annual_end = PL_FIRST - 1 + len(pl_rows) - 1        # row 15

    # ------------------------------------------------------------------
    # monthly table
    # ------------------------------------------------------------------
    mtop = annual_end + 2
    ws.set_row(r(mtop), 22)
    ws.merge_range(r(mtop), 1, r(mtop), ci("F"),
                   "  \U0001F4C5  MONTH BY MONTH", S.section_soft)
    hrow = mtop + 1
    ws.set_row(r(hrow), 26)
    for j, head in enumerate(["Month", "Cash in", "Expenses", "Profit",
                              "Events"]):
        ws.write(r(hrow), 1 + j, head, S.header(th.primary))
    ws.write_blank(r(hrow), 6, None, S.canvas)
    m_first = hrow + 1
    for i in range(12):
        row = m_first + i
        drow = C.DATA_MONTH_FIRST + i
        a = i % 2
        ws.set_row(r(row), 18)
        ws.write_formula(r(row), 1, "=%s!$H$%d" % (d, drow),
                         S.cell("center", a), months[i][0])
        ws.write_formula(r(row), 2, "=%s!$I$%d" % (d, drow),
                         S.cell("calc_money", a), months[i][1])
        bk.stats["formulas"] += 2
        ws.write_formula(r(row), 3, "=%s!$J$%d" % (d, drow),
                         S.cell("calc_money", a), months[i][2])
        ws.write_formula(r(row), 4, "=%s!$K$%d" % (d, drow),
                         S.cell("calc_money", a), months[i][3])
        ws.write_formula(r(row), 5, "=%s!$L$%d" % (d, drow),
                         S.cell("calc_num", a), months[i][4])
        bk.stats["formulas"] += 3
        ws.write_blank(r(row), 6, None, S.canvas)
    m_last = m_first + 11
    trow = m_last + 1
    ws.set_row(r(trow), 22)
    ws.write(r(trow), 1, "YEAR", S.header(th.accent))
    year_cached = [sum(mm[1] for mm in months), sum(mm[2] for mm in months),
                   sum(mm[3] for mm in months)]
    for j, colL in enumerate(["B", "C", "D"]):
        ws.write_formula(r(trow), 2 + j,
                         "=SUM($%s$%d:$%s$%d)" % (colL, m_first, colL,
                                                  m_last),
                         S.f(**S.base(font_size=11, bold=True,
                                      font_color=th.white, bg_color=th.accent,
                                      align="right", valign="vcenter",
                                      border=1, border_color=th.accent,
                                      num_format="#,##0.00")),
                         year_cached[j])
        bk.stats["formulas"] += 1
    ws.write_formula(r(trow), 5,
                     "=SUM($E$%d:$E$%d)" % (m_first, m_last),
                     S.f(**S.base(font_size=11, bold=True,
                                  font_color=th.white, bg_color=th.accent,
                                  align="center", valign="vcenter",
                                  border=1, border_color=th.accent,
                                  num_format="0")),
                     sum(mm[4] for mm in months))
    bk.stats["formulas"] += 1
    ws.write_blank(r(trow), 6, None, S.canvas)

    # CF: profit column red when negative
    bk.cond(KEY, m_first, 4, m_last, 4, {
        "type": "cell", "criteria": "<", "value": 0,
        "format": S.cf(bg=th.bad_soft, fg=th.bad, bold=True)})

    # ------------------------------------------------------------------
    # revenue by event type (right column) + top clients + menu items
    # ------------------------------------------------------------------
    ws.set_row(r(mtop), 22)
    ws.merge_range(r(mtop), 7, r(mtop), ci("I"),
                   "  \U0001F3AF  REVENUE BY EVENT TYPE", S.section_soft)
    ws.set_row(r(hrow), 26)
    ws.write(r(hrow), 7, "Event type", S.header(th.primary))
    ws.write(r(hrow), 8, "Booked value", S.header(th.primary))
    tf, tl = C.DATA_TYPE_FIRST, C.DATA_TYPE_FIRST + len(C.EVENT_TYPES) - 1
    types_cached = dict(m.agg.get("types") or [])
    for i in range(len(C.EVENT_TYPES)):
        row = m_first + i
        a = i % 2
        ws.set_row(r(row), 18)
        ws.write_formula(r(row), 7, "=%s!$N$%d" % (d, tf + i),
                         S.cell("center", a), C.EVENT_TYPES[i])
        ws.write_formula(r(row), 8, "=%s!$O$%d" % (d, tf + i),
                         S.cell("calc_money", a),
                         types_cached.get(C.EVENT_TYPES[i], 0))
        bk.stats["formulas"] += 2
    t_last_type = m_first + len(C.EVENT_TYPES) - 1

    ctop = t_last_type + 2
    ws.set_row(r(ctop), 22)
    ws.merge_range(r(ctop), 7, r(ctop), ci("I"),
                   "  \U0001F451  TOP CLIENTS (cash received)",
                   S.section_soft)
    ws.set_row(r(ctop + 1), 26)
    ws.write(r(ctop + 1), 7, "Client", S.header(th.primary))
    ws.write(r(ctop + 1), 8, "Received", S.header(th.primary))
    cf_, cl_ = C.DATA_CLIENT_FIRST, C.DATA_CLIENT_FIRST + 9
    clients_cached = (m.agg.get("clients_pool") or [])[:10]
    for i in range(10):
        row = ctop + 2 + i
        a = i % 2
        ws.set_row(r(row), 18)
        ws.write_formula(r(row), 7, "=%s!$Q$%d" % (d, cf_ + i),
                         S.cell("text", a),
                         clients_cached[i][0] if i < len(clients_cached)
                         else "")
        ws.write_formula(r(row), 8, "=%s!$R$%d" % (d, cf_ + i),
                         S.cell("calc_money", a),
                         clients_cached[i][1] if i < len(clients_cached)
                         else 0)
        bk.stats["formulas"] += 2
    c_last = ctop + 11

    # best-selling menu items (columns K..L)
    ws.merge_range(r(mtop), 10, r(mtop), ci("L"),
                   "  \U0001F37D\uFE0F  MOST-BOOKED DISHES", S.section_soft)
    ws.set_row(r(hrow), 26)
    ws.write(r(hrow), 10, "Menu item", S.header(th.primary))
    ws.write(r(hrow), 11, "On menus", S.header(th.primary))
    mf_, ml_ = C.DATA_MENU_FIRST, C.DATA_MENU_FIRST + 13
    menu_cached = (m.agg.get("menu_pool") or [])[:14]
    for i in range(14):
        row = m_first + i
        a = i % 2
        ws.set_row(r(row), 18)
        ws.write_formula(r(row), 10, "=%s!$T$%d" % (d, mf_ + i),
                         S.cell("text", a),
                         menu_cached[i][0] if i < len(menu_cached) else "")
        ws.write_formula(r(row), 11, "=%s!$U$%d" % (d, mf_ + i),
                         S.cell("calc_num", a),
                         menu_cached[i][1] if i < len(menu_cached) else 0)
        bk.stats["formulas"] += 2
    menu_last = m_first + 13

    # ------------------------------------------------------------------
    # charts
    # ------------------------------------------------------------------
    ch_top = max(c_last, menu_last, m_last, trow) + 2
    qname = bk.q(KEY)
    cats = "=%s!$B$%d:$B$%d" % (qname, m_first, m_last)

    ch = bk.chart("column")
    ch.add_series({
        "name": "Cash in", "categories": cats,
        "values": "=%s!$C$%d:$C$%d" % (qname, m_first, m_last),
        "fill": {"color": th.primary}, "border": {"color": th.primary},
        "gap": 60})
    ch.add_series({
        "name": "Expenses", "categories": cats,
        "values": "=%s!$D$%d:$D$%d" % (qname, m_first, m_last),
        "fill": {"color": th.accent}, "border": {"color": th.accent}})
    line = bk.chart("line")
    line.add_series({
        "name": "Profit", "categories": cats,
        "values": "=%s!$E$%d:$E$%d" % (qname, m_first, m_last),
        "line": {"color": th.ok, "width": 2.5},
        "marker": {"type": "circle", "size": 5,
                   "fill": {"color": th.ok}}})
    ch.combine(line)
    ch.set_title({"name": "Cash in vs expenses vs profit, by month",
                  "name_font": {"size": 12, "bold": True, "color": th.primary,
                                "name": th.body_font}})
    ch.set_y_axis({"num_format": "#,##0", "num_font": {"size": 9},
                   "major_gridlines": {"visible": True,
                                       "line": {"color": th.border}}})
    ch.set_x_axis({"num_font": {"size": 9}})
    ch.set_size({"width": 620, "height": 300})
    ch.set_legend({"position": "bottom", "font": {"size": 9}})
    ws.insert_chart(r(ch_top), 1, ch, {"x_offset": 6, "y_offset": 6})

    ch2 = bk.chart("bar")
    ch2.add_series({
        "name": "Booked value",
        "categories": "=%s!$H$%d:$H$%d" % (qname, m_first, t_last_type),
        "values": "=%s!$I$%d:$I$%d" % (qname, m_first, t_last_type),
        "fill": {"color": th.gold}, "border": {"color": th.gold},
        "gap": 45,
        "data_labels": {"value": True, "font": {"size": 8, "color": th.ink},
                        "num_format": "#,##0"}})
    ch2.set_title({"name": "Revenue by event type",
                   "name_font": {"size": 12, "bold": True,
                                 "color": th.primary, "name": th.title_font}})
    ch2.set_x_axis({"num_format": "#,##0", "num_font": {"size": 9},
                    "major_gridlines": {"visible": True,
                                        "line": {"color": th.border}}})
    ch2.set_y_axis({"num_font": {"size": 9}, "label_position": "low"})
    ch2.set_size({"width": 560, "height": 300})
    ch2.set_legend({"none": True})
    ws.insert_chart(r(ch_top), 7, ch2, {"x_offset": 6, "y_offset": 6})

    # ------------------------------------------------------------------
    # footer
    # ------------------------------------------------------------------
    foot = ch_top + 17
    ws.set_row(r(foot), 20)
    ws.set_row(r(foot + 1), 20)
    ws.merge_range(r(foot), 1, r(foot + 1), ci(LAST_COL),
                   "  \U0001F4A1  Cash basis: \u201CCash in\u201D counts "
                   "money actually received (deposits + payments) in the "
                   "month of the event; expenses count on the day you "
                   "spent them. Switch the reporting year on \u2699\uFE0F "
                   "Setup. For official accounts, always reconcile with "
                   "your accountant.", S.note)
    nav = foot + 3
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)
    bk.page(KEY, LAST_COL, nav + 1, freeze=(0, 0), zoom=85)
```

### `catering_tracker/sheets/tax.py`

```python
"""
\U0001F9FE Tax Tracker - sales tax collected, tax already paid, and what is
still due.  With the required "not professional advice" disclaimer.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "tax"
LAST_COL = "H"
TC = C.TAX_COLS


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY, {"A": 2.2, "B": 4.5, "C": 14, "D": 18, "E": 16, "F": 16,
                    "G": 15, "H": 16})
    log_last = C.TAX_LOG_FIRST + C.TAX_LOG_ROWS - 1
    bk.paint(KEY, 0, 0, log_last + 22, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F9FE  Tax Tracker",
        "  What you collected, what you already paid over, and what is "
        "still due \u2014 ready for filing day.",
        LAST_COL)

    lbl = S.f(**S.base(font_size=11, bold=True, font_color=th.ink,
                       bg_color=th.card, align="left", valign="vcenter",
                       border=1, border_color=th.border, indent=1))
    val = S.f(**S.base(font_size=13, bold=True, font_color=th.primary,
                       bg_color=th.primary_soft, align="center",
                       valign="vcenter", border=1, border_color=th.border,
                       num_format="#,##0.00"))
    val_pct = S.f(**S.base(font_size=13, bold=True, font_color=th.primary,
                           bg_color=th.primary_soft, align="center",
                           valign="vcenter", border=1,
                           border_color=th.border, num_format="0.0%"))
    hero = S.f(**S.base(font_name=th.title_font, font_size=16, bold=True,
                        font_color=th.white, bg_color=th.accent,
                        align="center", valign="vcenter", border=1,
                        border_color=th.accent, num_format="#,##0.00"))

    ws.set_row(r(8), 22)
    ws.merge_range(r(8), 1, r(8), ci(LAST_COL),
                   "  \U0001F4CA  SUMMARY", S.section_soft)

    inc_amt = bk.rng("income", "amount")
    taxable_cached = sum(i["amount"] for i in m.income)
    collected_cached = m.agg.get("tax_collected", 0)
    paid_cached = sum(t["amount"] for t in m.tax_paid)
    due_cached = m.agg.get("tax_due", 0)

    rows = [
        ("Taxable sales (all invoices)", "=SUM(%s)" % inc_amt,
         taxable_cached, val, "#,##0.00"),
        ("Tax rate (from \u2699\uFE0F Setup)", "=TaxRate",
         m.settings["tax"], val_pct, "0.0%"),
        ("Tax collected on sales", "=ROUND($G$%d*$G$%d,2)"
         % (C.TAX_SUM["sales"], C.TAX_SUM["rate"]), collected_cached,
         val, "#,##0.00"),
        ("Tax already paid (log below)", "=SUM($G$%d:$G$%d)"
         % (C.TAX_LOG_FIRST, log_last), paid_cached, val, "#,##0.00"),
        ("ESTIMATED TAX STILL DUE", "=ROUND($G$%d-$G$%d,2)"
         % (C.TAX_SUM["collected"], C.TAX_SUM["paid"]), due_cached,
         hero, "#,##0.00"),
    ]
    for j, (label, formula, cached, fmt, _nf) in enumerate(rows):
        row = C.TAX_SUM["sales"] + j
        ws.set_row(r(row), 30 if j == len(rows) - 1 else 24)
        ws.merge_range(r(row), 1, r(row), 5, label,
                       lbl if j < len(rows) - 1 else S.f(**S.base(
                           font_size=12, bold=True, font_color=th.white,
                           bg_color=th.accent, align="left",
                           valign="vcenter", border=1, border_color=th.accent,
                           indent=1)))
        ws.merge_range(r(row), 6, r(row), 7, "", fmt)
        ws.write_formula(r(row), 6, formula, fmt, cached)
        bk.stats["formulas"] += 1
    ws.set_row(r(14), 8)

    # ------------------------------------------------------------------
    # payments log
    # ------------------------------------------------------------------
    ws.set_row(r(C.TAX_LOG_HDR - 1), 22)
    ws.merge_range(r(C.TAX_LOG_HDR - 1), 1, r(C.TAX_LOG_HDR - 1),
                   ci(LAST_COL),
                   "  \U0001F4DD  TAX PAYMENTS LOG  \u2014  every payment "
                   "you make to the taxman", S.section_soft)
    ws.set_row(r(C.TAX_LOG_HDR), 24)
    ws.write(r(C.TAX_LOG_HDR), ci(TC["n"]), "#", S.header(th.primary))
    ws.write(r(C.TAX_LOG_HDR), ci(TC["date"]), "Date paid",
             S.header(th.primary))
    ws.merge_range(r(C.TAX_LOG_HDR), ci(TC["desc"]), r(C.TAX_LOG_HDR),
                   ci("F"), "What for", S.header(th.primary))
    ws.write(r(C.TAX_LOG_HDR), ci(TC["amount"]), "Amount",
             S.header(th.primary))
    ws.write(r(C.TAX_LOG_HDR), ci(TC["method"]), "Method",
             S.header(th.primary))

    for i in range(C.TAX_LOG_ROWS):
        row = C.TAX_LOG_FIRST + i
        a = i % 2
        ws.set_row(r(row), 20)
        ws.write(r(row), ci(TC["n"]), i + 1, S.idx(a))
        ws.write_blank(r(row), ci(TC["date"]), None, S.cell("date", a))
        ws.merge_range(r(row), ci(TC["desc"]), r(row), ci("F"), "",
                       S.cell("text", a))
        ws.write_blank(r(row), ci(TC["amount"]), None, S.cell("money", a))
        ws.write_blank(r(row), ci(TC["method"]), None, S.cell("center", a))
        if i < len(m.tax_paid):
            t = m.tax_paid[i]
            ws.write_datetime(r(row), ci(TC["date"]), t["date"],
                              S.cell("date", a))
            ws.write(r(row), ci(TC["desc"]), t["desc"], S.cell("text", a))
            ws.write(r(row), ci(TC["amount"]), t["amount"],
                     S.cell("money", a))

    ws.data_validation(r(C.TAX_LOG_FIRST), ci(TC["date"]), r(log_last),
                       ci(TC["date"]),
                       {"validate": "date", "criteria": ">=",
                        "value": "=DATE(2000,1,1)", "ignore_blank": True,
                        "input_title": "Date paid", "show_input": True,
                        "input_message": "When did you pay it?"})
    bk.stats["validations"] += 1
    ws.data_validation(r(C.TAX_LOG_FIRST), ci(TC["amount"]), r(log_last),
                       ci(TC["amount"]),
                       {"validate": "decimal", "criteria": ">=", "value": 0,
                        "ignore_blank": True})
    bk.stats["validations"] += 1
    bk.validate(KEY, C.TAX_LOG_FIRST, ci(TC["method"]), log_last,
                ci(TC["method"]), "=" + bk.listname("payment_methods"),
                title="Method")

    # ------------------------------------------------------------------
    # disclaimer + notes
    # ------------------------------------------------------------------
    drow = log_last + 2
    ws.set_row(r(drow), 20)
    ws.set_row(r(drow + 1), 20)
    ws.set_row(r(drow + 2), 20)
    ws.merge_range(r(drow), 1, r(drow + 2), ci(LAST_COL),
                   "  \u26A0\uFE0F  PLEASE NOTE: this tab is a simple "
                   "record-keeping helper only. It is NOT professional "
                   "tax, accounting or legal advice. Tax rules differ by "
                   "country, state and business structure \u2014 always "
                   "confirm your rates, registrations, deductions and "
                   "filing deadlines with a qualified accountant or tax "
                   "advisor.", S.note)

    nav = drow + 4
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)

    # CF: due > 0 glows
    bk.cond(KEY, C.TAX_SUM["due"], ci("G"), C.TAX_SUM["due"], ci("H"), {
        "type": "cell", "criteria": ">", "value": 0,
        "format": S.cf(bg=th.bad, fg=th.white, bold=True, size=16)})

    bk.page(KEY, LAST_COL, nav + 1, landscape=False, zoom=100)
```

### `catering_tracker/sheets/checklists.py`

```python
"""
✅ Checklists - the four lists from the spec:

  1. Event Prep (before the big day)
  2. Shopping Run
  3. Day-Of
  4. End-of-Event

Copy a block per event if you like, or just tick through them for the
event you are working on - the counters at the top update live.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "checklists"
LAST_COL = "H"

BLOCKS = [
    ("prep", "  \U0001F4CB  1 \u00B7 EVENT PREP  \u2014  the week before", [
        ("Confirm final guest count with client", "7 days out"),
        ("Confirm menu + dietary requirements in writing", "7 days out"),
        ("Confirm venue access time, power & parking", "5 days out"),
        ("Staff roster locked and briefed", "4 days out"),
        ("Equipment reserved and checked against tracker", "3 days out"),
        ("Supplier orders placed (lead times respected)", "3 days out"),
        ("Prep started: marinades, sauces, desserts", "2 days out"),
        ("Day-of timeline printed and shared with the team", "1 day out"),
    ]),
    ("shop", "  \U0001F6D2  2 \u00B7 SHOPPING RUN", [
        ("Shopping List tab reviewed (To buy column)", "day before"),
        ("Cold chain planned: coolers, ice packs", "day before"),
        ("Proteins bought first, kept cold", "morning of"),
        ("Produce picked over for quality", "morning of"),
        ("Receipts kept for the \U0001F4B8 Expenses tab", "every trip"),
        ("\U0001F4E6 Inventory quantities updated", "same day"),
        ("Crates labelled per event", "on return"),
        ("Perishables chilled / frozen immediately", "on return"),
    ]),
    ("day", "  \U0001F525  3 \u00B7 DAY-OF", [
        ("Load-out checked against Equipment tracker", "departure"),
        ("Venue walkthrough with the client contact", "on arrival"),
        ("Chafers / ranges set up and test-fired", "setup"),
        ("Hand-wash & sanitising station out", "setup"),
        ("Team briefing: allergies, timings, roles", "1 hr before"),
        ("Food temps logged (hot >63\u00B0C, cold <5\u00B0C)", "before service"),
        ("Backup supplies within reach", "before service"),
        ("Service start time confirmed with venue", "before service"),
    ]),
    ("end", "  \U0001F3C1  4 \u00B7 END-OF-EVENT", [
        ("Leftovers packed, labelled, handed to client", "at close"),
        ("Equipment counted back into the van", "at close"),
        ("Venue left as found - final walk with contact", "at close"),
        ("Gas off, power off, temps logged", "at close"),
        ("Wages & tips settled \u2192 \U0001F477 Staff tab", "same night"),
        ("Actual costs logged \u2192 \U0001F4B8 Expenses tab", "same night"),
        ("Event marked \U0001F389 Completed on Events tab", "same night"),
        ("Final invoice sent + review/referral requested", "next day"),
    ]),
]

BLOCK_TOP = 9


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY, {"A": 2.2, "B": 4.5, "C": 22, "D": 22, "E": 18, "F": 13,
                    "G": 8, "H": 20})
    n_items = sum(len(items) for _k, _t, items in BLOCKS)
    foot = BLOCK_TOP + len(BLOCKS) * 2 + n_items + 8
    bk.paint(KEY, 0, 0, foot, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \u2705  The Four Checklists",
        "  Prep \u2192 Shop \u2192 Day-of \u2192 Close.  Tick as you go; "
        "the counters above each block update live.",
        LAST_COL)

    # ------------------------------------------------------------------
    # counters (one chip per block)
    # ------------------------------------------------------------------
    tick_col = "G"
    row = BLOCK_TOP
    ranges = {}
    for bkey, _title, items in BLOCKS:
        first = row + 1
        last = first + len(items) - 1
        ranges[bkey] = (first, last)
        row = last + 2

    colors = {"prep": "primary", "shop": "ok", "day": "accent", "end": "info"}
    short = {"prep": "\U0001F4CB Prep", "shop": "\U0001F6D2 Shop",
             "day": "\U0001F525 Day-of", "end": "\U0001F3C1 Close"}
    spans = [2, 2, 2, 1]
    col = 1
    for (bkey, _title, items), span in zip(BLOCKS, spans):
        first, last = ranges[bkey]
        rng = "$%s$%d:$%s$%d" % (tick_col, first, tick_col, last)
        done = len(m.checks.get(bkey, ()))
        fmt = S.pill(th.soft(colors[bkey]), getattr(th, colors[bkey]),
                     size=10.5, bold=True, align="left")
        if span > 1:
            ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS),
                           col + span - 1, "", fmt)
        ws.write_formula(
            r(C.ROW_STATS), col,
            '="%s: "&COUNTIF(%s,"%s")&"/%d"' % (short[bkey], rng, C.TICK,
                                                len(items)),
            fmt, "%s: %d/%d" % (short[bkey], done, len(items)))
        bk.stats["formulas"] += 1
        col += span
    if col < ci(LAST_COL):
        ws.merge_range(r(C.ROW_STATS), col, r(C.ROW_STATS), ci(LAST_COL),
                       "", S.canvas)
    elif col == ci(LAST_COL):
        ws.write_blank(r(C.ROW_STATS), col, None, S.canvas)

    # ------------------------------------------------------------------
    # the four blocks
    # ------------------------------------------------------------------
    item_fmt = S.f(**S.base(font_size=10.5, font_color=th.ink,
                            bg_color=th.card, align="left", valign="vcenter",
                            border=1, border_color=th.border, indent=1))
    item_fmt_alt = S.f(**S.base(font_size=10.5, font_color=th.ink,
                                bg_color=th.alt, align="left",
                                valign="vcenter", border=1,
                                border_color=th.border, indent=1))
    when_fmt = S.f(**S.base(font_size=9.5, italic=True, font_color=th.muted,
                            bg_color=th.card, align="center",
                            valign="vcenter", border=1,
                            border_color=th.border))
    when_fmt_alt = S.f(**S.base(font_size=9.5, italic=True,
                                font_color=th.muted, bg_color=th.alt,
                                align="center", valign="vcenter", border=1,
                                border_color=th.border))
    block_colors = {"prep": th.primary, "shop": th.ok, "day": th.accent,
                    "end": th.info}
    for bi, (bkey, title, items) in enumerate(BLOCKS):
        first, last = ranges[bkey]
        hrow = first - 1
        ws.set_row(r(hrow), 24)
        ws.merge_range(r(hrow), 1, r(hrow), ci(LAST_COL), title,
                       S.f(**S.base(font_name=th.title_font, font_size=12,
                                    bold=True, font_color=th.white,
                                    bg_color=block_colors[bkey],
                                    align="left", valign="vcenter",
                                    indent=1)))
        checked = m.checks.get(bkey, set())
        for i, (text, when) in enumerate(items):
            row = first + i
            a = i % 2
            ws.set_row(r(row), 20)
            ws.write(r(row), ci("B"), i + 1, S.idx(a))
            ws.merge_range(r(row), ci("C"), r(row), ci("E"), text,
                           item_fmt_alt if a else item_fmt)
            ws.write(r(row), ci("F"), when, when_fmt_alt if a else when_fmt)
            tick = C.TICK if i in checked else ""
            ws.write(r(row), ci("G"), tick, S.cell("tick", a))
            ws.write_blank(r(row), ci("H"), None,
                           S.cell("text", a))
        # tick validation + glow for this block
        ws.data_validation(r(first), ci("G"), r(last), ci("G"), {
            "validate": "list", "source": "=Tick", "ignore_blank": True,
            "input_title": "Done?", "show_input": True,
            "input_message": "Pick \u2713 from the dropdown when done."})
        bk.stats["validations"] += 1
        bk.cond(KEY, first, ci("C"), last, ci("G"), {
            "type": "formula",
            "criteria": '=$G%d="%s"' % (first, C.TICK),
            "format": S.cf(bg=th.ok_soft, fg=th.ok, bold=True,
                           border=th.border)})
        bk.cond(KEY, first, ci("G"), last, ci("G"), {
            "type": "formula",
            "criteria": '=$G%d="%s"' % (first, C.TICK),
            "format": S.cf(bg=th.ok_soft, fg=th.ok, bold=True, size=13,
                           border=th.border)})

    # ------------------------------------------------------------------
    # footer
    # ------------------------------------------------------------------
    frow = row + 1
    ws.set_row(r(frow), 20)
    ws.set_row(r(frow + 1), 20)
    ws.merge_range(r(frow), 1, r(frow + 1), ci(LAST_COL),
                   "  \U0001F4A1  Running several events a week? Duplicate "
                   "this sheet (right-click the tab \u2192 Move or Copy "
                   "\u2192 Create a copy) and rename it per event. The "
                   "counters on each copy work independently.", S.note)

    nav = frow + 3
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)
    bk.page(KEY, LAST_COL, nav + 1, landscape=False, zoom=100)
```

### `catering_tracker/sheets/invoice.py`

```python
"""
\U0001F5A8️ Invoice & Proposal - pick an Event ID and the whole document
fills itself: client, event details, money, deposit, balance and a PAID /
BALANCE DUE stamp.  Prints on one A4 page.  The bottom half doubles as a
simple proposal (scope + terms) you can send before the event.
"""

from .. import config as C
from ..book import r, ci

KEY = "invoice"
LAST_COL = "H"
SEL_ROW = 10                 # Event ID selector
ev_first, ev_last = C.ROW_FIRST, C.ROW_FIRST + C.CAP["events"] - 1


def _idx(colletter):
    """INDEX/MATCH pull from the events sheet by the selected Event ID."""
    ev = "'%s'" % C.SHEET_NAMES["events"]
    ids = "%s!$C$%d:$C$%d" % (ev, ev_first, ev_last)
    src = "%s!$%s$%d:$%s$%d" % (ev, colletter, ev_first, colletter, ev_last)
    return 'IFERROR(INDEX(%s,MATCH($D$%d,%s,0)),"")' % (src, SEL_ROW, ids)


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY, {"A": 2.5, "B": 16, "C": 16, "D": 16, "E": 14, "F": 14,
                    "G": 14, "H": 14})
    bk.paint(KEY, 0, 0, 66, ci(LAST_COL), S.canvas)

    # demo selection
    sel = ""
    ev = None
    if m.events:
        ev = next((e for e in m.events if e["id"] == "EV-2026-005"),
                  m.events[0])
        sel = ev["id"]

    # ------------------------------------------------------------------
    # letterhead
    # ------------------------------------------------------------------
    ws.set_row(r(1), 8)
    ws.set_row(r(2), 40)
    ws.merge_range(r(2), 1, r(2), ci(LAST_COL), "", S.hero_title)
    ws.write_formula(r(2), 1, '="\U0001F37D\uFE0F  "&BusinessName',
                     S.hero_title,
                     "\U0001F37D\uFE0F  %s" % m.settings["business"])
    bk.stats["formulas"] += 1
    ws.set_row(r(3), 22)
    ws.merge_range(r(3), 1, r(3), 4, "  INVOICE & PROPOSAL",
                   S.f(**S.base(font_name=th.title_font, font_size=15,
                                bold=True, font_color=th.accent,
                                bg_color=th.bg, align="left",
                                valign="vcenter")))
    issued_f = S.f(**S.base(font_size=10, italic=True, font_color=th.muted,
                            bg_color=th.bg, align="right", valign="vcenter"))
    ws.merge_range(r(3), 5, r(3), ci(LAST_COL), "", issued_f)
    ws.write_formula(r(3), 5, '="Issued: "&TEXT(TODAY(),"dd mmm yyyy")',
                     issued_f,
                     "Issued: 12 Sep 2026" if m.events else "Issued: ")
    bk.stats["formulas"] += 1
    ws.set_row(r(4), 8)

    # ------------------------------------------------------------------
    # selector
    # ------------------------------------------------------------------
    sel_lbl = S.f(**S.base(font_size=11, bold=True, font_color=th.white,
                           bg_color=th.primary, align="left",
                           valign="vcenter", border=1,
                           border_color=th.primary, indent=1))
    sel_inp = S.f(**S.base(font_size=12, bold=True, font_color=th.primary,
                           bg_color=th.gold_soft, align="center",
                           valign="vcenter", border=2, border_color=th.gold,
                           locked=False))
    ws.set_row(r(6), 8)
    ws.set_row(r(SEL_ROW - 1), 22)
    ws.merge_range(r(SEL_ROW - 1), 1, r(SEL_ROW - 1), ci(LAST_COL),
                   "  \U0001F3AF  Pick the event - everything below fills "
                   "itself in", S.section_soft)
    ws.set_row(r(SEL_ROW), 32)
    ws.merge_range(r(SEL_ROW), 1, r(SEL_ROW), 2, "Event ID", sel_lbl)
    ws.merge_range(r(SEL_ROW), 3, r(SEL_ROW), 4, "", sel_inp)
    ws.write(r(SEL_ROW), 3, sel, sel_inp)
    ws.merge_range(r(SEL_ROW), 5, r(SEL_ROW), ci(LAST_COL),
                   "   \u2190 dropdown lists every Event ID from the "
                   "\U0001F4C5 Events tab", S.note_plain)
    bk.validate(KEY, SEL_ROW, 3, SEL_ROW, 3, "=EventList",
                title="Event ID",
                message="Which event is this document for?")

    # ------------------------------------------------------------------
    # pulled details
    # ------------------------------------------------------------------
    lbl = S.f(**S.base(font_size=10, bold=True, font_color=th.muted,
                       bg_color=th.card, align="left", valign="vcenter",
                       border=1, border_color=th.border, indent=1))
    valc = S.f(**S.base(font_size=11, bold=True, font_color=th.ink,
                        bg_color=th.card, align="left", valign="vcenter",
                        border=1, border_color=th.border, indent=1))
    valc_c = S.f(**S.base(font_size=11, bold=True, font_color=th.ink,
                          bg_color=th.card, align="center", valign="vcenter",
                          border=1, border_color=th.border))
    valc_date = S.f(**S.base(font_size=11, bold=True, font_color=th.ink,
                             bg_color=th.card, align="center",
                             valign="vcenter", border=1,
                             border_color=th.border,
                             num_format="dd mmm yyyy"))

    def pull(row, label, colletter, fmt, cached, lcols, vcols):
        ws.set_row(r(row), 22)
        ws.merge_range(r(row), lcols[0], r(row), lcols[1], label, lbl)
        if vcols[0] == vcols[1]:
            ws.write_formula(r(row), vcols[0], "=" + _idx(colletter), fmt,
                             cached)
        else:
            ws.merge_range(r(row), vcols[0], r(row), vcols[1], "", fmt)
            ws.write_formula(r(row), vcols[0], "=" + _idx(colletter), fmt,
                             cached)
        bk.stats["formulas"] += 1

    d1 = 12
    pull(d1, "Client", "D", valc, ev["client"] if ev else "", (1, 2), (3, 4))
    pull(d1, "Event date", "E", valc_date, ev["date"] if ev else "",
         (5, 6), (7, 7))
    pull(d1 + 1, "Event type", "F", valc_c, ev["type"] if ev else "",
         (1, 2), (3, 4))
    pull(d1 + 1, "Guests", "G", valc_c, ev["guests"] if ev else "",
         (5, 6), (7, 7))
    pull(d1 + 2, "Menu", "H", valc, ev["menu"] if ev else "", (1, 2), (3, 7))
    pull(d1 + 3, "Notes", "R", valc, ev["notes"] if ev else "", (1, 2),
         (3, 7))

    # ------------------------------------------------------------------
    # money box
    # ------------------------------------------------------------------
    box = d1 + 5                     # row 17
    ws.set_row(r(box - 1), 8)
    ws.set_row(r(box), 22)
    ws.merge_range(r(box), 1, r(box), ci(LAST_COL),
                   "  \U0001F4B0  THE MONEY", S.section_soft)
    r_quote = box + 1
    r_tax = box + 2
    r_total = box + 3
    r_dep = box + 4
    r_bal = box + 5

    def money_line(row, label, formula, cached, bg, big=False):
        ws.set_row(r(row), 26 if big else 22)
        lab_f = S.f(**S.base(font_size=12 if big else 11, bold=True,
                             font_color=th.primary if big else th.ink,
                             bg_color=bg, align="left", valign="vcenter",
                             border=1, border_color=th.border, indent=1))
        val_f = S.f(**S.base(font_size=14 if big else 12, bold=True,
                             font_color=th.ink, bg_color=bg, align="right",
                             valign="vcenter", border=1,
                             border_color=th.border, indent=1,
                             num_format="#,##0.00"))
        ws.merge_range(r(row), 1, r(row), 4, label, lab_f)
        ws.merge_range(r(row), 5, r(row), ci(LAST_COL), "", val_f)
        ws.write_formula(r(row), 5, formula, val_f, cached)
        bk.stats["formulas"] += 1

    price = ev["price"] if ev else ""
    dep = ev["deposit"] if ev else ""
    tax = m.settings["tax"]
    total_c = round(ev["price"] * (1 + tax), 2) if ev else ""
    bal_c = round(total_c - ev["deposit"], 2) if ev else ""

    money_line(r_quote, "Quote / invoice total", "=" + _idx("L"), price,
               th.card)
    money_line(r_tax, "Tax at your Setup rate",
               '=IF($F$%d="","",ROUND($F$%d*TaxRate,2))' % (r_quote, r_quote),
               round(ev["price"] * tax, 2) if ev else "", th.card)
    money_line(r_total, "TOTAL",
               '=IF($F$%d="","",$F$%d+$F$%d)' % (r_quote, r_quote, r_tax),
               total_c, th.primary_soft, big=True)
    money_line(r_dep, "Deposit held", "=" + _idx("O"), dep, th.card)
    money_line(r_bal, "BALANCE DUE",
               '=IF($F$%d="","",$F$%d-$F$%d)' % (r_quote, r_total, r_dep),
               bal_c, th.gold_soft, big=True)

    # PAID / BALANCE DUE stamp
    stamp_row = r_bal + 2
    ws.set_row(r(stamp_row), 32)
    stamp_fmt = S.f(**S.base(font_name=th.title_font, font_size=15,
                             bold=True, font_color=th.white, bg_color=th.ok,
                             align="center", valign="vcenter", border=1,
                             border_color=th.ok))
    ws.merge_range(r(stamp_row), 1, r(stamp_row), ci(LAST_COL), "",
                   stamp_fmt)
    ws.write_formula(
        r(stamp_row), 1,
        '=IF($D$%d="","",IF($F$%d<=0,"\u2705 PAID IN FULL",'
        '"\u23F3 BALANCE DUE: "&Currency&TEXT($F$%d,"#,##0")))'
        % (SEL_ROW, r_bal, r_bal), stamp_fmt,
        ("\u23F3 BALANCE DUE: %s" % m.money(bal_c)) if ev else "")
    bk.stats["formulas"] += 1
    bk.cond(KEY, stamp_row, 1, stamp_row, ci(LAST_COL), {
        "type": "formula",
        "criteria": '=ISNUMBER(SEARCH("PAID",$B$%d))' % stamp_row,
        "format": S.cf(bg=th.ok, fg=th.white, bold=True, size=15)})
    bk.cond(KEY, stamp_row, 1, stamp_row, ci(LAST_COL), {
        "type": "formula",
        "criteria": '=ISNUMBER(SEARCH("BALANCE",$B$%d))' % stamp_row,
        "format": S.cf(bg=th.accent, fg=th.white, bold=True, size=15)})

    # ------------------------------------------------------------------
    # proposal / terms
    # ------------------------------------------------------------------
    terms_top = stamp_row + 2
    ws.set_row(r(terms_top), 22)
    ws.merge_range(r(terms_top), 1, r(terms_top), ci(LAST_COL),
                   "  \U0001F4DC  PROPOSAL TERMS  (edit these to match "
                   "your business)", S.section_soft)
    terms = [
        "1.  This quote is valid for 30 days from the issue date.",
        "2.  A deposit (see above) confirms your date; the balance is due "
        "on or before the event day.",
        "3.  Final guest count is required 72 hours before service; "
        "catering is prepared to that number.",
        "4.  Please advise allergies and dietary requirements with the "
        "final count.",
        "5.  Cancellations inside 7 days forfeit the deposit; inside 72 "
        "hours, 100% of the quote is chargeable.",
        "6.  Menu substitutions of equal or greater value may be made if "
        "ingredients are unavailable.",
    ]
    t_fmt = S.f(**S.base(font_size=10.5, font_color=th.ink, bg_color=th.card,
                         align="left", valign="vcenter", border=1,
                         border_color=th.border, indent=1, locked=False))
    for j, t in enumerate(terms):
        row = terms_top + 1 + j
        ws.set_row(r(row), 20)
        ws.merge_range(r(row), 1, r(row), ci(LAST_COL), t, t_fmt)

    sign = terms_top + len(terms) + 2
    ws.set_row(r(sign), 34)
    sig_f = S.f(**S.base(font_size=10, italic=True, font_color=th.muted,
                         bg_color=th.bg, align="center", valign="bottom",
                         top=1, top_color=th.border_strong))
    ws.merge_range(r(sign), 1, r(sign), 3, "Signed (client)", sig_f)
    ws.write_blank(r(sign), 4, None, S.canvas)
    ws.merge_range(r(sign), 5, r(sign), 7, "Date", sig_f)

    thanks = sign + 2
    ws.set_row(r(thanks), 24)
    thanks_f = S.f(**S.base(font_name=th.title_font, font_size=12,
                            bold=True, italic=True, font_color=th.accent,
                            bg_color=th.bg, align="center",
                            valign="vcenter"))
    ws.merge_range(r(thanks), 1, r(thanks), ci(LAST_COL), "", thanks_f)
    ws.write_formula(
        r(thanks), 1,
        '="Thank you for choosing "&BusinessName&" \u2014 we cannot wait '
        'to feed your guests!"', thanks_f,
        "Thank you for choosing %s \u2014 we cannot wait to feed your "
        "guests!" % m.settings["business"])
    bk.stats["formulas"] += 1

    nav = thanks + 2
    ws.set_row(r(nav), 24)
    bk.nav_row(KEY, nav, max_col=LAST_COL)

    bk.page(KEY, LAST_COL, nav + 1, landscape=False, zoom=100, fit=True)
    ws.print_area("A1:%s%d" % (LAST_COL, thanks))
```

### `catering_tracker/sheets/guide.py`

```python
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
```

---

## 🧪 Appendix: QA Tooling Source

### `tools/verify_workbook.py`

```python
#!/usr/bin/env python3
"""
Audit a generated .xlsx without needing Excel.

Checks performed
    1. the workbook opens (openpyxl) and every expected tab exists / is hidden
    2. every sheet-qualified reference in every formula points at a real sheet
    3. no modern-only function leaked in without the ``_xlfn.`` prefix older
       Excel versions need
    4. no cell contains the literal strings "None", "#REF!" or "nan"
    5. nothing was written *under* a merged range (Excel only ever shows the
       top-left cell of a merge, so such writes would be invisible)

Usage
    python3 tools/verify_workbook.py products/*.xlsx
"""

import glob
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.utils import (column_index_from_string, get_column_letter,
                            range_boundaries)

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

# Functions that did not exist in Excel 2007 and therefore must be written as
# _xlfn.NAME by a file generator (XlsxWriter does this automatically for the
# names it knows about).
MODERN = {
    "TEXTJOIN", "CONCAT", "IFS", "SWITCH", "MAXIFS", "MINIFS", "XLOOKUP",
    "XMATCH", "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE", "LET",
    "LAMBDA", "RANDARRAY", "IMAGE", "TEXTSPLIT", "VSTACK", "HSTACK", "TOCOL",
    "TOROW", "TAKE", "DROP", "CHOOSEROWS", "CHOOSECOLS", "WRAPROWS",
    "WRAPCOLS", "GROUPBY", "PIVOTBY", "REGEXTEST", "REGEXEXTRACT",
    "REGEXREPLACE",
}

FUNC_RE = re.compile(r"(?<![A-Za-z0-9_.])([A-Z][A-Z0-9_.]*)\s*\(")
SHEET_REF_RE = re.compile(r"(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_.]*))!")


def sheet_xml(path):
    """{sheet name: worksheet xml} for low level checks."""
    zf = zipfile.ZipFile(path)
    wbxml = zf.read("xl/workbook.xml").decode()
    rels = dict(re.findall(
        r'Id="(rId\d+)"[^>]*Target="([^"]*worksheets/sheet\d+\.xml)"',
        zf.read("xl/_rels/workbook.xml.rels").decode()))
    out = {}
    for name, rid in re.findall(r'<sheet name="([^"]+)"[^>]*r:id="(rId\d+)"',
                                wbxml):
        out[name] = zf.read("xl/" + rels[rid].lstrip("/")).decode()
    return out, [n for n in zipfile.ZipFile(path).namelist()]


def hidden_writes(path):
    """Cells written inside a merged range but not at its top-left."""
    xmls, _ = sheet_xml(path)
    out = []
    for name, xml in xmls.items():
        covered = {}
        for m in re.findall(r'<mergeCell ref="([A-Z]+\d+:[A-Z]+\d+)"/>', xml):
            c1, r1, c2, r2 = range_boundaries(m)
            for rr in range(r1, r2 + 1):
                for cc in range(c1, c2 + 1):
                    if (rr, cc) != (r1, c1):
                        covered[(rr, cc)] = m
        root = ET.fromstring(xml)
        for row in root.iter(NS + "row"):
            for c in row.iter(NS + "c"):
                m = re.match(r"([A-Z]+)(\d+)", c.get("r"))
                key = (int(m.group(2)), column_index_from_string(m.group(1)))
                if key not in covered:
                    continue
                v = c.find(NS + "v")
                if c.find(NS + "f") is not None or (
                        v is not None and (v.text or "").strip()):
                    out.append((name, c.get("r"), covered[key]))
    return out


def check_formulas(path, problems):
    wb = openpyxl.load_workbook(path, data_only=False)
    sheetnames = set(wb.sheetnames)
    n_formula = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not (isinstance(v, str) and v.startswith("=")):
                    continue
                n_formula += 1
                if v in ("None", "nan"):
                    problems.append("%s!%s contains %r"
                                    % (ws.title, cell.coordinate, v))
                body = re.sub(r'"[^"]*"', '""', v)     # drop string literals
                for m in SHEET_REF_RE.finditer(body):
                    ref = m.group(1) or m.group(2)
                    if ref not in sheetnames:
                        problems.append("%s!%s references unknown sheet %r"
                                        % (ws.title, cell.coordinate, ref))
                for m in FUNC_RE.finditer(body):
                    fn = m.group(1).split(".")[-1]
                    if fn in MODERN and "_xlfn." not in m.group(1):
                        problems.append(
                            "%s!%s uses modern function %s without _xlfn."
                            % (ws.title, cell.coordinate, fn))
    return n_formula


def audit(path):
    problems = []
    wb = openpyxl.load_workbook(path, data_only=False)
    n_formula = check_formulas(path, problems)

    n_dv = sum(len(ws.data_validations.dataValidation)
               for ws in wb.worksheets)
    n_cf = sum(len(ws.conditional_formatting._cf_rules)
               for ws in wb.worksheets)
    _, names = sheet_xml(path)
    charts = len([n for n in names if n.startswith("xl/charts/chart")])

    for sheet, ref, merge in hidden_writes(path):
        problems.append("%s!%s is written inside merge %s and would never "
                        "be visible" % (sheet, ref, merge))

    return {
        "path": path,
        "sheets": wb.sheetnames,
        "hidden": [s.title for s in wb.worksheets
                   if s.sheet_state != "visible"],
        "formulas": n_formula,
        "validations": n_dv,
        "cond_formats": n_cf,
        "charts": charts,
        "problems": problems,
    }


def main(argv):
    paths = []
    for a in argv:
        paths.extend(sorted(glob.glob(a)) if any(c in a for c in "*?[") else [a])
    if not paths:
        print(__doc__)
        return 1
    bad = 0
    for path in paths:
        r = audit(path)
        print("=" * 78)
        print(path)
        print("  sheets   : %s" % ", ".join(r["sheets"]))
        print("  hidden   : %s" % (", ".join(r["hidden"]) or "-"))
        print("  formulas : %-6d validations %-4d cond. formats %-4d "
              "charts %d" % (r["formulas"], r["validations"],
                             r["cond_formats"], r["charts"]))
        if r["problems"]:
            bad += 1
            print("  PROBLEMS (%d):" % len(r["problems"]))
            for p in r["problems"][:25]:
                print("    - %s" % p)
            if len(r["problems"]) > 25:
                print("    ... and %d more" % (len(r["problems"]) - 25))
        else:
            print("  PROBLEMS : none \u2713")
    print("=" * 78)
    print("clean" if not bad else "%d file(s) need attention" % bad)
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
```

### `tools/calc_check.py`

```python
#!/usr/bin/env python3
"""
Recalculate a generated workbook with the ``formulas`` engine and report every
cell that evaluates to an Excel error (#REF!, #NAME?, #VALUE!, #DIV/0! ...).

This is the closest thing to opening the file in Excel without Excel: every
formula is parsed and evaluated for real, so broken references, typos and
unsupported functions show up here instead of in a customer's hands.

    /tmp/venv/bin/pip install formulas
    python3 tools/calc_check.py products/Christmas_Gift_Tracker_PREMIUM_Festive_EXAMPLE.xlsx

Known limitation: the ``formulas`` package cannot evaluate OFFSET-based defined
names (the auto-expanding dropdown lists), so those report #REF! here while
Excel and Google Sheets handle them normally.  They are listed separately as
"engine limitations" rather than failures.
"""

import glob
import logging
import re
import sys
import warnings

logging.disable(logging.CRITICAL)
warnings.simplefilter("ignore")

import formulas                                     # noqa: E402
import openpyxl                                     # noqa: E402


def check(path, verbose=True, compare=True):
    xl = formulas.ExcelModel().loads(path).finish()
    sol = xl.calculate()

    wb = openpyxl.load_workbook(path)
    wbv = openpyxl.load_workbook(path, data_only=True)
    book = path.split("/")[-1]
    book = re.escape("[%s]" % path.split("/")[-1].upper())

    def formula_of(key):
        m = re.match(r"^'\[.*\](.*)'!([A-Z]+\d+)$", key)
        if not m:
            return None, None
        sheet, coord = m.group(1), m.group(2)
        for ws in wb.worksheets:
            if ws.title.upper() == sheet:
                return ws.title, ws[coord].value
        return sheet, None

    mismatches = []
    if compare:
        for ws in wb.worksheets:
            wsv = wbv[ws.title]
            for row in ws.iter_rows():
                for cell in row:
                    v = cell.value
                    if not (isinstance(v, str) and v.startswith("=")):
                        continue
                    key = "'[%s]%s'!%s" % (book, ws.title.upper(),
                                           cell.coordinate)
                    got = sol.get(key)
                    if got is None:
                        continue
                    try:
                        calc = got.value[0, 0]
                    except Exception:
                        calc = got
                    cached = wsv[cell.coordinate].value
                    if calc is None and cached is None:
                        continue
                    if isinstance(calc, (int, float)) and \
                            isinstance(cached, (int, float)):
                        if abs(float(calc) - float(cached)) > 1e-6:
                            mismatches.append((ws.title, cell.coordinate,
                                               calc, cached))
                    elif str(calc).strip() != str(cached).strip():
                        mismatches.append((ws.title, cell.coordinate,
                                           calc, cached))

    errors, limits = [], []
    for key, value in sol.items():
        try:
            val = value.value[0, 0] if hasattr(value, "value") else value
        except Exception:
            continue
        text = str(val).strip()
        if not text.startswith("#"):
            continue
        if text == "#":            # a literal "#" header cell, not an error
            continue
        sheet, f = formula_of(key)
        if sheet is None:          # workbook-level defined name, not a cell
            limits.append((text, key, sheet, f))
            continue
        # things the engine cannot do but Excel and Google Sheets can:
        #   OFFSET() defined names (the auto-expanding dropdown lists)
        #   HYPERLINK()  (not implemented in the formulas package)
        if "OFFSET" in key.upper() or (f and "HYPERLINK(" in str(f).upper()):
            limits.append((text, key, sheet, f))
            continue
        errors.append((text, key, sheet, f))

    if verbose:
        print("=" * 78)
        print(path)
        print("  cells evaluated : %d" % len(sol))
        print("  engine limitations (fine in Excel / Sheets) : %d %s"
              % (len(limits), sorted({t for t, k, s2, f in limits})))
        if mismatches:
            print("  CACHED VALUE OUT OF DATE (%d):" % len(mismatches))
            for sheet, coord, calc, cached in mismatches[:15]:
                print("      %-20s %-6s recalcs to %r, file says %r"
                      % (sheet, coord, str(calc)[:30], str(cached)[:30]))
            if len(mismatches) > 15:
                print("      ... and %d more" % (len(mismatches) - 15))
        if errors:
            print("  ERRORS (%d):" % len(errors))
            for text, key, sheet, f in errors[:30]:
                print("    %-9s %-28s %s" % (text, (sheet or "?"),
                                             str(f)[:90]))
            if len(errors) > 30:
                print("    ... and %d more" % (len(errors) - 30))
        else:
            print("  ERRORS : none \u2713")
    return errors + mismatches, limits


def main(argv):
    paths = []
    for a in argv:
        paths.extend(sorted(glob.glob(a)) if any(c in a for c in "*?[") else [a])
    if not paths:
        print(__doc__)
        return 1
    total = 0
    for p in paths:
        errors, _ = check(p)
        total += len(errors)
    print("=" * 78)
    print("clean" if not total else "%d error cell(s) across %d file(s)"
          % (total, len(paths)))
    return 1 if total else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
```

### `tools/layout_check.py`

```python
#!/usr/bin/env python3
"""
Layout audit: finds text that will be clipped or overflow in Excel.

Excel never grows a merged cell to fit its text, so wrapped text in a merged
range needs an explicit row height.  This tool re-measures every cell in a
generated workbook the same way the builder did and reports:

    TOO TALL  wrapped text needs more height than the row(s) provide
    OVERFLOW  unwrapped text is wider than its cell and the neighbour is busy
    TINY ROW  an explicit row height smaller than one line of its font

    python3 tools/layout_check.py products/*.xlsx
"""

import glob
import math
import sys

import openpyxl
from openpyxl.utils import get_column_letter, range_boundaries

DEFAULT_COL_WIDTH = 8.43
CHARS_PER_UNIT = 1.05        # measured against Calibri 11 in Excel


def width_map(ws):
    """Column widths, expanding grouped ranges such as ``set_column("C:L")``."""
    widths = {}
    for dim in ws.column_dimensions.values():
        if not dim.width:
            continue
        lo = dim.min or 1
        hi = dim.max or lo
        for idx in range(lo, hi + 1):
            widths[idx] = dim.width
    return widths


def col_width(ws, idx, widths=None):
    if widths is None:
        widths = width_map(ws)
    return widths.get(idx, DEFAULT_COL_WIDTH)


def row_height(ws, idx, default=15.0):
    dim = ws.row_dimensions.get(idx)
    if dim is not None and dim.height:
        return dim.height
    return default


def check(path, verbose=True):
    wb = openpyxl.load_workbook(path)
    problems = []
    for ws in wb.worksheets:
        widths = width_map(ws)
        merged = {}
        for mr in ws.merged_cells.ranges:
            c1, r1, c2, r2 = range_boundaries(str(mr))
            merged[(r1, c1)] = (r1, c1, r2, c2)
            for rr in range(r1, r2 + 1):
                for cc in range(c1, c2 + 1):
                    if (rr, cc) != (r1, c1):
                        merged[(rr, cc)] = None      # swallowed cell

        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not isinstance(v, str) or not v.strip():
                    continue
                if v.startswith("="):
                    continue                       # formula: length unknown
                key = (cell.row, cell.column)
                if key in merged and merged[key] is None:
                    continue                       # hidden by a merge
                blocked = True
                if key in merged:
                    r1, c1, r2, c2 = merged[key]
                    width = sum(col_width(ws, c, widths)
                                for c in range(c1, c2 + 1))
                    height = sum(row_height(ws, r) for r in range(r1, r2 + 1))
                else:
                    width = col_width(ws, cell.column, widths)
                    height = row_height(ws, cell.row)
                    nxt = ws.cell(row=cell.row, column=cell.column + 1)
                    blocked = (nxt.value not in (None, "")
                               or (cell.row, cell.column + 1) in merged)

                size = cell.font.size or 11
                room = width * CHARS_PER_UNIT * (11.0 / float(size))
                text = v.strip()
                wrap = bool(cell.alignment.wrap_text)

                if wrap:
                    lines = 0
                    for part in text.split("\n"):
                        lines += max(1, int(math.ceil(len(part) / max(room, 1))))
                    line_h = size * 1.32 + 2.2
                    need = lines * line_h
                    if need > height + 1.5:
                        problems.append(
                            ("TOO TALL", ws.title, cell.coordinate,
                             "needs ~%.0fpt for %d line(s), has %.0fpt: %r"
                             % (need, lines, height, text[:60])))
                else:
                    if len(text) > room * 1.02 and key not in merged:
                        if blocked:
                            # a pasted web address is always longer than its
                            # column: the clickable button next to it is the
                            # intended way in, so this is by design
                            kind = ("CLIPPED LINK"
                                    if text.lower().startswith("http")
                                    else "OVERFLOW")
                            problems.append(
                                (kind, ws.title, cell.coordinate,
                                 "%d chars in %.0f wide cell: %r"
                                 % (len(text), width, text[:60])))
                if height < size * 1.15 and key not in merged:
                    problems.append(
                        ("TINY ROW", ws.title, cell.coordinate,
                         "row %.0fpt for %spt font" % (height, size)))
    if verbose:
        print("=" * 78)
        print(path)
        if problems:
            by_kind = {}
            for kind, sheet, coord, msg in problems:
                by_kind.setdefault(kind, []).append((sheet, coord, msg))
            for kind, items in sorted(by_kind.items()):
                print("  %-9s %d" % (kind, len(items)))
                for sheet, coord, msg in items[:12]:
                    print("      %-22s %-6s %s" % (sheet, coord, msg))
                if len(items) > 12:
                    print("      ... and %d more" % (len(items) - 12))
        else:
            print("  layout: no clipping or overflow found \u2713")
    return problems


def main(argv):
    paths = []
    for a in argv:
        paths.extend(sorted(glob.glob(a)) if any(c in a for c in "*?[") else [a])
    if not paths:
        print(__doc__)
        return 1
    total = 0
    for p in paths:
        total += len([x for x in check(p) if x[0] != "CLIPPED LINK"])
    print("=" * 78)
    print("clean" if not total else "%d layout issue(s)" % total)
    return 1 if total else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
```

### `tools/render_preview.py`

```python
#!/usr/bin/env python3
"""
Render a sheet of a generated .xlsx to a PNG so the design can be eyeballed
without Excel (there is no Excel or LibreOffice in this sandbox).

It is an approximation of Excel's rendering: cached formula values, solid
fills, borders, merged ranges, wrapped text, horizontal/vertical alignment,
number formats and text that spills into empty neighbours.  Emoji are drawn as
coloured tiles because no emoji font is installed - their position and size are
still visible, which is what matters for layout checking.

    /tmp/venv/bin/python tools/render_preview.py products/x.xlsx "🎄 Dashboard" \\
        /tmp/dash.png --rows 1-90
"""

import argparse
import datetime
import re
import sys
import unicodedata

import openpyxl
from openpyxl.utils import get_column_letter, range_boundaries
from PIL import Image, ImageDraw, ImageFont

FONT_DIR = "/usr/share/fonts/truetype/dejavu/"
PX_PER_WIDTH = 7.0
PT_TO_PX = 96.0 / 72.0
DEFAULT_ROW_PX = 20
EMOJI_RE = re.compile(
    "[\U0001F000-\U0001FAFF\u2600-\u27BF\u2B00-\u2BFF\uFE0F\u2190-\u21FF"
    "\u2700-\u27BF\u2764\u2699\u26A1\u2705\u274C\u2753\u2764]")


def is_emoji(ch):
    if ord(ch) < 0x2100:
        return False
    if ord(ch) in (0xFE0F, 0x2022, 0x2014, 0x2013, 0x201C, 0x201D, 0x2019,
                   0x2026, 0x00B7):
        return False
    cat = unicodedata.category(ch)
    return cat in ("So", "Sk", "Cs") or ord(ch) > 0x2500


def font(size, bold=False, italic=False):
    name = "DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf"
    return ImageFont.truetype(FONT_DIR + name, max(6, int(round(size))))


def color_of(c, default=None):
    if c is None:
        return default
    rgb = getattr(c, "rgb", None)
    if isinstance(rgb, str) and len(rgb) >= 6:
        rgb = rgb[-6:]
        try:
            return tuple(int(rgb[i:i + 2], 16) for i in (0, 2, 4))
        except ValueError:
            return default
    return default


def fmt_value(value, num_format):
    if value is None:
        return ""
    if isinstance(value, datetime.datetime):
        if value.hour or value.minute:
            return value.strftime("%d %b %Y %H:%M")
        return value.strftime("%d %b %Y")
    if isinstance(value, str):
        return value
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    nf = (num_format or "").lower()
    if isinstance(value, (int, float)):
        if "yy" in nf or "dd" in nf or "mmmm" in nf or "hh" in nf:
            try:
                dt = (datetime.datetime(1899, 12, 30)
                      + datetime.timedelta(days=float(value)))
            except (OverflowError, ValueError):
                return str(value)
            if "yyyy" in nf and "d" not in nf:
                return dt.strftime("%Y")
            if "dddd" in nf:
                return dt.strftime("%A %d %B %Y")
            if "ddd" in nf or "dd mmm" in nf or "d mmm" in nf:
                return dt.strftime("%d %b %Y")
            if "mm" in nf or "yy" in nf:
                return dt.strftime("%d/%m/%Y")
            return dt.strftime("%d %b")
        if "%" in nf:
            dp = 2 if "0.00%" in nf else (1 if "0.0%" in nf else 0)
            return ("%%.%df%%%%" % dp) % (value * 100)
        if "#,##0.00" in nf:
            return "{:,.2f}".format(value)
        if "#,##0" in nf or nf in ("0", "#,##0.0"):
            return "{:,.0f}".format(value) if "#,##0.0" not in nf \
                else "{:,.1f}".format(value)
        if value == int(value):
            return str(int(value))
        return "%.2f" % value
    return str(value)


def render(path, sheet_title, out, rows=None, cols=None, zoom=1.0):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet_title]

    # --- geometry -------------------------------------------------------
    widths = {}
    for dim in ws.column_dimensions.values():
        if not dim.width:
            continue
        for idx in range(dim.min or 1, (dim.max or dim.min or 1) + 1):
            widths[idx] = dim.width
    heights = {i: d.height for i, d in ws.row_dimensions.items() if d.height}

    r1, r2 = (rows or (1, ws.max_row))
    c1, c2 = (cols or (1, ws.max_column))

    def col_px(i):
        return int(round(widths.get(i, 8.43) * PX_PER_WIDTH)) + 5

    def row_px(i):
        return int(round(heights.get(i, 15.0) * PT_TO_PX)) or DEFAULT_ROW_PX

    xs = {}
    x = 0
    for i in range(c1, c2 + 1):
        xs[i] = x
        x += col_px(i)
    ys = {}
    y = 0
    for i in range(r1, r2 + 1):
        ys[i] = y
        y += row_px(i)
    W, H = int(x * zoom), int(y * zoom)
    img = Image.new("RGB", (max(W, 1), max(H, 1)), (255, 255, 255))
    d = ImageDraw.Draw(img)

    merges = {}
    swallowed = set()
    for mr in ws.merged_cells.ranges:
        mc1, mr1, mc2, mr2 = range_boundaries(str(mr))
        merges[(mr1, mc1)] = (mr1, mc1, mr2, mc2)
        for rr in range(mr1, mr2 + 1):
            for cc in range(mc1, mc2 + 1):
                if (rr, cc) != (mr1, mc1):
                    swallowed.add((rr, cc))

    def cell_box(rr, cc):
        if (rr, cc) in merges:
            a, b, c_, e = merges[(rr, cc)]
            x1, y1 = xs.get(b, x), ys.get(a, y)
            x2 = xs.get(e, x) + col_px(e)
            y2 = ys.get(c_, y) + row_px(c_)
        else:
            x1, y1 = xs.get(cc, x), ys.get(rr, y)
            x2, y2 = x1 + col_px(cc), y1 + row_px(rr)
        return x1, y1, x2, y2

    def occupied(rr, cc):
        c = ws.cell(row=rr, column=cc)
        return c.value not in (None, "")

    # --- paint ----------------------------------------------------------
    for rr in range(r1, r2 + 1):
        for cc in range(c1, c2 + 1):
            if (rr, cc) in swallowed:
                continue
            cell = ws.cell(row=rr, column=cc)
            x1, y1, x2, y2 = cell_box(rr, cc)
            if x2 <= 0 or y2 <= 0 or x1 >= x or y1 >= y:
                continue
            fill = color_of(cell.fill.start_color) if (
                cell.fill and cell.fill.patternType == "solid") else None
            if fill:
                d.rectangle([x1, y1, x2 - 1, y2 - 1], fill=fill)
            b = cell.border
            for side, edges in ((b.left, (x1, y1, x1, y2)),
                                (b.right, (x2 - 1, y1, x2 - 1, y2)),
                                (b.top, (x1, y1, x2, y1)),
                                (b.bottom, (x1, y2 - 1, x2, y2 - 1))):
                if side is not None and side.style:
                    col = color_of(side.color, (200, 200, 200))
                    thick = 2 if side.style in ("medium", "thick", "double") \
                        else 1
                    d.line(list(edges), fill=col, width=thick)

    # --- text -----------------------------------------------------------
    for rr in range(r1, r2 + 1):
        for cc in range(c1, c2 + 1):
            if (rr, cc) in swallowed:
                continue
            cell = ws.cell(row=rr, column=cc)
            raw = fmt_value(cell.value, cell.number_format)
            if not raw:
                continue
            x1, y1, x2, y2 = cell_box(rr, cc)
            if x2 <= 0 or y2 <= 0 or x1 >= x or y1 >= y:
                continue
            al = cell.alignment
            fnt = font((cell.font.size or 11) * 1.02,
                       bool(cell.font.bold), bool(cell.font.italic))
            col = color_of(cell.font.color, (0, 0, 0))
            indent = int((al.indent or 0) * 8)
            pad = 4 + indent

            # spill into empty neighbours when the text is not wrapped
            box_w = x2 - x1
            if not al.wrap_text:
                cc2 = cc + 1
                while (cc2 <= c2 and not occupied(rr, cc2)
                       and (rr, cc2) not in swallowed
                       and d.textlength(raw, font=fnt) > box_w - 2 * pad):
                    box_w += col_px(cc2)
                    cc2 += 1

            if al.wrap_text:
                words, lines, cur = raw.replace("\n", " \n ").split(), [], ""
                for w in words:
                    trial = (cur + " " + w).strip()
                    if d.textlength(trial, font=fnt) <= box_w - 2 * pad \
                            or not cur:
                        cur = trial
                    else:
                        lines.append(cur)
                        cur = w
                    if w == "\n":
                        lines.append(cur)
                        cur = ""
                if cur:
                    lines.append(cur)
            else:
                lines = raw.split("\n")

            lh = int((cell.font.size or 11) * 1.42)
            total_h = lh * len(lines)
            va = al.vertical or "bottom"
            if va == "center":
                ty = y1 + max(0, (y2 - y1 - total_h) // 2)
            elif va == "top":
                ty = y1 + 3
            else:
                ty = y1 + max(0, (y2 - y1 - total_h) // 2)

            for line in lines:
                drawn, cx = [], x1 + pad
                for ch in line:
                    if is_emoji(ch):
                        drawn.append(("tile", ch))
                    else:
                        drawn.append(("text", ch))
                # measure
                widths_px, tiles = [], []
                for kind, ch in drawn:
                    if kind == "tile":
                        wpx = int(lh * 0.95)
                        tiles.append(wpx)
                    else:
                        wpx = d.textlength(ch, font=fnt)
                    widths_px.append(wpx)
                line_w = sum(widths_px)
                ha = al.horizontal or ("right" if isinstance(
                    cell.value, (int, float)) and not cell.font.bold
                    else "left")
                if ha == "center":
                    cx = x1 + max(pad, (box_w - line_w) // 2)
                elif ha == "right":
                    cx = x1 + max(pad, box_w - line_w - pad)
                for (kind, ch), wpx in zip(drawn, widths_px):
                    if kind == "tile":
                        hue = (ord(ch) * 47) % 360
                        import colorsys
                        rgb = tuple(int(255 * v) for v in
                                    colorsys.hsv_to_rgb(hue / 360.0, 0.55, 0.85))
                        d.rounded_rectangle([cx, ty + 1, cx + wpx - 2,
                                             ty + lh - 2], radius=3, fill=rgb)
                    else:
                        d.text((cx, ty), ch, font=fnt, fill=col)
                    cx += wpx
                ty += lh
    # --- inserted pictures ---------------------------------------------
    from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, \
        TwoCellAnchor, AbsoluteAnchor
    for pic in getattr(ws, "_images", []):
        anc = pic.anchor
        frm = getattr(anc, "_from", None)
        if frm is None:
            continue
        col, row = frm.col + 1, frm.row + 1
        if not (c1 <= col <= c2 and r1 <= row <= r2):
            continue
        x1 = xs.get(col, 0) + int((frm.colOff or 0) / 9525.0)
        y1 = ys.get(row, 0) + int((frm.rowOff or 0) / 9525.0)
        to = getattr(anc, "to", None)
        if to is not None:
            x2 = xs.get(to.col + 1, x) + int((to.colOff or 0) / 9525.0)
            y2 = ys.get(to.row + 1, y) + int((to.rowOff or 0) / 9525.0)
            wpx, hpx = x2 - x1, y2 - y1
        else:
            try:
                ext = anc.ext
                wpx, hpx = int(ext.cx / 9525.0), int(ext.cy / 9525.0)
            except AttributeError:
                wpx, hpx = int(pic.width), int(pic.height)
        try:
            import io as _io
            sub = Image.open(_io.BytesIO(pic._data())).convert("RGBA")
            sub = sub.resize((max(wpx, 1), max(hpx, 1)), Image.LANCZOS)
            img.paste(sub, (x1, y1), sub)
        except Exception as exc:
            print("  (image skipped: %s)" % exc)

    if zoom != 1.0:
        img = img.resize((W, H), Image.LANCZOS)
    img.save(out)
    return out, img.size


def main(argv):
    p = argparse.ArgumentParser()
    p.add_argument("workbook")
    p.add_argument("sheet")
    p.add_argument("out")
    p.add_argument("--rows", default=None, help="e.g. 1-90")
    p.add_argument("--cols", default=None, help="e.g. A-M")
    p.add_argument("--zoom", type=float, default=1.0)
    a = p.parse_args(argv)
    rows = tuple(int(x) for x in a.rows.split("-")) if a.rows else None
    cols = None
    if a.cols:
        lo, hi = a.cols.split("-")
        cols = (openpyxl.utils.column_index_from_string(lo),
                openpyxl.utils.column_index_from_string(hi))
    out, size = render(a.workbook, a.sheet, a.out, rows, cols, a.zoom)
    print("%s -> %s (%dx%d)" % (a.sheet, out, size[0], size[1]))
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
```

### `tools/make_banner_alpha.py`

```python
#!/usr/bin/env python3
"""
Give the cover banners a transparent middle.

Excel draws inserted pictures *above* the cell grid, so text typed in the cells
underneath an opaque picture can never be seen.  The Start Here tab puts its
title in cells under the banner; for that to work the flat cream/white middle
of the artwork has to be see-through.

This script converts ``assets/banner_<theme>.jpg`` (flat background) into
``assets/banner_<theme>.png`` with an alpha channel: everything inside the
title window that matches the background colour becomes transparent, with a
one pixel feather so the watercolour edges stay soft.

    /tmp/venv/bin/python tools/make_banner_alpha.py
"""

import os

from PIL import Image, ImageFilter

HERE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ASSETS = os.path.join(HERE, "assets")

# title window as a fraction of the image (matches the C:H cell overlay)
WIN = (0.06, 0.02, 0.68, 0.98)
TOLERANCE = 26


def background_color(im):
    w, h = im.size
    px = im.load()
    samples = [px[int(w * 0.5), int(h * f)] for f in (0.02, 0.5, 0.98)]
    r = sum(s[0] for s in samples) // 3
    g = sum(s[1] for s in samples) // 3
    b = sum(s[2] for s in samples) // 3
    return r, g, b


def convert(src, dst):
    im = Image.open(src).convert("RGB")
    w, h = im.size
    bg = background_color(im)
    alpha = Image.new("L", (w, h), 255)
    ap = alpha.load()
    px = im.load()
    x1, y1, x2, y2 = (int(w * WIN[0]), int(h * WIN[1]),
                      int(w * WIN[2]), int(h * WIN[3]))
    for y in range(y1, y2):
        for x in range(x1, x2):
            r, g, b = px[x, y]
            if abs(r - bg[0]) + abs(g - bg[1]) + abs(b - bg[2]) < TOLERANCE * 3:
                ap[x, y] = 0
    alpha = alpha.filter(ImageFilter.GaussianBlur(1.1))
    out = im.convert("RGBA")
    out.putalpha(alpha)
    out.save(dst, optimize=True)
    return out.size, os.path.getsize(dst) // 1024


def main():
    import sys
    stems = sys.argv[1:] or ("banner_festive", "banner_minimal",
                             "banner_classic", "banner_fresh")
    for stem in stems:
        src = os.path.join(ASSETS, stem + ".jpg")
        dst = os.path.join(ASSETS, stem + ".png")
        if not os.path.exists(src):
            print("missing", src)
            continue
        size, kb = convert(src, dst)
        print("%-16s %s -> %s (%d KB)" % (stem, size, dst, kb))
        os.remove(src)


if __name__ == "__main__":
    main()
```

---

## 🗒️ Changelog

### v1.0.0 (initial release)

* 20-sheet Premium / 9-sheet Basic workbook in two themes (Classic, Fresh)
  and two fill states (blank, EXAMPLE).
* Dashboard: 12 KPI cards, food-cost & net-margin progress bars, six-stage
  pipeline strip, four charts (monthly revenue/profit combo, expense mix,
  revenue by event type, pipeline), next-events and money-due alert panels
  with cross-sheet OVERDUE / due-soon conditional formatting.
* Quote & Pricing Calculator with recommended price, price per guest,
  deposit and balance at the buyer's own margin.
* Menu & Recipe Costing with an ingredient-level recipe calculator that
  looks unit prices up from Inventory.
* Inventory LOW STOCK / REORDER flags; Shopping List auto-built per event
  (required minus in-stock) with supplier and estimated cost.
* Payments with automatic OVERDUE flag; Staff overtime & unpaid flags;
  Equipment service-due flags; month Calendar for any month/year.
* P&L monthly + annual, revenue by month/type, most-booked dishes,
  highest-value clients; Tax tracker with plain-English disclaimer.
* Print-ready Invoice & Proposal filled from a single dropdown.
* Start Here guide with watercolour cover banner (transparent-window PNG),
  five-step setup, tab tour and tips.
* Protection on by default (formula cells locked, inputs open, no password).
* QA gate on every build: `verify_workbook` (hidden writes / dangling refs),
  `calc_check` (full recalculation vs cached values) and `layout_check`
  (clipping / overflow / row heights).

---

*Generated from the working tree on 2026-09-13 (v1.0.0 source, includes QA tooling appendix).*
