# 🎄 Ultimate Christmas Gift Tracker - Python Source Code

> **Christmas Gift Command Center | Excel & Google Sheets Template**
> Author: **Novality Store** | Version: 1.0 | No macros, nothing to install

This document describes the complete Python source code that generates the
**Ultimate Christmas Gift Tracker** workbooks: up to 13 worksheets, 6 charts,
2,503 auto-calculating formulas, 56 drop-down validations and 46 conditional
formatting rules per premium workbook - in two editions (Basic / Premium),
two visual themes (Festive / Minimal) and two fill states (blank template /
filled example for listing screenshots).

---

## 📋 Table of Contents

1. [Overview](#-overview)
2. [Requirements](#-requirements)
3. [How to Run](#-how-to-run)
4. [Package Layout](#-package-layout)
5. [The Sheets](#-the-sheets)
6. [Design System](#-design-system)
7. [How the Automation Works](#-how-the-automation-works)
8. [Quality Checks](#-quality-checks)

---

## 📊 Overview

| Feature | Premium | Basic |
|---------|---------|-------|
| Worksheets | 13 (12 visible + hidden `_Data`) | 7 (6 visible + hidden `_Data`) |
| Charts | 6 | 6 |
| Formulas | 2,503 | 1,245 |
| Data validations | 56 | 24 |
| Conditional formats | 46 | 18 |
| Named ranges | 20 | 20 |
| File size | ~0.3-0.7 MB (incl. cover art) | ~0.3-0.7 MB (incl. cover art) |

Both editions open straight into Excel 2016+ or Google Sheets (File > Import >
Upload > Replace spreadsheet). There are no macros and no add-ins.

### 🎨 Colour palettes

| Role | Festive | Minimal |
|------|---------|---------|
| Canvas / card | `#FAF6EE` / `#FFFFFF` | `#FBFBFA` / `#FFFFFF` |
| Primary (pine / deep sage) | `#14432A` | `#33523F` |
| Accent (burgundy / clay) | `#8E2434` | `#B4654A` |
| Gold | `#C9A227` | `#B9974E` |
| Success / warn / danger | `#2E7D4F` / `#C77E1F` / `#B3372F` | sage / amber / rust |

---

## 🧰 Requirements

* Python 3.8+ (standard library only)
* The XlsxWriter library **vendored in this repository** (`xlsxwriter/`), so a
  fresh clone builds with no pip installs at all.

Optional, only for the quality tools in `tools/`:

* `openpyxl`, `formulas`, `Pillow` (any venv, e.g. `python3 -m venv /tmp/venv`)

---

## ▶️ How to Run

```bash
# the curated Etsy product set (6 files) into products/
python3 christmas_gift_tracker.py --all

# or one workbook at a time
python3 christmas_gift_tracker.py --edition premium --theme festive --mode demo
python3 christmas_gift_tracker.py --edition basic   --theme minimal
```

Flags: `--edition {basic,premium,both}`, `--theme {festive,minimal,both}`,
`--mode {blank,demo,both}`, `--out FILE`, `--outdir DIR`,
`--protect PASSWORD`, `--no-images`.

---

## 📦 Package Layout

```
christmas_gift_tracker.py        CLI entry point
christmas_tracker/
    config.py                    every row number, column map, list and
                                 dropdown value in one audited place
    theme.py                     the two palettes + per-tab colours
    styles.py                    the format factory (cached XlsxWriter
                                 formats) + wrapped-text height maths
    book.py                      Book: sheet registry, named ranges, page
                                 setup, nav buttons, charts, stats
    demo.py                      the example family + every aggregate the
                                 cached formula results are computed from
    workbook.py                  edition x theme x mode orchestrator
    sheets/
        common.py                table frames, headers, validations,
                                 conditional formats, totals, note blocks
        data.py  setup.py  gifts.py  budget.py  shopping.py  wishlist.py
        orders.py wrapping.py cards.py stockings.py todo.py
        dashboard.py  guide.py
tools/
    verify_workbook.py           structural audit (references, merges…)
    calc_check.py                recalculates every formula with the
                                 `formulas` engine and compares with the
                                 cached results stored in the file
    layout_check.py              finds text Excel would clip or overflow
    render_preview.py            renders a sheet to PNG for visual QA
    make_banner_alpha.py         gives the cover banners a see-through middle
assets/
    banner_festive.png           watercolour cover art (transparent middle)
    banner_minimal.png           line-art cover art (transparent middle)
```

---

## 📑 The Sheets

| # | Tab | Edition | Purpose |
|---|-----|---------|---------|
| 1 | 🎄 Dashboard | both | countdown, money + gift KPI cards, text progress bars, 4 charts, "what's left to do", next-five deadlines, per-recipient table |
| 2 | 🎁 Gift Tracker | both | one row per present; status pipeline 💡→🛒→️→✅→🎀→📦 with conditional formatting |
| 3 | 💰 Budget | both | planned vs actual per category, auto "pulled in" column, overspend alert at your chosen % |
| 4 | 💡 Wish List | premium | year-round idea parking with priorities and an "already on the gift list" check |
| 5 | 🛍️ Shopping List | both | non-gift spending (wrapping, baking, party…) feeding the budget |
| 6 | 📦 Order Tracker | premium | order numbers, expected vs actual delivery, tracking buttons, late alerts |
| 7 | 🎀 Wrapping & Hiding | premium | mirrors the gift list, adds hiding spots + gift tags, Secret Mode |
| 8 | 💌 Card Tracker | premium | bought → written → posted → replied, postage into the budget |
| 9 | 🧦 Stockings | premium | fillers per stocking with per-stocking budget rollup |
| 10 | ✅ To-Do List | premium | date-aware checklist; deadlines derived from the event date |
| 11 | ⚙️ Setup | both | event name/date, currency, alert %, every editable list |
| 12 | 📖 Start Here | both | watercolour cover, 3-minute tour, legend, FAQ, Google Sheets help |
| 13 | _Data (hidden) | both | the engine room: every dashboard number and chart series |

---

## 🎨 Design System

* Cream canvas, white cards, pine/burgundy/gold accents, subtle ❄ snowflake
  watermark on the dashboard, rounded KPI cards, emoji as iconography.
* Text progress bars (`REPT("█"…)&REPT("░"…)`) instead of in-cell charts so
  the bars survive printing, Google Sheets and Etsy thumbnails.
* Ticked boxes are a `✓` dropdown (named range `Tick`) rather than Excel
  checkbox controls, because those do not survive Google Sheets conversion.
* Every tab: frozen header rows, autofilter where useful, one-page-wide print
  setup with repeating titles, tab colours, and a button nav row at the foot.
* Merged cells never auto-fit: every wrapped block gets an explicit row height
  computed by `styles.wrap_height()`.

---

## 🤖 How the Automation Works

* `⚙️ Setup` holds the single source of truth. Named ranges (`EventName`,
  `EventDate`, `Currency`, `TotalBudget`, `AlertAt`, `DueSoonDays`,
  `SecretMode`, …) point at those cells so formulas read like sentences.
* Drop-down lists are `OFFSET(...COUNTA(...))` named ranges over the Setup
  lists, so adding a recipient extends every dropdown automatically.
* `_Data` mirrors Setup's lists, aggregates every tab with
  `SUMIF/COUNTIFS/SUMPRODUCT`, and builds the deadline pool the dashboard's
  "upcoming deadlines" panel reads with `SMALL/INDEX/MATCH`.
* Every formula is written **with a cached result**, so the file shows correct
  numbers in previews and third-party viewers before Excel recalculates.
* The example family lives in `demo.py`; the same numbers are computed twice -
  once in Python for the cached values, once in Excel for the live formulas -
  and `tools/calc_check.py` proves the two agree.

---

## 📜 Complete Source Code

Every file below is reproduced verbatim; together they are the whole
generator.  Create them with the paths shown and the product builds with
`python3 christmas_gift_tracker.py --all` (standard library only -
XlsxWriter ships inside this repository).

### `christmas_gift_tracker.py`  (95 lines)

```python
#!/usr/bin/env python3
"""
Build the Ultimate Christmas Gift Tracker workbooks.

Examples
--------
    python3 christmas_gift_tracker.py --all
    python3 christmas_gift_tracker.py --edition premium --theme festive \
        --mode demo --out products/hero.xlsx
    python3 christmas_gift_tracker.py --edition basic --theme minimal

The script needs nothing but the XlsxWriter library that lives in this
repository (it is imported from the checkout root), so it runs straight from
a fresh clone with a stock Python 3.
"""

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from christmas_tracker import build_all, build_workbook  # noqa: E402
from christmas_tracker.workbook import product_filename  # noqa: E402


def main(argv=None):
    p = argparse.ArgumentParser(
        description="Build Christmas / Ultimate Gift Tracker workbooks.")
    p.add_argument("--edition", choices=["basic", "premium", "both"],
                   default="premium")
    p.add_argument("--theme", choices=["festive", "minimal", "both"],
                   default="festive")
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
                   help="skip the watercolour cover banner")
    args = p.parse_args(argv)

    if args.all:
        stats = build_all(args.outdir, protect=args.protect,
                          images=not args.no_images)
        _report(stats)
        return 0

    editions = ["basic", "premium"] if args.edition == "both" else [args.edition]
    themes_ = ["festive", "minimal"] if args.theme == "both" else [args.theme]
    modes = ["blank", "demo"] if args.mode == "both" else [args.mode]

    stats = []
    for edition in editions:
        for theme_name in themes_:
            for mode in modes:
                if args.out and len(editions) == 1 and len(themes_) == 1 \
                        and len(modes) == 1:
                    path = args.out
                else:
                    os.makedirs(args.outdir, exist_ok=True)
                    path = os.path.join(
                        args.outdir, product_filename(edition, theme_name,
                                                      mode))
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

### `christmas_tracker/__init__.py`  (23 lines)

```python
"""
Christmas / Ultimate Gift Tracker - Etsy-ready Excel workbook builder.

Built on the XlsxWriter library vendored in this repository.  The package is
deliberately split into small modules so that every tab of the product can be
read (and tweaked) on its own:

    theme.py       colour palettes (festive + minimalist)
    config.py      sheet names, capacities, column maps, list contents
    styles.py      every cell format, derived from the active theme
    book.py        the build context: sheet registry, cross-sheet references
    demo.py        the optional "filled-in example" data set + aggregates
    sheets/        one module per worksheet
    workbook.py    orchestrates a full build

Entry point:  python christmas_gift_tracker.py --all
"""

from .workbook import build_workbook, build_all  # noqa: F401

__version__ = "1.0.0"
__product__ = "Ultimate Christmas Gift Tracker"
```

### `christmas_tracker/config.py`  (499 lines)

```python
"""
Single source of truth for the workbook layout.

Every sheet builder reads its geometry from here, which is what makes the
cross-sheet formulas (dashboard -> gift tracker -> budget -> ...) safe:
nothing hard-codes a cell address twice.

All row numbers in this module are **1-indexed** (spreadsheet style); the
sheet builders convert to XlsxWriter's 0-indexed API with ``r()``.
"""

# ===========================================================================
# PRODUCT
# ===========================================================================
PRODUCT = "Ultimate Christmas Gift Tracker"
PRODUCT_SHORT = "Christmas Gift Tracker"
TAGLINE = "Christmas \u2022 Birthdays \u2022 Holidays \u2022 Every Special Occasion"
AUTHOR = "Novality Store"
VERSION = "1.0"

# ===========================================================================
# SHEETS
# ===========================================================================
SHEET_NAMES = {
    "dashboard": "\U0001F384 Dashboard",
    "gifts":     "\U0001F381 Gift Tracker",
    "budget":    "\U0001F4B0 Budget",
    "wishlist":  "\U0001F4A1 Wish List",
    "shopping":  "\U0001F6CD\uFE0F Shopping List",
    "orders":    "\U0001F4E6 Order Tracker",
    "wrapping":  "\U0001F380 Wrapping & Hiding",
    "cards":     "\U0001F48C Card Tracker",
    "stockings": "\U0001F9E6 Stockings",
    "todo":      "\u2705 To-Do List",
    "setup":     "\u2699\uFE0F Setup",
    "guide":     "\U0001F4D6 Start Here",
    "data":      "_Data",
}

# Short labels used by the navigation buttons.
SHEET_SHORT = {
    "dashboard": "\U0001F3E0 Home",
    "gifts":     "\U0001F381 Gifts",
    "budget":    "\U0001F4B0 Budget",
    "wishlist":  "\U0001F4A1 Ideas",
    "shopping":  "\U0001F6CD\uFE0F Shop",
    "orders":    "\U0001F4E6 Orders",
    "wrapping":  "\U0001F380 Wrap",
    "cards":     "\U0001F48C Cards",
    "stockings": "\U0001F9E6 Stockings",
    "todo":      "\u2705 To-Do",
    "setup":     "\u2699\uFE0F Setup",
    "guide":     "\U0001F4D6 Help",
}

# Tab order per edition ("data" is always last and hidden).
EDITIONS = {
    "premium": ["dashboard", "gifts", "budget", "wishlist", "shopping",
                "orders", "wrapping", "cards", "stockings", "todo",
                "setup", "guide", "data"],
    "basic":   ["dashboard", "gifts", "budget", "shopping",
                "setup", "guide", "data"],
}

PREMIUM_ONLY = ("wishlist", "orders", "wrapping", "cards", "stockings", "todo")

# ===========================================================================
# DATA-SHEET GEOMETRY  (identical header band on every tracker tab)
# ===========================================================================
ROW_SPACER_1 = 1      # thin top margin
ROW_TITLE = 2         # big sheet title
ROW_SUBTITLE = 3      # one-line description + "back to dashboard" link
ROW_SPACER_2 = 4
ROW_STATS = 5         # quick-stat strip (formula driven)
ROW_SPACER_3 = 6
ROW_HEADER = 7        # column headers
ROW_FIRST = 8         # first data row

CAP = {
    "gifts": 60,
    "shopping": 45,
    "orders": 40,
    "cards": 40,
    "stockings": 45,
    "wishlist": 40,
    "todo": 40,
    "wrapping": 60,   # mirrors the gift tracker row-for-row
}

RECIPIENT_SLOTS = 16      # dashboard / chart capacity
CATEGORY_SLOTS = 16


def last_row(key):
    """Last data row (1-indexed) of a tracker sheet."""
    return ROW_FIRST + CAP[key] - 1


# ===========================================================================
# COLUMN MAPS  (letter per logical field)
# ===========================================================================
GIFT_COLS = {
    "n": "B", "recipient": "C", "relationship": "D", "idea": "E",
    "category": "F", "store": "G", "link": "H", "open": "I",
    "budget": "J", "cost": "K", "diff": "L", "pct": "M",
    "status": "N", "wrapped": "O", "delivered": "P",
    "deadline": "Q", "days": "R", "alert": "S", "notes": "T",
}

SHOP_COLS = {
    "n": "B", "item": "C", "category": "D", "store": "E", "qty": "F",
    "unit": "G", "total": "H", "bought": "I", "cost": "J", "spent": "K",
    "link": "L", "notes": "M",
}

ORDER_COLS = {
    "n": "B", "recipient": "C", "item": "D", "store": "E", "order_no": "F",
    "order_date": "G", "expected": "H", "actual": "I", "link": "J",
    "open": "K", "status": "L", "days": "M", "alert": "N",
    "return_by": "O", "returned": "P", "cost": "Q", "notes": "R",
}

WRAP_COLS = {
    "n": "B", "recipient": "C", "idea": "D", "bought": "E", "wrapped": "F",
    "hiding": "G", "tag": "H", "delivered": "I", "notes": "J",
}

CARD_COLS = {
    "n": "B", "name": "C", "relationship": "D", "address": "E",
    "bought": "F", "written": "G", "sent": "H", "date_sent": "I",
    "received": "J", "postage": "K", "status": "L", "notes": "M",
}

STOCK_COLS = {
    "n": "B", "owner": "C", "item": "D", "category": "E", "budget": "F",
    "cost": "G", "bought": "H", "wrapped": "I", "spent": "J",
    "hiding": "K", "notes": "L",
}

WISH_COLS = {
    "n": "B", "person": "C", "idea": "D", "category": "E", "link": "F",
    "open": "G", "price": "H", "priority": "I", "in_tracker": "J",
    "notes": "K",
}

TODO_COLS = {
    "done": "B", "task": "C", "category": "D", "deadline": "E",
    "days": "F", "status": "G", "notes": "H",
}

BUDGET_COLS = {
    "category": "B", "planned": "C", "manual": "D", "auto": "E",
    "actual": "F", "remaining": "G", "pct": "H", "status": "I",
    "bar": "J", "notes": "K",
}

COLS = {
    "gifts": GIFT_COLS,
    "shopping": SHOP_COLS,
    "orders": ORDER_COLS,
    "wrapping": WRAP_COLS,
    "cards": CARD_COLS,
    "stockings": STOCK_COLS,
    "wishlist": WISH_COLS,
    "todo": TODO_COLS,
    "budget": BUDGET_COLS,
}

# Column widths per sheet (letters are positional: A, B, C ...).
WIDTHS = {
    "gifts":     {"A": 2.2, "B": 4.5, "C": 18, "D": 12, "E": 27, "F": 13,
                  "G": 22, "H": 26, "I": 8.5, "J": 11, "K": 11, "L": 11,
                  "M": 9.5, "N": 16, "O": 9, "P": 9.5, "Q": 12, "R": 8.5,
                  "S": 13, "T": 30},
    "shopping":  {"A": 2.2, "B": 4.5, "C": 28, "D": 16, "E": 14, "F": 7,
                  "G": 11, "H": 11, "I": 10, "J": 11, "K": 11, "L": 20,
                  "M": 28},
    "orders":    {"A": 2.2, "B": 4.5, "C": 18, "D": 26, "E": 20, "F": 20,
                  "G": 12, "H": 13, "I": 13, "J": 26, "K": 8.5, "L": 17,
                  "M": 9, "N": 13, "O": 12, "P": 9.5, "Q": 11, "R": 26},
    "wrapping":  {"A": 2.2, "B": 4.5, "C": 16, "D": 28, "E": 13, "F": 13,
                  "G": 26, "H": 10, "I": 13, "J": 26},
    "cards":     {"A": 2.2, "B": 4.5, "C": 20, "D": 13, "E": 40, "F": 10,
                  "G": 10, "H": 9.5, "I": 12, "J": 10.5, "K": 10, "L": 15,
                  "M": 26},
    "stockings": {"A": 2.2, "B": 4.5, "C": 16, "D": 26, "E": 14, "F": 10.5,
                  "G": 10.5, "H": 10, "I": 10, "J": 10.5, "K": 24, "L": 24},
    "wishlist":  {"A": 2.2, "B": 4.5, "C": 16, "D": 30, "E": 14, "F": 28,
                  "G": 8.5, "H": 11, "I": 18, "J": 16, "K": 26},
    "todo":      {"A": 2.2, "B": 8, "C": 40, "D": 15, "E": 12, "F": 9,
                  "G": 15, "H": 30},
}

# ===========================================================================
# BUDGET TAB
# ===========================================================================
BUD_STATS_ROW = 5           # first row of the summary card block
BUD_HEADER = 14             # table header row
BUD_FIRST = 15              # first category row
BUD_TOTAL = 25              # totals row (BUD_FIRST + 10)

# (label, key, auto-link recipe, demo planned amount)
#   recipe terms: "gifts" = gift tracker actual cost,
#                 "shop:<Category>" = shopping list spend for that category,
#                 "stock" = stocking spend, "cards" = card postage spend
BUDGET_CATEGORIES = [
    ("\U0001F381 Gifts",              "gifts",     ["gifts", "shop:Gifts"], 850),
    ("\U0001F9E6 Stocking stuffers",  "stockings", ["stock", "shop:Stockings"], 140),
    ("\U0001F380 Wrapping & packaging", "wrapping", ["shop:Wrapping"], 55),
    ("\U0001F48C Cards & postage",    "cards",     ["shop:Cards", "shop:Postage", "cards"], 40),
    ("\U0001F35D Food & drink",       "food",      ["shop:Food"], 200),
    ("\U0001F36A Baking & treats",    "baking",    ["shop:Baking"], 35),
    ("\U0001F384 Decorations & tree", "decor",     ["shop:Decorations"], 60),
    ("\U0001F389 Party & entertaining", "party",   ["shop:Party Supplies"], 45),
    ("\U0001F697 Travel",             "travel",    ["shop:Travel"], 35),
    ("\u2764 Charity & other",        "other",     ["shop:Charity", "shop:Other"], 40),
]
BUDGET_ROWS = len(BUDGET_CATEGORIES)

# ===========================================================================
# SETUP TAB
# ===========================================================================
SU_EVENT_NAME = 6
SU_EVENT_DATE = 7
SU_DAYS = 8
SU_YEAR = 9
SU_WEEKS = 10
SU_MESSAGE = 11
SU_OCCASION = 14
SU_SUG_NAME = 15
SU_SUG_DATE = 16
SU_CURRENCY = 21
SU_BUDGET = 22
SU_ALERT = 23
SU_DUESOON = 24
SU_SECRET = 25
SU_COUNT_ORDERED = 26
SU_LIST_HEADER = 29
SU_LIST_FIRST = 30
SU_LIST_ROWS = 20
SU_FIXED_HEADER = 53
SU_FIXED_FIRST = 54

# Editable lists (dynamic named ranges) - column letter on the Setup sheet.
LIST_COLS = {
    "recipients": "B",
    "relationships": "C",
    "gift_categories": "D",
    "stores": "E",
    "shop_categories": "F",
    "hiding_spots": "G",
    "todo_categories": "H",
    "stocking_owners": "I",
    "stocking_items": "J",
}

LIST_TITLES = {
    "recipients": "\U0001F46A Recipients",
    "relationships": "\U0001F465 Relationships",
    "gift_categories": "\U0001F3F7\uFE0F Gift categories",
    "stores": "\U0001F3EC Stores",
    "shop_categories": "\U0001F6D2 Shopping categories",
    "hiding_spots": "\U0001F648 Hiding spots",
    "todo_categories": "\U0001F4CB To-do categories",
    "stocking_owners": "\U0001F9E6 Stocking owners",
    "stocking_items": "\U0001F36C Stocking items",
}

# Fixed lists (formulas and colours depend on this exact wording).
FIXED_COLS = {"statuses": "B", "order_statuses": "C", "priorities": "D",
              "tick": "E"}

# ===========================================================================
# LIST CONTENTS
# ===========================================================================
STATUSES = [
    "\U0001F4A1 Idea",
    "\U0001F6D2 Need to Buy",
    "\U0001F6CD\uFE0F Ordered",
    "\u2705 Purchased",
    "\U0001F380 Wrapped",
    "\U0001F4E6 Delivered",
]
ST_IDEA, ST_NEED, ST_ORDERED, ST_BOUGHT, ST_WRAPPED, ST_DELIVERED = STATUSES

ORDER_STATUSES = [
    "\U0001F6D2 Ordered",
    "\U0001F4E6 Shipped",
    "\U0001F69A Out for Delivery",
    "\u2705 Delivered",
    "\u21A9\uFE0F Return Requested",
    "\u274C Cancelled",
]
OS_ORDERED, OS_SHIPPED, OS_OUT, OS_DELIVERED, OS_RETURN, OS_CANCEL = ORDER_STATUSES

PRIORITIES = ["\u2B50\u2B50\u2B50 Must Have", "\u2B50\u2B50 Maybe",
              "\u2B50 Low Priority"]
PR_MUST, PR_MAYBE, PR_LOW = PRIORITIES

TICK = "\u2713"

CURRENCIES = ["$", "\u00A3", "\u20AC", "\u00A5", "\u20B9", "R$", "A$", "C$",
              "kr", "z\u0142", "CHF", "RM", "\u20AA", "R"]

YESNO = ["Yes", "No"]

RELATIONSHIPS = ["Mum", "Dad", "Grandma", "Grandpa", "Sister", "Brother",
                 "Partner", "Son", "Daughter", "Niece", "Nephew", "Cousin",
                 "Aunt", "Uncle", "Friend", "Colleague", "Neighbour",
                 "Teacher", "Secret Santa", "Other"]

GIFT_CATEGORIES = ["Clothing", "Beauty", "Electronics", "Toys", "Books",
                   "Home", "Food & Drink", "Gift Card", "Handmade",
                   "Jewellery", "Games", "Sports", "Music", "Kids", "Pets",
                   "Experience", "Other"]

STORES = ["Amazon", "Etsy", "Target", "Walmart", "John Lewis", "Tesco",
          "Argos", "ASOS", "Sephora", "Best Buy", "IKEA", "M&S",
          "Local Bookshop", "Not On The High Street", "Handmade by me",
          "Other"]

SHOP_CATEGORIES = ["Gifts", "Stockings", "Wrapping", "Food", "Baking",
                   "Decorations", "Cards", "Postage", "Party Supplies",
                   "Travel", "Charity", "Other"]

HIDING_SPOTS = ["Top shelf in closet", "Under the bed", "Garage shelf",
                "Suitcase in the attic", "Back of the pantry",
                "Linen closet", "Shoe box in wardrobe", "Behind the books",
                "Under the stairs box", "Car boot / trunk",
                "Decoration box", "At work / office", "Friend's house"]

TODO_CATEGORIES = ["Gifts", "Wrapping", "Cards", "Food", "Decorations",
                   "Travel", "Party", "Family", "Admin", "Other"]

STOCKING_ITEMS = ["Candy", "Chocolate", "Socks", "Beauty", "Hair clips",
                  "Stationery", "Small toy", "Gift card", "Snacks",
                  "Novelty", "Other"]

# Preset occasions: (name, month, day).  month=None -> "you choose the date".
OCCASIONS = [
    ("Christmas", 12, 25),
    ("Christmas Eve", 12, 24),
    ("Secret Santa", 12, 20),
    ("Hanukkah", 12, 15),
    ("New Year's Eve", 12, 31),
    ("Valentine's Day", 2, 14),
    ("Mother's Day", 5, 10),
    ("Father's Day", 6, 21),
    ("Easter", 4, 5),
    ("Halloween", 10, 31),
    ("Thanksgiving", 11, 27),
    ("Diwali", 11, 12),
    ("Birthday", None, None),
    ("Anniversary", None, None),
    ("Baby Shower", None, None),
    ("Wedding", None, None),
    ("Graduation", None, None),
    ("Housewarming", None, None),
]
OCCASION_ROWS = len(OCCASIONS)

# Default recipient / stocking-owner lists used when the workbook ships blank.
BLANK_RECIPIENTS = ["Mum", "Dad", "Grandma Rose", "Grandpa Joe",
                    "Sister Emily", "Brother Jack", "Niece Sophie",
                    "Nephew Leo", "Alex (partner)", "Maya (best friend)",
                    "Cousin Dan", "Aunt Carol"]
BLANK_OWNERS = ["Emma", "Oliver", "Mum", "Dad"]

# ===========================================================================
# HIDDEN "_Data" SHEET
# ===========================================================================
DATA_PRESET_FIRST = 2          # occasion table rows 2..2+len(OCCASIONS)-1
DATA_MATCH_MONTH = "E2"
DATA_MATCH_DAY = "F2"

DATA_REC_FIRST = 2             # recipient table H..P
DATA_REC_ROWS = RECIPIENT_SLOTS
DATA_CAT_FIRST = 2             # category table Q..R
DATA_CAT_ROWS = CATEGORY_SLOTS
DATA_STATUS_FIRST = 2          # status table T..U
DATA_BUD_FIRST = 2             # budget table W..Y
DATA_POOL_FIRST = 2            # deadline pool AA..AC

DATA_KPI_FIRST = 2             # KPI table AE(label) / AF(value)

# KPI registry: key -> (human label, number format).  Row order matters only
# for readability; the row number is derived from the position in this list.
KPI = [
    ("days_to_event",      "Days until the event", "0"),
    ("event_year",         "Event year", "0"),
    ("gifts_planned",      "Gifts planned", "0"),
    ("gifts_purchased",    "Gifts purchased", "0"),
    ("gifts_wrapped",      "Gifts wrapped", "0"),
    ("gifts_delivered",    "Gifts delivered / given", "0"),
    ("gifts_to_buy",       "Gifts still to buy", "0"),
    ("gifts_ordered",      "Gifts on order", "0"),
    ("wrap_to_do",         "Bought but not wrapped", "0"),
    ("deliver_to_do",      "Wrapped but not delivered", "0"),
    ("gift_completion",    "Gift completion %", "0%"),
    ("recipients_gifted",  "Recipients with gifts", "0"),
    ("avg_per_recipient",  "Average spend per recipient", "#,##0.00"),
    ("gift_budget_planned", "Gift budgets entered", "#,##0.00"),
    ("gift_actual_spent",  "Gift money spent", "#,##0.00"),
    ("budget_planned",     "Total planned budget", "#,##0.00"),
    ("budget_actual",      "Total spent (everything)", "#,##0.00"),
    ("budget_remaining",   "Budget remaining", "#,##0.00"),
    ("budget_pct",         "Budget used %", "0%"),
    ("shop_total",         "Shopping list items", "0"),
    ("shop_bought",        "Shopping items bought", "0"),
    ("shop_spent",         "Shopping list spend", "#,##0.00"),
    ("cards_total",        "Cards to send", "0"),
    ("cards_written",      "Cards written", "0"),
    ("cards_sent",         "Cards sent", "0"),
    ("stock_budget",       "Stocking budget", "#,##0.00"),
    ("stock_spent",        "Stocking spend", "#,##0.00"),
    ("stock_items",        "Stocking items", "0"),
    ("stock_bought",       "Stocking items bought", "0"),
    ("orders_total",       "Online orders", "0"),
    ("orders_outstanding", "Orders still coming", "0"),
    ("orders_late",        "Orders running late", "0"),
    ("orders_value",       "Value of orders", "#,##0.00"),
    ("todo_total",         "To-do tasks", "0"),
    ("todo_done",          "To-do tasks done", "0"),
    ("todo_overdue",       "To-do tasks overdue", "0"),
    ("wish_total",         "Gift ideas saved", "0"),
    ("wish_must",          "Must-have ideas", "0"),
    ("wish_value",         "Value of saved ideas", "#,##0.00"),
]

KPI_ROW = {key: DATA_KPI_FIRST + i for i, (key, _label, _fmt) in enumerate(KPI)}
KPI_LABEL = {key: label for key, label, _fmt in KPI}
KPI_FMT = {key: fmt for key, _label, fmt in KPI}

# Deadline pool: gift rows, then to-do rows, then order rows.
POOL_GIFT_ROWS = CAP["gifts"]
POOL_TODO_ROWS = CAP["todo"]
POOL_ORDER_ROWS = CAP["orders"]


def pool_size(edition):
    n = POOL_GIFT_ROWS
    if edition == "premium":
        n += POOL_TODO_ROWS + POOL_ORDER_ROWS
    return n


# ===========================================================================
# MISC
# ===========================================================================
DEFAULT_EVENT_NAME = "Christmas"
DEFAULT_EVENT_YEAR = 2026
DEFAULT_EVENT_DATE = (DEFAULT_EVENT_YEAR, 12, 25)
DEFAULT_BUDGET = 1500
DEFAULT_ALERT = 0.85
DEFAULT_DUESOON = 7
DEFAULT_CURRENCY = "$"

# To-do presets: (task, category, days BEFORE the event date, done?)
TODO_PRESETS = [
    ("Write the gift list & set the budget", "Gifts", 75, False),
    ("Order Christmas cards", "Cards", 60, False),
    ("Buy wrapping paper, ribbon & tags", "Wrapping", 45, False),
    ("Address & write all cards", "Cards", 35, False),
    ("Post the cards", "Cards", 25, False),
    ("Order gifts that need delivery time", "Gifts", 30, False),
    ("Buy the main gifts", "Gifts", 21, False),
    ("Prepare the stockings", "Gifts", 18, False),
    ("Buy the decorations & set up the tree", "Decorations", 20, False),
    ("Plan the Christmas dinner menu", "Food", 14, False),
    ("Book travel / check the car", "Travel", 21, False),
    ("Bake cookies & treats", "Food", 10, False),
    ("Buy the groceries", "Food", 5, False),
    ("Wrap the presents", "Wrapping", 7, False),
    ("Hide the wrapped gifts", "Wrapping", 6, False),
    ("Mail any packages", "Admin", 12, False),
    ("Confirm plans with family", "Family", 10, False),
    ("Charge cameras / buy batteries", "Admin", 4, False),
    ("Set the table & plan outfits", "Party", 3, False),
    ("Write the thank-you list", "Admin", -5, False),
]

# Shopping list presets used by the blank workbook (a helpful starter list).
SHOP_PRESETS = [
    ("Wrapping paper (3 rolls)", "Wrapping", "Target", 3, 4.00),
    ("Ribbon - red & gold", "Wrapping", "Amazon", 2, 3.50),
    ("Gift tags (pack of 50)", "Wrapping", "Etsy", 1, 6.00),
    ("Sticky tape & scissors", "Wrapping", "Target", 1, 5.00),
    ("Christmas cards (box of 30)", "Cards", "Amazon", 1, 14.00),
    ("Stamps", "Postage", "Post office", 30, 0.68),
    ("Plain flour", "Baking", "Grocery store", 2, 1.80),
    ("Cookie cutters", "Baking", "Amazon", 1, 8.00),
    ("Christmas crackers", "Party Supplies", "Target", 2, 5.50),
    ("Napkins & party plates", "Party Supplies", "Target", 1, 7.00),
    ("Tree baubles (gold)", "Decorations", "IKEA", 2, 9.00),
    ("Fairy lights", "Decorations", "Amazon", 1, 15.00),
    ("Wrapping tissue paper", "Wrapping", "Target", 1, 3.00),
    ("Chocolate for stockings", "Stockings", "Grocery store", 4, 3.20),
]
```

### `christmas_tracker/theme.py`  (153 lines)

```python
"""
Colour palettes / design systems.

Two themes ship with the product:

  festive   cream canvas, deep pine green, burgundy, antique gold
  minimal   white canvas, charcoal ink, sage green, clay accent

Everything else in the workbook is derived from the active theme, so a new
palette can be added by copying one of the dictionaries below.
"""


class Theme(object):
    """Small attribute bag so sheet code can read ``th.primary`` etc."""

    def __init__(self, **kw):
        self.__dict__.update(kw)

    def soft(self, key):
        """Return the pale 'tint' partner of a colour key."""
        return self.__dict__[key + "_soft"]


# ===========================================================================
# FESTIVE - "Christmas command centre"
# ===========================================================================
FESTIVE = Theme(
    key="festive",
    label="Festive",
    # canvas + surfaces
    bg="#FFF9EF",            # cream page background
    cover_bg="#FEF6E9",
    card="#FFFFFF",          # white cards
    alt="#FDFAF3",           # zebra stripe
    border="#E4D8C3",        # warm hairline
    border_strong="#CBB894",
    # brand
    primary="#14432A",       # deep pine green
    primary_2="#1E6B45",     # lighter green
    primary_2_soft="#DDEBE1",
    primary_soft="#E4EFE6",
    accent="#7B1E28",        # burgundy
    accent_soft="#F6E4E4",
    gold="#B8912F",          # antique gold
    gold_soft="#FBF2DA",
    # text
    ink="#33261C",
    muted="#8A7A6D",
    white="#FFFFFF",
    # semantic
    ok="#2F7D4F",
    ok_soft="#DFF0E3",
    warn="#B4761A",
    warn_soft="#FBF0D9",
    bad="#B02A2A",
    bad_soft="#F8E2E2",
    info="#2A5D8F",
    info_soft="#E3EDF7",
    plum="#6C3A6E",
    plum_soft="#F0E6F2",
    # typography
    title_font="Georgia",
    body_font="Calibri",
    mono_font="Consolas",
    # chart series colours (in draw order)
    series=["#14432A", "#7B1E28", "#B8912F", "#2A5D8F", "#2F7D4F",
            "#6C3A6E", "#C1443C", "#8A7A6D", "#1E6B45", "#B4761A"],
    # worksheet tab colours, per sheet key
    tabs={
        "dashboard": "#7B1E28",
        "gifts": "#14432A",
        "budget": "#B8912F",
        "wishlist": "#6C3A6E",
        "shopping": "#2A5D8F",
        "orders": "#1E6B45",
        "wrapping": "#C1443C",
        "cards": "#B4761A",
        "stockings": "#8C5A2B",
        "todo": "#2F7D4F",
        "setup": "#8A7A6D",
        "guide": "#7B1E28",
        "data": "#BFBFBF",
    },
)


# ===========================================================================
# MINIMAL - clean, modern, "not-Christmassy" (for the second product photo)
# ===========================================================================
MINIMAL = Theme(
    key="minimal",
    label="Minimal",
    bg="#FFFFFF",
    cover_bg="#FFFFFF",
    card="#FFFFFF",
    alt="#FAFAF8",
    border="#E3E3E1",
    border_strong="#C9C9C6",
    primary="#2F4F3E",       # deep sage
    primary_2="#4A6B59",
    primary_2_soft="#E8EEE9",
    primary_soft="#EDF1EE",
    accent="#A65A4A",        # clay
    accent_soft="#F7EDEA",
    gold="#B08D57",
    gold_soft="#F6F1E8",
    ink="#2E2E2E",
    muted="#8B8B8B",
    white="#FFFFFF",
    ok="#4B7A5A",
    ok_soft="#EDF3EE",
    warn="#A98036",
    warn_soft="#F8F2E6",
    bad="#A85450",
    bad_soft="#F8ECEB",
    info="#51708C",
    info_soft="#EDF2F6",
    plum="#6E5A78",
    plum_soft="#F2EFF4",
    title_font="Calibri Light",
    body_font="Calibri",
    mono_font="Consolas",
    series=["#2F4F3E", "#A65A4A", "#B08D57", "#51708C", "#4B7A5A",
            "#6E5A78", "#8C8C8C", "#A98036", "#7C9082", "#B8897E"],
    tabs={
        "dashboard": "#2F4F3E",
        "gifts": "#4A6B59",
        "budget": "#B08D57",
        "wishlist": "#6E5A78",
        "shopping": "#51708C",
        "orders": "#4B7A5A",
        "wrapping": "#A65A4A",
        "cards": "#A98036",
        "stockings": "#8C7355",
        "todo": "#4B7A5A",
        "setup": "#8C8C8C",
        "guide": "#2F4F3E",
        "data": "#BFBFBF",
    },
)


THEMES = {"festive": FESTIVE, "minimal": MINIMAL}


def get(name):
    try:
        return THEMES[name.lower()]
    except KeyError:
        raise SystemExit("Unknown theme '%s' (choose from: %s)"
                         % (name, ", ".join(sorted(THEMES))))
```

### `christmas_tracker/styles.py`  (319 lines)

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

### `christmas_tracker/book.py`  (491 lines)

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
            "category": "Christmas / Gift Planning Spreadsheet",
            "keywords": ("christmas gift tracker, gift list, budget planner, "
                         "holiday planner, etsy spreadsheet, excel template, "
                         "google sheets, stocking stuffer, card tracker"),
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
            "EventName": "%s!$C$%d" % (su, C.SU_EVENT_NAME),
            "EventDate": "%s!$C$%d" % (su, C.SU_EVENT_DATE),
            "Currency": "%s!$C$%d" % (su, C.SU_CURRENCY),
            "TotalBudget": "%s!$C$%d" % (su, C.SU_BUDGET),
            "AlertAt": "%s!$C$%d" % (su, C.SU_ALERT),
            "DueSoonDays": "%s!$C$%d" % (su, C.SU_DUESOON),
            "SecretMode": "%s!$C$%d" % (su, C.SU_SECRET),
            "CountOrdered": "%s!$C$%d" % (su, C.SU_COUNT_ORDERED),
            "Statuses": "%s!$%s$%d:$%s$%d" % (
                su, C.FIXED_COLS["statuses"], C.SU_FIXED_FIRST,
                C.FIXED_COLS["statuses"], C.SU_FIXED_FIRST + 5),
            "OrderStatuses": "%s!$%s$%d:$%s$%d" % (
                su, C.FIXED_COLS["order_statuses"], C.SU_FIXED_FIRST,
                C.FIXED_COLS["order_statuses"], C.SU_FIXED_FIRST + 5),
            "Priorities": "%s!$%s$%d:$%s$%d" % (
                su, C.FIXED_COLS["priorities"], C.SU_FIXED_FIRST,
                C.FIXED_COLS["priorities"], C.SU_FIXED_FIRST + 2),
            "Tick": "%s!$%s$%d:$%s$%d" % (
                su, C.FIXED_COLS["tick"], C.SU_FIXED_FIRST,
                C.FIXED_COLS["tick"], C.SU_FIXED_FIRST + 1),
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

        # Occasion presets live on the hidden data sheet.
        d = self.q("data")
        self.wb.define_name(
            "Occasions",
            "=OFFSET(%s!$A$%d,0,0,COUNTA(%s!$A$%d:$A$%d),1)"
            % (d, C.DATA_PRESET_FIRST, d, C.DATA_PRESET_FIRST,
               C.DATA_PRESET_FIRST + C.OCCASION_ROWS - 1))

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
    "recipients": "Recipients",
    "relationships": "Relationships",
    "gift_categories": "GiftCategories",
    "stores": "Stores",
    "shop_categories": "ShopCategories",
    "hiding_spots": "HidingSpots",
    "todo_categories": "TodoCategories",
    "stocking_owners": "StockingOwners",
    "stocking_items": "StockingItems",
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

### `christmas_tracker/demo.py`  (760 lines)

```python
"""
Sample data ("demo" mode) + the Python twin of every workbook calculation.

Two jobs:

1.  Provide a realistic filled-in example so the product can be screenshotted
    for the Etsy listing (``mode="demo"``).
2.  Compute the *cached values* that XlsxWriter stores next to each formula.
    Excel and Google Sheets both recalculate on open (``fullCalcOnLoad``), but
    the cached values mean the dashboard also looks right in previewers,
    thumbnailers and print-to-PDF tools that never calculate.

The functions in :func:`compute` mirror the worksheet formulas exactly - if
you change a formula in a sheet builder, change its twin here.
"""

from datetime import date, timedelta

from . import config as C

TODAY = date.today()
EVENT_DATE = date(*C.DEFAULT_EVENT_DATE)


def _d(offset_days):
    """A date ``offset_days`` from the event date (negative = before)."""
    return EVENT_DATE + timedelta(days=offset_days)


# ===========================================================================
# GIFT TRACKER
# ===========================================================================
# (recipient, relationship, idea, category, store, link, budget, cost,
#  status, wrapped, delivered, deadline offset, notes)
GIFTS = [
    ("Mum", "Mum", "Cashmere scarf", "Clothing", "John Lewis",
     "https://www.johnlewis.com", 60, 55.00, C.ST_BOUGHT, True, False, -10,
     "Wanted burgundy or cream"),
    ("Mum", "Mum", "Spa gift set", "Beauty", "Sephora",
     "https://www.sephora.com", 45, 42.99, C.ST_WRAPPED, True, False, -10, ""),
    ("Dad", "Dad", "Leather wallet", "Clothing", "Amazon",
     "https://www.amazon.com", 40, 38.50, C.ST_DELIVERED, True, True, -14,
     "Engraved initials"),
    ("Dad", "Dad", "Whisky tasting set", "Food & Drink",
     "Not On The High Street", "https://www.notonthehighstreet.com", 55, 55.00,
     C.ST_BOUGHT, True, False, -8, ""),
    ("Grandma Rose", "Grandma", "Digital photo frame", "Electronics",
     "Amazon", "https://www.amazon.com", 90, 84.99, C.ST_ORDERED, False,
     False, -5, "Pre-load family photos"),
    ("Grandma Rose", "Grandma", "Chunky knitted blanket", "Handmade", "Etsy",
     "https://www.etsy.com", 45, 45.00, C.ST_WRAPPED, True, False, -12, ""),
    ("Grandpa Joe", "Grandpa", "Crossword book bundle", "Books",
     "Local Bookshop", "", 25, 22.00, C.ST_DELIVERED, True, True, -16, ""),
    ("Grandpa Joe", "Grandpa", "Heated slippers", "Clothing", "Amazon",
     "https://www.amazon.com", 30, 27.50, C.ST_DELIVERED, True, True, -16, ""),
    ("Sister Emily", "Sister", "Gold hoop earrings", "Jewellery", "Etsy",
     "https://www.etsy.com", 50, 48.00, C.ST_WRAPPED, True, False, -6, ""),
    ("Sister Emily", "Sister", "Bath bomb gift box", "Beauty", "Sephora",
     "https://www.sephora.com", 25, None, C.ST_NEED, False, False, -6,
     "Check the winter collection"),
    ("Brother Jack", "Brother", "Wireless earbuds", "Electronics",
     "Best Buy", "https://www.bestbuy.com", 80, 74.99, C.ST_ORDERED, False,
     False, -4, "Express delivery paid"),
    ("Brother Jack", "Brother", "Craft beer hamper", "Food & Drink", "Tesco",
     "", 35, 33.00, C.ST_BOUGHT, False, False, -9, ""),
    ("Niece Sophie", "Niece", "LEGO Friends set", "Toys", "Amazon",
     "https://www.amazon.com", 45, 39.99, C.ST_DELIVERED, True, True, -18,
     "Age 8 - checked the age guide"),
    ("Niece Sophie", "Niece", "Art supplies case", "Kids", "Etsy",
     "https://www.etsy.com", 22, 20.00, C.ST_WRAPPED, True, False, -7, ""),
    ("Nephew Leo", "Nephew", "Dinosaur museum tickets", "Experience", "Other",
     "", 40, 36.00, C.ST_BOUGHT, False, False, -3, "Print the voucher"),
    ("Nephew Leo", "Nephew", "Graphic novel bundle", "Books",
     "Local Bookshop", "", 25, None, C.ST_IDEA, False, False, -3,
     "Ask Emily what he is reading"),
    ("Alex (partner)", "Partner", "Weekend getaway", "Experience", "Other",
     "", 180, 165.00, C.ST_BOUGHT, False, False, -2, "Booked, no wrapping!"),
    ("Alex (partner)", "Partner", "Engraved watch", "Jewellery", "Etsy",
     "https://www.etsy.com", 95, 89.00, C.ST_ORDERED, False, False, -5, ""),
    ("Maya (best friend)", "Friend", "Scented candle trio", "Home",
     "Not On The High Street", "https://www.notonthehighstreet.com", 30,
     28.50, C.ST_DELIVERED, True, True, -20, ""),
    ("Maya (best friend)", "Friend", "Baking cookbook", "Books", "Amazon",
     "https://www.amazon.com", 22, None, C.ST_NEED, False, False, -6, ""),
    ("Cousin Dan", "Cousin", "Strategy board game", "Games", "Target",
     "https://www.target.com", 35, 32.00, C.ST_BOUGHT, False, False, -8, ""),
    ("Cousin Dan", "Cousin", "Coffee bean subscription", "Food & Drink",
     "Etsy", "https://www.etsy.com", 30, None, C.ST_NEED, False, False, -8,
     "3-month subscription"),
    ("Aunt Carol", "Aunt", "Silk scarf", "Clothing", "M&S",
     "https://www.marksandspencer.com", 40, 37.50, C.ST_DELIVERED, True, True,
     -21, ""),
    ("Aunt Carol", "Aunt", "Garden gloves & tool set", "Home", "Amazon",
     "https://www.amazon.com", 18, 16.00, C.ST_WRAPPED, True, False, -7, ""),
    ("Secret Santa", "Secret Santa", "Novelty reindeer mug", "Other",
     "Target", "https://www.target.com", 15, 12.99, C.ST_BOUGHT, False, False,
     -4, "Office Secret Santa - $15 limit"),
]

GIFT_KEYS = ("recipient", "relationship", "idea", "category", "store", "link",
             "budget", "cost", "status", "wrapped", "delivered", "deadline",
             "notes")

# Where the wrapped presents are hiding (used by the Wrapping & Hiding tab).
# gift index -> (hiding spot, gift tag attached?)
WRAP_DETAILS = {
    1: ("Top shelf in closet", True),
    2: ("Car boot / trunk", True),
    5: ("Suitcase in the attic", True),
    6: ("Behind the books", True),
    7: ("Under the bed", True),
    8: ("Linen closet", True),
    12: ("Garage shelf", True),
    13: ("Top shelf in closet", False),
    18: ("Under the stairs box", True),
    22: ("Shoe box in wardrobe", True),
    23: ("Back of the pantry", True),
}


def gift_rows():
    out = []
    for i, g in enumerate(GIFTS):
        row = dict(zip(GIFT_KEYS, g))
        row["deadline"] = _d(row["deadline"])
        row["wrapped"] = C.TICK if row["wrapped"] else ""
        row["delivered"] = C.TICK if row["delivered"] else ""
        spot, tag = WRAP_DETAILS.get(i, ("", False))
        row["hiding"] = spot
        row["tag"] = C.TICK if tag else ""
        out.append(row)
    return out


# ===========================================================================
# SHOPPING LIST
# ===========================================================================
# (item, category, store, qty, unit cost, bought, actual cost, notes)
SHOPPING = [
    ("Wrapping paper - 3 roll pack", "Wrapping", "Target", 3, 4.00, True,
     11.50, "Gold foil design"),
    ("Ribbon - red & gold", "Wrapping", "Amazon", 2, 3.50, True, 6.80, ""),
    ("Gift tags (pack of 50)", "Wrapping", "Etsy", 1, 6.00, True, 6.00,
     "Hand-lettered"),
    ("Sticky tape & scissors", "Wrapping", "Target", 1, 5.00, False, None, ""),
    ("Tissue paper bundle", "Wrapping", "Target", 1, 3.00, False, None, ""),
    ("Christmas cards (box of 30)", "Cards", "Amazon", 1, 14.00, True, 13.50,
     ""),
    ("Stamps - book of 30", "Postage", "Post office", 1, 20.40, True, 20.40,
     ""),
    ("Thank-you notes", "Cards", "Etsy", 1, 8.00, False, None, "For January"),
    ("Plain flour", "Baking", "Grocery store", 2, 1.80, True, 3.40, ""),
    ("Cookie cutters set", "Baking", "Amazon", 1, 8.00, True, 7.50, ""),
    ("Cake decorating sprinkles", "Baking", "Grocery store", 1, 4.50, False,
     None, ""),
    ("Chocolate for stockings", "Stockings", "Grocery store", 4, 3.20, True,
     12.00, ""),
    ("Candy canes", "Stockings", "Grocery store", 2, 2.00, True, 3.80, ""),
    ("Christmas crackers", "Party Supplies", "Target", 2, 5.50, False, None,
     ""),
    ("Napkins & party plates", "Party Supplies", "Target", 1, 7.00, False,
     None, ""),
    ("Tree baubles - gold", "Decorations", "IKEA", 2, 9.00, True, 16.00, ""),
    ("Fairy lights (warm white)", "Decorations", "Amazon", 1, 15.00, True,
     14.99, ""),
    ("Wreath for the front door", "Decorations", "Local market", 1, 22.00,
     False, None, ""),
    ("Turkey & trimmings", "Food", "Grocery store", 1, 65.00, False, None,
     "Order by 18 Dec"),
    ("Sparkling cider", "Food", "Grocery store", 2, 6.50, False, None, ""),
    ("Chocolate selection boxes", "Food", "Grocery store", 3, 8.00, True,
     22.50, "For neighbours & teachers"),
    ("Petrol / parking for the trip", "Travel", "Shell", 1, 38.40, False,
     None, ""),
    ("Food bank donation", "Charity", "Local charity", 1, 25.00, True, 25.00,
     ""),
    ("Gift bags (spares)", "Other", "Target", 1, 4.00, False, None, ""),
]

SHOP_KEYS = ("item", "category", "store", "qty", "unit", "bought", "cost",
             "notes")


def shop_rows():
    out = []
    for s in SHOPPING:
        row = dict(zip(SHOP_KEYS, s))
        row["bought"] = C.TICK if row["bought"] else ""
        out.append(row)
    return out


def shop_presets():
    """Starter rows shipped with the blank workbook."""
    out = []
    for item, cat, store, qty, unit in C.SHOP_PRESETS:
        out.append({"item": item, "category": cat, "store": store, "qty": qty,
                    "unit": unit, "bought": "", "cost": None, "notes": ""})
    return out


# ===========================================================================
# ORDER TRACKER
# ===========================================================================
# (recipient, item, store, order no, order date, expected, actual, link,
#  status, cost, return deadline, notes)
ORDERS = [
    ("Grandma Rose", "Digital photo frame", "Amazon", "112-4471902-5563810",
     _d(-20), _d(-3), None, "https://www.amazon.com/gp/your-account/order-details?order=112-4471902",
     C.OS_SHIPPED, 84.99, _d(25), "Tracking updated twice a day"),
    ("Brother Jack", "Wireless earbuds", "Best Buy", "BB-9938-2211", _d(-18),
     _d(-4), None, "https://www.bestbuy.com/orders", C.OS_OUT, 74.99, _d(24),
     "Signature required"),
    ("Alex (partner)", "Engraved watch", "Etsy", "ETSY-778201", _d(-25),
     _d(-6), None, "https://www.etsy.com/your/orders", C.OS_SHIPPED, 89.00,
     _d(22), "Engraving takes 5 days"),
    ("Dad", "Leather wallet", "Amazon", "112-8830042-1120394", _d(-32),
     _d(-14), _d(-15), "https://www.amazon.com/gp/your-account/order-details?order=112-8830042",
     C.OS_DELIVERED, 38.50, _d(16), "Left with neighbour"),
    ("Niece Sophie", "LEGO Friends set", "Amazon", "112-2291047-9930112",
     _d(-40), _d(-18), _d(-18), "https://www.amazon.com/gp/your-account/order-details?order=112-2291047",
     C.OS_DELIVERED, 39.99, _d(12), ""),
    ("Maya (best friend)", "Scented candle trio", "Not On The High Street",
     "NOTHS-55231", _d(-45), _d(-20), _d(-19),
     "https://www.notonthehighstreet.com/my-account", C.OS_DELIVERED, 28.50,
     _d(10), ""),
    ("Aunt Carol", "Silk scarf", "M&S", "MS-4419287", _d(-50), _d(-21),
     _d(-21), "https://www.marksandspencer.com/my-account", C.OS_DELIVERED,
     37.50, _d(9), ""),
    ("Sister Emily", "Bath bomb gift box", "Sephora", "SEPH-88213", _d(-6),
     _d(-1), None, "https://www.sephora.com/account/orders", C.OS_ORDERED,
     24.00, _d(29), "Waiting for dispatch"),
]

ORDER_KEYS = ("recipient", "item", "store", "order_no", "order_date",
              "expected", "actual", "link", "status", "cost", "return_by",
              "notes")


def order_rows():
    out = []
    for o in ORDERS:
        row = dict(zip(ORDER_KEYS, o))
        row["returned"] = ""
        out.append(row)
    return out


# ===========================================================================
# CARD TRACKER
# ===========================================================================
CARDS = [
    ("Grandma Rose", "Grandma", "14 Rosewood Lane, Harrogate HG1 2AB, UK",
     True, True, True, _d(-12), True, 0.68, "Big handwriting card"),
    ("Grandpa Joe", "Grandpa", "14 Rosewood Lane, Harrogate HG1 2AB, UK",
     True, True, True, _d(-12), True, 0.68, "Same envelope as Grandma"),
    ("Aunt Carol", "Aunt", "88 Maple Drive, Leeds LS6 4QT, UK", True, True,
     True, _d(-10), False, 0.68, ""),
    ("Uncle Bill", "Uncle", "88 Maple Drive, Leeds LS6 4QT, UK", True, True,
     True, _d(-10), False, 0.68, ""),
    ("Cousin Dan", "Cousin", "3 Ash Court, Manchester M1 5AN, UK", True, True,
     True, _d(-9), False, 0.68, ""),
    ("Cousin Priya", "Cousin", "3 Ash Court, Manchester M1 5AN, UK", True,
     True, False, None, False, 0.00, "Card written, needs posting"),
    ("Maya (best friend)", "Friend", "27b Kingsland Road, London E2 8AA, UK",
     True, True, True, _d(-8), True, 0.68, "Hand delivered a gift too"),
    ("Tom & Jess", "Friend", "5 Cedar Walk, Bristol BS1 4TR, UK", True, True,
     False, None, False, 0.00, ""),
    ("Mrs Higgins (neighbour)", "Neighbour", "9 Birch Hill, York YO1 7HD, UK",
     True, False, False, None, False, 0.00, "Write after the 15th"),
    ("Ms Patel (teacher)", "Teacher", "St Anne's Primary, School Office", True,
     False, False, None, False, 0.00, "From Sophie"),
    ("Office team", "Colleague", "Hand delivered", True, True, True, _d(-5),
     False, 0.00, "Left on the team table"),
    ("Nana Flo", "Grandma", "2 Willow Cottages, Derby DE1 3RT, UK", False,
     False, False, None, False, 0.00, "Buy a photo card"),
    ("The Robinsons", "Neighbour", "11 Oakfield Road, Sheffield S10 2QN, UK",
     False, False, False, None, False, 0.00, ""),
    ("Cousin Sam", "Cousin", "44 Station Road, Newcastle NE1 5BR, UK", False,
     False, False, None, False, 0.00, ""),
]

CARD_KEYS = ("name", "relationship", "address", "bought", "written", "sent",
             "date_sent", "received", "postage", "notes")


def card_rows():
    out = []
    for c in CARDS:
        row = dict(zip(CARD_KEYS, c))
        for k in ("bought", "written", "sent", "received"):
            row[k] = C.TICK if row[k] else ""
        out.append(row)
    return out


# ===========================================================================
# STOCKINGS
# ===========================================================================
STOCKINGS = [
    ("Emma", "Lip balm trio", "Beauty", 8, 7.50, True, True,
     "Top shelf in closet", "Peppermint - her favourite"),
    ("Emma", "Hair clips set", "Hair clips", 6, 5.00, True, False,
     "Top shelf in closet", ""),
    ("Emma", "Chocolate coins", "Chocolate", 5, 4.20, True, True,
     "Under the bed", ""),
    ("Emma", "Mini diary & pen", "Stationery", 9, 8.00, False, False, "", ""),
    ("Emma", "Fuzzy socks", "Socks", 7, 6.50, True, False, "Linen closet", ""),
    ("Oliver", "Dinosaur figures", "Small toy", 12, 11.00, True, True,
     "Garage shelf", ""),
    ("Oliver", "Candy cane pack", "Candy", 4, 3.80, True, True,
     "Back of the pantry", ""),
    ("Oliver", "Football stickers", "Novelty", 6, 5.50, True, False,
     "Shoe box in wardrobe", ""),
    ("Oliver", "Comic book", "Stationery", 7, None, False, False, "",
     "Ask what he is reading"),
    ("Oliver", "Bath bubble bombs", "Beauty", 8, 7.00, True, False,
     "Linen closet", ""),
    ("Mum", "Silk eye mask", "Beauty", 15, 14.00, True, True,
     "Suitcase in the attic", ""),
    ("Mum", "Hand cream & nails set", "Beauty", 12, 11.50, True, False,
     "Under the stairs box", ""),
    ("Mum", "Favourite tea blend", "Snacks", 9, 8.50, True, True,
     "Back of the pantry", ""),
    ("Dad", "Leather keyring", "Novelty", 14, 13.00, True, False,
     "Car boot / trunk", "Initials J.M."),
    ("Dad", "Golf balls (3 pack)", "Other", 12, 11.00, True, True,
     "Garage shelf", ""),
    ("Dad", "Hot sauce bundle", "Snacks", 10, 9.50, False, False, "",
     "Check the heat levels!"),
]

STOCK_KEYS = ("owner", "item", "category", "budget", "cost", "bought",
              "wrapped", "hiding", "notes")


def stock_rows():
    out = []
    for s in STOCKINGS:
        row = dict(zip(STOCK_KEYS, s))
        row["bought"] = C.TICK if row["bought"] else ""
        row["wrapped"] = C.TICK if row["wrapped"] else ""
        out.append(row)
    return out


# ===========================================================================
# WISH LIST
# ===========================================================================
WISHLIST = [
    ("Alex (partner)", "Engraved watch", "Jewellery",
     "https://www.etsy.com/listing/watch", 95.00, C.PR_MUST,
     "Mentioned it in October"),
    ("Alex (partner)", "Cast iron skillet", "Home",
     "https://www.amazon.com", 45.00, C.PR_MAYBE, "For the new kitchen"),
    ("Mum", "Cashmere scarf", "Clothing", "https://www.johnlewis.com", 60.00,
     C.PR_MUST, "Burgundy or cream"),
    ("Mum", "Gardening book", "Books", "", 18.00, C.PR_LOW, ""),
    ("Sister Emily", "Bath bomb gift box", "Beauty",
     "https://www.sephora.com", 25.00, C.PR_MUST, "On the wish list she sent"),
    ("Sister Emily", "Silk pillowcase", "Home", "https://www.amazon.com",
     32.00, C.PR_MAYBE, ""),
    ("Nephew Leo", "Graphic novel bundle", "Books", "", 25.00, C.PR_MAYBE,
     "Ages 10-12"),
    ("Niece Sophie", "Watercolour set", "Kids", "https://www.etsy.com", 20.00,
     C.PR_LOW, ""),
    ("Grandma Rose", "Digital photo frame", "Electronics",
     "https://www.amazon.com", 90.00, C.PR_MUST, "Big buttons, simple menu"),
    ("Brother Jack", "Wireless earbuds", "Electronics",
     "https://www.bestbuy.com", 80.00, C.PR_MUST, ""),
    ("Maya (best friend)", "Pottery class for two", "Experience", "", 70.00,
     C.PR_MAYBE, "Birthday idea for March"),
    ("Cousin Dan", "Coffee bean subscription", "Food & Drink",
     "https://www.etsy.com", 30.00, C.PR_LOW, ""),
]

WISH_KEYS = ("person", "idea", "category", "link", "price", "priority",
             "notes")


def wish_rows():
    return [dict(zip(WISH_KEYS, w)) for w in WISHLIST]


# ===========================================================================
# BUDGET (planned amounts) + manual actuals
# ===========================================================================
BUDGET_PLANNED = {
    "gifts": 1050, "stockings": 120, "wrapping": 50, "cards": 35,
    "food": 130, "baking": 25, "decor": 40, "party": 25, "travel": 0,
    "other": 25,
}
BUDGET_MANUAL = {"travel": 38.40}


# ===========================================================================
# TO-DO LIST (presets, with which ones are already ticked in the demo)
# ===========================================================================
TODO_DONE = {
    "Write the gift list & set the budget",
    "Order Christmas cards",
    "Buy wrapping paper, ribbon & tags",
    "Order gifts that need delivery time",
    "Buy the main gifts",
    "Buy the decorations & set up the tree",
    "Mail any packages",
}


def todo_rows(mode):
    """To-do presets.  ``offset`` = days BEFORE the event date."""
    out = []
    for task, cat, offset, _done in C.TODO_PRESETS:
        done = (mode == "demo" and task in TODO_DONE)
        out.append({"done": C.TICK if done else "", "task": task,
                    "category": cat, "offset": offset,
                    "deadline": EVENT_DATE + timedelta(days=-offset),
                    "notes": ""})
    return out


# ===========================================================================
# SETTINGS
# ===========================================================================
SETTINGS = {
    "event_name": C.DEFAULT_EVENT_NAME,
    "currency": C.DEFAULT_CURRENCY,
    "budget": C.DEFAULT_BUDGET,
    "alert": C.DEFAULT_ALERT,
    "duesoon": C.DEFAULT_DUESOON,
    "secret": "No",
    "count_ordered": "No",
    "occasion": "Christmas",
}


# ===========================================================================
# THE CALCULATOR
# ===========================================================================
class Model(object):
    """Everything the workbook would calculate if Excel did it for us."""

    def __init__(self, mode="blank", edition="premium"):
        self.mode = mode
        self.edition = edition
        self.demo = (mode == "demo")
        self.settings = dict(SETTINGS) if self.demo else {
            "event_name": C.DEFAULT_EVENT_NAME,
            "currency": C.DEFAULT_CURRENCY, "budget": 0,
            "alert": C.DEFAULT_ALERT, "duesoon": C.DEFAULT_DUESOON,
            "secret": "No", "count_ordered": "No", "occasion": "Christmas"}
        self.event_date = EVENT_DATE
        self.recipients = (GIFT_RECIPIENTS if self.demo
                           else list(C.BLANK_RECIPIENTS))
        self.owners = (STOCK_OWNERS if self.demo else list(C.BLANK_OWNERS))
        self.categories = list(C.GIFT_CATEGORIES)

        self.gifts = gift_rows() if self.demo else []
        self.shopping = shop_rows() if self.demo else shop_presets()
        self.orders = order_rows() if (self.demo and edition == "premium") else []
        self.cards = card_rows() if (self.demo and edition == "premium") else []
        self.stockings = stock_rows() if (self.demo and edition == "premium") else []
        self.wishlist = wish_rows() if (self.demo and edition == "premium") else []
        self.todos = todo_rows(mode)
        self.budget_planned = dict(BUDGET_PLANNED) if self.demo else {}
        self.budget_manual = dict(BUDGET_MANUAL) if self.demo else {}
        self.agg = {}
        self._compute()

    # ------------------------------------------------------------------
    def _compute(self):
        a = self.agg
        g = self.gifts

        # --- gifts -----------------------------------------------------
        a["gifts_planned"] = len([x for x in g if x["idea"]])
        bought_states = (C.ST_BOUGHT, C.ST_WRAPPED, C.ST_DELIVERED)
        a["gifts_purchased"] = len([x for x in g if x["status"] in bought_states])
        if self.settings["count_ordered"] == "Yes":
            a["gifts_purchased"] += len([x for x in g
                                         if x["status"] == C.ST_ORDERED])
        a["gifts_wrapped"] = len([
            x for x in g
            if x["wrapped"] == C.TICK
            or (x["wrapped"] != C.TICK
                and x["status"] in (C.ST_WRAPPED, C.ST_DELIVERED))])
        a["gifts_delivered"] = len([
            x for x in g
            if x["delivered"] == C.TICK
            or (x["delivered"] != C.TICK and x["status"] == C.ST_DELIVERED)])
        a["gifts_to_buy"] = len([x for x in g
                                 if x["status"] in (C.ST_IDEA, C.ST_NEED)])
        a["gifts_ordered"] = len([x for x in g if x["status"] == C.ST_ORDERED])
        a["wrap_to_do"] = max(0, a["gifts_purchased"] - a["gifts_wrapped"])
        a["deliver_to_do"] = max(0, a["gifts_wrapped"] - a["gifts_delivered"])
        a["gift_completion"] = (a["gifts_purchased"] / float(a["gifts_planned"])
                                if a["gifts_planned"] else 0)
        a["gift_budget_planned"] = sum(x["budget"] or 0 for x in g)
        a["gift_actual_spent"] = sum(x["cost"] or 0 for x in g)

        # --- recipients ------------------------------------------------
        rec = []
        for i in range(C.RECIPIENT_SLOTS):
            name = self.recipients[i] if i < len(self.recipients) else ""
            rows = [x for x in g if x["recipient"] == name] if name else []
            gifts_n = len(rows)
            bought_n = len([x for x in rows if x["status"] in bought_states])
            wrapped_n = len([x for x in rows
                             if x["wrapped"] == C.TICK
                             or x["status"] in (C.ST_WRAPPED, C.ST_DELIVERED)])
            delivered_n = len([x for x in rows
                               if x["delivered"] == C.TICK
                               or x["status"] == C.ST_DELIVERED])
            rec.append({
                "name": name,
                "planned": sum(x["budget"] or 0 for x in rows),
                "spent": sum(x["cost"] or 0 for x in rows),
                "gifts": gifts_n, "bought": bought_n, "wrapped": wrapped_n,
                "delivered": delivered_n,
                "to_buy": max(0, gifts_n - bought_n),
                "pct": (bought_n / float(gifts_n)) if gifts_n else "",
            })
        a["recipients"] = rec
        a["recipients_gifted"] = len([x for x in rec if x["gifts"] > 0])
        a["avg_per_recipient"] = (a["gift_actual_spent"] /
                                  float(a["recipients_gifted"])
                                  if a["recipients_gifted"] else 0)

        # --- categories ------------------------------------------------
        cats = []
        for i in range(C.CATEGORY_SLOTS):
            name = self.categories[i] if i < len(self.categories) else ""
            cats.append({
                "name": name,
                "spent": sum(x["cost"] or 0 for x in g
                             if name and x["category"] == name)})
        a["categories"] = cats

        # --- status breakdown ------------------------------------------
        a["status_counts"] = [{"name": s,
                               "count": len([x for x in g if x["status"] == s])}
                              for s in C.STATUSES]

        # --- shopping --------------------------------------------------
        sh = self.shopping
        a["shop_total"] = len([x for x in sh if x["item"]])
        a["shop_bought"] = len([x for x in sh if x["bought"] == C.TICK])
        a["shop_spent"] = sum(self._shop_spent(x) for x in sh)

        # --- budget ----------------------------------------------------
        rows = []
        for label, key, recipe, demo_planned in C.BUDGET_CATEGORIES:
            planned = self.budget_planned.get(key, demo_planned if self.demo else 0)
            manual = self.budget_manual.get(key, 0)
            auto = self._budget_auto(recipe)
            actual = manual + auto
            rows.append({"label": label, "key": key, "planned": planned,
                         "manual": manual, "auto": auto, "actual": actual,
                         "remaining": planned - actual,
                         "pct": (actual / float(planned)) if planned else ""})
        a["budget_rows"] = rows
        a["budget_planned"] = sum(x["planned"] for x in rows)
        a["budget_actual"] = sum(x["actual"] for x in rows)
        a["budget_remaining"] = a["budget_planned"] - a["budget_actual"]
        a["budget_pct"] = (a["budget_actual"] / float(a["budget_planned"])
                           if a["budget_planned"] else 0)

        # --- cards -----------------------------------------------------
        a["cards_total"] = len([x for x in self.cards if x["name"]])
        a["cards_written"] = len([x for x in self.cards
                                  if x["written"] == C.TICK])
        a["cards_sent"] = len([x for x in self.cards if x["sent"] == C.TICK])
        a["cards_received"] = len([x for x in self.cards
                                   if x["received"] == C.TICK])
        a["cards_postage"] = sum(x["postage"] or 0 for x in self.cards)

        # --- stockings -------------------------------------------------
        st = self.stockings
        a["stock_items"] = len([x for x in st if x["item"]])
        a["stock_bought"] = len([x for x in st if x["bought"] == C.TICK])
        a["stock_budget"] = sum(x["budget"] or 0 for x in st)
        a["stock_spent"] = sum(self._stock_spent(x) for x in st)
        a["stock_by_owner"] = self._stock_by_owner()

        # --- orders ----------------------------------------------------
        od = self.orders
        a["orders_total"] = len([x for x in od if x["item"]])
        a["orders_outstanding"] = len([
            x for x in od
            if x["status"] not in (C.OS_DELIVERED, C.OS_CANCEL, C.OS_RETURN)])
        a["orders_late"] = len([
            x for x in od
            if x["expected"] and x["expected"] < TODAY
            and x["status"] not in (C.OS_DELIVERED, C.OS_CANCEL)])
        a["orders_value"] = sum(x["cost"] or 0 for x in od)

        # --- to-do -----------------------------------------------------
        td = self.todos
        a["todo_total"] = len([x for x in td if x["task"]])
        a["todo_done"] = len([x for x in td if x["done"] == C.TICK])
        a["todo_overdue"] = len([x for x in td
                                 if x["done"] != C.TICK and x["deadline"]
                                 and x["deadline"] < TODAY])
        a["todo_soon"] = len([x for x in td
                              if x["done"] != C.TICK and x["deadline"]
                              and TODAY <= x["deadline"]
                              <= TODAY + timedelta(days=self.settings["duesoon"])])

        # --- wish list -------------------------------------------------
        wl = self.wishlist
        a["wish_total"] = len([x for x in wl if x["idea"]])
        a["wish_must"] = len([x for x in wl if x["priority"] == C.PR_MUST])
        a["wish_value"] = sum(x["price"] or 0 for x in wl)

        # --- countdown -------------------------------------------------
        a["days_to_event"] = (self.event_date - TODAY).days
        a["event_year"] = self.event_date.year
        a["weeks_to_event"] = max(0, -(-(self.event_date - TODAY).days // 7))

        # --- deadline pool ---------------------------------------------
        a["pool"] = self._deadline_pool()

        # --- text ------------------------------------------------------
        a.update(self._text())

    # ------------------------------------------------------------------
    def _shop_spent(self, row):
        if row["bought"] != C.TICK:
            return 0
        if row.get("cost"):
            return row["cost"]
        return (row.get("qty") or 0) * (row.get("unit") or 0)

    def _stock_spent(self, row):
        if row["bought"] != C.TICK:
            return 0
        if row.get("cost") is not None:
            return row["cost"]
        return row.get("budget") or 0

    def _budget_auto(self, recipe):
        total = 0.0
        for term in recipe:
            if term == "gifts":
                total += sum(x["cost"] or 0 for x in self.gifts)
            elif term == "stock":
                total += sum(self._stock_spent(x) for x in self.stockings)
            elif term == "cards":
                total += sum(x["postage"] or 0 for x in self.cards)
            elif term.startswith("shop:"):
                cat = term.split(":", 1)[1]
                total += sum(self._shop_spent(x) for x in self.shopping
                             if x["category"] == cat)
        return round(total, 2)

    def _stock_by_owner(self):
        out = []
        for i in range(C.RECIPIENT_SLOTS):
            name = self.owners[i] if i < len(self.owners) else ""
            rows = [x for x in self.stockings if x["owner"] == name] if name else []
            budget = sum(x["budget"] or 0 for x in rows)
            spent = sum(self._stock_spent(x) for x in rows)
            out.append({"name": name, "items": len(rows), "budget": budget,
                        "spent": spent, "remaining": budget - spent,
                        "bought": len([x for x in rows
                                       if x["bought"] == C.TICK])})
        return out

    def _deadline_pool(self):
        pool = []
        for x in self.gifts:
            if x["deadline"] and x["deadline"] >= TODAY and \
                    x["status"] != C.ST_DELIVERED:
                pool.append((x["deadline"],
                             "\U0001F381 %s \u2014 %s" % (x["idea"],
                                                         x["recipient"])))
        for x in self.todos:
            if x["done"] != C.TICK and x["deadline"] and x["deadline"] >= TODAY:
                pool.append((x["deadline"], "\u2705 %s" % x["task"]))
        for x in self.orders:
            if x["expected"] and x["expected"] >= TODAY and x["status"] not in (
                    C.OS_DELIVERED, C.OS_CANCEL):
                pool.append((x["expected"],
                             "\U0001F4E6 %s \u2014 %s" % (x["item"],
                                                          x["store"])))
        pool.sort(key=lambda t: t[0])
        return pool

    def _text(self):
        a = self.agg
        cur = self.settings["currency"]
        name = self.settings["event_name"]
        days = a["days_to_event"]
        t = {}
        if days > 1:
            t["countdown"] = "\u23F3  %d Days Until %s  \U0001F381" % (days, name)
        elif days == 1:
            t["countdown"] = "\u23F3  Tomorrow is %s!  \U0001F381" % name
        elif days == 0:
            t["countdown"] = "\U0001F389  It's %s today!  \U0001F384" % name
        else:
            t["countdown"] = ("\U0001F384  %s %d has been and gone \u2014 "
                              "update the date in \u2699\uFE0F Setup to plan "
                              "next year" % (name, self.event_date.year))

        pct = a["budget_pct"]
        if a["budget_planned"] <= 0:
            t["budget_alert"] = ("\U0001F4DD  Set your planned budgets on the "
                                 "\U0001F4B0 Budget tab to turn on the alerts.")
        elif a["budget_actual"] > a["budget_planned"]:
            over = a["budget_actual"] - a["budget_planned"]
            t["budget_alert"] = ("\U0001F534  Over budget by %s%s \u2014 pause "
                                 "the non-essentials."
                                 % (cur, format(over, ",.2f")))
        elif pct >= self.settings["alert"]:
            t["budget_alert"] = ("\u26A0\uFE0F  You've spent %.0f%% of your "
                                 "%s budget \u2014 %s%s left."
                                 % (pct * 100, name, cur,
                                    format(a["budget_remaining"], ",.2f")))
        else:
            t["budget_alert"] = ("\U0001F7E2  On track \u2014 %.0f%% of the "
                                 "budget used, %s%s still to spend."
                                 % (pct * 100, cur,
                                    format(a["budget_remaining"], ",.2f")))

        t["bar_budget"] = _bar(a["budget_pct"])
        t["bar_gifts"] = _bar(a["gift_completion"])
        t["bar_cards"] = _bar(a["cards_sent"] / float(a["cards_total"])
                              if a["cards_total"] else 0)
        t["bar_stock"] = _bar(a["stock_bought"] / float(a["stock_items"])
                              if a["stock_items"] else 0)
        t["bar_shop"] = _bar(a["shop_bought"] / float(a["shop_total"])
                             if a["shop_total"] else 0)
        t["bar_todo"] = _bar(a["todo_done"] / float(a["todo_total"])
                             if a["todo_total"] else 0)
        return t

    # ------------------------------------------------------------------
    def money(self, value, decimals=0):
        fmt = ",.%df" % decimals
        return self.settings["currency"] + format(value or 0, fmt)


GIFT_RECIPIENTS = [
    "Mum", "Dad", "Grandma Rose", "Grandpa Joe", "Sister Emily",
    "Brother Jack", "Niece Sophie", "Nephew Leo", "Alex (partner)",
    "Maya (best friend)", "Cousin Dan", "Aunt Carol", "Uncle Bill",
    "Secret Santa",
]
STOCK_OWNERS = ["Emma", "Oliver", "Mum", "Dad"]


def _bar(pct, blocks=18):
    try:
        pct = min(1.0, max(0.0, float(pct)))
    except (TypeError, ValueError):
        pct = 0.0
    filled = int(round(pct * blocks))
    return "\u2588" * filled + "\u2591" * (blocks - filled)
```

### `christmas_tracker/workbook.py`  (93 lines)

```python
"""
Orchestrator: turns (edition, theme, mode) into a finished .xlsx.

Editions
    basic    dashboard + gift tracker + budget + shopping + setup + guide
    premium  everything: wish list, orders, wrapping, cards, stockings, to-do

Themes
    festive  cream / pine / burgundy / gold
    minimal  white / sage / clay

Modes
    blank    clean, ready-to-use template (lists + checklist pre-seeded)
    demo     filled-in example family, for listing screenshots
"""

import os

from . import theme as themes
from .book import Book
from .demo import Model
from .sheets import (budget, cards, dashboard, data, gifts, guide, orders,
                     setup, shopping, stockings, todo, wishlist, wrapping)

BUILDERS = {
    "data": data.build,
    "setup": setup.build,
    "gifts": gifts.build,
    "budget": budget.build,
    "shopping": shopping.build,
    "wishlist": wishlist.build,
    "orders": orders.build,
    "wrapping": wrapping.build,
    "cards": cards.build,
    "stockings": stockings.build,
    "todo": todo.build,
    "dashboard": dashboard.build,
    "guide": guide.build,
}

# Build order: the hidden engine first, then the tabs left-to-right.
BUILD_ORDER = ["data", "setup", "gifts", "budget", "shopping", "wishlist",
               "orders", "wrapping", "cards", "stockings", "todo",
               "dashboard", "guide"]


def build_workbook(path, edition="premium", theme_name="festive",
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
    return ("Christmas_Gift_Tracker_%s_%s%s.xlsx"
            % (edition.upper(), theme_name.capitalize(), suffix))


def build_all(outdir="products", protect=None, images=True):
    """The set of files an Etsy listing actually ships / screenshots."""
    combos = [
        ("premium", "festive", "demo"),    # listing hero + screenshots
        ("premium", "festive", "blank"),   # the product
        ("premium", "minimal", "blank"),   # second style
        ("basic", "festive", "blank"),     # the cheaper tier
        ("basic", "minimal", "blank"),
        ("basic", "festive", "demo"),      # cheaper tier screenshot
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

### `christmas_tracker/sheets/__init__.py`  (2 lines)

```python
"""Worksheet builders - one module per tab."""
```

### `christmas_tracker/sheets/common.py`  (304 lines)

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
    for field in fields:
        bk.validate(key, C.ROW_FIRST, ci(bk.col(key, field)),
                    C.last_row(key), ci(bk.col(key, field)), "=Tick",
                    title="Tick it off", message=message)


def money_dv(bk, key, fields, label="money"):
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

### `christmas_tracker/sheets/data.py`  (495 lines)

```python
"""
The hidden ``_Data`` worksheet.

Everything the dashboard, the charts and the "what's left to do" panel need is
calculated once, here:

  A:C   occasion presets (feeds the Setup "reuse for any event" helper)
  E:F   the preset the user picked, resolved to a month/day
  H:P   per-recipient gift + money summary
  Q:R   spend per gift category            (pie / bar charts)
  T:U   gift status breakdown              (doughnut chart)
  AA:AC the upcoming-deadline pool         (dashboard "next 5 deadlines")
  AE:AF the KPI table - every headline number in the workbook

Keeping the maths on one hidden sheet means the visible tabs stay clean and
the dashboard formulas stay readable (``=_Data!$AF$5`` instead of a
nine-line SUMPRODUCT).
"""

from .. import config as C
from ..book import r, ci

TODAY_FN = "TODAY()"


def _setup_cell(col, row):
    return "'%s'!$%s$%d" % (C.SHEET_NAMES["setup"], col, row)


def build(bk):
    ws = bk.ws("data")
    S = bk.S
    m = bk.demo
    agg = m.agg

    hdr = S.f(**S.base(font_size=9, bold=True, font_color=bk.th.white,
                       bg_color=bk.th.muted, align="center", valign="vcenter",
                       border=1, border_color=bk.th.muted))
    txt = S.f(**S.base(font_size=9, font_color=bk.th.ink, align="left"))
    num = S.f(**S.base(font_size=9, num_format="#,##0.00"))
    integer = S.f(**S.base(font_size=9, num_format="0"))
    dt = S.f(**S.base(font_size=9, num_format="dd mmm yyyy"))
    pct = S.f(**S.base(font_size=9, num_format="0%"))

    ws.set_column("A:A", 20)
    ws.set_column("B:C", 8)
    ws.set_column("D:D", 2)
    ws.set_column("E:F", 10)
    ws.set_column("G:G", 2)
    ws.set_column("H:H", 20)
    ws.set_column("I:P", 11)
    ws.set_column("Q:Q", 18)
    ws.set_column("R:R", 11)
    ws.set_column("S:S", 2)
    ws.set_column("T:T", 18)
    ws.set_column("U:U", 9)
    ws.set_column("V:Z", 2)
    ws.set_column("AA:AA", 13)
    ws.set_column("AB:AB", 44)
    ws.set_column("AC:AC", 7)
    ws.set_column("AD:AD", 2)
    ws.set_column("AE:AE", 30)
    ws.set_column("AF:AF", 14)
    ws.set_column("AG:AH", 40)

    # ------------------------------------------------------------------
    # occasion presets
    # ------------------------------------------------------------------
    for col, label in (("A", "Occasion"), ("B", "Month"), ("C", "Day")):
        ws.write(r(1), ci(col), label, hdr)
    for i, (name, month, day) in enumerate(C.OCCASIONS):
        row = C.DATA_PRESET_FIRST + i
        ws.write(r(row), ci("A"), name, txt)
        ws.write(r(row), ci("B"), month if month else "", integer)
        ws.write(r(row), ci("C"), day if day else "", integer)

    ws.write(r(1), ci("E"), "Preset month", hdr)
    ws.write(r(1), ci("F"), "Preset day", hdr)
    occ_first = C.DATA_PRESET_FIRST
    occ_last = C.DATA_PRESET_FIRST + C.OCCASION_ROWS - 1
    ws.write_formula(
        r(2), ci("E"),
        "=IFERROR(INDEX($B$%d:$B$%d,MATCH(%s,$A$%d:$A$%d,0)),\"\")"
        % (occ_first, occ_last, _setup_cell("C", C.SU_OCCASION),
           occ_first, occ_last), integer, "")
    ws.write_formula(
        r(2), ci("F"),
        "=IFERROR(INDEX($C$%d:$C$%d,MATCH(%s,$A$%d:$A$%d,0)),\"\")"
        % (occ_first, occ_last, _setup_cell("C", C.SU_OCCASION),
           occ_first, occ_last), integer, "")
    bk.stats["formulas"] += 2

    # ------------------------------------------------------------------
    # per-recipient summary  (H:P)
    # ------------------------------------------------------------------
    rec_headers = ["Recipient", "Planned", "Spent", "Gifts", "Bought",
                   "Wrapped", "Delivered", "To buy", "Done %"]
    for i, label in enumerate(rec_headers):
        ws.write(r(1), ci("H") + i, label, hdr)

    g = bk.q("gifts")
    gt_recip = bk.rng("gifts", "recipient")
    gt_budget = bk.rng("gifts", "budget")
    gt_cost = bk.rng("gifts", "cost")
    gt_status = bk.rng("gifts", "status")
    gt_wrap = bk.rng("gifts", "wrapped")
    gt_deliv = bk.rng("gifts", "delivered")

    for i in range(C.DATA_REC_ROWS):
        row = C.DATA_REC_FIRST + i
        setup_row = C.SU_LIST_FIRST + i
        name_ref = "$H%d" % row
        cached_rec = agg["recipients"][i]
        ws.write_formula(r(row), ci("H"),
                         "=IF(%s=\"\",\"\",%s)"
                         % (_setup_cell("B", setup_row),
                            _setup_cell("B", setup_row)),
                         txt, cached_rec["name"])
        ws.write_formula(r(row), ci("I"),
                         "=IF(%s=\"\",0,SUMIF(%s,%s,%s))"
                         % (name_ref, gt_recip, name_ref, gt_budget),
                         num, cached_rec["planned"])
        ws.write_formula(r(row), ci("J"),
                         "=IF(%s=\"\",0,SUMIF(%s,%s,%s))"
                         % (name_ref, gt_recip, name_ref, gt_cost),
                         num, cached_rec["spent"])
        ws.write_formula(r(row), ci("K"),
                         "=IF(%s=\"\",0,COUNTIF(%s,%s))"
                         % (name_ref, gt_recip, name_ref),
                         integer, cached_rec["gifts"])
        bought = "".join(
            "+COUNTIFS(%s,%s,%s,\"%s\")" % (gt_recip, name_ref, gt_status, s)
            for s in (C.ST_BOUGHT, C.ST_WRAPPED, C.ST_DELIVERED))
        ws.write_formula(r(row), ci("L"),
                         "=IF(%s=\"\",0%s+IF(CountOrdered=\"Yes\","
                         "COUNTIFS(%s,%s,%s,\"%s\"),0))"
                         % (name_ref, bought, gt_recip, name_ref, gt_status,
                            C.ST_ORDERED),
                         integer, cached_rec["bought"])
        ws.write_formula(r(row), ci("M"),
                         "=IF(%s=\"\",0,COUNTIFS(%s,%s,%s,\"%s\")"
                         "+SUMPRODUCT((%s=%s)*(%s<>\"%s\")*((%s=\"%s\")"
                         "+(%s=\"%s\")>0)))"
                         % (name_ref, gt_recip, name_ref, gt_wrap, C.TICK,
                            gt_recip, name_ref, gt_wrap, C.TICK, gt_status,
                            C.ST_WRAPPED, gt_status, C.ST_DELIVERED),
                         integer, cached_rec["wrapped"])
        ws.write_formula(r(row), ci("N"),
                         "=IF(%s=\"\",0,COUNTIFS(%s,%s,%s,\"%s\")"
                         "+SUMPRODUCT((%s=%s)*(%s<>\"%s\")*(%s=\"%s\")))"
                         % (name_ref, gt_recip, name_ref, gt_deliv, C.TICK,
                            gt_recip, name_ref, gt_deliv, C.TICK, gt_status,
                            C.ST_DELIVERED),
                         integer, cached_rec["delivered"])
        ws.write_formula(r(row), ci("O"),
                         "=IF(%s=\"\",0,MAX(0,$K%d-$L%d))" % (name_ref, row,
                                                              row),
                         integer, cached_rec["to_buy"])
        ws.write_formula(r(row), ci("P"),
                         "=IF(%s=\"\",\"\",IFERROR($L%d/$K%d,0))"
                         % (name_ref, row, row), pct, cached_rec["pct"])
        bk.stats["formulas"] += 9

    # ------------------------------------------------------------------
    # spend by gift category  (Q:R)
    # ------------------------------------------------------------------
    ws.write(r(1), ci("Q"), "Gift category", hdr)
    ws.write(r(1), ci("R"), "Spent", hdr)
    gt_cat = bk.rng("gifts", "category")
    for i in range(C.DATA_CAT_ROWS):
        row = C.DATA_CAT_FIRST + i
        setup_row = C.SU_LIST_FIRST + i
        cat_ref = "$Q%d" % row
        cached_cat = agg["categories"][i]
        ws.write_formula(r(row), ci("Q"),
                         "=IF(%s=\"\",\"\",%s)"
                         % (_setup_cell("D", setup_row),
                            _setup_cell("D", setup_row)),
                         txt, cached_cat["name"])
        ws.write_formula(r(row), ci("R"),
                         "=IF(%s=\"\",0,SUMIF(%s,%s,%s))"
                         % (cat_ref, gt_cat, cat_ref, gt_cost),
                         num, cached_cat["spent"])
        bk.stats["formulas"] += 2

    # ------------------------------------------------------------------
    # gift status breakdown  (T:U)
    # ------------------------------------------------------------------
    ws.write(r(1), ci("T"), "Status", hdr)
    ws.write(r(1), ci("U"), "Gifts", hdr)
    for i, status in enumerate(C.STATUSES):
        row = C.DATA_STATUS_FIRST + i
        ws.write_formula(r(row), ci("T"),
                         "=%s" % _setup_cell("B", C.SU_FIXED_FIRST + i),
                         txt, status)
        ws.write_formula(r(row), ci("U"),
                         "=COUNTIF(%s,$T%d)" % (gt_status, row),
                         integer, agg["status_counts"][i]["count"])
        bk.stats["formulas"] += 2

    # ------------------------------------------------------------------
    # deadline pool  (AA:AC)
    # ------------------------------------------------------------------
    ws.write(r(1), ci("AA"), "Due", hdr)
    ws.write(r(1), ci("AB"), "What", hdr)
    ws.write(r(1), ci("AC"), "Days", hdr)
    pool = _pool(bk)
    for i, (key, label, days) in enumerate(pool):
        row = C.DATA_POOL_FIRST + i
        ws.write_formula(r(row), ci("AA"), key[0], dt, key[1])
        ws.write_formula(r(row), ci("AB"), label[0], txt, label[1])
        ws.write_formula(r(row), ci("AC"), days[0], integer, days[1])
        bk.stats["formulas"] += 3

    # ------------------------------------------------------------------
    # KPI table  (AE:AF)
    # ------------------------------------------------------------------
    ws.write(r(1), ci("AE"), "Metric", hdr)
    ws.write(r(1), ci("AF"), "Value", hdr)
    for key, formula, numfmt in _kpis(bk):
        row = C.KPI_ROW[key]
        fmt = S.f(**S.base(font_size=9, num_format=numfmt,
                           font_color=bk.th.primary, bold=True))
        cached = agg.get(key, 0)
        if isinstance(cached, bool):
            cached = int(cached)
        ws.write(r(row), ci("AE"), C.KPI_LABEL[key], txt)
        ws.write_formula(r(row), ci("AF"), formula, fmt,
                         cached if cached is not None else 0)
        bk.stats["formulas"] += 1

    ws.write(r(1), ci("AH"),
             "This sheet is hidden on purpose: it feeds the dashboard, the "
             "charts and every automatic number in the workbook. Please do "
             "not delete or rename it (unhide with: right-click any tab > "
             "Unhide).", S.f(**S.base(font_size=10, italic=True,
                                      font_color=bk.th.muted, align="left",
                                      valign="vcenter")))
    ws.set_row(r(1), 22)
    return ws


# ---------------------------------------------------------------------------
# deadline pool
# ---------------------------------------------------------------------------
def _pool(bk):
    """Return [(date_formula, cached), (label_formula, cached),
    (days_formula, cached)] for every pool row."""
    m = bk.demo
    rows = []

    gt_dl = bk.col("gifts", "deadline")
    gt_st = bk.col("gifts", "status")
    gt_idea = bk.col("gifts", "idea")
    gt_rec = bk.col("gifts", "recipient")
    for i in range(C.POOL_GIFT_ROWS):
        srow = C.ROW_FIRST + i
        prow = C.DATA_POOL_FIRST + i
        f_date = ("=IF(OR('{g}'!${d}${s}=\"\",'{g}'!${d}${s}<TODAY(),"
                  "'{g}'!${st}${s}=\"{dl}\"),\"\",'{g}'!${d}${s})"
                  ).format(g=C.SHEET_NAMES["gifts"], d=gt_dl, s=srow,
                           st=gt_st, dl=C.ST_DELIVERED)
        f_lab = ("=IF($AA${p}=\"\",\"\",\"\U0001F381 \"&'{g}'!${i}${s}&"
                 "\" \u2014 \"&'{g}'!${c}${s})"
                 ).format(p=prow, g=C.SHEET_NAMES["gifts"], i=gt_idea, s=srow,
                          c=gt_rec)
        f_days = "=IF($AA${p}=\"\",\"\",$AA${p}-TODAY())".format(p=prow)
        cached = ("", "", "")
        if i < len(m.gifts):
            g = m.gifts[i]
            dl = g["deadline"]
            if dl and dl >= _today() and g["status"] != C.ST_DELIVERED:
                cached = (dl, "\U0001F381 %s \u2014 %s" % (g["idea"],
                                                           g["recipient"]),
                          (dl - _today()).days)
        rows.append(((f_date, cached[0]), (f_lab, cached[1]),
                     (f_days, cached[2])))

    if bk.has("todo"):
        td_dl = bk.col("todo", "deadline")
        td_done = bk.col("todo", "done")
        td_task = bk.col("todo", "task")
        for i in range(C.POOL_TODO_ROWS):
            srow = C.ROW_FIRST + i
            prow = C.DATA_POOL_FIRST + C.POOL_GIFT_ROWS + i
            f_date = ("=IF(OR('{t}'!${d}${s}=\"\",'{t}'!${dn}${s}=\"{tick}\","
                      "'{t}'!${d}${s}<TODAY()),\"\",'{t}'!${d}${s})"
                      ).format(t=C.SHEET_NAMES["todo"], d=td_dl, s=srow,
                               dn=td_done, tick=C.TICK)
            f_lab = ("=IF($AA${p}=\"\",\"\",\"\u2705 \"&'{t}'!${c}${s})"
                     ).format(p=prow, t=C.SHEET_NAMES["todo"], c=td_task,
                              s=srow)
            f_days = "=IF($AA${p}=\"\",\"\",$AA${p}-TODAY())".format(p=prow)
            cached = ("", "", "")
            if i < len(m.todos):
                t = m.todos[i]
                dl = t["deadline"]
                if t["done"] != C.TICK and dl and dl >= _today():
                    cached = (dl, "\u2705 %s" % t["task"],
                              (dl - _today()).days)
            rows.append(((f_date, cached[0]), (f_lab, cached[1]),
                         (f_days, cached[2])))

    if bk.has("orders"):
        od_ex = bk.col("orders", "expected")
        od_st = bk.col("orders", "status")
        od_item = bk.col("orders", "item")
        od_store = bk.col("orders", "store")
        for i in range(C.POOL_ORDER_ROWS):
            srow = C.ROW_FIRST + i
            prow = (C.DATA_POOL_FIRST + C.POOL_GIFT_ROWS + C.POOL_TODO_ROWS
                    + i)
            f_date = ("=IF(OR('{o}'!${d}${s}=\"\",'{o}'!${d}${s}<TODAY(),"
                      "'{o}'!${st}${s}=\"{del1}\",'{o}'!${st}${s}=\"{can}\"),"
                      "\"\",'{o}'!${d}${s})").format(
                          o=C.SHEET_NAMES["orders"], d=od_ex, s=srow,
                          st=od_st, del1=C.OS_DELIVERED, can=C.OS_CANCEL)
            f_lab = ("=IF($AA${p}=\"\",\"\",\"\U0001F4E6 \"&'{o}'!${i}${s}&"
                     "\" \u2014 \"&'{o}'!${st}${s})").format(
                         p=prow, o=C.SHEET_NAMES["orders"], i=od_item, s=srow,
                         st=od_store)
            f_days = "=IF($AA${p}=\"\",\"\",$AA${p}-TODAY())".format(p=prow)
            cached = ("", "", "")
            if i < len(m.orders):
                o = m.orders[i]
                ex = o["expected"]
                if ex and ex >= _today() and o["status"] not in (
                        C.OS_DELIVERED, C.OS_CANCEL):
                    cached = (ex,
                              "\U0001F4E6 %s \u2014 %s" % (o["item"],
                                                           o["store"]),
                              (ex - _today()).days)
            rows.append(((f_date, cached[0]), (f_lab, cached[1]),
                         (f_days, cached[2])))
    return rows


def _today():
    from datetime import date
    return date.today()


# ---------------------------------------------------------------------------
# KPI formulas
# ---------------------------------------------------------------------------
def _kpis(bk):
    """Yield (key, formula, number_format) for the whole KPI table."""
    m = bk.demo
    out = []

    def k(name):
        return "$AF$%d" % C.KPI_ROW[name]

    def kfmt(name):
        return C.KPI_FMT[name]

    gt = bk.q("gifts")
    out.append(("days_to_event", "=EventDate-TODAY()", kfmt("days_to_event")))
    out.append(("event_year", "=YEAR(EventDate)", kfmt("event_year")))

    st = bk.rng("gifts", "status")
    idea = bk.rng("gifts", "idea")
    wrap = bk.rng("gifts", "wrapped")
    deliv = bk.rng("gifts", "delivered")

    out.append(("gifts_planned", "=COUNTA(%s)" % idea, kfmt("gifts_planned")))
    bought = "+".join('COUNTIF(%s,"%s")' % (st, s)
                      for s in (C.ST_BOUGHT, C.ST_WRAPPED, C.ST_DELIVERED))
    bought += '+IF(CountOrdered="Yes",COUNTIF(%s,"%s"),0)' % (st, C.ST_ORDERED)
    out.append(("gifts_purchased", "=" + bought, kfmt("gifts_purchased")))
    out.append(("gifts_wrapped",
                '=COUNTIF(%s,"%s")+SUMPRODUCT((%s<>"%s")*((%s="%s")+(%s="%s")>0))'
                % (wrap, C.TICK, wrap, C.TICK, st, C.ST_WRAPPED, st,
                   C.ST_DELIVERED), kfmt("gifts_wrapped")))
    out.append(("gifts_delivered",
                '=COUNTIF(%s,"%s")+SUMPRODUCT((%s<>"%s")*(%s="%s"))'
                % (deliv, C.TICK, deliv, C.TICK, st, C.ST_DELIVERED),
                kfmt("gifts_delivered")))
    out.append(("gifts_to_buy",
                '=COUNTIF(%s,"%s")+COUNTIF(%s,"%s")'
                % (st, C.ST_IDEA, st, C.ST_NEED), kfmt("gifts_to_buy")))
    out.append(("gifts_ordered", '=COUNTIF(%s,"%s")' % (st, C.ST_ORDERED),
                kfmt("gifts_ordered")))
    out.append(("wrap_to_do",
                "=MAX(0,%s-%s)" % (k("gifts_purchased"), k("gifts_wrapped")),
                kfmt("wrap_to_do")))
    out.append(("deliver_to_do",
                "=MAX(0,%s-%s)" % (k("gifts_wrapped"), k("gifts_delivered")),
                kfmt("deliver_to_do")))
    out.append(("gift_completion",
                "=IFERROR(%s/%s,0)" % (k("gifts_purchased"),
                                       k("gifts_planned")),
                kfmt("gift_completion")))
    out.append(("recipients_gifted",
                "=COUNTIF($K$%d:$K$%d,\">0\")"
                % (C.DATA_REC_FIRST, C.DATA_REC_FIRST + C.DATA_REC_ROWS - 1),
                kfmt("recipients_gifted")))
    out.append(("avg_per_recipient",
                "=IFERROR(%s/%s,0)" % (k("gift_actual_spent"),
                                       k("recipients_gifted")),
                kfmt("avg_per_recipient")))
    out.append(("gift_budget_planned", "=SUM(%s)" % bk.rng("gifts", "budget"),
                kfmt("gift_budget_planned")))
    out.append(("gift_actual_spent", "=SUM(%s)" % bk.rng("gifts", "cost"),
                kfmt("gift_actual_spent")))

    bud = bk.q("budget")
    out.append(("budget_planned",
                "=SUM(%s!$%s$%d:$%s$%d)"
                % (bud, C.BUDGET_COLS["planned"], C.BUD_FIRST,
                   C.BUDGET_COLS["planned"], C.BUD_FIRST + C.BUDGET_ROWS - 1),
                kfmt("budget_planned")))
    out.append(("budget_actual",
                "=SUM(%s!$%s$%d:$%s$%d)"
                % (bud, C.BUDGET_COLS["actual"], C.BUD_FIRST,
                   C.BUDGET_COLS["actual"], C.BUD_FIRST + C.BUDGET_ROWS - 1),
                kfmt("budget_actual")))
    out.append(("budget_remaining",
                "=%s-%s" % (k("budget_planned"), k("budget_actual")),
                kfmt("budget_remaining")))
    out.append(("budget_pct",
                "=IFERROR(%s/%s,0)" % (k("budget_actual"), k("budget_planned")),
                kfmt("budget_pct")))

    def simple(key, formula, zero="0"):
        if bk.has(_sheet_for(key)):
            return (key, formula, kfmt(key))
        return (key, "=" + zero, kfmt(key))

    sh = bk.q("shopping")
    out.append(simple("shop_total", "=COUNTA(%s)" % bk.rng("shopping", "item")))
    out.append(simple("shop_bought", '=COUNTIF(%s,"%s")'
                      % (bk.rng("shopping", "bought"), C.TICK)))
    out.append(simple("shop_spent", "=SUM(%s)" % bk.rng("shopping", "spent")))

    out.append(simple("cards_total", "=COUNTA(%s)" % bk.rng("cards", "name")))
    out.append(simple("cards_written", '=COUNTIF(%s,"%s")'
                      % (bk.rng("cards", "written"), C.TICK)))
    out.append(simple("cards_sent", '=COUNTIF(%s,"%s")'
                      % (bk.rng("cards", "sent"), C.TICK)))

    out.append(simple("stock_budget", "=SUM(%s)"
                      % bk.rng("stockings", "budget")))
    out.append(simple("stock_spent", "=SUM(%s)" % bk.rng("stockings", "spent")))
    out.append(simple("stock_items", "=COUNTA(%s)"
                      % bk.rng("stockings", "item")))
    out.append(simple("stock_bought", '=COUNTIF(%s,"%s")'
                      % (bk.rng("stockings", "bought"), C.TICK)))

    ost = bk.rng("orders", "status")
    oex = bk.rng("orders", "expected")
    out.append(simple("orders_total", "=COUNTA(%s)"
                      % bk.rng("orders", "item")))
    out.append(simple("orders_outstanding",
                      '=SUMPRODUCT((%s<>"")*(%s<>"%s")*(%s<>"%s")*(%s<>"%s"))'
                      % (ost, ost, C.OS_DELIVERED, ost, C.OS_CANCEL, ost,
                         C.OS_RETURN)))
    out.append(simple("orders_late",
                      '=SUMPRODUCT((%s<>"")*(%s<TODAY())*(%s<>"%s")*(%s<>"%s"))'
                      % (oex, oex, ost, C.OS_DELIVERED, ost, C.OS_CANCEL)))
    out.append(simple("orders_value", "=SUM(%s)" % bk.rng("orders", "cost")))

    td = bk.rng("todo", "deadline")
    tdn = bk.rng("todo", "done")
    out.append(simple("todo_total", "=COUNTA(%s)" % bk.rng("todo", "task")))
    out.append(simple("todo_done", '=COUNTIF(%s,"%s")' % (tdn, C.TICK)))
    out.append(simple("todo_overdue",
                      '=SUMPRODUCT((%s<>"%s")*(%s<>"")*(%s<TODAY()))'
                      % (tdn, C.TICK, td, td)))

    out.append(simple("wish_total", "=COUNTA(%s)" % bk.rng("wishlist", "idea")))
    out.append(simple("wish_must", '=COUNTIF(%s,"%s")'
                      % (bk.rng("wishlist", "priority"), C.PR_MUST)))
    out.append(simple("wish_value", "=SUM(%s)" % bk.rng("wishlist", "price")))

    return out


_SHEET_FOR = {
    "shop_total": "shopping", "shop_bought": "shopping",
    "shop_spent": "shopping", "cards_total": "cards",
    "cards_written": "cards", "cards_sent": "cards",
    "stock_budget": "stockings", "stock_spent": "stockings",
    "stock_items": "stockings", "stock_bought": "stockings",
    "orders_total": "orders", "orders_outstanding": "orders",
    "orders_late": "orders", "orders_value": "orders",
    "todo_total": "todo", "todo_done": "todo", "todo_overdue": "todo",
    "wish_total": "wishlist", "wish_must": "wishlist",
    "wish_value": "wishlist",
}


def _sheet_for(key):
    return _SHEET_FOR.get(key, "gifts")
```

### `christmas_tracker/sheets/setup.py`  (384 lines)

```python
"""
The ⚙️ Setup tab: event details, money/alert settings and every editable
dropdown list in the workbook.
"""

from datetime import date

from .. import config as C
from ..book import r, ci
from . import common as K

LAST_COL = "J"


def _suggested_date(occasion):
    """Python twin of the suggested-date formula on the Setup tab."""
    today = date.today()
    for name, month, day in C.OCCASIONS:
        if name != occasion:
            continue
        if not month:
            return ""
        year = today.year
        try:
            candidate = date(year, month, day)
        except ValueError:
            return ""
        if candidate < today:
            candidate = date(year + 1, month, day)
        return candidate
    return ""


def build(bk):
    key = "setup"
    ws = bk.ws(key)
    S = bk.S
    th = bk.th
    m = bk.demo

    bk.widths(key, {"A": 2.2, "B": 34, "C": 24, "D": 18, "E": 18, "F": 18,
                    "G": 22, "H": 18, "I": 18, "J": 18})
    bk.paint(key, 0, 0, 84, ci(LAST_COL), S.canvas)

    # ------------------------------------------------------------------
    # title band
    # ------------------------------------------------------------------
    ws.set_row(r(1), 7)
    ws.set_row(r(2), 32)
    ws.set_row(r(3), 18)
    ws.merge_range(r(2), 1, r(2), ci(LAST_COL) - 3,
                   "\u2699\uFE0F  Setup \u2014 make it yours", S.sheet_title)
    ws.merge_range(r(3), 1, r(3), ci(LAST_COL) - 3,
                   "  Change these settings once and every tab, countdown, "
                   "alert and dropdown follows automatically.", S.sheet_sub)
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
                          bg_color=th.card, align="left", valign="vcenter",
                          border=1, border_color=th.border, indent=1))
    input_f = S.f(**S.base(font_size=11, bold=True, font_color=th.primary,
                           bg_color=th.gold_soft, align="center",
                           valign="vcenter", border=2, border_color=th.gold,
                           locked=False))
    auto_f = S.f(**S.base(font_size=11, bold=True, font_color=th.primary_2,
                          bg_color=th.primary_soft, align="center",
                          valign="vcenter", border=1, border_color=th.border))

    def setting(row, text, note, value, fmt, cached=None, height=24):
        ws.set_row(r(row), height)
        ws.write(r(row), ci("B"), text, label)
        if isinstance(value, str) and value.startswith("="):
            ws.write_formula(r(row), ci("C"), value, fmt,
                             cached if cached is not None else 0)
            bk.stats["formulas"] += 1
        elif isinstance(value, date):
            ws.write_datetime(r(row), ci("C"), value, fmt)
        else:
            ws.write(r(row), ci("C"), value, fmt)
        ws.merge_range(r(row), ci("D"), r(row), ci(LAST_COL), note, note_f)

    def section(row, text, style=None):
        ws.set_row(r(row), 24)
        ws.merge_range(r(row), 1, r(row), ci(LAST_COL), text,
                       style or S.section)

    # ------------------------------------------------------------------
    # 1. your event
    # ------------------------------------------------------------------
    section(5, "  \U0001F384  YOUR EVENT")
    date_fmt = S.f(**S.base(font_size=11, bold=True, font_color=th.primary,
                            bg_color=th.gold_soft, align="center",
                            valign="vcenter", border=2, border_color=th.gold,
                            num_format="dd mmm yyyy", locked=False))
    money_fmt = S.f(**S.base(font_size=11, bold=True, font_color=th.primary,
                             bg_color=th.gold_soft, align="center",
                             valign="vcenter", border=2, border_color=th.gold,
                             num_format="#,##0.00", locked=False))
    pct_fmt = S.f(**S.base(font_size=11, bold=True, font_color=th.primary,
                           bg_color=th.gold_soft, align="center",
                           valign="vcenter", border=2, border_color=th.gold,
                           num_format="0%", locked=False))
    num_fmt = S.f(**S.base(font_size=11, bold=True, font_color=th.primary,
                           bg_color=th.gold_soft, align="center",
                           valign="vcenter", border=2, border_color=th.gold,
                           num_format="0", locked=False))

    setting(C.SU_EVENT_NAME, "Event name",
            "Christmas, 40th Birthday, Baby Shower\u2026 it appears on the "
            "dashboard and in every message.",
            m.settings["event_name"], input_f)
    setting(C.SU_EVENT_DATE, "Event date",
            "The countdown, the deadline alerts and the to-do dates all work "
            "from this one cell.",
            date(*C.DEFAULT_EVENT_DATE), date_fmt)
    setting(C.SU_DAYS, "Days to go  (auto)",
            "Recalculates every day by itself.",
            "=IF($C$%d=\"\",\"\",$C$%d-TODAY())"
            % (C.SU_EVENT_DATE, C.SU_EVENT_DATE), auto_f,
            cached=m.agg["days_to_event"])
    setting(C.SU_YEAR, "Event year  (auto)", "",
            "=IF($C$%d=\"\",\"\",YEAR($C$%d))"
            % (C.SU_EVENT_DATE, C.SU_EVENT_DATE), auto_f,
            cached=m.agg["event_year"])
    setting(C.SU_WEEKS, "Weeks to go  (auto)",
            "Rough planning guide: gifts by week 4, cards by week 3, "
            "wrapping by week 1.",
            "=IF($C$%d=\"\",\"\",MAX(0,ROUNDUP(($C$%d-TODAY())/7,0)))"
            % (C.SU_EVENT_DATE, C.SU_EVENT_DATE), auto_f,
            cached=m.agg["weeks_to_event"])
    ws.set_row(r(C.SU_MESSAGE), 26)
    ws.write(r(C.SU_MESSAGE), ci("B"), "Countdown message  (auto)", label)
    msg = ("=IF($C$%d=\"\",\"\U0001F449  Enter your event date to start the "
           "countdown\",IF($C$%d-TODAY()>1,\"\u23F3  \"&($C$%d-TODAY())&\" "
           "days until \"&$C$%d&\"  \u2022  \"&TEXT($C$%d,\"dddd d mmmm "
           "yyyy\"),IF($C$%d-TODAY()=1,\"\u23F3  Tomorrow is \"&$C$%d&\"!  "
           "\U0001F381\",IF($C$%d=TODAY(),\"\U0001F389  It's \"&$C$%d&\" "
           "today!  \U0001F384\",\"\U0001F384  \"&$C$%d&\" \"&TEXT($C$%d,"
           "\"yyyy\")&\" has been and gone \u2014 change the date above to "
           "plan the next one.\"))))"
           % ((C.SU_EVENT_DATE, C.SU_EVENT_DATE, C.SU_EVENT_DATE,
               C.SU_EVENT_NAME, C.SU_EVENT_DATE, C.SU_EVENT_DATE,
               C.SU_EVENT_NAME, C.SU_EVENT_DATE, C.SU_EVENT_NAME,
               C.SU_EVENT_NAME, C.SU_EVENT_DATE)))
    ws.merge_range(r(C.SU_MESSAGE), ci("C"), r(C.SU_MESSAGE), ci(LAST_COL),
                   "", S.f(**S.base(font_size=11, bold=True,
                                    font_color=th.white, bg_color=th.accent,
                                    align="left", valign="vcenter", indent=1)))
    ws.write_formula(r(C.SU_MESSAGE), ci("C"), msg,
                     S.f(**S.base(font_size=11, bold=True, font_color=th.white,
                                  bg_color=th.accent, align="left",
                                  valign="vcenter", indent=1)),
                     m.agg["countdown"])
    bk.stats["formulas"] += 1
    ws.set_row(r(12), 8)

    # ------------------------------------------------------------------
    # 2. reuse for any occasion
    # ------------------------------------------------------------------
    section(13, "  \u267B\uFE0F  REUSE IT FOR ANY OCCASION  "
                "(birthdays, Valentine's, Mother's Day, weddings, "
                "Secret Santa\u2026)")
    setting(C.SU_OCCASION, "Pick an occasion",
            "Just a helper \u2014 it suggests a name and a date for you.",
            m.settings["occasion"], input_f)
    setting(C.SU_SUG_NAME, "Suggested event name  (auto)", "",
            "=IF($C$%d=\"\",\"\",$C$%d)" % (C.SU_OCCASION, C.SU_OCCASION),
            auto_f, cached=m.settings["occasion"])
    d = bk.q("data")
    auto_date = S.f(**S.base(font_size=11, bold=True, font_color=th.primary_2,
                             bg_color=th.primary_soft, align="center",
                             valign="vcenter", border=1,
                             border_color=th.border,
                             num_format="dd mmm yyyy"))
    suggested = _suggested_date(m.settings["occasion"])
    setting(C.SU_SUG_DATE, "Suggested date  (auto)",
            "Copy the two suggestions above into Event name / Event date and "
            "the whole workbook switches occasion.",
            "=IF(%s!$E$2=\"\",\"\",DATE(YEAR(TODAY())+IF(DATE(YEAR(TODAY()),"
            "%s!$E$2,%s!$F$2)<TODAY(),1,0),%s!$E$2,%s!$F$2))"
            % (d, d, d, d, d), auto_date, cached=suggested)
    ws.set_row(r(17), 18)
    ws.merge_range(r(17), 1, r(17), ci(LAST_COL),
                   "  \U0001F4A1  The workbook is date-driven, not "
                   "year-driven: nothing is hard-coded to 2026, so you can "
                   "keep using it every year (and for every other gift-giving "
                   "occasion) just by changing the event date.",
                   S.f(**S.base(font_size=10, italic=True, font_color=th.ink,
                                bg_color=th.gold_soft, align="left",
                                valign="vcenter", indent=1)))
    ws.set_row(r(18), 8)

    # ------------------------------------------------------------------
    # 3. money, alerts & privacy
    # ------------------------------------------------------------------
    section(20, "  \U0001F4B0  MONEY, ALERTS & PRIVACY")
    setting(C.SU_CURRENCY, "Currency symbol",
            "Used by every dashboard and summary readout. To show it inside "
            "the tables too: select the money columns \u2192 Home \u2192 "
            "Number format \u2192 Currency.",
            m.settings["currency"], input_f)
    setting(C.SU_BUDGET, "Total budget target",
            "Your headline number for the whole event (the \U0001F4B0 Budget "
            "tab breaks it down by category).",
            m.settings["budget"] or 0, money_fmt)
    setting(C.SU_ALERT, "Warn me when spending reaches",
            "e.g. 85% \u2014 the dashboard and the Budget tab start shouting "
            "at this point.",
            m.settings["alert"], pct_fmt)
    setting(C.SU_DUESOON, "\u201CDue soon\u201D window (days)",
            "Deadlines inside this many days turn orange everywhere.",
            m.settings["duesoon"], num_fmt)
    setting(C.SU_SECRET, "\U0001F648 Secret Mode",
            "Yes = hiding spots on the \U0001F380 Wrapping and \U0001F9E6 "
            "Stockings tabs go invisible (white on white) so nobody can "
            "spoil the surprise over your shoulder.",
            m.settings["secret"], input_f)
    setting(C.SU_COUNT_ORDERED, "Count \u201COrdered\u201D as purchased?",
            "Yes = gifts that are ordered but not yet delivered still count "
            "as bought on the dashboard.",
            m.settings["count_ordered"], input_f)
    ws.set_row(r(27), 8)

    # ------------------------------------------------------------------
    # 4. editable lists
    # ------------------------------------------------------------------
    section(28, "  \U0001F4CB  YOUR LISTS  \u2014  edit any time, every "
                "dropdown in the workbook updates itself")
    ws.set_row(r(C.SU_LIST_HEADER), 30)
    for lk, col in sorted(C.LIST_COLS.items(), key=lambda kv: kv[1]):
        ws.write(r(C.SU_LIST_HEADER), ci(col), C.LIST_TITLES[lk],
                 S.header(th.primary))
    values = {
        "recipients": m.recipients,
        "relationships": C.RELATIONSHIPS,
        "gift_categories": C.GIFT_CATEGORIES,
        "stores": C.STORES,
        "shop_categories": C.SHOP_CATEGORIES,
        "hiding_spots": C.HIDING_SPOTS,
        "todo_categories": C.TODO_CATEGORIES,
        "stocking_owners": m.owners,
        "stocking_items": C.STOCKING_ITEMS,
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
    ws.set_row(r(50), 26)
    ws.merge_range(r(50), 1, r(51), ci(LAST_COL),
                   "  \u2022  Type over the example names with your own \u2014 "
                   "add as many as you like (20 rows each).\n"
                   "  \u2022  Don't leave a blank row in the middle of a list: "
                   "the dropdown stops at the first gap.\n"
                   "  \u2022  Deleting a name here does NOT delete your gifts "
                   "\u2014 it only removes it from the dropdown.",
                   S.note)
    ws.set_row(r(51), 26)
    ws.set_row(r(52), 8)

    # ------------------------------------------------------------------
    # 5. fixed lists
    # ------------------------------------------------------------------
    section(52, "  \U0001F512  FIXED LISTS  \u2014  the colours, counts and "
                "dashboard maths depend on this exact wording, so please "
                "leave them alone", S.section_accent)
    ws.set_row(r(C.SU_FIXED_HEADER), 26)
    fixed_titles = {"statuses": "Gift status", "order_statuses": "Order status",
                    "priorities": "Priority", "tick": "Tick box"}
    for fk, col in sorted(C.FIXED_COLS.items(), key=lambda kv: kv[1]):
        ws.write(r(C.SU_FIXED_HEADER), ci(col), fixed_titles[fk],
                 S.header(th.accent))
    ws.merge_range(r(C.SU_FIXED_HEADER), ci("F"), r(C.SU_FIXED_HEADER),
                   ci(LAST_COL),
                   "  Status order = your gift pipeline: idea \u2192 need to "
                   "buy \u2192 ordered \u2192 purchased \u2192 wrapped "
                   "\u2192 delivered.", S.note)
    fixed_values = {"statuses": C.STATUSES, "order_statuses": C.ORDER_STATUSES,
                    "priorities": C.PRIORITIES, "tick": [C.TICK, ""]}
    fixed_fmt = S.f(**S.base(font_size=10, font_color=th.ink, bg_color=th.alt,
                             align="left", valign="vcenter", border=1,
                             border_color=th.border, indent=1))
    for i in range(6):
        row = C.SU_FIXED_FIRST + i
        ws.set_row(r(row), 18)
        for fk, col in C.FIXED_COLS.items():
            vals = fixed_values[fk]
            value = vals[i] if i < len(vals) else ""
            ws.write(r(row), ci(col), value, fixed_fmt)
        for col in ("F", "G", "H", "I", "J"):
            ws.write_blank(r(row), ci(col), None, S.canvas)
    ws.set_row(r(60), 20)
    ws.merge_range(r(60), 1, r(61), ci(LAST_COL),
                   "  \u2713  The tick column holds the \u201Ccheckbox\u201D "
                   "used all over the workbook: pick \u2713 from a dropdown to "
                   "tick something off, pick the empty option to untick it.\n"
                   "  \u2713  Want different wording? Change it here AND in "
                   "the cells that already use it (Find & Replace does the "
                   "job in two clicks).",
                   S.note)
    ws.set_row(r(61), 20)
    ws.set_row(r(62), 8)

    # ------------------------------------------------------------------
    # 6. navigation
    # ------------------------------------------------------------------
    section(63, "  \U0001F9ED  WHERE TO GO NEXT")
    ws.set_row(r(64), 26)
    ws.set_row(r(65), 26)
    row = bk.nav_row(key, 64, first_col=1, span=1,
                     max_col=LAST_COL) + 2
    ws.set_row(r(row - 1), 10)
    ws.merge_range(r(row), 1, r(row), ci(LAST_COL),
                   "  %s v%s  \u2022  %s  \u2022  works in Excel 2016+ and "
                   "Google Sheets  \u2022  no macros, nothing to install"
                   % (C.PRODUCT, C.VERSION, C.TAGLINE), S.footer)

    # ------------------------------------------------------------------
    # validation
    # ------------------------------------------------------------------
    dv = ws.data_validation
    dv(r(C.SU_OCCASION), ci("C"), r(C.SU_OCCASION), ci("C"),
       {"validate": "list", "source": "=Occasions", "ignore_blank": True})
    dv(r(C.SU_CURRENCY), ci("C"), r(C.SU_CURRENCY), ci("C"),
       {"validate": "list", "source": C.CURRENCIES, "ignore_blank": True})
    for row in (C.SU_SECRET, C.SU_COUNT_ORDERED):
        dv(r(row), ci("C"), r(row), ci("C"),
           {"validate": "list", "source": C.YESNO, "ignore_blank": True,
            "show_error": True, "error_title": "Yes or No",
            "error_message": "Please choose Yes or No."})
    dv(r(C.SU_EVENT_DATE), ci("C"), r(C.SU_EVENT_DATE), ci("C"),
       {"validate": "date", "criteria": "between", "minimum": date(2000, 1, 1),
        "maximum": date(2100, 12, 31), "ignore_blank": True,
        "show_error": True, "error_title": "Enter a date",
        "error_message": "Please enter a real date, e.g. 25/12/2026.",
        "error_type": "warning"})
    dv(r(C.SU_BUDGET), ci("C"), r(C.SU_BUDGET), ci("C"),
       {"validate": "decimal", "criteria": ">=", "value": 0,
        "ignore_blank": True, "show_error": True, "error_title": "Budget",
        "error_message": "Enter a positive number (no currency symbol).",
        "error_type": "warning"})
    dv(r(C.SU_ALERT), ci("C"), r(C.SU_ALERT), ci("C"),
       {"validate": "decimal", "criteria": "between", "minimum": 0.05,
        "maximum": 1, "ignore_blank": True, "show_input": True,
        "input_title": "Alert threshold",
        "input_message": "0.85 = warn me when I have spent 85%.",
        "show_error": True, "error_title": "Percentage",
        "error_message": "Enter a percentage between 5% and 100% (0.05-1).",
        "error_type": "warning"})
    dv(r(C.SU_DUESOON), ci("C"), r(C.SU_DUESOON), ci("C"),
       {"validate": "whole", "criteria": "between", "minimum": 1,
        "maximum": 60, "ignore_blank": True, "show_error": True,
        "error_title": "Days", "error_message": "Enter a number of days "
        "between 1 and 60.", "error_type": "warning"})
    bk.stats["validations"] += 8

    # ------------------------------------------------------------------
    # page setup
    # ------------------------------------------------------------------
    bk.page(key, LAST_COL, row + 3, landscape=True, freeze=(r(5), 0),
            zoom=100,
            title_rows=(0, 4))
    return ws
```

### `christmas_tracker/sheets/gifts.py`  (309 lines)

```python
"""
\U0001F381 Gift Tracker - the heart of the product.

Inputs (white cells): recipient, idea, category, store, link, budget, actual
cost, status, wrapped, delivered, deadline, notes.
Automatic (tinted cells): row number, open-link button, difference, % of
budget, days left and the deadline alert.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "gifts"
LAST_COL = "T"

COLUMNS = [
    ("n", "#", "idx", None),
    ("recipient", "Recipient", "text", None),
    ("relationship", "Relationship", "center", None),
    ("idea", "Gift idea", "text", None),
    ("category", "Category", "center", None),
    ("store", "Store / where from", "text", None),
    ("link", "Link", "link", None),
    ("open", "\U0001F517", "calc_c", "gold"),
    ("budget", "Budget", "money", None),
    ("cost", "Actual cost", "money", None),
    ("diff", "Under / over", "calc_money", "primary_2"),
    ("pct", "% of budget", "calc_pct", "primary_2"),
    ("status", "Status", "center", None),
    ("wrapped", "\U0001F380 Wrapped", "tick", None),
    ("delivered", "\U0001F4E6 Given", "tick", None),
    ("deadline", "Buy-by date", "date", None),
    ("days", "Days left", "calc_num", "primary_2"),
    ("alert", "Deadline alert", "calc_c", "primary_2"),
    ("notes", "Notes", "text", None),
]


def build(bk):
    ws = bk.ws(KEY)
    S = bk.S
    th = bk.th
    m = bk.demo

    bk.widths(KEY)
    bk.paint(KEY, 0, 0, C.last_row(KEY) + 28, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F381  Gift Tracker",
        "  Every present in one place \u2014 type in the white cells, the "
        "tinted cells do the maths.",
        LAST_COL)

    # ------------------------------------------------------------------
    # quick-stat strip
    # ------------------------------------------------------------------
    chips = [
        ("B", "D", "\U0001F381 Planned: ", "gifts_planned", "primary"),
        ("E", "G", "\u2705 Purchased: ", "gifts_purchased", "ok"),
        ("H", "J", "\U0001F380 Wrapped: ", "gifts_wrapped", "accent"),
        ("K", "M", "\U0001F4E6 Given: ", "gifts_delivered", "info"),
        ("N", "P", "\U0001F6D2 Still to buy: ", "gifts_to_buy", "warn"),
        ("Q", "S", "\U0001F4B0 Spent on gifts: ", "gift_actual_spent",
         "primary_2"),
    ]
    for c1, c2, label, kpi_key, color in chips:
        fmt = S.pill(th.soft(color), getattr(th, color), size=10.5,
                     bold=True, align="left")
        ws.merge_range(r(C.ROW_STATS), ci(c1), r(C.ROW_STATS), ci(c2), "", fmt)
        if kpi_key == "gift_actual_spent":
            formula = '="%s"&Currency&TEXT(%s,"#,##0")' % (label,
                                                           bk.kpi(kpi_key))
            cached = "%s%s" % (label, m.money(m.agg[kpi_key]))
        else:
            formula = '="%s"&%s' % (label, bk.kpi(kpi_key))
            cached = "%s%d" % (label, m.agg[kpi_key])
        ws.write_formula(r(C.ROW_STATS), ci(c1), formula, fmt, cached)
        bk.stats["formulas"] += 1
    if LAST_COL != "T":
        ws.merge_range(r(C.ROW_STATS), ci("T"), r(C.ROW_STATS), ci(LAST_COL),
                       "", S.canvas)

    # ------------------------------------------------------------------
    # table
    # ------------------------------------------------------------------
    K.table_frame(bk, KEY, COLUMNS, height=20)
    for i in range(K.n_rows(KEY)):
        rownum = C.ROW_FIRST + i
        cached = _row_cached(m, i)
        values = _row_formulas(bk, rownum)
        if i < len(m.gifts):
            g = m.gifts[i]
            values.update({k: g[k] for k in
                           ("recipient", "relationship", "idea", "category",
                            "store", "link", "budget", "cost", "status",
                            "wrapped", "delivered", "deadline", "notes")})
        K.write_row(bk, KEY, COLUMNS, rownum, values, cached)

    ws.autofilter(r(C.ROW_HEADER), ci("B"), r(C.last_row(KEY)), ci(LAST_COL))
    ws.freeze_panes(r(C.ROW_FIRST), ci("C"))

    # ------------------------------------------------------------------
    # validation
    # ------------------------------------------------------------------
    K.list_dv(bk, KEY, "recipient", "recipients", title="Who is it for?",
              message="Pick a name from your recipient list (\u2699\uFE0F "
                      "Setup \u2192 Your lists). Type a new name to add one.",
              error="Use a name from your list, or type a new one.")
    K.list_dv(bk, KEY, "relationship", "relationships", title="Relationship")
    K.list_dv(bk, KEY, "category", "gift_categories", title="Gift category",
              message="Drives the \u201Cspend by category\u201D chart on the "
                      "dashboard.")
    K.list_dv(bk, KEY, "store", "stores", title="Store")
    K.fixed_dv(bk, KEY, "status", "Statuses", title="Gift status",
               message="The pipeline: \U0001F4A1 Idea \u2192 \U0001F6D2 Need "
                       "to Buy \u2192 \U0001F6CD\uFE0F Ordered \u2192 "
                       "\u2705 Purchased \u2192 \U0001F380 Wrapped \u2192 "
                       "\U0001F4E6 Delivered.",
               error="Pick one of the six statuses \u2014 the colours, the "
                     "counts and the dashboard all depend on them.")
    K.tick_dv(bk, KEY, ["wrapped", "delivered"])
    K.date_dv(bk, KEY, ["deadline"],
              message="The last day you can buy this and still be ready. "
                      "Turns red when it passes.")
    K.money_dv(bk, KEY, ["budget", "cost"])

    # ------------------------------------------------------------------
    # conditional formatting
    # ------------------------------------------------------------------
    n = C.ROW_FIRST
    st = bk.col(KEY, "status")
    K.status_cf(bk, KEY, "status", {
        C.ST_IDEA: (th.plum_soft, th.plum),
        C.ST_NEED: (th.warn_soft, th.warn),
        C.ST_ORDERED: (th.info_soft, th.info),
        C.ST_BOUGHT: (th.ok_soft, th.ok),
        C.ST_WRAPPED: (th.gold_soft, th.gold),
        C.ST_DELIVERED: (th.primary_soft, th.primary),
    })
    # money: under / over budget
    diff = bk.col(KEY, "diff")
    bk.cond(KEY, n, ci(diff), C.last_row(KEY), ci(diff), {
        "type": "formula", "criteria": '=$%s%d<0' % (diff, n),
        "format": S.cf(bg=th.bad_soft, fg=th.bad, bold=True)})
    bk.cond(KEY, n, ci(diff), C.last_row(KEY), ci(diff), {
        "type": "formula", "criteria": '=$%s%d>0' % (diff, n),
        "format": S.cf(bg=th.ok_soft, fg=th.ok, bold=True)})
    # % of budget: data bar + red over 100%
    pct = bk.col(KEY, "pct")
    K.databar(bk, KEY, "pct", color=th.gold)
    bk.cond(KEY, n, ci(pct), C.last_row(KEY), ci(pct), {
        "type": "formula", "criteria": '=$%s%d>1' % (pct, n),
        "format": S.cf(bg=th.bad_soft, fg=th.bad, bold=True)})
    # buy-by date
    K.deadline_cf(bk, KEY, "deadline", "delivered")
    # ticks
    K.tick_cf(bk, KEY, ["wrapped", "delivered"])
    # deadline alert column
    alert = bk.col(KEY, "alert")
    for text, (bg, fg) in (("\u2705 Done", (th.primary_soft, th.primary)),
                           ("\U0001F534 Overdue", (th.bad_soft, th.bad)),
                           ("\U0001F7E0 Due soon", (th.warn_soft, th.warn)),
                           ("\U0001F7E2 On track", (th.ok_soft, th.ok))):
        bk.cond(KEY, n, ci(alert), C.last_row(KEY), ci(alert), {
            "type": "formula",
            "criteria": '=$%s%d="%s"' % (alert, n, text),
            "format": S.cf(bg=bg, fg=fg, bold=True)})
    # nudge: a gift with no budget yet
    budget = bk.col(KEY, "budget")
    idea = bk.col(KEY, "idea")
    bk.cond(KEY, n, ci(budget), C.last_row(KEY), ci(budget), {
        "type": "formula",
        "criteria": '=AND($%s%d<>"",$%s%d="")' % (idea, n, budget, n),
        "format": S.cf(bg=th.gold_soft, fg=th.warn, bold=True)})
    # whole-row tint once a gift is wrapped / handed over (lowest priority)
    dlv = bk.col(KEY, "delivered")
    bk.cond(KEY, n, ci("B"), C.last_row(KEY), ci(LAST_COL), {
        "type": "formula",
        "criteria": '=OR($%s%d="%s",$%s%d="%s")' % (st, n, C.ST_DELIVERED,
                                                    dlv, n, C.TICK),
        "format": S.cf(bg=th.primary_soft, fg=th.primary)})

    # ------------------------------------------------------------------
    # totals
    # ------------------------------------------------------------------
    total_row = C.last_row(KEY) + 2
    j, k, l = (bk.col(KEY, "budget"), bk.col(KEY, "cost"), bk.col(KEY, "diff"))
    K.totals_row(
        bk, KEY, total_row,
        {
            "budget": ("=SUM($%s$%d:$%s$%d)" % (j, C.ROW_FIRST, j,
                                                C.last_row(KEY)),
                       "#,##0.00", m.agg["gift_budget_planned"]),
            "cost": ("=SUM($%s$%d:$%s$%d)" % (k, C.ROW_FIRST, k,
                                              C.last_row(KEY)),
                     "#,##0.00", m.agg["gift_actual_spent"]),
            "diff": ("=SUM($%s$%d:$%s$%d)" % (l, C.ROW_FIRST, l,
                                              C.last_row(KEY)),
                     "#,##0.00",
                     m.agg["gift_budget_planned"] - m.agg["gift_actual_spent"]),
            "pct": ("=IFERROR(SUM($%s$%d:$%s$%d)/SUM($%s$%d:$%s$%d),\"\")"
                    % (k, C.ROW_FIRST, k, C.last_row(KEY),
                       j, C.ROW_FIRST, j, C.last_row(KEY)),
                    "0%",
                    (m.agg["gift_actual_spent"] / m.agg["gift_budget_planned"])
                    if m.agg["gift_budget_planned"] else ""),
        },
        label="  TOTALS \u2014 gift list", label_span=("B", "I"),
        last_col=LAST_COL)
    summary_fmt = S.f(**S.base(font_size=11, bold=True, font_color=th.white,
                               bg_color=th.primary, align="center",
                               valign="vcenter", border=1,
                               border_color=th.primary))
    ws.merge_range(r(total_row), ci("N"), r(total_row), ci(LAST_COL), "",
                   summary_fmt)
    text = ('="\u2705 "&%s&" of "&%s&" gifts bought   \u2022   \U0001F380 "'
            '&%s&" wrapped   \u2022   \U0001F4E6 "&%s&" handed over"'
            % (bk.kpi("gifts_purchased"), bk.kpi("gifts_planned"),
               bk.kpi("gifts_wrapped"), bk.kpi("gifts_delivered")))
    cached = ("\u2705 %d of %d gifts bought   \u2022   \U0001F380 %d wrapped"
              "   \u2022   \U0001F4E6 %d handed over"
              % (m.agg["gifts_purchased"], m.agg["gifts_planned"],
                 m.agg["gifts_wrapped"], m.agg["gifts_delivered"]))
    ws.write_formula(r(total_row), ci("N"), text, summary_fmt, cached)
    bk.stats["formulas"] += 1

    # ------------------------------------------------------------------
    # tips + navigation
    # ------------------------------------------------------------------
    row = total_row + 2
    row = K.note_block(
        bk, KEY, row, "B", LAST_COL,
        ["\u2022  One row = one present. Need more than 60? Click row 67, "
         "drag down, and the formulas, dropdowns and colours come with it.",
         "\u2022  Status is the pipeline (\U0001F4A1 \u2192 \U0001F4E6). The "
         "\U0001F380 Wrapped and \U0001F4E6 Given tick boxes are quick "
         "shortcuts \u2014 ticking one, or moving the status on, both count.",
         "\u2022  Paste a product link in the Link column and the "
         "\U0001F517 column turns into a clickable button.",
         "\u2022  Everything on the \U0001F384 Dashboard, the \U0001F4B0 "
         "Budget tab and the \U0001F380 Wrapping tab reads from this sheet "
         "\u2014 you never re-type a gift twice."],
        title="  \U0001F4A1  How this tab works")
    row += 1
    ws.set_row(r(row), 26)
    row = bk.nav_row(KEY, row, first_col=1, span=2,
                     max_col=LAST_COL) + 2
    ws.merge_range(r(row), 1, r(row), ci(LAST_COL),
                   "  %s v%s  \u2022  %s" % (C.PRODUCT, C.VERSION, C.TAGLINE),
                   S.footer)

    bk.page(KEY, LAST_COL, row + 2, landscape=True,
            title_rows=(0, C.ROW_HEADER - 1))
    return ws


# ---------------------------------------------------------------------------
def _row_formulas(bk, n):
    def L(f):
        return bk.col(KEY, f)

    return {
        "n": '=IF($%s%d="","",ROW()-%d)' % (L("recipient"), n, C.ROW_FIRST - 1),
        "open": '=IF($%s%d="","",HYPERLINK($%s%d,"\U0001F517"))'
                % (L("link"), n, L("link"), n),
        "diff": '=IF($%s%d="","",IF($%s%d="",0,$%s%d)-$%s%d)'
                % (L("cost"), n, L("budget"), n, L("budget"), n, L("cost"), n),
        "pct": '=IF(OR($%s%d="",$%s%d="",$%s%d=0),"",$%s%d/$%s%d)'
               % (L("cost"), n, L("budget"), n, L("budget"), n,
                  L("cost"), n, L("budget"), n),
        "days": '=IF($%s%d="","",$%s%d-TODAY())'
                % (L("deadline"), n, L("deadline"), n),
        "alert": ('=IF($%s%d="","",IF(OR($%s%d="%s",$%s%d="%s"),"\u2705 Done",'
                 'IF($%s%d="","",IF($%s%d<TODAY(),"\U0001F534 Overdue",'
                 'IF($%s%d-TODAY()<=DueSoonDays,"\U0001F7E0 Due soon",'
                 '"\U0001F7E2 On track")))))'
                 % (L("recipient"), n, L("status"), n, C.ST_DELIVERED,
                    L("delivered"), n, C.TICK, L("deadline"), n,
                    L("deadline"), n, L("deadline"), n)),
    }


def _row_cached(m, i):
    """Cached values for the automatic columns (demo mode)."""
    from datetime import date as _date
    today = _date.today()
    if i >= len(m.gifts):
        return {"n": "", "open": "", "diff": "", "pct": "", "days": "",
                "alert": ""}
    g = m.gifts[i]
    budget, cost = g["budget"], g["cost"]
    dl = g["deadline"]
    diff = "" if cost is None else (budget or 0) - cost
    pct = "" if (cost is None or not budget) else cost / float(budget)
    days = "" if not dl else (dl - today).days
    if g["status"] == C.ST_DELIVERED or g["delivered"] == C.TICK:
        alert = "\u2705 Done"
    elif not dl:
        alert = ""
    elif dl < today:
        alert = "\U0001F534 Overdue"
    elif (dl - today).days <= m.settings["duesoon"]:
        alert = "\U0001F7E0 Due soon"
    else:
        alert = "\U0001F7E2 On track"
    return {"n": i + 1, "open": "\U0001F517" if g["link"] else "",
            "diff": diff, "pct": pct, "days": days, "alert": alert}
```

### `christmas_tracker/sheets/budget.py`  (420 lines)

```python
"""
\U0001F4B0 Budget - planned vs. actual for every corner of the season.

The "Auto from trackers" column is what makes this tab different from a plain
budget sheet: gift money, shopping-list money, stocking money and postage are
pulled in from the other tabs, so the total is always true.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "budget"
LAST_COL = "K"
FIRST = C.BUD_FIRST
LAST = C.BUD_FIRST + C.BUDGET_ROWS - 1
TOTAL = C.BUD_TOTAL

COLUMNS = [
    ("category", "Category", "text", None),
    ("planned", "Planned budget", "money", None),
    ("manual", "Extra spend (type here)", "money", None),
    ("auto", "Pulled in automatically", "calc_money", "primary_2"),
    ("actual", "Total spent", "calc_money", "primary_2"),
    ("remaining", "Remaining", "calc_money", "primary_2"),
    ("pct", "% used", "calc_pct", "primary_2"),
    ("status", "Status", "calc_c", "primary_2"),
    ("bar", "Progress", "calc", "primary_2"),
    ("notes", "Notes", "text", None),
]


def build(bk):
    ws = bk.ws(KEY)
    S = bk.S
    th = bk.th
    m = bk.demo

    bk.widths(KEY, {"A": 2.2, "B": 30, "C": 13, "D": 13, "E": 14, "F": 13,
                    "G": 13, "H": 10, "I": 17, "J": 24, "K": 26})
    bk.paint(KEY, 0, 0, 72, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F4B0  Christmas Budget",
        "  Plan it, watch it, and get told the moment you are about to "
        "overspend.",
        LAST_COL)

    # ------------------------------------------------------------------
    # overview cards
    # ------------------------------------------------------------------
    ws.set_row(r(7), 24)
    ws.merge_range(r(7), 1, r(7), ci(LAST_COL),
                   "  \U0001F4B0  MONEY OVERVIEW", S.section)

    planned_sum = "SUM($%s$%d:$%s$%d)" % (bk.col(KEY, "planned"), FIRST,
                                          bk.col(KEY, "planned"), LAST)
    actual_sum = "SUM($%s$%d:$%s$%d)" % (bk.col(KEY, "actual"), FIRST,
                                         bk.col(KEY, "actual"), LAST)
    cards = [
        ("B", "C", "TOTAL PLANNED BUDGET", "primary",
         '=Currency&TEXT(%s,"#,##0")' % planned_sum,
         m.money(m.agg["budget_planned"])),
        ("D", "E", "SPENT SO FAR", "accent",
         '=Currency&TEXT(%s,"#,##0")' % actual_sum,
         m.money(m.agg["budget_actual"])),
        ("F", "G", "STILL TO SPEND", "ok",
         '=Currency&TEXT(%s-%s,"#,##0")' % (planned_sum, actual_sum),
         m.money(m.agg["budget_remaining"])),
        ("H", "I", "BUDGET USED", "gold",
         "=IFERROR(%s/%s,0)" % (actual_sum, planned_sum),
         m.agg["budget_pct"]),
    ]
    ws.set_row(r(8), 18)
    ws.set_row(r(9), 36)
    for c1, c2, label, color, formula, cached in cards:
        ws.merge_range(r(8), ci(c1), r(8), ci(c2), "  " + label,
                       S.kpi_label(getattr(th, color)))
        if c1 == "H":
            fmt = S.kpi_value(getattr(th, color), num_format="0%", size=22)
        else:
            fmt = S.kpi_value(getattr(th, color), size=19, align="center")
        ws.merge_range(r(9), ci(c1), r(9), ci(c2), "", fmt)
        ws.write_formula(r(9), ci(c1), formula, fmt, cached)
        bk.stats["formulas"] += 1

    # alert card (two rows tall)
    ws.merge_range(r(8), ci("J"), r(8), ci(LAST_COL), "",
                   S.kpi_label(th.bad))
    ws.write(r(8), ci("J"), "  \u26A0\uFE0F  BUDGET ALERT",
             S.kpi_label(th.bad))
    alert_fmt = S.f(**S.base(font_size=10.5, bold=True, font_color=th.white,
                             bg_color=th.accent, align="left",
                             valign="vcenter", text_wrap=True, indent=1))
    ws.merge_range(r(9), ci("J"), r(9), ci(LAST_COL), "", alert_fmt)
    alert = ('=IF({p}<=0,"\U0001F4DD  Add your planned budgets below to '
             'switch on the alerts.",IF({a}>{p},"\U0001F534  Over budget by "'
             '&Currency&TEXT({a}-{p},"#,##0.00")&" \u2014 pause the '
             'non-essentials.",IF({a}/{p}>=AlertAt,"\u26A0\uFE0F  You\'ve '
             'spent "&TEXT({a}/{p},"0%")&" of your "&EventName&" budget '
             '\u2014 "&Currency&TEXT({p}-{a},"#,##0.00")&" left to '
             'play with.","\U0001F7E2  On track \u2014 "&TEXT({a}/{p},"0%")'
             '&" used, "&Currency&TEXT({p}-{a},"#,##0.00")&" still to '
             'spend.")))').format(p=planned_sum, a=actual_sum)
    ws.write_formula(r(9), ci("J"), alert, alert_fmt, m.agg["budget_alert"])
    bk.stats["formulas"] += 1

    # progress bar
    ws.set_row(r(10), 22)
    ws.merge_range(r(10), ci("B"), r(10), ci("C"), "  Budget progress",
                   S.bar_label)
    ws.merge_range(r(10), ci("D"), r(10), ci(LAST_COL), "", S.bar_text)
    bar = ('=IFERROR(REPT("\u2588",ROUND(MIN(1,{a}/{p})*34,0))&REPT("\u2591",'
           '34-ROUND(MIN(1,{a}/{p})*34,0))&"   "&TEXT({a}/{p},"0%"),"'
           '\u2591\u2591\u2591")').format(p=planned_sum, a=actual_sum)
    ws.write_formula(r(10), ci("D"), bar, S.bar_text,
                     m.agg["bar_budget"] + "   %d%%"
                     % round(m.agg["budget_pct"] * 100))
    bk.stats["formulas"] += 1

    # quick facts
    ws.set_row(r(11), 24)
    facts = [
        ("B", "C", "primary",
         '="\U0001F381 Gifts are "&TEXT(IFERROR($%s%d/%s,0),"0%%")&" of '
         'everything"' % (bk.col(KEY, "actual"), FIRST, actual_sum),
         "\U0001F381 Gifts are %d%% of everything"
         % round(100 * (m.agg["budget_rows"][0]["actual"] /
                        m.agg["budget_actual"])) if m.agg["budget_actual"]
         else "\U0001F381 Gifts are 0% of everything"),
        ("D", "F", "info",
         '="\U0001F4C8 Biggest spend: "&IFERROR(INDEX($%s$%d:$%s$%d,'
         'MATCH(MAX($%s$%d:$%s$%d),$%s$%d:$%s$%d,0)),"\u2014")'
         % ((bk.col(KEY, "category"), FIRST, bk.col(KEY, "category"), LAST)
            + (bk.col(KEY, "actual"), FIRST, bk.col(KEY, "actual"), LAST)
            + (bk.col(KEY, "actual"), FIRST, bk.col(KEY, "actual"), LAST)),
         "\U0001F4C8 Biggest spend: %s"
         % max(m.agg["budget_rows"], key=lambda x: x["actual"])["label"]
         if m.agg["budget_actual"] else "\U0001F4C8 Biggest spend: \u2014"),
        ("G", "H", "bad",
         '=COUNTIF($%s$%d:$%s$%d,"<0")&" over budget"'
         % (bk.col(KEY, "remaining"), FIRST, bk.col(KEY, "remaining"), LAST),
         "%d over budget" % len([x for x in m.agg["budget_rows"]
                                 if x["remaining"] < 0])),
        ("I", LAST_COL, "gold",
         '="\U0001F3AF Target "&Currency&TEXT(TotalBudget,"#,##0")&"   '
         '\u2022   planned "&Currency&TEXT(%s,"#,##0")' % planned_sum,
         "\U0001F3AF Target %s   \u2022   planned %s"
         % (m.money(m.settings["budget"]), m.money(m.agg["budget_planned"]))),
    ]
    for c1, c2, color, formula, cached in facts:
        fmt = S.pill(th.soft(color), getattr(th, color), size=10, bold=True,
                     align="left")
        ws.merge_range(r(11), ci(c1), r(11), ci(c2), "", fmt)
        ws.write_formula(r(11), ci(c1), formula, fmt, cached)
        bk.stats["formulas"] += 1
    ws.set_row(r(12), 8)

    # ------------------------------------------------------------------
    # the table
    # ------------------------------------------------------------------
    ws.set_row(r(13), 24)
    ws.merge_range(r(13), 1, r(13), ci(LAST_COL),
                   "  \U0001F4CB  BUDGET BY CATEGORY   (white = you type, "
                   "tinted = automatic)", S.section_soft)
    K.header_row(bk, KEY, COLUMNS, row=C.BUD_HEADER, height=34)
    for i in range(C.BUDGET_ROWS):
        ws.set_row(r(FIRST + i), 22)
    for i in range(C.BUDGET_ROWS):
        rownum = FIRST + i
        a = K.alt(rownum)
        for field, label, kind, color in COLUMNS:
            ws.write_blank(r(rownum), ci(bk.col(KEY, field)), None,
                           S.cell(kind, a) if not kind.startswith("calc")
                           else S.cell(kind))
    for i, (label, bkey, recipe, demo_planned) in enumerate(C.BUDGET_CATEGORIES):
        rownum = FIRST + i
        cached_row = m.agg["budget_rows"][i]
        bar_fmt = S.f(**S.base(font_name=th.mono_font, font_size=10,
                               bold=True, font_color=th.primary_2,
                               bg_color=th.card, align="left",
                               valign="vcenter", border=1,
                               border_color=th.border, locked=True))
        values = {
            "category": label,
            "planned": demo_planned if bk.mode == "demo" else None,
            "manual": m.budget_manual.get(bkey) if bk.mode == "demo" else None,
            "auto": _auto_formula(bk, recipe, rownum),
            "actual": "=$%s%d+$%s%d" % (bk.col(KEY, "manual"), rownum,
                                        bk.col(KEY, "auto"), rownum),
            "remaining": "=$%s%d-$%s%d" % (bk.col(KEY, "planned"), rownum,
                                           bk.col(KEY, "actual"), rownum),
            "pct": '=IF($%s%d=0,"",$%s%d/$%s%d)'
                   % (bk.col(KEY, "planned"), rownum, bk.col(KEY, "actual"),
                      rownum, bk.col(KEY, "planned"), rownum),
            "status": _status_formula(bk, rownum),
            "bar": _bar_formula(bk, rownum),
            "notes": "",
        }
        cached = {
            "auto": cached_row["auto"], "actual": cached_row["actual"],
            "remaining": cached_row["remaining"], "pct": cached_row["pct"],
            "status": _status_cached(cached_row, m),
            "bar": _bar_cached(cached_row),
        }
        K.write_row(bk, KEY, COLUMNS, rownum, values, cached)
        # the progress column wants its own mono format
        ws.write_formula(r(rownum), ci(bk.col(KEY, "bar")),
                         values["bar"], bar_fmt, cached["bar"])

    # totals row
    tcols = {}
    for field in ("planned", "manual", "auto", "actual", "remaining"):
        col = bk.col(KEY, field)
        cached_key = {"planned": "budget_planned", "actual": "budget_actual",
                      "remaining": "budget_remaining"}.get(field)
        cached_val = m.agg[cached_key] if cached_key else sum(
            x[field] for x in m.agg["budget_rows"])
        tcols[field] = ("=SUM($%s$%d:$%s$%d)" % (col, FIRST, col, LAST),
                        "#,##0.00", cached_val)
    tcols["pct"] = ("=IFERROR(%s/%s,\"\")" % (actual_sum, planned_sum), "0%",
                    m.agg["budget_pct"])
    K.totals_row(bk, KEY, TOTAL, tcols, label="  TOTAL",
                 label_span=("B", "B"), last_col=LAST_COL)
    ws.merge_range(r(TOTAL), ci("I"), r(TOTAL), ci(LAST_COL), "",
                   S.f(**S.base(font_size=11, bold=True, font_color=th.white,
                                bg_color=th.primary, align="center",
                                valign="vcenter", border=1,
                                border_color=th.primary)))
    ws.write_formula(r(TOTAL), ci("I"),
                     '=IF(%s=0,"\u2014","Budget "&TEXT(%s/%s,"0%%")&" used  '
                     '\u2022  "&Currency&TEXT(%s-%s,"#,##0")&" left")'
                     % (planned_sum, actual_sum, planned_sum, planned_sum,
                        actual_sum),
                     S.f(**S.base(font_size=11, bold=True, font_color=th.white,
                                  bg_color=th.primary, align="center",
                                  valign="vcenter", border=1,
                                  border_color=th.primary)),
                     "Budget %d%% used  \u2022  %s left"
                     % (round(100 * m.agg["budget_pct"]),
                        m.money(m.agg["budget_remaining"]))
                     if m.agg["budget_planned"] else "\u2014")
    bk.stats["formulas"] += 1

    # ------------------------------------------------------------------
    # conditional formatting
    # ------------------------------------------------------------------
    n = FIRST
    K.databar(bk, KEY, "pct", color=th.gold, first=FIRST, last=LAST)
    pct_col = ci(bk.col(KEY, "pct"))
    bk.cond(KEY, n, pct_col, LAST, pct_col, {
        "type": "formula", "criteria": '=$%s%d>1' % (bk.col(KEY, "pct"), n),
        "format": S.cf(bg=th.bad_soft, fg=th.bad, bold=True)})
    bk.cond(KEY, n, pct_col, LAST, pct_col, {
        "type": "formula",
        "criteria": '=AND($%s%d<=1,$%s%d>=AlertAt)'
                   % (bk.col(KEY, "pct"), n, bk.col(KEY, "pct"), n),
        "format": S.cf(bg=th.warn_soft, fg=th.warn, bold=True)})
    rem_col = ci(bk.col(KEY, "remaining"))
    bk.cond(KEY, n, rem_col, LAST, rem_col, {
        "type": "formula",
        "criteria": '=$%s%d<0' % (bk.col(KEY, "remaining"), n),
        "format": S.cf(bg=th.bad_soft, fg=th.bad, bold=True)})
    status_col = ci(bk.col(KEY, "status"))
    for text, (bg, fg) in (("\U0001F534 Over budget", (th.bad_soft, th.bad)),
                           ("\U0001F7E0 Nearly there", (th.warn_soft, th.warn)),
                           ("\U0001F7E2 On track", (th.ok_soft, th.ok)),
                           ("\u26AA Not started", (th.alt, th.muted)),
                           ("\u26AA Set a plan", (th.alt, th.muted))):
        bk.cond(KEY, n, status_col, LAST, status_col, {
            "type": "formula",
            "criteria": '=$%s%d="%s"' % (bk.col(KEY, "status"), n, text),
            "format": S.cf(bg=bg, fg=fg, bold=True)})
    bk.cond(KEY, n, ci("B"), LAST, ci(LAST_COL), {
        "type": "formula",
        "criteria": '=$%s%d<0' % (bk.col(KEY, "remaining"), n),
        "format": S.cf(bg=th.bad_soft, fg=th.bad)})

    # ------------------------------------------------------------------
    # charts
    # ------------------------------------------------------------------
    ws.set_row(r(26), 8)
    ws.set_row(r(27), 24)
    ws.merge_range(r(27), 1, r(27), ci(LAST_COL),
                   "  \U0001F4CA  BUDGET vs ACTUAL", S.section)
    cats = "=%s!$%s$%d:$%s$%d" % (bk.q(KEY), bk.col(KEY, "category"), FIRST,
                                  bk.col(KEY, "category"), LAST)
    planned_vals = "=%s!$%s$%d:$%s$%d" % (bk.q(KEY), bk.col(KEY, "planned"),
                                          FIRST, bk.col(KEY, "planned"), LAST)
    actual_vals = "=%s!$%s$%d:$%s$%d" % (bk.q(KEY), bk.col(KEY, "actual"),
                                         FIRST, bk.col(KEY, "actual"), LAST)

    ch1 = bk.chart("column")
    ch1.add_series({"name": "Planned", "categories": cats,
                    "values": planned_vals,
                    "fill": {"color": th.primary_soft and "#BFD3C2"},
                    "border": {"color": th.primary},
                    "gap": 60})
    ch1.add_series({"name": "Actually spent", "categories": cats,
                    "values": actual_vals, "fill": {"color": th.accent},
                    "border": {"color": th.accent}})
    ch1.set_title({"name": "Planned vs actually spent, by category",
                   "name_font": {"size": 12, "bold": True, "color": th.primary,
                                 "name": th.body_font}})
    ch1.set_y_axis({"num_format": "#,##0", "num_font": {"size": 9},
                    "major_gridlines": {"visible": True,
                                        "line": {"color": th.border}}})
    ch1.set_x_axis({"num_font": {"size": 9}, "label_position": "low"})
    ch1.set_size({"width": 560, "height": 290})
    ch1.set_legend({"position": "bottom", "font": {"size": 9}})
    ch1.set_style(2)
    ws.insert_chart(r(28), ci("B"), ch1, {"x_offset": 6, "y_offset": 6})

    ch2 = bk.chart("doughnut")
    ch2.add_series({
        "name": "Spent", "categories": cats, "values": actual_vals,
        "points": [{"fill": {"color": c}} for c in th.series],
        "data_labels": {"percentage": True, "font": {"size": 9,
                                                     "color": th.white}},
    })
    ch2.set_title({"name": "Where the money actually goes",
                   "name_font": {"size": 12, "bold": True,
                                 "color": th.primary, "name": th.body_font}})
    ch2.set_hole_size(58)
    ch2.set_size({"width": 540, "height": 290})
    ch2.set_legend({"position": "right", "font": {"size": 9}})
    ws.insert_chart(r(28), ci("G"), ch2, {"x_offset": 6, "y_offset": 6})

    # ------------------------------------------------------------------
    # notes + nav
    # ------------------------------------------------------------------
    row = 44
    row = K.note_block(
        bk, KEY, row, "B", LAST_COL,
        ["\u2022  \u201CPlanned budget\u201D and \u201CExtra spend\u201D are "
         "the only two columns you type in.",
         "\u2022  \u201CPulled in automatically\u201D adds up the money from "
         "the \U0001F381 Gift Tracker, \U0001F6CD\uFE0F Shopping List, "
         "\U0001F9E6 Stockings and \U0001F48C Card postage \u2014 so this "
         "total is always honest.",
         "\u2022  Set the warning level in \u2699\uFE0F Setup (\u201CWarn me "
         "when spending reaches\u201D). At that point the alert card, the "
         "status column and the dashboard all change colour.",
         "\u2022  Add your own category? Insert a row above the TOTAL row and "
         "everything (including the charts) picks it up."],
        title="  \U0001F4A1  How this tab works")
    row += 1
    ws.set_row(r(row), 26)
    row = bk.nav_row(KEY, row, first_col=1, span=2,
                     max_col=LAST_COL) + 2
    ws.merge_range(r(row), 1, r(row), ci(LAST_COL),
                   "  %s v%s  \u2022  %s" % (C.PRODUCT, C.VERSION, C.TAGLINE),
                   S.footer)
    bk.page(KEY, LAST_COL, row + 2, landscape=True, title_rows=(0, 13))
    return ws


# ---------------------------------------------------------------------------
def _auto_formula(bk, recipe, rownum):
    terms = []
    for term in recipe:
        if term == "gifts":
            terms.append("SUM(%s)" % bk.rng("gifts", "cost"))
        elif term == "stock":
            if bk.has("stockings"):
                terms.append("SUM(%s)" % bk.rng("stockings", "spent"))
        elif term == "cards":
            if bk.has("cards"):
                terms.append("SUM(%s)" % bk.rng("cards", "postage"))
        elif term.startswith("shop:"):
            cat = term.split(":", 1)[1]
            if bk.has("shopping"):
                terms.append('SUMIFS(%s,%s,"%s")'
                             % (bk.rng("shopping", "spent"),
                                bk.rng("shopping", "category"), cat))
    if not terms:
        return "=0"
    return "=" + "+".join(terms)


def _status_formula(bk, n):
    planned = "$%s%d" % (bk.col(KEY, "planned"), n)
    actual = "$%s%d" % (bk.col(KEY, "actual"), n)
    return ('=IF({p}=0,"\u26AA Set a plan",IF({a}>{p},"\U0001F534 Over '
            'budget",IF({a}/{p}>=AlertAt,"\U0001F7E0 Nearly there",'
            'IF({a}=0,"\u26AA Not started","\U0001F7E2 On track"))))'
            ).format(p=planned, a=actual)


def _status_cached(row, m):
    if not row["planned"]:
        return "\u26AA Set a plan"
    if row["actual"] > row["planned"]:
        return "\U0001F534 Over budget"
    if row["pct"] and row["pct"] >= m.settings["alert"]:
        return "\U0001F7E0 Nearly there"
    if not row["actual"]:
        return "\u26AA Not started"
    return "\U0001F7E2 On track"


def _bar_formula(bk, n):
    planned = "$%s%d" % (bk.col(KEY, "planned"), n)
    actual = "$%s%d" % (bk.col(KEY, "actual"), n)
    blocks = 12
    pct = "MIN(1,IFERROR({a}/{p},0))".format(a=actual, p=planned)
    filled = "ROUND(%s*%d,0)" % (pct, blocks)
    return ('=IF({p}=0,"\u2014",REPT("\u2588",{f})&REPT("\u2591",{b}-{f})'
            '&" "&TEXT(IFERROR({a}/{p},0),"0%"))').format(p=planned, a=actual,
                                                          f=filled, b=blocks)


def _bar_cached(row, blocks=12):
    if not row["planned"]:
        return "\u2014"
    pct = min(1.0, max(0.0, (row["actual"] or 0) / float(row["planned"])))
    filled = int(round(pct * blocks))
    return ("\u2588" * filled + "\u2591" * (blocks - filled) + " %d%%"
            % round(pct * 100))
```

### `christmas_tracker/sheets/shopping.py`  (177 lines)

```python
"""
\U0001F6CD\uFE0F Shopping List - everything that is not a present:
wrapping, cards, baking, decorations, party bits.  Money spent here flows
straight into the \U0001F4B0 Budget tab.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "shopping"
LAST_COL = "M"

COLUMNS = [
    ("n", "#", "idx", None),
    ("item", "Item", "text", None),
    ("category", "Category", "center", None),
    ("store", "Store", "text", None),
    ("qty", "Qty", "qty", None),
    ("unit", "Unit cost", "money", None),
    ("total", "Total (auto)", "calc_money", "primary_2"),
    ("bought", "Bought", "tick", None),
    ("cost", "Actual cost", "money", None),
    ("spent", "In budget (auto)", "calc_money", "primary_2"),
    ("link", "Link", "link", None),
    ("notes", "Notes", "text", None),
]


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY)
    bk.paint(KEY, 0, 0, C.last_row(KEY) + 28, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F6CD\uFE0F  Shopping List",
        "  Wrapping, cards, baking, decorations, party bits \u2014 the stuff "
        "that quietly eats the budget.",
        LAST_COL)

    chips = [
        ("B", "D", "\U0001F4DD Items: ", "shop_total", "primary", "0"),
        ("E", "F", "\u2705 Bought: ", "shop_bought", "ok", "0"),
        ("G", "H", "\U0001F4B0 Spent: ", "shop_spent", "accent", "$"),
        ("I", "J", "\U0001F4B5 Estimated: ", None, "gold", "$"),
        ("K", LAST_COL, "\U0001F4CA Progress: ", None, "info", "%"),
    ]
    for c1, c2, label, kpi_key, color, kind in chips:
        fmt = S.pill(th.soft(color), getattr(th, color), size=10.5, bold=True,
                     align="left")
        ws.merge_range(r(C.ROW_STATS), ci(c1), r(C.ROW_STATS), ci(c2), "", fmt)
        est = "SUM(%s)" % bk.rng(KEY, "total")
        if kpi_key:
            if kind == "$":
                formula = '="%s"&Currency&TEXT(%s,"#,##0")' % (label,
                                                               bk.kpi(kpi_key))
                cached = "%s%s" % (label, m.money(m.agg[kpi_key]))
            else:
                formula = '="%s"&%s' % (label, bk.kpi(kpi_key))
                cached = "%s%d" % (label, m.agg[kpi_key])
        elif kind == "$":
            formula = '="%s"&Currency&TEXT(%s,"#,##0")' % (label, est)
            cached = "%s%s" % (label, m.money(sum(
                (x["qty"] or 0) * (x["unit"] or 0) for x in m.shopping)))
        else:
            formula = '="%s"&TEXT(IFERROR(%s/%s,0),"0%%")' % (
                label, bk.kpi("shop_bought"), bk.kpi("shop_total"))
            pct = (m.agg["shop_bought"] / float(m.agg["shop_total"])
                   if m.agg["shop_total"] else 0)
            cached = "%s%d%%" % (label, round(pct * 100))
        ws.write_formula(r(C.ROW_STATS), ci(c1), formula, fmt, cached)
        bk.stats["formulas"] += 1

    K.table_frame(bk, KEY, COLUMNS, height=20)
    for i in range(K.n_rows(KEY)):
        rownum = C.ROW_FIRST + i
        values = _formulas(bk, rownum)
        if i < len(m.shopping):
            row = m.shopping[i]
            values.update({"item": row["item"], "category": row["category"],
                           "store": row["store"], "qty": row["qty"],
                           "unit": row["unit"], "bought": row["bought"],
                           "cost": row["cost"], "notes": row["notes"]})
        K.write_row(bk, KEY, COLUMNS, rownum, values, _cached(m, i))

    ws.autofilter(r(C.ROW_HEADER), ci("B"), r(C.last_row(KEY)), ci(LAST_COL))
    ws.freeze_panes(r(C.ROW_FIRST), ci("C"))

    K.list_dv(bk, KEY, "category", "shop_categories", title="Category",
              message="This is what links the item to a line on the "
                      "\U0001F4B0 Budget tab.")
    K.list_dv(bk, KEY, "store", "stores", title="Store")
    K.tick_dv(bk, KEY, ["bought"],
               message="Tick \u2713 when it is in the basket \u2014 the cost "
                       "then flows into your budget automatically.")
    K.money_dv(bk, KEY, ["unit", "cost"])
    K.whole_dv(bk, KEY, ["qty"])

    n = C.ROW_FIRST
    K.tick_cf(bk, KEY, ["bought"])
    bought = bk.col(KEY, "bought")
    bk.cond(KEY, n, ci("B"), C.last_row(KEY), ci(LAST_COL), {
        "type": "formula", "criteria": '=$%s%d="%s"' % (bought, n, C.TICK),
        "format": S.cf(bg=th.ok_soft, fg=th.muted, strike=True)})

    total_row = C.last_row(KEY) + 2
    est_total = sum((x["qty"] or 0) * (x["unit"] or 0) for x in m.shopping)
    K.totals_row(
        bk, KEY, total_row,
        {"total": ("=SUM($%s$%d:$%s$%d)" % (bk.col(KEY, "total"),
                                            C.ROW_FIRST, bk.col(KEY, "total"),
                                            C.last_row(KEY)),
                   "#,##0.00", est_total),
         "bought": ('=COUNTIF($%s$%d:$%s$%d,"%s")&" / "&COUNTA($%s$%d:$%s$%d)'
                    % (bk.col(KEY, "bought"), C.ROW_FIRST,
                       bk.col(KEY, "bought"), C.last_row(KEY), C.TICK,
                       bk.col(KEY, "item"), C.ROW_FIRST, bk.col(KEY, "item"),
                       C.last_row(KEY)), "@",
                    "%d / %d" % (m.agg["shop_bought"], m.agg["shop_total"])),
         "cost": ("=SUM($%s$%d:$%s$%d)" % (bk.col(KEY, "cost"), C.ROW_FIRST,
                                           bk.col(KEY, "cost"),
                                           C.last_row(KEY)),
                  "#,##0.00", sum(x["cost"] or 0 for x in m.shopping)),
         "spent": ("=SUM($%s$%d:$%s$%d)" % (bk.col(KEY, "spent"),
                                            C.ROW_FIRST, bk.col(KEY, "spent"),
                                            C.last_row(KEY)),
                   "#,##0.00", m.agg["shop_spent"])},
        label="  TOTALS", label_span=("B", "F"), last_col=LAST_COL)

    row = total_row + 2
    row = K.note_block(
        bk, KEY, row, "B", LAST_COL,
        ["\u2022  Tick \u2713 Bought and the \u201CIn budget\u201D column "
         "picks up the actual cost (or the estimate if you have not typed one "
         "yet).",
         "\u2022  The category decides which \U0001F4B0 Budget line the money "
         "lands on \u2014 Wrapping, Cards, Postage, Food, Baking, "
         "Decorations, Party Supplies, Travel, Charity, Other.",
         "\u2022  Starter rows are pre-filled so you can see how it works "
         "\u2014 overwrite or delete them freely."],
        title="  \U0001F4A1  How this tab works")
    row += 1
    ws.set_row(r(row), 26)
    row = bk.nav_row(KEY, row, first_col=1, span=2,
                     max_col=LAST_COL) + 2
    ws.merge_range(r(row), 1, r(row), ci(LAST_COL),
                   "  %s v%s  \u2022  %s" % (C.PRODUCT, C.VERSION, C.TAGLINE),
                   S.footer)
    bk.page(KEY, LAST_COL, row + 2, landscape=True,
            title_rows=(0, C.ROW_HEADER - 1))
    return ws


def _formulas(bk, n):
    def L(f):
        return bk.col(KEY, f)

    return {
        "n": '=IF($%s%d="","",ROW()-%d)' % (L("item"), n, C.ROW_FIRST - 1),
        "total": '=IF($%s%d="","",IF($%s%d="",1,$%s%d)*$%s%d)'
                 % (L("unit"), n, L("qty"), n, L("qty"), n, L("unit"), n),
        "spent": '=IF($%s%d="%s",IF($%s%d<>"",$%s%d,IF($%s%d<>"",$%s%d,0)),0)'
                 % (L("bought"), n, C.TICK, L("cost"), n, L("cost"), n,
                    L("total"), n, L("total"), n),
    }


def _cached(m, i):
    if i >= len(m.shopping):
        return {"n": "", "total": "", "spent": 0}
    row = m.shopping[i]
    qty, unit = row["qty"] or 0, row["unit"] or 0
    total = "" if row["unit"] is None else qty * unit
    return {"n": i + 1, "total": total, "spent": m._shop_spent(row)}
```

### `christmas_tracker/sheets/wishlist.py`  (177 lines)

```python
"""
\U0001F4A1 Wish List / Gift Ideas - the "parking bay" for ideas.

Capture hints all year, rank them, and see at a glance which ones have already
made it onto the \U0001F381 Gift Tracker (that check is a COUNTIFS against
the gift list, so it stays true as you work).
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "wishlist"
LAST_COL = "K"

COLUMNS = [
    ("n", "#", "idx", None),
    ("person", "Person", "text", None),
    ("idea", "Gift idea", "text", None),
    ("category", "Category", "center", None),
    ("link", "Link", "link", None),
    ("open", "\U0001F517", "calc_c", "gold"),
    ("price", "Price seen", "money", None),
    ("priority", "Priority", "center", None),
    ("in_tracker", "On the gift list?", "calc_c", "primary_2"),
    ("notes", "Notes", "text", None),
]


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY)
    bk.paint(KEY, 0, 0, C.last_row(KEY) + 28, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F4A1  Wish List & Gift Ideas",
        "  Hints all year round, ranked by how much they really want them.",
        LAST_COL)

    chips = [
        ("B", "C", "\U0001F4A1 Ideas saved: ", "wish_total", "plum"),
        ("D", "E", "\u2B50 Must-haves: ", "wish_must", "gold"),
        ("F", "G", "\U0001F4B0 Value of ideas: ", "wish_value", "accent"),
        ("H", "I", "\u2714 Already on the gift list: ", None, "ok"),
        ("J", LAST_COL, "\u2795 Still to move across: ", None, "warn"),
    ]
    for c1, c2, label, kpi_key, color in chips:
        fmt = S.pill(th.soft(color), getattr(th, color), size=10.5, bold=True,
                     align="left")
        ws.merge_range(r(C.ROW_STATS), ci(c1), r(C.ROW_STATS), ci(c2), "", fmt)
        if kpi_key == "wish_value":
            formula = '="%s"&Currency&TEXT(%s,"#,##0")' % (label,
                                                           bk.kpi(kpi_key))
            cached = "%s%s" % (label, m.money(m.agg[kpi_key]))
        elif kpi_key:
            formula = '="%s"&%s' % (label, bk.kpi(kpi_key))
            cached = "%s%d" % (label, m.agg[kpi_key])
        elif label.startswith("\u2714"):
            formula = '="%s"&COUNTIF(%s,"\u2714 In gift tracker")' % (
                label, bk.rng(KEY, "in_tracker"))
            cached = "%s%d" % (label, _in_tracker_count(m))
        else:
            formula = '="%s"&COUNTIF(%s,"\u2795 Not yet")' % (
                label, bk.rng(KEY, "in_tracker"))
            cached = "%s%d" % (label, m.agg["wish_total"] - _in_tracker_count(m))
        ws.write_formula(r(C.ROW_STATS), ci(c1), formula, fmt, cached)
        bk.stats["formulas"] += 1

    K.table_frame(bk, KEY, COLUMNS, height=20)
    for i in range(K.n_rows(KEY)):
        rownum = C.ROW_FIRST + i
        values = _formulas(bk, rownum)
        if i < len(m.wishlist):
            w = m.wishlist[i]
            values.update({k: w[k] for k in ("person", "idea", "category",
                                             "link", "price", "priority",
                                             "notes")})
        K.write_row(bk, KEY, COLUMNS, rownum, values, _cached(m, i))

    ws.autofilter(r(C.ROW_HEADER), ci("B"), r(C.last_row(KEY)), ci(LAST_COL))
    ws.freeze_panes(r(C.ROW_FIRST), ci("C"))

    K.list_dv(bk, KEY, "person", "recipients", title="Who is it for?")
    K.list_dv(bk, KEY, "category", "gift_categories", title="Category")
    K.fixed_dv(bk, KEY, "priority", "Priorities", title="Priority",
               message="\u2B50\u2B50\u2B50 Must have \u2014 they have "
                       "dropped a serious hint.\n\u2B50\u2B50 Maybe.\n"
                       "\u2B50 Low priority / filler.")
    K.money_dv(bk, KEY, ["price"], label="a price")

    n = C.ROW_FIRST
    K.status_cf(bk, KEY, "priority", {
        C.PR_MUST: (th.gold_soft, th.gold),
        C.PR_MAYBE: (th.info_soft, th.info),
        C.PR_LOW: (th.alt, th.muted),
    })
    K.status_cf(bk, KEY, "in_tracker", {
        "\u2714 In gift tracker": (th.ok_soft, th.ok),
        "\u2795 Not yet": (th.warn_soft, th.warn),
    })
    K.databar(bk, KEY, "price", color=th.plum)

    total_row = C.last_row(KEY) + 2
    K.totals_row(
        bk, KEY, total_row,
        {"price": ("=SUM($%s$%d:$%s$%d)" % (bk.col(KEY, "price"),
                                            C.ROW_FIRST, bk.col(KEY, "price"),
                                            C.last_row(KEY)),
                   "#,##0.00", m.agg["wish_value"]),
         "in_tracker": ('=COUNTIF($%s$%d:$%s$%d,"\u2795 Not yet")&" idea(s) '
                        'still to move across"'
                        % (bk.col(KEY, "in_tracker"), C.ROW_FIRST,
                           bk.col(KEY, "in_tracker"), C.last_row(KEY)),
                        "@", "%d idea(s) still to move across"
                        % (m.agg["wish_total"] - _in_tracker_count(m)))},
        label="  TOTALS", label_span=("B", "G"), last_col=LAST_COL)

    row = total_row + 2
    row = K.note_block(
        bk, KEY, row, "B", LAST_COL,
        ["\u2022  Use this tab all year: when somebody mentions something in "
         "March, write it here and forget about it until November.",
         "\u2022  \u201COn the gift list?\u201D checks the \U0001F381 Gift "
         "Tracker for the same person + same idea, so you always know what "
         "still needs moving across.",
         "\u2022  To move an idea over: copy the person and idea cells, paste "
         "them into a new row on the \U0001F381 Gift Tracker, then set the "
         "status to \U0001F6D2 Need to Buy.",
         "\u2022  Sort by Priority (Data \u2192 Sort, or use the filter "
         "arrows) to build the shopping list in the right order."],
        title="  \U0001F4A1  How this tab works")
    row += 1
    ws.set_row(r(row), 26)
    row = bk.nav_row(KEY, row, first_col=1, span=2,
                     max_col=LAST_COL) + 2
    ws.merge_range(r(row), 1, r(row), ci(LAST_COL),
                   "  %s v%s  \u2022  %s" % (C.PRODUCT, C.VERSION, C.TAGLINE),
                   S.footer)
    bk.page(KEY, LAST_COL, row + 2, landscape=True,
            title_rows=(0, C.ROW_HEADER - 1))
    return ws


def _formulas(bk, n):
    def L(f):
        return bk.col(KEY, f)

    return {
        "n": '=IF($%s%d="","",ROW()-%d)' % (L("idea"), n, C.ROW_FIRST - 1),
        "open": '=IF($%s%d="","",HYPERLINK($%s%d,"\U0001F517"))'
                % (L("link"), n, L("link"), n),
        "in_tracker": ('=IF(OR($%s%d="",$%s%d=""),"",IF(COUNTIFS(%s,$%s%d,%s,'
                       '$%s%d)>0,"\u2714 In gift tracker","\u2795 Not yet"))'
                       % (L("person"), n, L("idea"), n,
                          bk.rng("gifts", "recipient"), L("person"), n,
                          bk.rng("gifts", "idea"), L("idea"), n)),
    }


def _in_tracker_count(m):
    pairs = set((g["recipient"], g["idea"]) for g in m.gifts)
    return len([w for w in m.wishlist
                if (w["person"], w["idea"]) in pairs])


def _cached(m, i):
    if i >= len(m.wishlist):
        return {"n": "", "open": "", "in_tracker": ""}
    w = m.wishlist[i]
    pairs = set((g["recipient"], g["idea"]) for g in m.gifts)
    return {"n": i + 1, "open": "\U0001F517" if w["link"] else "",
            "in_tracker": ("\u2714 In gift tracker"
                           if (w["person"], w["idea"]) in pairs
                           else "\u2795 Not yet")}
```

### `christmas_tracker/sheets/orders.py`  (230 lines)

```python
"""
\U0001F4E6 Order Tracker - the "will it arrive in time?" tab.

Every online order in one place with its expected date, tracking link, the
number of days you have left and an alert that turns red the moment a parcel
is late.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "orders"
LAST_COL = "R"

COLUMNS = [
    ("n", "#", "idx", None),
    ("recipient", "Recipient", "text", None),
    ("item", "Gift / item", "text", None),
    ("store", "Store", "text", None),
    ("order_no", "Order number", "text", None),
    ("order_date", "Ordered on", "date", None),
    ("expected", "Expected", "date", None),
    ("actual", "Arrived", "date", None),
    ("link", "Tracking link", "link", None),
    ("open", "\U0001F517", "calc_c", "gold"),
    ("status", "Order status", "center", None),
    ("days", "Days left", "calc_num", "primary_2"),
    ("alert", "Alert", "calc_c", "primary_2"),
    ("return_by", "Return by", "date", None),
    ("returned", "Returned", "tick", None),
    ("cost", "Cost", "money", None),
    ("notes", "Notes", "text", None),
]


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY)
    bk.paint(KEY, 0, 0, C.last_row(KEY) + 28, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F4E6  Online Order Tracker",
        "  Parcels, couriers and cut-off dates \u2014 nothing arrives on "
        "the 27th on your watch.",
        LAST_COL)

    chips = [
        ("B", "D", "\U0001F4E6 Orders: ", "orders_total", "primary", "n"),
        ("E", "G", "\U0001F69A Still coming: ", "orders_outstanding", "info",
         "n"),
        ("H", "I", "\u2705 Arrived: ", None, "ok", "n"),
        ("J", "L", "\U0001F534 Running late: ", "orders_late", "bad", "n"),
        ("M", "N", "\U0001F4B5 Value: ", "orders_value", "accent", "$"),
        ("O", LAST_COL, "\U0001F504 Returns open: ", None, "warn", "n"),
    ]
    for c1, c2, label, kpi_key, color, kind in chips:
        fmt = S.pill(th.soft(color), getattr(th, color), size=10.5, bold=True,
                     align="left")
        ws.merge_range(r(C.ROW_STATS), ci(c1), r(C.ROW_STATS), ci(c2), "", fmt)
        if kpi_key is None and label.startswith("\u2705"):
            formula = '="%s"&COUNTIF(%s,"%s")' % (label,
                                                  bk.rng(KEY, "status"),
                                                  C.OS_DELIVERED)
            cached = "%s%d" % (label, len([x for x in m.orders
                                           if x["status"] == C.OS_DELIVERED]))
        elif kpi_key is None:
            formula = '="%s"&COUNTIF(%s,"%s")' % (label,
                                                  bk.rng(KEY, "returned"),
                                                  C.TICK)
            cached = "%s%d" % (label, len([x for x in m.orders
                                           if x["returned"] == C.TICK]))
        elif kind == "$":
            formula = '="%s"&Currency&TEXT(%s,"#,##0")' % (label,
                                                           bk.kpi(kpi_key))
            cached = "%s%s" % (label, m.money(m.agg[kpi_key]))
        else:
            formula = '="%s"&%s' % (label, bk.kpi(kpi_key))
            cached = "%s%d" % (label, m.agg[kpi_key])
        ws.write_formula(r(C.ROW_STATS), ci(c1), formula, fmt, cached)
        bk.stats["formulas"] += 1

    K.table_frame(bk, KEY, COLUMNS, height=20)
    for i in range(K.n_rows(KEY)):
        rownum = C.ROW_FIRST + i
        values = _formulas(bk, rownum)
        if i < len(m.orders):
            o = m.orders[i]
            values.update({k: o[k] for k in
                           ("recipient", "item", "store", "order_no",
                            "order_date", "expected", "actual", "link",
                            "status", "return_by", "returned", "cost",
                            "notes")})
        K.write_row(bk, KEY, COLUMNS, rownum, values, _cached(m, i))

    ws.autofilter(r(C.ROW_HEADER), ci("B"), r(C.last_row(KEY)), ci(LAST_COL))
    ws.freeze_panes(r(C.ROW_FIRST), ci("C"))

    K.list_dv(bk, KEY, "recipient", "recipients", title="Who is it for?")
    K.list_dv(bk, KEY, "store", "stores", title="Store")
    K.fixed_dv(bk, KEY, "status", "OrderStatuses", title="Order status",
               message="\U0001F6D2 Ordered \u2192 \U0001F4E6 Shipped \u2192 "
                       "\U0001F69A Out for delivery \u2192 \u2705 Delivered "
                       "(or \u21A9\uFE0F Return requested / \u274C "
                       "Cancelled).")
    K.tick_dv(bk, KEY, ["returned"],
               message="Tick \u2713 once the parcel is back with the seller.")
    K.date_dv(bk, KEY, ["order_date", "expected", "actual", "return_by"])
    K.money_dv(bk, KEY, ["cost"])

    n = C.ROW_FIRST
    st = bk.col(KEY, "status")
    K.status_cf(bk, KEY, "status", {
        C.OS_ORDERED: (th.plum_soft, th.plum),
        C.OS_SHIPPED: (th.info_soft, th.info),
        C.OS_OUT: (th.warn_soft, th.warn),
        C.OS_DELIVERED: (th.ok_soft, th.ok),
        C.OS_RETURN: (th.gold_soft, th.gold),
        C.OS_CANCEL: (th.bad_soft, th.bad),
    })
    K.tick_cf(bk, KEY, ["returned"])
    K.deadline_cf(bk, KEY, "expected")
    alert = bk.col(KEY, "alert")
    for text, (bg, fg) in (("\u2705 Arrived", (th.ok_soft, th.ok)),
                           ("\U0001F534 Late", (th.bad_soft, th.bad)),
                           ("\U0001F7E1 Arriving soon", (th.warn_soft,
                                                         th.warn)),
                           ("\U0001F535 On the way", (th.info_soft, th.info)),
                           ("\u274C Cancelled", (th.alt, th.muted)),
                           ("\u21A9\uFE0F Returning", (th.gold_soft, th.gold))):
        bk.cond(KEY, n, ci(alert), C.last_row(KEY), ci(alert), {
            "type": "formula",
            "criteria": '=$%s%d="%s"' % (alert, n, text),
            "format": S.cf(bg=bg, fg=fg, bold=True)})
    bk.cond(KEY, n, ci("B"), C.last_row(KEY), ci(LAST_COL), {
        "type": "formula",
        "criteria": '=$%s%d="%s"' % (st, n, C.OS_DELIVERED),
        "format": S.cf(bg=th.ok_soft, fg=th.primary)})

    total_row = C.last_row(KEY) + 2
    K.totals_row(
        bk, KEY, total_row,
        {"cost": ("=SUM($%s$%d:$%s$%d)" % (bk.col(KEY, "cost"), C.ROW_FIRST,
                                           bk.col(KEY, "cost"),
                                           C.last_row(KEY)),
                  "#,##0.00", m.agg["orders_value"]),
         "expected": ('=COUNTIF($%s$%d:$%s$%d,"%s")&" of "&COUNTA($%s$%d:'
                      '$%s$%d)&" arrived"'
                      % (st, C.ROW_FIRST, st, C.last_row(KEY), C.OS_DELIVERED,
                         bk.col(KEY, "item"), C.ROW_FIRST, bk.col(KEY, "item"),
                         C.last_row(KEY)), "@",
                      "%d of %d arrived" % (
                          len([x for x in m.orders
                               if x["status"] == C.OS_DELIVERED]),
                          m.agg["orders_total"]))},
        label="  TOTALS", label_span=("B", "F"), last_col=LAST_COL)

    row = total_row + 2
    row = K.note_block(
        bk, KEY, row, "B", LAST_COL,
        ["\u2022  \u201CDays left\u201D counts down to the expected delivery "
         "date; the alert turns \U0001F534 red the day after it was due.",
         "\u2022  Paste the courier link in \u201CTracking link\u201D and the "
         "\U0001F517 button opens it in one click.",
         "\u2022  Keep an eye on the \u201CReturn by\u201D column \u2014 "
         "January returns sneak up fast.",
         "\u2022  Order money is NOT added to the budget twice: the gift "
         "itself is counted on the \U0001F381 Gift Tracker. This tab is for "
         "logistics."],
        title="  \U0001F4A1  How this tab works")
    row += 1
    ws.set_row(r(row), 26)
    row = bk.nav_row(KEY, row, first_col=1, span=2,
                     max_col=LAST_COL) + 2
    ws.merge_range(r(row), 1, r(row), ci(LAST_COL),
                   "  %s v%s  \u2022  %s" % (C.PRODUCT, C.VERSION, C.TAGLINE),
                   S.footer)
    bk.page(KEY, LAST_COL, row + 2, landscape=True,
            title_rows=(0, C.ROW_HEADER - 1))
    return ws


def _formulas(bk, n):
    def L(f):
        return bk.col(KEY, f)

    return {
        "n": '=IF($%s%d="","",ROW()-%d)' % (L("item"), n, C.ROW_FIRST - 1),
        "open": '=IF($%s%d="","",HYPERLINK($%s%d,"\U0001F517"))'
                % (L("link"), n, L("link"), n),
        "days": '=IF($%s%d="","",$%s%d-TODAY())' % (L("expected"), n,
                                                    L("expected"), n),
        "alert": ('=IF($%s%d="","",IF($%s%d="%s","\u2705 Arrived",'
                 'IF($%s%d="%s","\u274C Cancelled",IF($%s%d="%s",'
                 '"\u21A9\uFE0F Returning",IF($%s%d="","\u2014",'
                 'IF($%s%d<TODAY(),"\U0001F534 Late",'
                 'IF($%s%d-TODAY()<=DueSoonDays,"\U0001F7E1 Arriving soon",'
                 '"\U0001F535 On the way")))))))'
                 % (L("item"), n, L("status"), n, C.OS_DELIVERED,
                    L("status"), n, C.OS_CANCEL, L("status"), n, C.OS_RETURN,
                    L("expected"), n, L("expected"), n, L("expected"), n)),
    }


def _cached(m, i):
    from datetime import date
    today = date.today()
    if i >= len(m.orders):
        return {"n": "", "open": "", "days": "", "alert": ""}
    o = m.orders[i]
    ex = o["expected"]
    if o["status"] == C.OS_DELIVERED:
        alert = "\u2705 Arrived"
    elif o["status"] == C.OS_CANCEL:
        alert = "\u274C Cancelled"
    elif o["status"] == C.OS_RETURN:
        alert = "\u21A9\uFE0F Returning"
    elif not ex:
        alert = "\u2014"
    elif ex < today:
        alert = "\U0001F534 Late"
    elif (ex - today).days <= m.settings["duesoon"]:
        alert = "\U0001F7E1 Arriving soon"
    else:
        alert = "\U0001F535 On the way"
    return {"n": i + 1, "open": "\U0001F517" if o["link"] else "",
            "days": (ex - today).days if ex else "", "alert": alert}
```

### `christmas_tracker/sheets/wrapping.py`  (230 lines)

```python
"""
\U0001F380 Wrapping & Hiding - the fun one (and the one parents love).

Everything on the left is pulled live from the \U0001F381 Gift Tracker, so
there is nothing to re-type: this tab is where you record *where the wrapped
present is hiding* and whether the gift tag is on it.

Secret Mode (\u2699\uFE0F Setup) makes the hiding-spot text invisible so a
nosy helper cannot spoil the surprise by glancing at your screen.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "wrapping"
LAST_COL = "J"

COLUMNS = [
    ("n", "#", "calc_c", "primary_2"),
    ("recipient", "For", "calc", "primary_2"),
    ("idea", "Gift", "calc", "primary_2"),
    ("bought", "Bought?", "calc_c", "primary_2"),
    ("wrapped", "Wrapped?", "calc_c", "primary_2"),
    ("hiding", "\U0001F648 Hiding spot", "text", None),
    ("tag", "\U0001F3F7\uFE0F Tag on", "tick", None),
    ("delivered", "Handed over?", "calc_c", "primary_2"),
    ("notes", "Notes", "text", None),
]


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo
    gt = bk.q("gifts")

    bk.widths(KEY)
    bk.paint(KEY, 0, 0, C.last_row(KEY) + 28, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F380  Wrapping & Hiding",
        "  Mirrors your gift list automatically \u2014 add the hiding spot "
        "and the gift tag here.",
        LAST_COL)

    chips = [
        ("B", "C", "\U0001F380 To wrap: ", "wrap_to_do", "warn", "n"),
        ("D", "E", "\u2705 Wrapped: ", "gifts_wrapped", "ok", "n"),
        ("F", "G", "\U0001F3F7\uFE0F Tags done: ", None, "info", "n"),
        ("H", "H", "\U0001F648 Spots noted: ", None, "plum", "n"),
        ("I", "I", "\U0001F4E6 Handed over: ", "gifts_delivered", "primary",
         "n"),
        ("J", LAST_COL, None, None, "accent", "secret"),
    ]
    for c1, c2, label, kpi_key, color, kind in chips:
        fmt = S.pill(th.soft(color), getattr(th, color), size=10.5, bold=True,
                     align="left")
        if c1 == c2:
            ws.write(r(C.ROW_STATS), ci(c1), "", fmt)
        else:
            ws.merge_range(r(C.ROW_STATS), ci(c1), r(C.ROW_STATS), ci(c2),
                           "", fmt)
        if kind == "secret":
            formula = ('=IF(SecretMode="Yes","\U0001F648 Secret Mode ON '
                       '\u2014 hiding spots are invisible","\U0001F440 '
                       'Secret Mode off \u2014 switch it on in '
                       '\u2699\uFE0F Setup")')
            cached = ("\U0001F648 Secret Mode ON \u2014 hiding spots are "
                      "invisible" if m.settings["secret"] == "Yes"
                      else "\U0001F440 Secret Mode off \u2014 switch it on "
                           "in \u2699\uFE0F Setup")
        elif kpi_key:
            formula = '="%s"&%s' % (label, bk.kpi(kpi_key))
            cached = "%s%d" % (label, m.agg[kpi_key])
        elif label.startswith("\U0001F3F7"):
            formula = '="%s"&COUNTIF(%s,"%s")' % (label, bk.rng(KEY, "tag"),
                                                  C.TICK)
            cached = "%s%d" % (label, _tags_done(m))
        else:
            formula = '="%s"&COUNTA(%s)' % (label, bk.rng(KEY, "hiding"))
            cached = "%s%d" % (label, _spots_noted(m))
        ws.write_formula(r(C.ROW_STATS), ci(c1), formula, fmt, cached)
        bk.stats["formulas"] += 1

    K.table_frame(bk, KEY, COLUMNS, height=20)
    for i in range(K.n_rows(KEY)):
        rownum = C.ROW_FIRST + i
        values = _formulas(bk, rownum, gt)
        if i < len(m.gifts) and m.gifts[i]["idea"]:
            values["hiding"] = m.gifts[i].get("hiding", "")
            values["tag"] = m.gifts[i].get("tag", "")
            values["notes"] = ""
        K.write_row(bk, KEY, COLUMNS, rownum, values, _cached(m, i, gt))

    ws.freeze_panes(r(C.ROW_FIRST), ci("C"))

    bk.ws(KEY).data_validation(
        r(C.ROW_FIRST), ci(bk.col(KEY, "hiding")), r(C.last_row(KEY)),
        ci(bk.col(KEY, "hiding")),
        {"validate": "list", "source": "=" + bk.listname("hiding_spots"),
         "ignore_blank": True, "show_input": True,
         "input_title": "Where is it hiding?",
         "input_message": "Pick a spot or type your own.",
         "show_error": True, "error_type": "warning",
         "error_title": "Free text is fine",
         "error_message": "Keep it short \u2014 or add your own spots to the "
                          "list in \u2699\uFE0F Setup."})
    K.tick_dv(bk, KEY, ["tag"], message="Tick \u2713 when the name tag is on "
                                        "the parcel.")

    n = C.ROW_FIRST
    for field, mapping in (
            ("bought", {"\u2705 Bought": (th.ok_soft, th.ok),
                        "\U0001F69A On the way": (th.info_soft, th.info),
                        "\u23F3 Not yet": (th.alt, th.muted)}),
            ("wrapped", {"\U0001F380 Wrapped": (th.gold_soft, th.gold),
                         "\u2014 to wrap": (th.warn_soft, th.warn)}),
            ("delivered", {"\U0001F4E6 Given": (th.primary_soft, th.primary),
                           "\u2014": (th.alt, th.muted)})):
        K.status_cf(bk, KEY, field, mapping)
    K.tick_cf(bk, KEY, ["tag"])
    K.secret_cf(bk, KEY, "hiding")
    wrapped = bk.col(KEY, "wrapped")
    bk.cond(KEY, n, ci("B"), C.last_row(KEY), ci(LAST_COL), {
        "type": "formula",
        "criteria": '=$%s%d="\U0001F380 Wrapped"' % (wrapped, n),
        "format": S.cf(bg=th.gold_soft, fg=th.ink)})

    total_row = C.last_row(KEY) + 2
    K.totals_row(
        bk, KEY, total_row,
        {"tag": ('=COUNTIF(%s,"%s")' % (bk.rng(KEY, "tag"), C.TICK), "0",
                 _tags_done(m)),
         "hiding": ("=COUNTA(%s)" % bk.rng(KEY, "hiding"), "0",
                    _spots_noted(m))},
        label="  WRAPPING ROOM", label_span=("B", "E"), last_col=LAST_COL)
    fmt = S.f(**S.base(font_size=11, bold=True, font_color=th.white,
                       bg_color=th.primary, align="center", valign="vcenter",
                       border=1, border_color=th.primary))
    ws.merge_range(r(total_row), ci("I"), r(total_row), ci(LAST_COL), "", fmt)
    ws.write_formula(r(total_row), ci("I"),
                     '="\U0001F380 "&%s&" wrapped   \u2022   \u23F3 "&%s'
                     '&" still to wrap"'
                     % (bk.kpi("gifts_wrapped"), bk.kpi("wrap_to_do")), fmt,
                     "\U0001F380 %d wrapped   \u2022   \u23F3 %d still to wrap"
                     % (m.agg["gifts_wrapped"], m.agg["wrap_to_do"]))
    bk.stats["formulas"] += 1

    row = total_row + 2
    row = K.note_block(
        bk, KEY, row, "B", LAST_COL,
        ["\u2022  The left-hand columns are live mirrors of the "
         "\U0001F381 Gift Tracker \u2014 tick \U0001F380 Wrapped or move the "
         "status on over there and this tab updates itself.",
         "\u2022  Type the hiding spot here (\U0001F648). Parents: this is "
         "the column that stops you wrapping the same present twice or "
         "losing it in the loft.",
         "\u2022  Someone about to glance at your screen? \u2699\uFE0F Setup "
         "\u2192 Secret Mode = Yes and every hiding spot turns white-on-white "
         "instantly. For total privacy, right-click column G \u2192 Hide.",
         "\u2022  \U0001F3F7\uFE0F Tag on = the name label is attached, so "
         "nothing ends up under the wrong tree."],
        title="  \U0001F4A1  How this tab works")
    row += 1
    ws.set_row(r(row), 26)
    row = bk.nav_row(KEY, row, first_col=1, span=2,
                     max_col=LAST_COL) + 2
    ws.merge_range(r(row), 1, r(row), ci(LAST_COL),
                   "  %s v%s  \u2022  %s" % (C.PRODUCT, C.VERSION, C.TAGLINE),
                   S.footer)
    bk.page(KEY, LAST_COL, row + 2, landscape=True,
            title_rows=(0, C.ROW_HEADER - 1))
    return ws


def _formulas(bk, n, gt):
    def G(f):
        return "%s!$%s$%d" % (gt, bk.col("gifts", f), n)

    bought_states = 'OR(%s="%s",%s="%s",%s="%s")' % (
        G("status"), C.ST_BOUGHT, G("status"), C.ST_WRAPPED, G("status"),
        C.ST_DELIVERED)
    wrapped_states = 'OR(%s="%s",%s="%s",%s="%s")' % (
        G("wrapped"), C.TICK, G("status"), C.ST_WRAPPED, G("status"),
        C.ST_DELIVERED)
    return {
        "n": '=IF(%s="","",ROW()-%d)' % (G("recipient"), C.ROW_FIRST - 1),
        "recipient": '=IF(%s="","",%s)' % (G("recipient"), G("recipient")),
        "idea": '=IF(%s="","",%s)' % (G("recipient"), G("idea")),
        "bought": '=IF(%s="","",IF(%s,"\u2705 Bought",IF(%s="%s",'
                  '"\U0001F69A On the way","\u23F3 Not yet")))'
                  % (G("recipient"), bought_states, G("status"), C.ST_ORDERED),
        "wrapped": '=IF(%s="","",IF(%s,"\U0001F380 Wrapped","\u2014 to wrap"))'
                   % (G("recipient"), wrapped_states),
        "delivered": '=IF(%s="","",IF(OR(%s="%s",%s="%s"),"\U0001F4E6 Given",'
                     '"\u2014"))' % (G("recipient"), G("delivered"), C.TICK,
                                     G("status"), C.ST_DELIVERED),
    }


def _cached(m, i, gt):
    if i >= len(m.gifts) or not m.gifts[i]["idea"]:
        return {"n": "", "recipient": "", "idea": "", "bought": "",
                "wrapped": "", "delivered": ""}
    g = m.gifts[i]
    st = g["status"]
    if st in (C.ST_BOUGHT, C.ST_WRAPPED, C.ST_DELIVERED):
        bought = "\u2705 Bought"
    elif st == C.ST_ORDERED:
        bought = "\U0001F69A On the way"
    else:
        bought = "\u23F3 Not yet"
    wrapped = ("\U0001F380 Wrapped"
               if (g["wrapped"] == C.TICK or st in (C.ST_WRAPPED,
                                                    C.ST_DELIVERED))
               else "\u2014 to wrap")
    delivered = ("\U0001F4E6 Given"
                 if (g["delivered"] == C.TICK or st == C.ST_DELIVERED)
                 else "\u2014")
    return {"n": i + 1, "recipient": g["recipient"], "idea": g["idea"],
            "bought": bought, "wrapped": wrapped, "delivered": delivered}


def _tags_done(m):
    return len([g for g in m.gifts if g.get("tag") == C.TICK])


def _spots_noted(m):
    return len([g for g in m.gifts if g.get("hiding")])
```

### `christmas_tracker/sheets/cards.py`  (177 lines)

```python
"""
\U0001F48C Card Tracker - who got a card, who still needs one, and what the
post office took from you.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "cards"
LAST_COL = "M"

COLUMNS = [
    ("n", "#", "idx", None),
    ("name", "Name / family", "text", None),
    ("relationship", "Relationship", "center", None),
    ("address", "Address", "text", None),
    ("bought", "Card bought", "tick", None),
    ("written", "Written", "tick", None),
    ("sent", "Sent", "tick", None),
    ("date_sent", "Date sent", "date", None),
    ("received", "Reply back", "tick", None),
    ("postage", "Postage", "money", None),
    ("status", "Status (auto)", "calc_c", "primary_2"),
    ("notes", "Notes", "text", None),
]


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY)
    bk.paint(KEY, 0, 0, C.last_row(KEY) + 28, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F48C  Christmas Card Tracker",
        "  Bought \u2192 written \u2192 posted \u2192 replied. Nobody gets "
        "forgotten, and postage lands in your budget.",
        LAST_COL)

    chips = [
        ("B", "C", "\U0001F4DD Cards: ", "cards_total", "primary"),
        ("D", "E", "\U0001F6D2 Bought: ", None, "info"),
        ("F", "G", "\u270D\uFE0F Written: ", "cards_written", "warn"),
        ("H", "I", "\U0001F4EE Sent: ", "cards_sent", "ok"),
        ("J", "K", "\U0001F4EC Replies: ", None, "plum"),
        ("L", LAST_COL, "\U0001F4B0 Postage: ", None, "accent"),
    ]
    for c1, c2, label, kpi_key, color in chips:
        fmt = S.pill(th.soft(color), getattr(th, color), size=10.5, bold=True,
                     align="left")
        ws.merge_range(r(C.ROW_STATS), ci(c1), r(C.ROW_STATS), ci(c2), "", fmt)
        if kpi_key:
            formula = '="%s"&%s' % (label, bk.kpi(kpi_key))
            cached = "%s%d" % (label, m.agg[kpi_key])
        elif label.startswith("\U0001F6D2"):
            formula = '="%s"&COUNTIF(%s,"%s")' % (label, bk.rng(KEY, "bought"),
                                                  C.TICK)
            cached = "%s%d" % (label, len([x for x in m.cards
                                           if x["bought"] == C.TICK]))
        elif label.startswith("\U0001F4EC"):
            formula = '="%s"&COUNTIF(%s,"%s")' % (label,
                                                  bk.rng(KEY, "received"),
                                                  C.TICK)
            cached = "%s%d" % (label, m.agg["cards_received"])
        else:
            formula = '="%s"&Currency&TEXT(SUM(%s),"#,##0.00")' % (
                label, bk.rng(KEY, "postage"))
            cached = "%s%s" % (label, m.money(m.agg["cards_postage"], 2))
        ws.write_formula(r(C.ROW_STATS), ci(c1), formula, fmt, cached)
        bk.stats["formulas"] += 1

    K.table_frame(bk, KEY, COLUMNS, height=20)
    for i in range(K.n_rows(KEY)):
        rownum = C.ROW_FIRST + i
        values = {"n": '=IF($%s%d="","",ROW()-%d)'
                         % (bk.col(KEY, "name"), rownum, C.ROW_FIRST - 1),
                  "status": _status_formula(bk, rownum)}
        if i < len(m.cards):
            c = m.cards[i]
            values.update({k: c[k] for k in
                           ("name", "relationship", "address", "bought",
                            "written", "sent", "date_sent", "received",
                            "postage", "notes")})
        K.write_row(bk, KEY, COLUMNS, rownum, values, _cached(m, i))

    ws.autofilter(r(C.ROW_HEADER), ci("B"), r(C.last_row(KEY)), ci(LAST_COL))
    ws.freeze_panes(r(C.ROW_FIRST), ci("C"))

    K.list_dv(bk, KEY, "relationship", "relationships", title="Relationship")
    K.tick_dv(bk, KEY, ["bought", "written", "sent", "received"])
    K.date_dv(bk, KEY, ["date_sent"])
    K.money_dv(bk, KEY, ["postage"], label="postage")

    n = C.ROW_FIRST
    K.tick_cf(bk, KEY, ["bought", "written", "sent", "received"])
    K.status_cf(bk, KEY, "status", {
        "\U0001F4EE Sent": (th.ok_soft, th.ok),
        "\u270D\uFE0F Written": (th.info_soft, th.info),
        "\U0001F6D2 Bought": (th.warn_soft, th.warn),
        "\U0001F4DD To do": (th.alt, th.muted),
    })
    sent = bk.col(KEY, "sent")
    bk.cond(KEY, n, ci("B"), C.last_row(KEY), ci(LAST_COL), {
        "type": "formula", "criteria": '=$%s%d="%s"' % (sent, n, C.TICK),
        "format": S.cf(bg=th.ok_soft, fg=th.primary)})

    total_row = C.last_row(KEY) + 2
    fmt = S.f(**S.base(font_size=11, bold=True, font_color=th.white,
                       bg_color=th.primary, align="center", valign="vcenter",
                       border=1, border_color=th.primary))
    K.totals_row(
        bk, KEY, total_row,
        {"postage": ("=SUM($%s$%d:$%s$%d)" % (bk.col(KEY, "postage"),
                                              C.ROW_FIRST,
                                              bk.col(KEY, "postage"),
                                              C.last_row(KEY)),
                     "#,##0.00", m.agg["cards_postage"])},
        label="  TOTALS", label_span=("B", "I"), last_col=LAST_COL)
    ws.merge_range(r(total_row), ci("J"), r(total_row), ci(LAST_COL), "", fmt)
    ws.write_formula(
        r(total_row), ci("J"),
        '="\U0001F4EE Cards sent: "&%s&" / "&%s'
        % (bk.kpi("cards_sent"), bk.kpi("cards_total")), fmt,
        "\U0001F4EE Cards sent: %d / %d" % (m.agg["cards_sent"],
                                            m.agg["cards_total"]))
    bk.stats["formulas"] += 1

    row = total_row + 2
    row = K.note_block(
        bk, KEY, row, "B", LAST_COL,
        ["\u2022  Tick the boxes as you go: bought \u2192 written \u2192 sent "
         "\u2192 reply. The status column reads the ticks for you.",
         "\u2022  Postage you type here is added to the \u201CCards & "
         "postage\u201D line on the \U0001F4B0 Budget tab automatically.",
         "\u2022  Posting cut-off: aim for %d days before the big day "
         "(the dashboard shows the exact date)."
         % 14],
        title="  \U0001F4A1  How this tab works")
    row += 1
    ws.set_row(r(row), 26)
    row = bk.nav_row(KEY, row, first_col=1, span=2,
                     max_col=LAST_COL) + 2
    ws.merge_range(r(row), 1, r(row), ci(LAST_COL),
                   "  %s v%s  \u2022  %s" % (C.PRODUCT, C.VERSION, C.TAGLINE),
                   S.footer)
    bk.page(KEY, LAST_COL, row + 2, landscape=True,
            title_rows=(0, C.ROW_HEADER - 1))
    return ws


def _status_formula(bk, n):
    def L(f):
        return "$%s%d" % (bk.col(KEY, f), n)

    return ('=IF({name}="","",IF({sent}="{t}","\U0001F4EE Sent",'
            'IF({written}="{t}","\u270D\uFE0F Written",'
            'IF({bought}="{t}","\U0001F6D2 Bought","\U0001F4DD To do"))))'
            ).format(name=L("name"), sent=L("sent"), written=L("written"),
                     bought=L("bought"), t=C.TICK)


def _cached(m, i):
    if i >= len(m.cards):
        return {"n": "", "status": ""}
    c = m.cards[i]
    if c["sent"] == C.TICK:
        status = "\U0001F4EE Sent"
    elif c["written"] == C.TICK:
        status = "\u270D\uFE0F Written"
    elif c["bought"] == C.TICK:
        status = "\U0001F6D2 Bought"
    else:
        status = "\U0001F4DD To do"
    return {"n": i + 1, "status": status}
```

### `christmas_tracker/sheets/stockings.py`  (279 lines)

```python
"""
\U0001F9E6 Stocking Stuffers - one row per little present, plus a
per-stocking summary so every stocking ends up equally spoiled (and equally
affordable).
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "stockings"
LAST_COL = "L"
OWNER_ROWS = 8

COLUMNS = [
    ("n", "#", "idx", None),
    ("owner", "Stocking", "text", None),
    ("item", "Item", "text", None),
    ("category", "Type", "center", None),
    ("budget", "Budget", "money", None),
    ("cost", "Cost", "money", None),
    ("bought", "Bought", "tick", None),
    ("wrapped", "Wrapped", "tick", None),
    ("spent", "In budget (auto)", "calc_money", "primary_2"),
    ("hiding", "\U0001F648 Hiding spot", "text", None),
    ("notes", "Notes", "text", None),
]


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY)
    bk.paint(KEY, 0, 0, C.last_row(KEY) + 34, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \U0001F9E6  Stocking Stuffer Tracker",
        "  The little things add up fastest \u2014 budget them per stocking "
        "and stop the January regret.",
        LAST_COL)

    chips = [
        ("B", "C", "\U0001F9E6 Stockings: ", None, "primary"),
        ("D", "E", "\U0001F381 Items: ", "stock_items", "info"),
        ("F", "G", "\u2705 Bought: ", "stock_bought", "ok"),
        ("H", "I", "\U0001F4B0 Budget: ", "stock_budget", "gold"),
        ("J", "K", "\U0001F4B8 Spent: ", "stock_spent", "accent"),
        ("L", LAST_COL, "\U0001F3AF Left: ", None, "primary_2"),
    ]
    for c1, c2, label, kpi_key, color in chips:
        fmt = S.pill(th.soft(color), getattr(th, color), size=10.5, bold=True,
                     align="left")
        if c1 == c2:
            ws.write(r(C.ROW_STATS), ci(c1), "", fmt)
        else:
            ws.merge_range(r(C.ROW_STATS), ci(c1), r(C.ROW_STATS), ci(c2),
                           "", fmt)
        if kpi_key in ("stock_budget", "stock_spent"):
            formula = '="%s"&Currency&TEXT(%s,"#,##0")' % (label,
                                                           bk.kpi(kpi_key))
            cached = "%s%s" % (label, m.money(m.agg[kpi_key]))
        elif kpi_key:
            formula = '="%s"&%s' % (label, bk.kpi(kpi_key))
            cached = "%s%d" % (label, m.agg[kpi_key])
        elif label.startswith("\U0001F9E6"):
            formula = '="%s"&COUNTA(%s)' % (label, bk.rng(KEY, "owner"))
            cached = "%s%d" % (label, len(set(x["owner"] for x in m.stockings)))
        else:
            formula = '="%s"&Currency&TEXT(%s-%s,"#,##0")' % (
                label, bk.kpi("stock_budget"), bk.kpi("stock_spent"))
            cached = "%s%s" % (label, m.money(m.agg["stock_budget"] -
                                              m.agg["stock_spent"]))
        ws.write_formula(r(C.ROW_STATS), ci(c1), formula, fmt, cached)
        bk.stats["formulas"] += 1

    K.table_frame(bk, KEY, COLUMNS, height=20)
    for i in range(K.n_rows(KEY)):
        rownum = C.ROW_FIRST + i
        values = {"n": '=IF($%s%d="","",ROW()-%d)'
                         % (bk.col(KEY, "item"), rownum, C.ROW_FIRST - 1),
                  "spent": '=IF($%s%d="","",IF($%s%d="%s",IF($%s%d<>"",$%s%d,'
                           'IF($%s%d<>"",$%s%d,0)),0))'
                           % (bk.col(KEY, "owner"), rownum,
                              bk.col(KEY, "bought"), rownum, C.TICK,
                              bk.col(KEY, "cost"), rownum, bk.col(KEY, "cost"),
                              rownum, bk.col(KEY, "budget"), rownum,
                              bk.col(KEY, "budget"), rownum)}
        if i < len(m.stockings):
            s = m.stockings[i]
            values.update({k: s[k] for k in
                           ("owner", "item", "category", "budget", "cost",
                            "bought", "wrapped", "hiding", "notes")})
        K.write_row(bk, KEY, COLUMNS, rownum, values, _cached(m, i))

    ws.autofilter(r(C.ROW_HEADER), ci("B"), r(C.last_row(KEY)), ci(LAST_COL))
    ws.freeze_panes(r(C.ROW_FIRST), ci("C"))

    K.list_dv(bk, KEY, "owner", "stocking_owners", title="Whose stocking?",
              message="Add or change stockings in \u2699\uFE0F Setup \u2192 "
                      "Stocking owners.")
    K.list_dv(bk, KEY, "category", "stocking_items", title="Type of filler")
    K.tick_dv(bk, KEY, ["bought", "wrapped"])
    K.money_dv(bk, KEY, ["budget", "cost"])
    ws.data_validation(r(C.ROW_FIRST), ci(bk.col(KEY, "hiding")),
                       r(C.last_row(KEY)), ci(bk.col(KEY, "hiding")),
                       {"validate": "list",
                        "source": "=" + bk.listname("hiding_spots"),
                        "ignore_blank": True, "show_input": True,
                        "input_title": "Where is it hiding?",
                        "input_message": "Pick a spot or type your own.",
                        "show_error": True, "error_type": "warning",
                        "error_title": "Free text is fine",
                        "error_message": "Add your own spots to the list in "
                                         "\u2699\uFE0F Setup."})
    bk.stats["validations"] += 1

    n = C.ROW_FIRST
    K.tick_cf(bk, KEY, ["bought", "wrapped"])
    K.secret_cf(bk, KEY, "hiding")
    bought = bk.col(KEY, "bought")
    bk.cond(KEY, n, ci("B"), C.last_row(KEY), ci(LAST_COL), {
        "type": "formula", "criteria": '=$%s%d="%s"' % (bought, n, C.TICK),
        "format": S.cf(bg=th.ok_soft, fg=th.muted, strike=True)})

    total_row = C.last_row(KEY) + 2
    K.totals_row(
        bk, KEY, total_row,
        {"budget": ("=SUM($%s$%d:$%s$%d)" % (bk.col(KEY, "budget"),
                                             C.ROW_FIRST,
                                             bk.col(KEY, "budget"),
                                             C.last_row(KEY)),
                    "#,##0.00", m.agg["stock_budget"]),
         "cost": ("=SUM($%s$%d:$%s$%d)" % (bk.col(KEY, "cost"), C.ROW_FIRST,
                                           bk.col(KEY, "cost"),
                                           C.last_row(KEY)),
                  "#,##0.00", sum(x["cost"] or 0 for x in m.stockings)),
         "spent": ("=SUM($%s$%d:$%s$%d)" % (bk.col(KEY, "spent"),
                                            C.ROW_FIRST,
                                            bk.col(KEY, "spent"),
                                            C.last_row(KEY)),
                   "#,##0.00", m.agg["stock_spent"])},
        label="  TOTALS", label_span=("B", "E"), last_col=LAST_COL)
    fmt = S.f(**S.base(font_size=11, bold=True, font_color=th.white,
                       bg_color=th.primary, align="center", valign="vcenter",
                       border=1, border_color=th.primary))
    ws.merge_range(r(total_row), ci("H"), r(total_row), ci(LAST_COL), "", fmt)
    ws.write_formula(
        r(total_row), ci("H"),
        '="\U0001F9E6 Stocking budget: "&Currency&TEXT(%s,"#,##0")&'
        '"   \u2022   spent: "&Currency&TEXT(%s,"#,##0")&"   \u2022   '
        'remaining: "&Currency&TEXT(%s-%s,"#,##0")'
        % (bk.kpi("stock_budget"), bk.kpi("stock_spent"),
           bk.kpi("stock_budget"), bk.kpi("stock_spent")), fmt,
        "\U0001F9E6 Stocking budget: %s   \u2022   spent: %s   \u2022   "
        "remaining: %s" % (m.money(m.agg["stock_budget"]),
                           m.money(m.agg["stock_spent"]),
                           m.money(m.agg["stock_budget"] -
                                   m.agg["stock_spent"])))
    bk.stats["formulas"] += 1

    # ------------------------------------------------------------------
    # per-stocking summary
    # ------------------------------------------------------------------
    row = total_row + 2
    ws.set_row(r(row), 24)
    ws.merge_range(r(row), 1, r(row), ci(LAST_COL),
                   "  \U0001F9E6  EACH STOCKING AT A GLANCE  (automatic)",
                   S.section_soft)
    row += 1
    heads = [("C", "Stocking"), ("D", "Items"), ("E", "Budget"),
             ("F", "Spent"), ("G", "Remaining"), ("H", "Bought"),
             ("I", "Wrapped"), ("J", "Progress")]
    ws.set_row(r(row), 26)
    for col, label in heads:
        c1 = ci(col)
        c2 = ci(LAST_COL) if col == "J" else c1
        if c2 > c1:
            ws.merge_range(r(row), c1, r(row), c2, label, S.header(th.primary))
        else:
            ws.write(r(row), c1, label, S.header(th.primary))
    for i in range(OWNER_ROWS):
        row += 1
        ws.set_row(r(row), 20)
        setup_row = C.SU_LIST_FIRST + i
        owner_ref = "'%s'!$I$%d" % (C.SHEET_NAMES["setup"], setup_row)
        name_ref = "$C%d" % row
        a = K.alt(row)
        owner_rng = bk.rng(KEY, "owner")
        spent_rng = bk.rng(KEY, "spent")
        budget_rng = bk.rng(KEY, "budget")
        bought_rng = bk.rng(KEY, "bought")
        wrapped_rng = bk.rng(KEY, "wrapped")
        cells = [
            ("C", '=IF(%s="","",%s)' % (owner_ref, owner_ref),
             S.cell("text", a, bold=True), m.owners[i] if i < len(m.owners)
             else ""),
            ("D", '=IF(%s="",0,COUNTIF(%s,%s))' % (name_ref, owner_rng,
                                                   name_ref),
             S.cell("calc_num"), _owner(m, i, "items")),
            ("E", '=IF(%s="",0,SUMIF(%s,%s,%s))' % (name_ref, owner_rng,
                                                    name_ref, budget_rng),
             S.cell("calc_money"), _owner(m, i, "budget")),
            ("F", '=IF(%s="",0,SUMIF(%s,%s,%s))' % (name_ref, owner_rng,
                                                    name_ref, spent_rng),
             S.cell("calc_money"), _owner(m, i, "spent")),
            ("G", '=IF(%s="","",%s)' % (name_ref, "$E%d-$F%d" % (row, row)),
             S.cell("calc_money"), _owner(m, i, "remaining")),
            ("H", '=IF(%s="",0,COUNTIFS(%s,%s,%s,"%s"))'
                  % (name_ref, owner_rng, name_ref, bought_rng, C.TICK),
             S.cell("calc_num"), _owner(m, i, "bought")),
            ("I", '=IF(%s="",0,COUNTIFS(%s,%s,%s,"%s"))'
                  % (name_ref, owner_rng, name_ref, wrapped_rng, C.TICK),
             S.cell("calc_num"),
             len([x for x in m.stockings
                  if i < len(m.owners) and x["owner"] == m.owners[i]
                  and x["wrapped"] == C.TICK])),
        ]
        for col, formula, fmt, cached in cells:
            ws.write_formula(r(row), ci(col), formula, fmt,
                             cached if cached is not None else 0)
            bk.stats["formulas"] += 1
        bar_fmt = S.f(**S.base(font_name=th.mono_font, font_size=10, bold=True,
                               font_color=th.primary_2, bg_color=th.card,
                               align="left", valign="vcenter", border=1,
                               border_color=th.border))
        ws.merge_range(r(row), ci("J"), r(row), ci(LAST_COL), "", bar_fmt)
        bar = ('=IF($C%d="","",REPT("\u2588",ROUND(MIN(1,IFERROR($F%d/$E%d,0))'
               '*16,0))&REPT("\u2591",16-ROUND(MIN(1,IFERROR($F%d/$E%d,0))*16,'
               '0))&"   "&TEXT(IFERROR($F%d/$E%d,0),"0%%")&" spent")'
               % ((row,) * 7))
        pct = 0
        bud = _owner(m, i, "budget") or 0
        spent = _owner(m, i, "spent") or 0
        pct = min(1.0, spent / float(bud)) if bud else 0
        filled = int(round(pct * 16))
        cached_bar = ""
        if i < len(m.owners):
            cached_bar = ("\u2588" * filled + "\u2591" * (16 - filled)
                          + "   %d%% spent" % round(pct * 100))
        ws.write_formula(r(row), ci("J"), bar, bar_fmt, cached_bar)
        bk.stats["formulas"] += 1

    row += 2
    row = K.note_block(
        bk, KEY, row, "B", LAST_COL,
        ["\u2022  One stocking = one name in \u2699\uFE0F Setup \u2192 "
         "Stocking owners. The summary above rebuilds itself as you add rows.",
         "\u2022  Tick \u2713 Bought and the money moves into the "
         "\u201CStocking stuffers\u201D line of the \U0001F4B0 Budget tab.",
         "\u2022  Hiding spots respond to Secret Mode (\u2699\uFE0F Setup) "
         "just like the \U0001F380 Wrapping tab."],
        title="  \U0001F4A1  How this tab works")
    row += 1
    ws.set_row(r(row), 26)
    row = bk.nav_row(KEY, row, first_col=1, span=2,
                     max_col=LAST_COL) + 2
    ws.merge_range(r(row), 1, r(row), ci(LAST_COL),
                   "  %s v%s  \u2022  %s" % (C.PRODUCT, C.VERSION, C.TAGLINE),
                   S.footer)
    bk.page(KEY, LAST_COL, row + 2, landscape=True,
            title_rows=(0, C.ROW_HEADER - 1))
    return ws


def _owner(m, i, field):
    rows = m.agg["stock_by_owner"]
    if i >= len(rows):
        return ""
    val = rows[i][field]
    return val if val != "" else 0


def _cached(m, i):
    if i >= len(m.stockings):
        return {"n": "", "spent": ""}
    s = m.stockings[i]
    return {"n": i + 1, "spent": m._stock_spent(s)}
```

### `christmas_tracker/sheets/todo.py`  (219 lines)

```python
"""
\u2705 To-Do List - date-aware, so it shouts about the things that are about
to bite you.

Deadlines ship as formulas relative to your event date (``=EventDate-45``),
which means the whole plan re-dates itself when you move the event.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "todo"
LAST_COL = "H"

COLUMNS = [
    ("done", "Done", "tick", None),
    ("task", "Task", "text", None),
    ("category", "Category", "center", None),
    ("deadline", "Deadline", "date", None),
    ("days", "Days left", "calc_num", "primary_2"),
    ("status", "Status (auto)", "calc_c", "primary_2"),
    ("notes", "Notes", "text", None),
]


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    bk.widths(KEY)
    bk.paint(KEY, 0, 0, C.last_row(KEY) + 28, ci(LAST_COL), S.canvas)
    bk.title_block(
        KEY,
        "  \u2705  Christmas To-Do List",
        "  Pre-loaded with the classic holiday checklist \u2014 every "
        "deadline is calculated from your event date.",
        LAST_COL)

    chips = [
        ("B", "C", "\U0001F4CB Tasks: ", "todo_total", "primary"),
        ("D", "D", "\u2705 Done: ", "todo_done", "ok"),
        ("E", "E", "\U0001F534 Overdue: ", "todo_overdue", "bad"),
        ("F", "F", "\U0001F7E0 Due soon: ", None, "warn"),
        ("G", LAST_COL, "\U0001F4CA Progress: ", None, "gold"),
    ]
    for c1, c2, label, kpi_key, color in chips:
        fmt = S.pill(th.soft(color), getattr(th, color), size=10.5, bold=True,
                     align="left")
        if c1 == c2:
            ws.write(r(C.ROW_STATS), ci(c1), "", fmt)
        else:
            ws.merge_range(r(C.ROW_STATS), ci(c1), r(C.ROW_STATS), ci(c2),
                           "", fmt)
        if kpi_key:
            formula = '="%s"&%s' % (label, bk.kpi(kpi_key))
            cached = "%s%d" % (label, m.agg[kpi_key])
        elif label.startswith("\U0001F7E0"):
            formula = _due_soon_formula(bk, label)
            cached = "%s%d" % (label, m.agg["todo_soon"])
        else:
            formula = ('="%s"&REPT("\u2588",ROUND(MIN(1,IFERROR(%s/%s,0))*14,0))'
                       '&REPT("\u2591",14-ROUND(MIN(1,IFERROR(%s/%s,0))*14,0))'
                       '&"  "&TEXT(IFERROR(%s/%s,0),"0%%")'
                       % ((label, bk.kpi("todo_done"), bk.kpi("todo_total"))
                          + (bk.kpi("todo_done"), bk.kpi("todo_total"))
                          + (bk.kpi("todo_done"), bk.kpi("todo_total"))))
            pct = (m.agg["todo_done"] / float(m.agg["todo_total"])
                   if m.agg["todo_total"] else 0)
            filled = int(round(min(1.0, pct) * 14))
            cached = "%s%s  %d%%" % (label, "\u2588" * filled +
                                     "\u2591" * (14 - filled),
                                     round(pct * 100))
        ws.write_formula(r(C.ROW_STATS), ci(c1), formula, fmt, cached)
        bk.stats["formulas"] += 1

    K.table_frame(bk, KEY, COLUMNS, height=20)
    for i in range(K.n_rows(KEY)):
        rownum = C.ROW_FIRST + i
        values = _formulas(bk, rownum)
        if i < len(m.todos):
            t = m.todos[i]
            values.update({"done": t["done"], "task": t["task"],
                           "category": t["category"],
                           "deadline": "=EventDate-%d" % t["offset"]
                           if t["offset"] >= 0 else
                           "=EventDate+%d" % abs(t["offset"]),
                           "notes": ""})
        K.write_row(bk, KEY, COLUMNS, rownum, values, _cached(m, i))

    ws.autofilter(r(C.ROW_HEADER), ci("B"), r(C.last_row(KEY)), ci(LAST_COL))
    ws.freeze_panes(r(C.ROW_FIRST), 0)

    K.list_dv(bk, KEY, "category", "todo_categories", title="Category")
    K.tick_dv(bk, KEY, ["done"], message="Tick \u2713 when it is done \u2014 "
                                         "the row greys out and the status "
                                         "turns green.")
    K.date_dv(bk, KEY, ["deadline"],
              message="Type a date, or a formula like =EventDate-14 so it "
                      "moves with your event date.")

    n = C.ROW_FIRST
    done = bk.col(KEY, "done")
    status = bk.col(KEY, "status")
    K.tick_cf(bk, KEY, ["done"])
    for text, (bg, fg) in (("\U0001F7E2 Complete", (th.ok_soft, th.ok)),
                           ("\U0001F534 OVERDUE", (th.bad_soft, th.bad)),
                           ("\U0001F7E0 Due soon", (th.warn_soft, th.warn)),
                           ("\U0001F535 Upcoming", (th.info_soft, th.info))):
        bk.cond(KEY, n, ci(status), C.last_row(KEY), ci(status), {
            "type": "formula",
            "criteria": '=$%s%d="%s"' % (status, n, text),
            "format": S.cf(bg=bg, fg=fg, bold=True)})
    K.deadline_cf(bk, KEY, "deadline", "done")
    bk.cond(KEY, n, ci("B"), C.last_row(KEY), ci(LAST_COL), {
        "type": "formula", "criteria": '=$%s%d="%s"' % (done, n, C.TICK),
        "format": S.cf(bg=th.ok_soft, fg=th.muted, strike=True)})
    bk.cond(KEY, n, ci("B"), C.last_row(KEY), ci(LAST_COL), {
        "type": "formula",
        "criteria": '=AND($%s%d<>"%s",$%s%d<>"",$%s%d<TODAY())'
                   % (done, n, C.TICK, bk.col(KEY, "deadline"), n,
                      bk.col(KEY, "deadline"), n),
        "format": S.cf(bg=th.bad_soft, fg=th.bad)})

    total_row = C.last_row(KEY) + 2
    fmt = S.f(**S.base(font_size=11, bold=True, font_color=th.white,
                       bg_color=th.primary, align="center", valign="vcenter",
                       border=1, border_color=th.primary))
    K.totals_row(bk, KEY, total_row, {}, label="  CHECKLIST",
                 label_span=("B", "B"), last_col=LAST_COL)
    ws.merge_range(r(total_row), ci("C"), r(total_row), ci(LAST_COL), "", fmt)
    ws.write_formula(
        r(total_row), ci("C"),
        '="\u2705 "&%s&" of "&%s&" tasks done   \u2022   \U0001F534 "&%s'
        '&" overdue   \u2022   \U0001F7E0 "&%s&" due in the next "'
        '&DueSoonDays&" days"'
        % (bk.kpi("todo_done"), bk.kpi("todo_total"), bk.kpi("todo_overdue"),
           _due_soon_ref(bk)), fmt,
        "\u2705 %d of %d tasks done   \u2022   \U0001F534 %d overdue   "
        "\u2022   \U0001F7E0 %d due in the next %d days"
        % (m.agg["todo_done"], m.agg["todo_total"], m.agg["todo_overdue"],
           m.agg["todo_soon"], m.settings["duesoon"]))
    bk.stats["formulas"] += 1

    row = total_row + 2
    row = K.note_block(
        bk, KEY, row, "B", LAST_COL,
        ["\u2022  Deadlines are formulas like ``=EventDate-45`` so the whole "
         "plan re-dates itself when you change the event date in "
         "\u2699\uFE0F Setup. Want a fixed date? Just type over it.",
         "\u2022  Status is automatic: \U0001F7E2 Complete, \U0001F534 "
         "OVERDUE, \U0001F7E0 Due soon (inside your Setup window) or "
         "\U0001F535 Upcoming.",
         "\u2022  Overdue and due-soon tasks also appear on the dashboard's "
         "\u201CWhat's left to do\u201D and \u201CUpcoming deadlines\u201D "
         "panels."],
        title="  \U0001F4A1  How this tab works")
    row += 1
    ws.set_row(r(row), 26)
    row = bk.nav_row(KEY, row, first_col=1, span=2,
                     max_col=LAST_COL) + 2
    ws.merge_range(r(row), 1, r(row), ci(LAST_COL),
                   "  %s v%s  \u2022  %s" % (C.PRODUCT, C.VERSION, C.TAGLINE),
                   S.footer)
    bk.page(KEY, LAST_COL, row + 2, landscape=True,
            title_rows=(0, C.ROW_HEADER - 1))
    return ws


def _formulas(bk, n):
    def L(f):
        return bk.col(KEY, f)

    return {
        "days": '=IF(OR($%s%d="",$%s%d="%s"),"",$%s%d-TODAY())'
                % (L("deadline"), n, L("done"), n, C.TICK, L("deadline"), n),
        "status": ('=IF($%s%d="","",IF($%s%d="%s","\U0001F7E2 Complete",'
                   'IF($%s%d="","",IF($%s%d<TODAY(),"\U0001F534 OVERDUE",'
                   'IF($%s%d-TODAY()<=DueSoonDays,"\U0001F7E0 Due soon",'
                   '"\U0001F535 Upcoming")))))'
                   % (L("task"), n, L("done"), n, C.TICK, L("deadline"), n,
                      L("deadline"), n, L("deadline"), n)),
    }


def _due_soon_formula(bk, label):
    done = bk.rng(KEY, "done")
    dl = bk.rng(KEY, "deadline")
    return ('="%s"&SUMPRODUCT((%s<>"%s")*(%s<>"")*(%s>=TODAY())*'
            '(%s-TODAY()<=DueSoonDays))' % (label, done, C.TICK, dl, dl, dl))


def _due_soon_ref(bk):
    done = bk.rng(KEY, "done")
    dl = bk.rng(KEY, "deadline")
    return ('SUMPRODUCT((%s<>"%s")*(%s<>"")*(%s>=TODAY())*'
            '(%s-TODAY()<=DueSoonDays))' % (done, C.TICK, dl, dl, dl))


def _cached(m, i):
    from datetime import date
    today = date.today()
    if i >= len(m.todos):
        return {"days": "", "status": ""}
    t = m.todos[i]
    dl = t["deadline"]
    if t["done"] == C.TICK:
        return {"days": "", "status": "\U0001F7E2 Complete"}
    if not dl:
        return {"days": "", "status": ""}
    days = (dl - today).days
    if days < 0:
        status = "\U0001F534 OVERDUE"
    elif days <= m.settings["duesoon"]:
        status = "\U0001F7E0 Due soon"
    else:
        status = "\U0001F535 Upcoming"
    return {"days": days, "status": status}
```

### `christmas_tracker/sheets/dashboard.py`  (872 lines)

```python
"""
\U0001F384 Dashboard - the "wow" page and the reason the product feels
premium.

Everything here is a formula: countdown, money, gift progress, text progress
bars, four charts, the "what's left to do" panel, the next five deadlines and
the per-recipient table.  Nothing is typed on this tab.
"""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "dashboard"
LAST_COL = "M"
CARDS = [("B", "D"), ("E", "G"), ("H", "J"), ("K", "M")]

ROW_HERO_1, ROW_HERO_2, ROW_HERO_3 = 2, 3, 4
ROW_BUD_SEC, ROW_BUD_LBL, ROW_BUD_VAL, ROW_BUD_BAR = 6, 7, 8, 9
ROW_GIF_SEC = 11
ROW_GIF_LBL1, ROW_GIF_VAL1 = 12, 13
ROW_GIF_LBL2, ROW_GIF_VAL2 = 14, 15
ROW_GIF_BAR = 16
ROW_CHART_SEC = 18
ROW_CHART_1 = 19
ROW_CHART_2 = 34
ROW_PANEL_SEC = 49
ROW_PANEL = 50
ROW_REC_SEC = 61
ROW_REC_HDR = 62
ROW_REC_FIRST = 63
ROW_REC_TOTAL = 79
ROW_NAV_SEC = 81
ROW_NAV = 82
ROW_FOOTER = 88


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo
    agg = m.agg

    ws.set_column("A:A", 2.2)
    ws.set_column("B:M", 13)
    bk.paint(KEY, 0, 0, ROW_FOOTER + 2, ci(LAST_COL), S.canvas)

    # ------------------------------------------------------------------
    # hero
    # ------------------------------------------------------------------
    ws.set_row(r(1), 8)
    ws.set_row(r(ROW_HERO_1), 40)
    ws.set_row(r(ROW_HERO_2), 34)
    ws.set_row(r(ROW_HERO_3), 20)
    title = ('="\U0001F384  "&EventName&" "&TEXT(EventDate,"yyyy")&'
             '"   \u2022   Gift Command Center"')
    ws.merge_range(r(ROW_HERO_1), 1, r(ROW_HERO_1), ci(LAST_COL), "",
                   S.hero_title)
    ws.write_formula(r(ROW_HERO_1), 1, title, S.hero_title,
                     "\U0001F384  %s %d   \u2022   Gift Command Center"
                     % (m.settings["event_name"], m.agg["event_year"]))
    ws.merge_range(r(ROW_HERO_2), 1, r(ROW_HERO_2), ci(LAST_COL), "",
                   S.hero_count)
    ws.write_formula(r(ROW_HERO_2), 1, "=" + _countdown(bk), S.hero_count,
                     agg["countdown"])
    ws.merge_range(r(ROW_HERO_3), 1, r(ROW_HERO_3), ci(LAST_COL), "",
                   S.hero_meta)
    meta = ('=TEXT(TODAY(),"dddd d mmmm yyyy")&"   \u2022   "&EventName&'
            '" falls on "&TEXT(EventDate,"dddd d mmmm yyyy")&"   \u2022   "'
            '&%s&" of "&%s&" gifts bought   \u2022   "&%s&" wrapped   '
            '\u2022   "&%s&" handed over"'
            % (bk.kpi("gifts_purchased"), bk.kpi("gifts_planned"),
               bk.kpi("gifts_wrapped"), bk.kpi("gifts_delivered")))
    ws.write_formula(r(ROW_HERO_3), 1, meta, S.hero_meta,
                     "%s   \u2022   %s falls on %s   \u2022   %d of %d gifts "
                     "bought   \u2022   %d wrapped   \u2022   %d handed over"
                     % (_today_text(), m.settings["event_name"],
                        m.event_date.strftime("%A %d %B %Y"),
                        agg["gifts_purchased"], agg["gifts_planned"],
                        agg["gifts_wrapped"], agg["gifts_delivered"]))
    bk.stats["formulas"] += 3
    ws.set_row(r(5), 8)

    # ------------------------------------------------------------------
    # budget cards
    # ------------------------------------------------------------------
    _section(ws, S, ROW_BUD_SEC, "  \U0001F4B0  BUDGET AT A GLANCE")
    budget_cards = [
        ("TOTAL %s BUDGET" % m.settings["event_name"].upper(), "primary",
         '=Currency&TEXT(%s,"#,##0")' % bk.kpi("budget_planned"),
         m.money(agg["budget_planned"])),
        ("SPENT SO FAR", "accent",
         '=Currency&TEXT(%s,"#,##0")' % bk.kpi("budget_actual"),
         m.money(agg["budget_actual"])),
        ("STILL TO SPEND", "ok",
         '=Currency&TEXT(%s,"#,##0")' % bk.kpi("budget_remaining"),
         m.money(agg["budget_remaining"])),
        ("BUDGET USED", "gold", "=%s" % bk.kpi("budget_pct"),
         agg["budget_pct"]),
    ]
    _cards(bk, ROW_BUD_LBL, ROW_BUD_VAL, budget_cards,
           num_formats=[None, None, None, "0%"])

    # budget progress bar
    ws.set_row(r(ROW_BUD_BAR), 24)
    ws.merge_range(r(ROW_BUD_BAR), 1, r(ROW_BUD_BAR), ci("C"),
                   "  Budget progress", S.bar_label)
    ws.merge_range(r(ROW_BUD_BAR), ci("D"), r(ROW_BUD_BAR), ci("K"), "",
                   S.bar_text)
    ws.write_formula(r(ROW_BUD_BAR), ci("D"),
                     _bar_formula(bk, bk.kpi("budget_actual"),
                                  bk.kpi("budget_planned"), 40), S.bar_text,
                     _bar(agg["budget_pct"], 40))
    ws.merge_range(r(ROW_BUD_BAR), ci("L"), r(ROW_BUD_BAR), ci(LAST_COL), "",
                   S.bar_pct)
    ws.write_formula(r(ROW_BUD_BAR), ci("L"),
                     '=TEXT(%s,"0%%")&" used"' % bk.kpi("budget_pct"),
                     S.bar_pct, "%d%% used" % round(agg["budget_pct"] * 100))
    bk.stats["formulas"] += 2
    ws.set_row(r(10), 8)

    # ------------------------------------------------------------------
    # gift cards
    # ------------------------------------------------------------------
    _section(ws, S, ROW_GIF_SEC, "  \U0001F381  GIFT PROGRESS",
             style=S.section_accent)
    gift_cards_1 = [
        ("\U0001F381 GIFTS PLANNED", "primary", "=%s" % bk.kpi("gifts_planned"),
         agg["gifts_planned"]),
        ("\u2705 PURCHASED", "ok",
         '=%s&" / "&%s' % (bk.kpi("gifts_purchased"), bk.kpi("gifts_planned")),
         "%d / %d" % (agg["gifts_purchased"], agg["gifts_planned"])),
        ("\U0001F380 WRAPPED", "accent", "=%s" % bk.kpi("gifts_wrapped"),
         agg["gifts_wrapped"]),
        ("\U0001F4E6 DELIVERED / GIVEN", "info",
         "=%s" % bk.kpi("gifts_delivered"), agg["gifts_delivered"]),
    ]
    gift_cards_2 = [
        ("\U0001F6D2 STILL TO BUY", "warn", "=%s" % bk.kpi("gifts_to_buy"),
         agg["gifts_to_buy"]),
        ("\U0001F6CD\uFE0F ON ORDER", "plum", "=%s" % bk.kpi("gifts_ordered"),
         agg["gifts_ordered"]),
        ("\U0001F4C8 COMPLETION", "gold", "=%s" % bk.kpi("gift_completion"),
         agg["gift_completion"]),
        ("\U0001F465 AVG SPEND PER PERSON", "primary_2",
         '=Currency&TEXT(%s,"#,##0")' % bk.kpi("avg_per_recipient"),
         m.money(agg["avg_per_recipient"])),
    ]
    _cards(bk, ROW_GIF_LBL1, ROW_GIF_VAL1, gift_cards_1)
    _cards(bk, ROW_GIF_LBL2, ROW_GIF_VAL2, gift_cards_2,
           num_formats=[None, None, "0%", None])

    ws.set_row(r(ROW_GIF_BAR), 24)
    ws.merge_range(r(ROW_GIF_BAR), 1, r(ROW_GIF_BAR), ci("C"),
                   "  Gift progress", S.bar_label)
    ws.merge_range(r(ROW_GIF_BAR), ci("D"), r(ROW_GIF_BAR), ci("K"), "",
                   S.bar_text)
    ws.write_formula(r(ROW_GIF_BAR), ci("D"),
                     _bar_formula(bk, bk.kpi("gifts_purchased"),
                                  bk.kpi("gifts_planned"), 40), S.bar_text,
                     _bar(agg["gift_completion"], 40))
    ws.merge_range(r(ROW_GIF_BAR), ci("L"), r(ROW_GIF_BAR), ci(LAST_COL), "",
                   S.bar_pct)
    ws.write_formula(r(ROW_GIF_BAR), ci("L"),
                     '=%s&" / "&%s&" gifts"'
                     % (bk.kpi("gifts_purchased"), bk.kpi("gifts_planned")),
                     S.bar_pct, "%d / %d gifts" % (agg["gifts_purchased"],
                                                   agg["gifts_planned"]))
    bk.stats["formulas"] += 2
    ws.set_row(r(17), 8)

    # ------------------------------------------------------------------
    # charts
    # ------------------------------------------------------------------
    _section(ws, S, ROW_CHART_SEC, "  \U0001F4CA  WHERE THE MONEY GOES")
    for row in range(ROW_CHART_1, ROW_CHART_2 + 15):
        ws.set_row(r(row), 20)
    _charts(bk)
    ws.set_row(r(48), 8)

    # ------------------------------------------------------------------
    # panels: what's left + deadlines
    # ------------------------------------------------------------------
    ws.set_row(r(ROW_PANEL_SEC), 22)
    ws.merge_range(r(ROW_PANEL_SEC), 1, r(ROW_PANEL_SEC), ci("G"),
                   "  \U0001F3AF  WHAT'S LEFT TO DO", S.section)
    ws.merge_range(r(ROW_PANEL_SEC), ci("H"), r(ROW_PANEL_SEC), ci(LAST_COL),
                   "  \u23F0  UPCOMING DEADLINES", S.section_gold)
    _left_panel(bk)
    _right_panel(bk)
    ws.set_row(r(60), 8)

    # ------------------------------------------------------------------
    # per-recipient table
    # ------------------------------------------------------------------
    _section(ws, S, ROW_REC_SEC,
             "  \U0001F385  PER-RECIPIENT SUMMARY   (automatic)",
             style=S.section_soft)
    _recipients(bk)
    ws.set_row(r(80), 8)

    # ------------------------------------------------------------------
    # navigation + footer
    # ------------------------------------------------------------------
    _section(ws, S, ROW_NAV_SEC, "  \U0001F9ED  JUMP STRAIGHT TO",
             style=S.section_soft)
    ws.set_row(r(ROW_NAV), 26)
    ws.set_row(r(ROW_NAV + 1), 26)
    foot = bk.nav_row(KEY, ROW_NAV, first_col=1, span=2,
                      max_col=LAST_COL) + 2
    ws.set_row(r(foot - 1), 8)
    ws.merge_range(r(foot), 1, r(foot), ci(LAST_COL),
                   "  %s v%s  \u2022  %s  \u2022  Excel 2016+ and Google "
                   "Sheets  \u2022  no macros  \u2022  change the event date "
                   "in \u2699\uFE0F Setup and the whole workbook follows"
                   % (C.PRODUCT, C.VERSION, C.TAGLINE), S.footer)

    bk.page(KEY, LAST_COL, foot + 2, landscape=True, zoom=85)
    return ws


# ===========================================================================
# pieces
# ===========================================================================
def _section(ws, S, row, text, style=None):
    ws.set_row(r(row), 22)
    ws.merge_range(r(row), 1, r(row), ci(LAST_COL), text, style or S.section)


def _cards(bk, label_row, value_row, cards, num_formats=None,
           label_formulas=None):
    ws, S, th = bk.ws(KEY), bk.S, bk.th
    num_formats = num_formats or [None] * len(cards)
    label_formulas = label_formulas or [False] * len(cards)
    ws.set_row(r(label_row), 18)
    ws.set_row(r(value_row), 38)
    for i, (label, color, formula, cached) in enumerate(cards):
        c1, c2 = CARDS[i]
        lbl_fmt = S.kpi_label(getattr(th, color), size=9.5,
                              align="center" if not label_formulas[i] else
                              "center")
        ws.merge_range(r(label_row), ci(c1), r(label_row), ci(c2), "",
                       lbl_fmt)
        if label_formulas[i] or (isinstance(label, str) and
                                 label.startswith("=")):
            ws.write_formula(r(label_row), ci(c1), label, lbl_fmt, "")
            bk.stats["formulas"] += 1
        else:
            ws.write(r(label_row), ci(c1), label, lbl_fmt)
        val_fmt = S.kpi_value(getattr(th, color),
                              num_format=num_formats[i], size=22)
        ws.merge_range(r(value_row), ci(c1), r(value_row), ci(c2), "",
                       val_fmt)
        ws.write_formula(r(value_row), ci(c1), formula, val_fmt, cached)
        bk.stats["formulas"] += 1


def _charts(bk):
    ws, th = bk.ws(KEY), bk.th
    d = bk.q("data")

    # 1 - spending per person (horizontal bar)
    ch = bk.chart("bar")
    ch.add_series({
        "name": "Spent",
        "categories": "=%s!$H$%d:$H$%d" % (d, C.DATA_REC_FIRST,
                                           C.DATA_REC_FIRST +
                                           C.DATA_REC_ROWS - 1),
        "values": "=%s!$J$%d:$J$%d" % (d, C.DATA_REC_FIRST,
                                       C.DATA_REC_FIRST + C.DATA_REC_ROWS - 1),
        "fill": {"color": th.accent},
        "border": {"color": th.accent},
        "gap": 45,
    })
    ch.set_title({"name": "Spending per person",
                  "name_font": {"size": 12, "bold": True, "color": th.primary,
                                "name": th.body_font}})
    ch.set_x_axis({"num_format": "#,##0", "num_font": {"size": 9},
                   "major_gridlines": {"visible": True,
                                       "line": {"color": th.border}}})
    ch.set_y_axis({"num_font": {"size": 9}, "label_position": "low"})
    ch.set_size({"width": 555, "height": 275})
    ch.set_legend({"none": True})
    ws.insert_chart(r(ROW_CHART_1), ci("B"), ch, {"x_offset": 8,
                                                  "y_offset": 6})

    # 2 - planned vs spent by budget category
    b = bk.q("budget")
    first, last = C.BUD_FIRST, C.BUD_FIRST + C.BUDGET_ROWS - 1
    cats = "=%s!$B$%d:$B$%d" % (b, first, last)
    ch = bk.chart("column")
    ch.add_series({"name": "Planned", "categories": cats,
                   "values": "=%s!$C$%d:$C$%d" % (b, first, last),
                   "fill": {"color": th.border_strong},
                   "border": {"color": th.border_strong},
                   "gap": 55})
    ch.add_series({"name": "Spent", "categories": cats,
                   "values": "=%s!$F$%d:$F$%d" % (b, first, last),
                   "fill": {"color": th.primary},
                   "border": {"color": th.primary}})
    ch.set_title({"name": "Budget vs actual, by category",
                  "name_font": {"size": 12, "bold": True, "color": th.primary,
                                "name": th.body_font}})
    ch.set_y_axis({"num_format": "#,##0", "num_font": {"size": 9},
                   "major_gridlines": {"visible": True,
                                       "line": {"color": th.border}}})
    ch.set_x_axis({"num_font": {"size": 8}})
    ch.set_size({"width": 555, "height": 275})
    ch.set_legend({"position": "bottom", "font": {"size": 9}})
    ws.insert_chart(r(ROW_CHART_1), ci("H"), ch, {"x_offset": 8,
                                                  "y_offset": 6})

    # 3 - gift status doughnut
    ch = bk.chart("doughnut")
    ch.add_series({
        "name": "Gifts",
        "categories": "=%s!$T$%d:$T$%d" % (d, C.DATA_STATUS_FIRST,
                                           C.DATA_STATUS_FIRST + 5),
        "values": "=%s!$U$%d:$U$%d" % (d, C.DATA_STATUS_FIRST,
                                       C.DATA_STATUS_FIRST + 5),
        "points": [{"fill": {"color": c}} for c in
                   [th.plum, th.warn, th.info, th.ok, th.gold, th.primary]],
        "data_labels": {"value": True, "font": {"size": 9,
                                                "color": th.white}},
    })
    ch.set_title({"name": "Where the gifts are at",
                  "name_font": {"size": 12, "bold": True, "color": th.primary,
                                "name": th.body_font}})
    ch.set_hole_size(58)
    ch.set_size({"width": 555, "height": 275})
    ch.set_legend({"position": "right", "font": {"size": 9}})
    ws.insert_chart(r(ROW_CHART_2), ci("B"), ch, {"x_offset": 8,
                                                  "y_offset": 6})

    # 4 - bought vs still to buy, per person
    ch = bk.chart("column", subtype="stacked")
    ch.add_series({
        "name": "Bought",
        "categories": "=%s!$H$%d:$H$%d" % (d, C.DATA_REC_FIRST,
                                           C.DATA_REC_FIRST +
                                           C.DATA_REC_ROWS - 1),
        "values": "=%s!$L$%d:$L$%d" % (d, C.DATA_REC_FIRST,
                                       C.DATA_REC_FIRST + C.DATA_REC_ROWS - 1),
        "fill": {"color": th.primary}, "border": {"color": th.primary},
        "gap": 60, "overlap": 100})
    ch.add_series({
        "name": "Still to buy",
        "categories": "=%s!$H$%d:$H$%d" % (d, C.DATA_REC_FIRST,
                                           C.DATA_REC_FIRST +
                                           C.DATA_REC_ROWS - 1),
        "values": "=%s!$O$%d:$O$%d" % (d, C.DATA_REC_FIRST,
                                       C.DATA_REC_FIRST + C.DATA_REC_ROWS - 1),
        "fill": {"color": th.gold}, "border": {"color": th.gold}})
    ch.set_title({"name": "Gifts per person: bought vs still to buy",
                  "name_font": {"size": 12, "bold": True, "color": th.primary,
                                "name": th.body_font}})
    ch.set_y_axis({"num_font": {"size": 9}, "major_unit": 1,
                   "major_gridlines": {"visible": True,
                                       "line": {"color": th.border}}})
    ch.set_x_axis({"num_font": {"size": 8}})
    ch.set_size({"width": 555, "height": 275})
    ch.set_legend({"position": "bottom", "font": {"size": 9}})
    ws.insert_chart(r(ROW_CHART_2), ci("H"), ch, {"x_offset": 8,
                                                  "y_offset": 6})


def _left_panel(bk):
    ws, S, th, m = bk.ws(KEY), bk.S, bk.th, bk.demo
    lines = _todo_lines(bk)
    panel = S.panel(color=th.border, size=10.5, bold=False, valign="vcenter",
                    wrap=False)
    for i, (formula, cached) in enumerate(lines[:10]):
        row = ROW_PANEL + i
        ws.set_row(r(row), 19)
        ws.merge_range(r(row), 1, r(row), ci("G"), "", panel)
        ws.write_formula(r(row), 1, formula, panel, cached)
        bk.stats["formulas"] += 1
    for i in range(len(lines), 10):
        row = ROW_PANEL + i
        ws.set_row(r(row), 19)
        ws.merge_range(r(row), 1, r(row), ci("G"), "", panel)
    # colour the panic lines
    bk.cond(KEY, ROW_PANEL, 1, ROW_PANEL + 9, ci("G"), {
        "type": "formula",
        "criteria": '=ISNUMBER(SEARCH("\U0001F534",$B%d))' % ROW_PANEL,
        "format": S.cf(bg=th.bad_soft, fg=th.bad, bold=True)})
    bk.cond(KEY, ROW_PANEL, 1, ROW_PANEL + 9, ci("G"), {
        "type": "formula",
        "criteria": '=ISNUMBER(SEARCH("\u26A0\uFE0F",$B%d))' % ROW_PANEL,
        "format": S.cf(bg=th.warn_soft, fg=th.warn, bold=True)})
    bk.cond(KEY, ROW_PANEL, 1, ROW_PANEL + 9, ci("G"), {
        "type": "formula",
        "criteria": '=ISNUMBER(SEARCH("\u2705",$B%d))' % ROW_PANEL,
        "format": S.cf(bg=th.ok_soft, fg=th.ok, bold=True)})


def _right_panel(bk):
    ws, S, th, m = bk.ws(KEY), bk.S, bk.th, bk.demo
    pool_first = C.DATA_POOL_FIRST
    pool_last = C.DATA_POOL_FIRST + C.pool_size(bk.edition) - 1
    dates = bk.data_rng("AA", pool_first, pool_last)
    labels = bk.data_rng("AB", pool_first, pool_last)

    date_fmt = S.f(**S.base(font_size=10, bold=True, font_color=th.accent,
                            bg_color=th.card, align="center",
                            valign="vcenter", border=1,
                            border_color=th.border,
                            num_format="dd mmm"))
    label_fmt = S.f(**S.base(font_size=10, font_color=th.ink,
                             bg_color=th.card, align="left", valign="vcenter",
                             border=1, border_color=th.border, indent=1))
    days_fmt = S.f(**S.base(font_size=10, bold=True, font_color=th.primary,
                            bg_color=th.card, align="center",
                            valign="vcenter", border=1,
                            border_color=th.border, num_format='0 "days"'))
    pool = m.agg["pool"]
    for i in range(5):
        row = ROW_PANEL + i
        ws.set_row(r(row), 19)
        small = "SMALL(%s,%d)" % (dates, i + 1)
        ws.write_formula(r(row), ci("H"), "=IFERROR(%s,\"\")" % small,
                         date_fmt, pool[i][0] if i < len(pool) else "")
        ws.merge_range(r(row), ci("I"), r(row), ci("K"), "", label_fmt)
        ws.write_formula(
            r(row), ci("I"),
            '=IF($H%d="","\u2014 nothing scheduled \u2014",IFERROR(INDEX(%s,'
            'MATCH(%s,%s,0)),""))' % (row, labels, small, dates),
            label_fmt, pool[i][1] if i < len(pool)
            else "\u2014 nothing scheduled \u2014")
        ws.merge_range(r(row), ci("L"), r(row), ci(LAST_COL), "", days_fmt)
        ws.write_formula(r(row), ci("L"),
                         '=IF($H%d="","",$H%d-TODAY())' % (row, row),
                         days_fmt, (pool[i][0] - _today()).days
                         if i < len(pool) else "")
        bk.stats["formulas"] += 3

    # key dates
    key_hdr = S.f(**S.base(font_size=10, bold=True, font_color=th.ink,
                           bg_color=th.gold_soft, align="left",
                           valign="vcenter", border=1, border_color=th.border,
                           indent=1))
    row = ROW_PANEL + 5
    ws.set_row(r(row), 19)
    ws.merge_range(r(row), ci("H"), r(row), ci(LAST_COL),
                   "  \U0001F5D3\uFE0F  KEY DATES (calculated from your "
                   "event date)", key_hdr)
    key_dates = [
        ("\U0001F4EE Last online order date", 10),
        ("\U0001F48C Cards in the post by", 14),
        ("\U0001F380 Wrapping day", 2),
        ("\U0001F384 The big day", 0),
    ]
    for i, (label, back) in enumerate(key_dates):
        row += 1
        ws.set_row(r(row), 19)
        when = "EventDate" if back == 0 else "EventDate-%d" % back
        from datetime import timedelta
        ws.write_formula(r(row), ci("H"), "=" + when, date_fmt,
                         m.event_date - timedelta(days=back))
        ws.merge_range(r(row), ci("I"), r(row), ci("K"), label, label_fmt)
        ws.merge_range(r(row), ci("L"), r(row), ci(LAST_COL), "", days_fmt)
        ws.write_formula(r(row), ci("L"),
                         '=IF($H%d-TODAY()<0,"\u2713 passed",$H%d-TODAY())'
                         % (row, row), days_fmt,
                         "\u2713 passed"
                         if (m.event_date - timedelta(days=back)) < _today()
                         else ((m.event_date - timedelta(days=back)) -
                               _today()).days)
        bk.stats["formulas"] += 2


def _recipients(bk):
    ws, S, th, m = bk.ws(KEY), bk.S, bk.th, bk.demo
    heads = [("B", "C", "Recipient"), ("D", None, "Budget"),
             ("E", None, "Spent"), ("F", None, "Remaining"),
             ("G", None, "Gifts"), ("H", None, "Bought"),
             ("I", None, "Wrapped"), ("J", None, "Given"),
             ("K", None, "To buy"), ("L", None, "Done %"),
             ("M", None, "Progress")]
    ws.set_row(r(ROW_REC_HDR), 26)
    for c1, c2, label in heads:
        if c2:
            ws.merge_range(r(ROW_REC_HDR), ci(c1), r(ROW_REC_HDR), ci(c2),
                           label, S.header(th.primary))
        else:
            ws.write(r(ROW_REC_HDR), ci(c1), label, S.header(th.primary))

    name_fmt = S.f(**S.base(font_size=10.5, bold=True, font_color=th.ink,
                            bg_color=th.card, align="left", valign="vcenter",
                            border=1, border_color=th.border, indent=1))
    money_fmt = S.f(**S.base(font_size=10.5, font_color=th.ink,
                             bg_color=th.card, align="right",
                             valign="vcenter", border=1,
                             border_color=th.border, num_format="#,##0"))
    num_fmt = S.f(**S.base(font_size=10.5, font_color=th.ink,
                           bg_color=th.card, align="center", valign="vcenter",
                           border=1, border_color=th.border,
                           num_format="0"))
    pct_fmt = S.f(**S.base(font_size=10.5, bold=True, font_color=th.primary,
                           bg_color=th.card, align="center", valign="vcenter",
                           border=1, border_color=th.border, num_format="0%"))
    bar_fmt = S.f(**S.base(font_name=th.mono_font, font_size=10, bold=True,
                           font_color=th.primary_2, bg_color=th.card,
                           align="left", valign="vcenter", border=1,
                           border_color=th.border))
    data_cols = ["D", "E", "F", "G", "H", "I", "J", "K", "L"]
    src = ["I", "J", None, "K", "L", "M", "N", "O", "P"]
    for i in range(C.DATA_REC_ROWS):
        row = ROW_REC_FIRST + i
        drow = C.DATA_REC_FIRST + i
        rec = m.agg["recipients"][i]
        blank = not rec["name"]
        ws.set_row(r(row), 18)
        ws.merge_range(r(row), ci("B"), r(row), ci("C"), "", name_fmt)
        ws.write_formula(r(row), ci("B"), "=%s!$H$%d" % (bk.q("data"), drow),
                         name_fmt, rec["name"])
        for col, sc in zip(data_cols, src):
            guard = '=IF($B%d="","",%%s)' % row
            if sc is None:                       # remaining = budget - spent
                formula = '=IF(OR($B%d="",%s!$H$%d=""),"",$D%d-$E%d)' % (
                    row, bk.q("data"), drow, row, row)
                cached = "" if blank else rec["planned"] - rec["spent"]
                fmt = money_fmt
            elif col == "L":
                formula = guard % ("%s!$%s$%d" % (bk.q("data"), sc, drow))
                cached = "" if blank else rec["pct"]
                fmt = pct_fmt
            elif col in ("D", "E"):
                formula = guard % ("%s!$%s$%d" % (bk.q("data"), sc, drow))
                cached = "" if blank else rec["planned" if col == "D"
                                            else "spent"]
                fmt = money_fmt
            else:
                formula = guard % ("%s!$%s$%d" % (bk.q("data"), sc, drow))
                key = {"G": "gifts", "H": "bought", "I": "wrapped",
                       "J": "delivered", "K": "to_buy"}[col]
                cached = "" if blank else rec[key]
                fmt = num_fmt
            ws.write_formula(r(row), ci(col), formula, fmt,
                             cached if cached != "" else "")
            bk.stats["formulas"] += 1
        ws.write_formula(r(row), ci("M"),
                         '=IF($B%d="","",REPT("\u2588",ROUND(IFERROR($L%d,0)'
                         '*10,0))&REPT("\u2591",10-ROUND(IFERROR($L%d,0)*10,0)))'
                         % ((row,) * 3), bar_fmt,
                         _bar(rec["pct"] or 0, 10) if not blank else "")
        bk.stats["formulas"] += 1

    # totals
    total_fmt = S.f(**S.base(font_size=11, bold=True, font_color=th.white,
                             bg_color=th.primary, align="center",
                             valign="vcenter", border=1,
                             border_color=th.primary))
    money_total = S.f(**S.base(font_size=11, bold=True, font_color=th.white,
                               bg_color=th.primary, align="right",
                               valign="vcenter", border=1,
                               border_color=th.primary, num_format="#,##0"))
    bar_total = S.f(**S.base(font_name=th.mono_font, font_size=10, bold=True,
                             font_color=th.gold_soft, bg_color=th.primary,
                             align="center", valign="vcenter", border=1,
                             border_color=th.primary))
    ws.set_row(r(ROW_REC_TOTAL), 24)
    ws.merge_range(r(ROW_REC_TOTAL), 1, r(ROW_REC_TOTAL), ci("C"),
                   "  EVERYBODY", total_fmt)
    pct_total = S.f(**S.base(font_size=11, bold=True, font_color=th.white,
                             bg_color=th.primary, align="center",
                             valign="vcenter", border=1,
                             border_color=th.primary, num_format="0%"))
    for col, fmt in (("D", money_total), ("E", money_total),
                     ("F", money_total), ("G", total_fmt), ("H", total_fmt),
                     ("I", total_fmt), ("J", total_fmt), ("K", total_fmt),
                     ("L", pct_total), ("M", bar_total)):
        ws.write_formula(
            r(ROW_REC_TOTAL), ci(col),
            "=SUM(%s%d:%s%d)" % (col, ROW_REC_FIRST, col,
                                 ROW_REC_FIRST + C.DATA_REC_ROWS - 1)
            if col != "L" else
            '=IFERROR(%s/%s,0)' % (bk.kpi("gifts_purchased"),
                                   bk.kpi("gifts_planned"))
            if col != "M" else
            '=REPT("\u2588",ROUND(IFERROR(%s/%s,0)*10,0))&REPT("\u2591",'
            '10-ROUND(IFERROR(%s/%s,0)*10,0))'
            % ((bk.kpi("gifts_purchased"), bk.kpi("gifts_planned")) * 2),
            fmt,
            (sum(x["planned"] for x in m.agg["recipients"]) if col == "D" else
             sum(x["spent"] for x in m.agg["recipients"]) if col == "E" else
             sum(x["planned"] - x["spent"] for x in m.agg["recipients"])
             if col == "F" else
             m.agg["gifts_planned"] if col == "G" else
             m.agg["gifts_purchased"] if col == "H" else
             m.agg["gifts_wrapped"] if col == "I" else
             m.agg["gifts_delivered"] if col == "J" else
             m.agg["gifts_to_buy"] if col == "K" else
             m.agg["gift_completion"] if col == "L" else
             _bar(m.agg["gift_completion"] or 0, 10)))
        bk.stats["formulas"] += 1

    # hide the unused rows of the table
    bk.cond(KEY, ROW_REC_FIRST, 1, ROW_REC_FIRST + C.DATA_REC_ROWS - 1,
            ci(LAST_COL), {
                "type": "formula",
                "criteria": '=$B%d=""' % ROW_REC_FIRST,
                "format": S.cf(bg=th.card, fg=th.card, border=th.card)})
    bk.cond(KEY, ROW_REC_FIRST, ci("L"),
            ROW_REC_FIRST + C.DATA_REC_ROWS - 1, ci("L"), {
                "type": "data_bar", "bar_color": th.primary_2,
                "bar_solid": True, "min_type": "num", "min_value": 0,
                "max_type": "num", "max_value": 1})


# ===========================================================================
# formula fragments
# ===========================================================================
def _countdown(bk):
    days = bk.kpi("days_to_event")
    return ('IF({d}>1,"\u23F3  "&{d}&" Days Until "&EventName&"  \U0001F381",'
            'IF({d}=1,"\u23F3  Tomorrow is "&EventName&"!  \U0001F381",'
            'IF({d}=0,"\U0001F389  It\'s "&EventName&" today!  \U0001F384",'
            '"\U0001F384  "&EventName&" "&TEXT(EventDate,"yyyy")&" has been '
            'and gone \u2014 set a new date in \u2699\uFE0F Setup")))'
            ).format(d=days)


def _bar_formula(bk, numerator, denominator, blocks):
    pct = "MIN(1,IFERROR((%s)/(%s),0))" % (numerator, denominator)
    filled = "ROUND(%s*%d,0)" % (pct, blocks)
    return ('=IFERROR(REPT("\u2588",%s)&REPT("\u2591",%d-%s),"%s")'
            % (filled, blocks, filled, "\u2591" * blocks))


def _bar(pct, blocks):
    try:
        pct = min(1.0, max(0.0, float(pct)))
    except (TypeError, ValueError):
        pct = 0.0
    filled = int(round(pct * blocks))
    return "\u2588" * filled + "\u2591" * (blocks - filled)


def _today():
    from datetime import date
    return date.today()


def _today_text():
    return _today().strftime("%A %d %B %Y")


def _lit(text):
    """A quoted Excel string literal (with inner quotes escaped)."""
    return '"%s"' % str(text).replace('"', '""')


def _cat(*parts):
    return "&".join(parts)


def _iff(cond, then, other):
    return "IF(%s,%s,%s)" % (cond, then, other)


def _todo_lines(bk):
    """The 'what's left to do' sentences: (formula, cached text) x 10."""
    m = bk.demo
    agg = m.agg
    out = []

    def K(name):
        return bk.kpi(name)

    # 1. gifts still to buy
    tb, od = K("gifts_to_buy"), K("gifts_ordered")
    out.append((
        "=" + _iff(tb + "=0",
                   _lit("\u2705  Every gift is bought \u2014 nice work!"),
                   _cat(_lit("\U0001F6D2  "), tb,
                        _lit(" gift(s) still to buy"),
                        _iff(od + ">0",
                             _cat(_lit(" ("), od, _lit(" already on order)")),
                             _lit("")))),
        ("\u2705  Every gift is bought \u2014 nice work!"
         if agg["gifts_to_buy"] == 0 else
         "\U0001F6D2  %d gift(s) still to buy%s"
         % (agg["gifts_to_buy"],
            " (%d already on order)" % agg["gifts_ordered"]
            if agg["gifts_ordered"] else ""))))

    # 2. wrapping
    w = K("wrap_to_do")
    out.append((
        "=" + _iff(w + "=0", _lit("\u2705  Everything is wrapped"),
                   _cat(_lit("\U0001F380  "), w,
                        _lit(" present(s) still to wrap"))),
        ("\u2705  Everything is wrapped" if agg["wrap_to_do"] == 0
         else "\U0001F380  %d present(s) still to wrap" % agg["wrap_to_do"])))

    # 3. handing over
    d = K("deliver_to_do")
    out.append((
        "=" + _iff(d + "=0", _lit("\u2705  All wrapped gifts handed over"),
                   _cat(_lit("\U0001F4E6  "), d,
                        _lit(" wrapped gift(s) not given yet"))),
        ("\u2705  All wrapped gifts handed over" if agg["deliver_to_do"] == 0
         else "\U0001F4E6  %d wrapped gift(s) not given yet"
         % agg["deliver_to_do"])))

    # 4. budget
    p, rem, pct = K("budget_planned"), K("budget_remaining"), K("budget_pct")
    out.append((
        "=" + _iff(
            p + "=0",
            _lit("\U0001F4DD  Set your category budgets on the "
                 "\U0001F4B0 Budget tab"),
            _iff(rem + "<0",
                 _cat(_lit("\U0001F534  Over budget by "), "Currency",
                      "TEXT(-" + rem + ',"#,##0")'),
                 _iff(pct + ">=AlertAt",
                      _cat(_lit("\u26A0\uFE0F  "), "TEXT(" + pct + ',"0%")',
                           _lit(" of the budget used \u2014 "), "Currency",
                           "TEXT(" + rem + ',"#,##0")',
                           _lit(" left to play with")),
                      _cat(_lit("\U0001F4B0  "), "Currency",
                           "TEXT(" + rem + ',"#,##0")',
                           _lit(" left to spend ("), "TEXT(" + pct + ',"0%")',
                           _lit(" used)"))))),
        _budget_line(m)))

    # 5. shopping list
    t, b = K("shop_total"), K("shop_bought")
    left = agg["shop_total"] - agg["shop_bought"]
    out.append((
        "=" + _iff(t + "=0", _lit("\U0001F6CD\uFE0F  Your shopping list is "
                                 "empty"),
                   _iff(t + "=" + b,
                        _cat(_lit("\u2705  Shopping list complete ("), t,
                             _lit(" items)")),
                        _cat(_lit("\U0001F6CD\uFE0F  "), "(" + t + "-" + b
                             + ")", _lit(" of "), t,
                             _lit(" shopping items still to buy")))),
        ("\U0001F6CD\uFE0F  Your shopping list is empty"
         if agg["shop_total"] == 0 else
         "\u2705  Shopping list complete (%d items)" % agg["shop_total"]
         if left == 0 else
         "\U0001F6CD\uFE0F  %d of %d shopping items still to buy"
         % (left, agg["shop_total"]))))

    # 6. parcels
    if bk.has("orders"):
        t, l, o = K("orders_total"), K("orders_late"), K("orders_outstanding")
        out.append((
            "=" + _iff(t + "=0",
                       _lit("\U0001F4E6  No online orders yet \u2014 log "
                            "them on the \U0001F4E6 Order Tracker"),
                       _iff(l + ">0",
                            _cat(_lit("\U0001F534  "), l,
                                 _lit(" parcel(s) are late \u2014 chase them "
                                      "up")),
                            _cat(_lit("\U0001F69A  "), o,
                                 _lit(" parcel(s) still on the way")))),
            ("\U0001F4E6  No online orders yet" if agg["orders_total"] == 0
             else "\U0001F534  %d parcel(s) are late \u2014 chase them up"
             % agg["orders_late"] if agg["orders_late"]
             else "\U0001F69A  %d parcel(s) still on the way"
             % agg["orders_outstanding"])))

    # 7. cards
    if bk.has("cards"):
        t, s = K("cards_total"), K("cards_sent")
        left = agg["cards_total"] - agg["cards_sent"]
        out.append((
            "=" + _iff(t + "=0",
                       _lit("\U0001F48C  Add your card list to the "
                            "\U0001F48C Card Tracker"),
                       _iff(t + "=" + s,
                            _cat(_lit("\u2705  All "), t,
                                 _lit(" cards are posted")),
                            _cat(_lit("\U0001F48C  "), "(" + t + "-" + s + ")",
                                 _lit(" of "), t,
                                 _lit(" cards still to post")))),
            ("\U0001F48C  Add your card list" if agg["cards_total"] == 0
             else "\u2705  All %d cards are posted" % agg["cards_total"]
             if left == 0 else
             "\U0001F48C  %d of %d cards still to post"
             % (left, agg["cards_total"]))))

    # 8. stockings
    if bk.has("stockings"):
        t, b = K("stock_items"), K("stock_bought")
        left = agg["stock_items"] - agg["stock_bought"]
        out.append((
            "=" + _iff(t + "=0",
                       _lit("\U0001F9E6  Stockings not started \u2014 add "
                            "fillers on the \U0001F9E6 Stockings tab"),
                       _iff(t + "=" + b,
                            _cat(_lit("\u2705  All "), t,
                                 _lit(" stocking fillers bought")),
                            _cat(_lit("\U0001F9E6  "), "(" + t + "-" + b + ")",
                                 _lit(" of "), t,
                                 _lit(" stocking fillers still to buy")))),
            ("\U0001F9E6  Stockings not started" if agg["stock_items"] == 0
             else "\u2705  All %d stocking fillers bought" % agg["stock_items"]
             if left == 0 else
             "\U0001F9E6  %d of %d stocking fillers still to buy"
             % (left, agg["stock_items"]))))

    # 9. checklist
    if bk.has("todo"):
        t, o, dn = K("todo_total"), K("todo_overdue"), K("todo_done")
        out.append((
            "=" + _iff(t + "=0", _lit("\u2705  No tasks on the checklist"),
                       _iff(o + ">0",
                            _cat(_lit("\U0001F534  "), o,
                                 _lit(" overdue task(s) on the \u2705 To-Do "
                                      "list")),
                            _cat(_lit("\u2705  "), dn, _lit(" of "), t,
                                 _lit(" checklist tasks done")))),
            ("\u2705  No tasks on the checklist" if agg["todo_total"] == 0
             else "\U0001F534  %d overdue task(s) on the \u2705 To-Do list"
             % agg["todo_overdue"] if agg["todo_overdue"]
             else "\u2705  %d of %d checklist tasks done"
             % (agg["todo_done"], agg["todo_total"]))))

    # 10. wish list
    if bk.has("wishlist"):
        t, mu = K("wish_total"), K("wish_must")
        out.append((
            "=" + _iff(t + "=0",
                       _lit("\U0001F4A1  Save gift ideas on the \U0001F4A1 "
                            "Wish List all year round"),
                       _cat(_lit("\U0001F4A1  "), t, _lit(" idea(s) saved "
                                                           "\u2014 "), mu,
                            _lit(" of them must-haves"))),
            ("\U0001F4A1  Save gift ideas on the \U0001F4A1 Wish List"
             if agg["wish_total"] == 0 else
             "\U0001F4A1  %d idea(s) saved \u2014 %d of them must-haves"
             % (agg["wish_total"], agg["wish_must"]))))

    tips = [
        "\U0001F4A1  Tip: sort the \U0001F381 Gift Tracker by \u201CStore"
        "\u201D and you shop the whole list in one trip.",
        "\U0001F4A1  Tip: the \U0001F380 Wrapping tab remembers where every "
        "present is hiding.",
        "\U0001F4A1  Tip: change the event date in \u2699\uFE0F Setup and "
        "every deadline re-dates itself.",
        "\U0001F4A1  Tip: use this file for birthdays too \u2014 pick an "
        "occasion in \u2699\uFE0F Setup.",
        "\U0001F4A1  Tip: \U0001F648 Secret Mode hides the hiding spots when "
        "somebody is watching.",
    ]
    for tip in tips:
        if len(out) >= 10:
            break
        out.append(("=" + _lit(tip), tip))
    while len(out) < 10:
        out.append(('=""', ""))
    return out[:10]


def _budget_line(m):
    agg, cur = m.agg, m.settings["currency"]
    if agg["budget_planned"] == 0:
        return ("\U0001F4DD  Set your category budgets on the \U0001F4B0 "
                "Budget tab")
    if agg["budget_remaining"] < 0:
        return "\U0001F534  Over budget by %s" % m.money(-agg["budget_remaining"])
    if agg["budget_pct"] >= m.settings["alert"]:
        return ("\u26A0\uFE0F  %d%% of the budget used \u2014 %s left to play "
                "with" % (round(agg["budget_pct"] * 100),
                          m.money(agg["budget_remaining"])))
    return ("\U0001F4B0  %s left to spend (%d%% used)"
            % (m.money(agg["budget_remaining"]),
               round(agg["budget_pct"] * 100)))
```

### `christmas_tracker/sheets/guide.py`  (430 lines)

```python
"""
\U0001F4D6 Start Here - the cover page and instruction manual.

Built as a sequence of content blocks (headings, paragraphs, bullet lists and
tables) so the copy can be edited without touching row numbers.  A watercolour
cover banner is inserted when ``assets/`` contains the image for the active
theme.
"""

import os
import struct

from .. import config as C
from ..book import r, ci, cl
from ..styles import wrap_height

KEY = "guide"
LAST_COL = "L"
CONTENT = ("C", "L")          # paragraph span
TABLE_SPANS = {
    "tabs": [("C", "E"), ("F", "H"), ("I", "J"), ("K", "L")],
    "legend": [("C", "D"), ("E", "G"), ("H", "L")],
    "auto": [("C", "E"), ("F", "L")],
}


def build(bk):
    ws = bk.ws(KEY)
    S, th, m = bk.S, bk.th, bk.demo

    ws.set_column("A:A", 2.2)
    ws.set_column("B:B", 3)
    ws.set_column("C:L", 12)
    width = 120

    body = S.guide_text(size=10.5, bg=th.card)
    body_noborder = S.f(**S.base(font_size=10.5, font_color=th.ink,
                                 bg_color=th.card, align="left",
                                 valign="top", text_wrap=True, indent=1))
    bullet = S.f(**S.base(font_size=10.5, font_color=th.ink, bg_color=th.card,
                          align="left", valign="top", text_wrap=True,
                          border=1, border_color=th.border, indent=1))
    step_no = S.f(**S.base(font_size=14, bold=True, font_color=th.white,
                           bg_color=th.accent, align="center",
                           valign="vcenter", border=1, border_color=th.accent))

    row = 1
    ws.set_row(r(row), 7)
    row += 1

    # ------------------------------------------------------------------
    # cover banner + overlaid title
    # ------------------------------------------------------------------
    image = _banner_path(bk)
    span = 0
    if image:
        w, h = _png_size(image)
        target = 860
        scale = target / float(w)
        height = h * scale
        span = int(height / 20.0) + 1
        bg = th.cover_bg                # the cells under the see-through
        cover = S.f(**S.base(bg_color=bg))   # middle must match the artwork
        for i in range(span):
            ws.set_row(r(row + i), 20)
            for col in range(1, ci(LAST_COL) + 1):
                ws.write(r(row + i), col, "", cover)
        ws.insert_image(r(row), 1, image,
                        {"x_scale": scale, "y_scale": scale,
                         "object_position": 1})
        title_row = row + max(1, span // 2 - 1)
        ws.merge_range(r(title_row), ci("C"), r(title_row), ci("H"),
                       "\U0001F4D6  START HERE",
                       S.f(**S.base(font_name=th.title_font, font_size=28,
                                    bold=True, font_color=th.primary,
                                    bg_color=bg, align="center",
                                    valign="vcenter")))
        ws.merge_range(r(title_row + 1), ci("C"), r(title_row + 1), ci("H"),
                       "  your %s-minute tour of the %s"
                       % ("3", C.PRODUCT_SHORT),
                       S.f(**S.base(font_size=12, italic=True,
                                    font_color=th.muted, bg_color=bg,
                                    align="center", valign="vcenter")))
        row += span + 1
    else:
        ws.set_row(r(row), 40)
        ws.merge_range(r(row), ci("B"), r(row), ci(LAST_COL),
                       "  \U0001F4D6  START HERE", S.sheet_title)
        row += 1

    ws.set_row(r(row), 6)
    row += 1

    def para(text, style=None, span_cols=CONTENT, minimum=18):
        nonlocal row
        c1, c2 = span_cols
        w = sum(12 for c in range(ci(c1), ci(c2) + 1))
        h = wrap_height(text, w, minimum=minimum)
        ws.merge_range(r(row), ci(c1), r(row), ci(c2), text,
                       style or body)
        ws.set_row(r(row), h)
        row += 1
        return row

    def section(text, style=None):
        nonlocal row
        ws.set_row(r(row), 24)
        ws.merge_range(r(row), ci("B"), r(row), ci(LAST_COL), text,
                       style or S.section)
        row += 1

    def bullets(items, prefix="\u2022  "):
        nonlocal row
        for item in items:
            para(prefix + item, bullet)

    def spacer(h=8):
        nonlocal row
        ws.set_row(r(row), h)
        row += 1

    def table(kind, headers, rows, header_color=None):
        nonlocal row
        spans = TABLE_SPANS[kind]
        ws.set_row(r(row), 24)
        for (c1, c2), label in zip(spans, headers):
            ws.merge_range(r(row), ci(c1), r(row), ci(c2), label,
                           S.header(header_color or th.primary))
        row += 1
        for i, cells in enumerate(rows):
            h = 18
            for (c1, c2), text in zip(spans, cells):
                w = sum(12 for c in range(ci(c1), ci(c2) + 1))
                h = max(h, wrap_height(text, w, minimum=18))
                fmt = S.f(**S.base(font_size=9.5, font_color=th.ink,
                                   bg_color=th.alt if i % 2 else th.card,
                                   align="left", valign="top", text_wrap=True,
                                   border=1, border_color=th.border,
                                   indent=1))
                ws.merge_range(r(row), ci(c1), r(row), ci(c2), text, fmt)
            ws.set_row(r(row), h)
            row += 1

    def steps(items):
        nonlocal row
        for i, (title, text) in enumerate(items, 1):
            ws.merge_range(r(row), ci("B"), r(row + 1), ci("B"), str(i),
                           step_no)
            w = sum(12 for c in range(ci("C"), ci(LAST_COL) + 1))
            h = wrap_height(text, w, minimum=20)
            ws.merge_range(r(row), ci("C"), r(row), ci(LAST_COL),
                           title, S.f(**S.base(font_size=11, bold=True,
                                               font_color=th.primary,
                                               bg_color=th.primary_soft,
                                               align="left", valign="vcenter",
                                               border=1, border_color=th.border,
                                               indent=1)))
            ws.merge_range(r(row + 1), ci("C"), r(row + 1), ci(LAST_COL),
                           text, bullet)
            ws.set_row(r(row), 20)
            ws.set_row(r(row + 1), h)
            row += 2

    # ------------------------------------------------------------------
    # 1. quick start
    # ------------------------------------------------------------------
    section("  \u26A1  THREE-MINUTE QUICK START")
    spacer(6)
    steps([
        ("Set your event.", "\u2699\uFE0F Setup \u2192 Event name and Event "
         "date. Every countdown, deadline and key date in the workbook is "
         "calculated from that one date, so nothing is locked to a year."),
        ("Make the money yours.", "On the same tab pick your currency symbol, "
         "your total budget target, the % that triggers the budget warning, "
         "and how many days counts as \u201Cdue soon\u201D."),
        ("Type your people.", "\u2699\uFE0F Setup \u2192 Your lists: recipients, "
         "relationships, stores, categories, hiding spots. Every dropdown in "
         "every tab reads from these lists and updates instantly."),
        ("Add your gifts.", "\U0001F381 Gift Tracker: one row per present. "
         "Use the Status dropdown as your pipeline (\U0001F4A1 idea \u2192 "
         "\U0001F4E6 delivered). The tinted columns (difference, % of budget, "
         "days left, alert) fill themselves in."),
        ("Plan the money.", "\U0001F4B0 Budget: type a planned amount per "
         "category. The \u201Cpulled in automatically\u201D column adds up the "
         "real money from the gift, shopping, stocking and postage columns."),
        ("Enjoy the dashboard.", "\U0001F384 Dashboard needs no typing at all. "
         "It recalculates the moment you change anything anywhere else."),
    ])
    spacer()

    # ------------------------------------------------------------------
    # 2. tab guide
    # ------------------------------------------------------------------
    section("  \U0001F5C2\uFE0F  WHAT EVERY TAB DOES", style=S.section_accent)
    spacer(6)
    tab_rows = [
        ["\U0001F384 Dashboard", "The command centre: countdown, money, gift "
         "progress, bars, 4 charts, deadlines.", "nothing", "everything"],
        ["\U0001F381 Gift Tracker", "One row per present \u2014 the heart of "
         "the workbook.", "the gift", "difference, %, alerts"],
        ["\U0001F4B0 Budget", "Planned vs actual per category with a warning "
         "level you choose.", "planned amounts", "actuals + charts"],
    ]
    if bk.edition == "premium":
        tab_rows += [
            ["\U0001F4A1 Wish List", "Ideas parked all year, ranked by "
             "priority; detects ideas already on the gift list.",
             "ideas + priority", "in-tracker check"],
            ["\U0001F6CD\uFE0F Shopping List", "Everything that is not a "
             "present: wrapping, cards, baking, decorations.", "the items",
             "totals + budget feed"],
            ["\U0001F4E6 Order Tracker", "Parcels: order numbers, expected "
             "dates, tracking links, late alerts.", "order details",
             "days left + alerts"],
            ["\U0001F380 Wrapping & Hiding", "Mirrors the gift list and adds "
             "hiding spots + gift tags (with Secret Mode).", "hiding spot",
             "mirror of gifts"],
            ["\U0001F48C Card Tracker", "Bought \u2192 written \u2192 posted "
             "\u2192 replied, plus postage costs.", "names + ticks",
             "status + totals"],
            ["\U0001F9E6 Stockings", "Fillers per stocking with a budget and a "
             "per-stocking summary.", "items + costs", "budget vs spent"],
            ["\u2705 To-Do List", "Date-aware checklist; deadlines calculated "
             "from your event date.", "tick when done", "status + overdue"],
        ]
    tab_rows += [
        ["\u2699\uFE0F Setup", "Event, money settings, and every editable "
         "list in the workbook.", "settings + lists", "suggestions"],
        ["\U0001F4D6 Start Here", "This page: tour, legend, automation list, "
         "Google Sheets help, FAQ.", "\u2014", "\u2014"],
        ["_Data (hidden)", "The engine room: every dashboard number, chart "
         "series and deadline is computed here.", "never", "all the maths"],
    ]
    table("tabs", ["Tab", "What it's for", "You type", "Automatic"], tab_rows)
    spacer()

    # ------------------------------------------------------------------
    # 3. legend
    # ------------------------------------------------------------------
    section("  \U0001F3A8  COLOUR & SYMBOL LEGEND", style=S.section_gold)
    spacer(6)
    table("legend", ["Looks like", "Means", "Where you'll see it"], [
        ["\U0001F534 red", "Overdue, over budget, or running late",
         "deadline columns, budget status, order alerts, to-do status"],
        ["\U0001F7E0 amber", "Due soon (inside your Setup window) or nearly "
         "at the budget limit", "deadlines, % used, budget status"],
        ["\U0001F7E2 green", "On track / complete / ticked off",
         "tick boxes, statuses, to-do list, wrapped rows"],
        ["\U0001F535 blue", "Ordered / on the way",
         "gift status, order status"],
        ["\U0001F380 gold", "Wrapped, or a priority must-have",
         "gift status, wrapping tab, wish list"],
        ["\u26AA grey", "Not started yet", "budget rows, card status"],
        ["\u2713 tick", "Pick \u2713 from the dropdown to tick, pick the blank "
         "option to untick",
         "wrapped, given, bought, sent, done, returned"],
        ["\u2588\u2591 bars", "Text progress bars (screenshot-friendly and "
         "print-friendly)", "dashboard, budget, stockings"],
    ])
    spacer()

    # ------------------------------------------------------------------
    # 4. automation list
    # ------------------------------------------------------------------
    section("  \U0001F916  EVERYTHING THAT IS AUTOMATIC",
            style=S.section_accent)
    spacer(6)
    auto = [
        ["Countdown", "Days and weeks until the event, recalculated every "
         "time the file opens."],
        ["Money", "Total budget, spent, remaining, % used, average spend per "
         "person."],
        ["Budget alert", "A plain-English warning at the % you chose in "
         "Setup (\u201Cyou've spent 87% of your Christmas budget\u201D)."],
        ["Gift counts", "Planned, purchased, wrapped, delivered, still to "
         "buy, on order, completion %."],
        ["Per person", "Budget, spend, gift counts and a progress bar for "
         "every recipient."],
        ["Per gift", "Under/over budget, % of budget used, days left, "
         "deadline alert."],
        ["Buttons", "Any pasted link becomes a clickable \U0001F517 button."],
        ["Shopping", "Qty \u00D7 unit = total; tick Bought and the actual (or "
         "estimated) cost flows into the budget."],
        ["Budget feed", "Gift, shopping, stocking and postage money is pulled "
         "into the Budget tab by category."],
        ["Stockings", "Budget, spent, remaining and a progress bar for each "
         "stocking."],
        ["Cards", "Status from the ticks; \u201Ccards sent x / y\u201D; "
         "postage into the budget."],
        ["Orders", "Days until expected; late / arriving-soon / on-the-way "
         "alerts."],
        ["Wish list", "Tells you whether an idea is already on the gift list."],
        ["To-do", "Deadlines from the event date; status from the date; "
         "overdue counts."],
        ["Wrapping", "Mirrors the gift list; Secret Mode blanks the hiding "
         "spots."],
        ["Dashboard", "Four charts, the next five deadlines across every tab, "
         "key dates and the \u201Cwhat's left to do\u201D sentences."],
    ]
    table("auto", ["Feature", "What it does"], auto, header_color=th.accent)
    spacer()

    # ------------------------------------------------------------------
    # 5. google sheets + versions
    # ------------------------------------------------------------------
    section("  \U0001F4E5  USING IT IN GOOGLE SHEETS (AND WHICH EXCEL)")
    spacer(6)
    bullets([
        "Google Sheets: sheets.new \u2192 File \u2192 Import \u2192 Upload "
        "\u2192 choose this .xlsx \u2192 Import location: Replace spreadsheet. "
        "Everything (formulas, dropdowns, colours, charts) converts.",
        "After import, click any tab once so Sheets recalculates; then File "
        "\u2192 Save as Google Sheets to keep your copy in Drive.",
        "Excel: 2016, 2019, 2021 and Microsoft 365 on Windows or Mac. The "
        "file contains no macros, so there are no security warnings and it "
        "opens straight away.",
        "Phone / tablet: opens in the Excel and Google Sheets apps \u2014 "
        "fine for checking the dashboard and ticking things off; do the big "
        "data entry on a computer.",
        "If a number ever looks stale in Excel: Formulas \u2192 Calculation "
        "Options \u2192 Automatic, or press F9.",
    ])
    spacer()

    # ------------------------------------------------------------------
    # 6. FAQ
    # ------------------------------------------------------------------
    section("  \u2753  FREQUENTLY ASKED", style=S.section_gold)
    spacer(6)
    faq = [
        ["How do I add more rows?", "Select the last data row and drag the "
         "little square in its bottom-right corner down (or right-click "
         "\u2192 Insert above the TOTALS row). Formulas, dropdowns and "
         "colours come with the row."],
        ["Can I use it for birthdays?", "Yes \u2014 \u2699\uFE0F Setup \u2192 "
         "pick an occasion or type any event name and date. The dashboard "
         "re-words itself around it."],
        ["How do I change currency?", "\u2699\uFE0F Setup \u2192 Currency "
         "symbol changes every dashboard and summary readout. To show it "
         "inside the tables too: select the money columns \u2192 Home \u2192 "
         "Number format \u2192 Currency."],
        ["Why are the checkboxes dropdowns?", "Because \u2713 dropdowns work "
         "identically in Excel and Google Sheets, whereas Excel's built-in "
         "checkbox controls don't survive conversion. Pick \u2713 to tick, "
         "the blank option to untick."],
        ["Can I delete tabs I don't use?", "Please keep _Data (hidden) \u2014 "
         "the dashboard is computed there. For the rest: leaving a tab empty "
         "is safer than deleting it, but deleting premium tabs won't break "
         "the rest as long as _Data stays."],
        ["How do I hide the hiding spots?", "\u2699\uFE0F Setup \u2192 Secret "
         "Mode = Yes turns them white-on-white instantly. For total privacy, "
         "right-click the column \u2192 Hide."],
        ["Will it print nicely?", "Yes \u2014 every tab is set up landscape, "
         "fitted to one page wide, with the title and column headers repeated "
         "on each page."],
        ["Can we share it in the family?", "Save it to OneDrive, Dropbox or "
         "Drive and share it like any workbook; or import it to Google "
         "Sheets and share a link so everyone ticks together."],
        ["Is it password protected?", "Not by default \u2014 you can edit "
         "everything. If you want it locked down: Review \u2192 Protect "
         "Sheet / Protect Workbook in Excel."],
        ["I changed the event date \u2014 will deadlines move?", "The to-do "
         "list and the dashboard key dates move automatically. Dates you "
         "typed yourself on gift rows and orders stay where you put them."],
    ]
    table("auto", ["Question", "Answer"], faq, header_color=th.gold)
    spacer()

    # ------------------------------------------------------------------
    # 7. licence
    # ------------------------------------------------------------------
    section("  \U0001F4C4  LICENCE & SUPPORT", style=S.section_soft)
    spacer(6)
    para("Personal licence: use it for your own household, every year, for "
         "every occasion. Please don't resell, redistribute or share the file "
         "itself \u2014 send people here instead. If something looks wrong or "
         "you'd love a feature added, message the shop: small improvements "
         "usually ship within days, free, to everyone who already bought it.",
         S.f(**S.base(font_size=10, italic=True, font_color=th.muted,
                      bg_color=th.alt, align="left", valign="top",
                      text_wrap=True, border=1, border_color=th.border,
                      indent=1)))
    spacer()
    ws.set_row(r(row), 26)
    row = bk.nav_row(KEY, row, first_col=1, span=2,
                     max_col=LAST_COL) + 2
    ws.merge_range(r(row), ci("B"), r(row), ci(LAST_COL),
                   "  %s v%s  \u2022  %s  \u2022  made with love (and "
                   "formulas)  \u2022  %s"
                   % (C.PRODUCT, C.VERSION, C.TAGLINE, C.AUTHOR), S.footer)
    bk.page(KEY, LAST_COL, row + 2, landscape=False, zoom=100)
    return ws


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
        data = fh.read(4096)
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        w, h = struct.unpack(">II", data[16:24])
        return w, h
    if data[:2] == b"\xff\xd8":
        i = 2
        while i < len(data) - 9:
            if data[i] != 0xFF:
                i += 1
                continue
            marker = data[i + 1]
            if marker in (0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7,
                          0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF):
                h, w = struct.unpack(">HH", data[i + 5:i + 9])
                return w, h
            length = struct.unpack(">H", data[i + 2:i + 4])[0]
            i += 2 + length
    return 1600, 640
```

_(The quality tools in `tools/` - `verify_workbook.py`,
`calc_check.py`, `layout_check.py`, `render_preview.py` and
`make_banner_alpha.py` - are development aids, not part of the
product; they live in the repository alongside this document.)_


```bash
/tmp/venv/bin/python tools/verify_workbook.py products/*.xlsx   # structure
/tmp/venv/bin/python tools/calc_check.py   products/*.xlsx      # recalculation
/tmp/venv/bin/python tools/layout_check.py products/*.xlsx      # clipping
/tmp/venv/bin/python tools/render_preview.py products/X.xlsx "🎄 Dashboard" /tmp/d.png
```

`calc_check.py` uses the open-source `formulas` engine to evaluate all ~5,400
cells and fails the build on any `#REF!/#NAME?/#VALUE!` or on any cached value
that disagrees with the live recalculation. Known engine limitations (OFFSET
named ranges, `HYPERLINK`) are reported separately, not as failures.
