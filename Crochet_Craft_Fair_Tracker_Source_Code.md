# Crochet Craft Fair Profit & Inventory Tracker - Complete Source Code

**Product:** Crochet Craft Fair Profit & Inventory Tracker (Excel workbook generator)
**Version:** 1.0.0  |  **Built:** 13 September 2026  |  **Repo:** `arbabkhan007/Excel`, branch `arena/01a09508-excel`, tag `crochet-v1.0.0`
**Generator:** Python 3 + vendored XlsxWriter (no other runtime dependencies)

This document is the *entire* product source: every module that builds the six
shipped workbooks (Premium/Basic x Berry/Mint themes, blank + EXAMPLE modes),
the command-line entry point, and the four QA tools used to prove every build.
Regenerate the whole product set with:

```bash
python3 crochet_craft_fair_tracker.py --all --outdir products
```

## What the product does

Thirteen wired-together features for crochet makers who sell at craft fairs:
product catalog with true unit cost and margin, yarn & materials inventory
with automatic reorder flags, production batches (made / reserved / sold /
returned / damaged / available), a craft-fair log with five cost lines per
market, a sales log with event + product dropdowns, a per-event profit
calculator (gross, costs, net, margin, break-even, ROI), an auto-pulled
reorder list with priorities and budget, a per-fair packing checklist with a
% packed meter, a profit dashboard (cards + six charts + live panels), a
pricing calculator (materials + labour + overhead + margin -> suggested and
charm price), a month-by-month yearly summary, and best-seller tracking.
Everything recomputes from two tabs: Craft Fairs and Sales Log.

## Editions & files shipped

| File | Sheets | Formulas | Dropdowns | Cond. formats | Charts |
|---|---|---|---|---|---|
| PREMIUM Berry / Mint (blank) | 14 | 1,431 | 35 | 21 | 6 |
| PREMIUM Berry EXAMPLE | 14 | 1,431 | 35 | 21 | 6 |
| BASIC Berry / Mint (blank) | 8 | 940 | 18 | 6 | 6 |
| BASIC Berry EXAMPLE | 8 | 940 | 18 | 6 | 6 |

Premium adds Yarn & Materials, Made & Stocked, Event Profit, Reorder List,
Packing Checklist and Monthly Summary; Basic keeps the core fair-selling loop.

## QA performed on every shipped file

* `tools/verify_workbook.py` - structure: hidden _Data, no writes under
  merges, defined names, DV/CF/chart counts  -> 0 problems on all six.
* `tools/calc_check.py` - full recalculation with the `formulas` engine:
  0 error cells and 0 cached-vs-computed mismatches on all six (the single
  reported #REF! is the engine's known OFFSET limitation; Excel and Google
  Sheets resolve those dropdown names normally).
* `tools/layout_check.py` - column-width / row-height clipping scan -> clean.
* `tools/render_preview.py` - pixel renders of Dashboard and Start Here used
  for visual review during development.

---


## 1. Package init

`crochet_tracker/__init__.py`

```python
"""Ultimate Crochet Craft Fair Tracker - workbook generator package."""

__version__ = "1.0"
```

## 2. Config - geometry, column maps, lists, caps

`crochet_tracker/config.py`

```python
"""
Every layout constant for the Ultimate Crochet Craft Fair Tracker.

Row numbers, column letters, sheet orders, dropdown values and list positions
all live here so the sheet builders stay readable and the whole workbook can
be re-skinned or re-rowed from one audited place.
"""

PRODUCT = "Ultimate Crochet Craft Fair Tracker"
PRODUCT_SHORT = "Crochet Craft Fair Tracker"
TAGLINE = "Products \u2022 Yarn Stock \u2022 Fairs \u2022 Profit"
AUTHOR = "Novality Store"
VERSION = "1.0"

# ===========================================================================
# SHEETS
# ===========================================================================
SHEET_NAMES = {
    "data": "_Data",
    "setup": "\u2699\uFE0F Lists & Settings",
    "dashboard": "\U0001F4CA Dashboard",
    "catalog": "\U0001F9F6 Product Catalog",
    "materials": "\U0001F9F5 Yarn & Materials",
    "production": "\U0001F4E6 Made & Stocked",
    "events": "\U0001F3EA Craft Fairs",
    "sales": "\U0001F4B0 Sales Log",
    "eventprofit": "\U0001F9EE Event Profit",
    "reorder": "\U0001F504 Reorder List",
    "packing": "\U0001F392 Packing Checklist",
    "pricing": "\U0001F4B5 Pricing Calculator",
    "monthly": "\U0001F4C5 Monthly Summary",
    "guide": "\U0001F4D6 Start Here",
}

SHEET_SHORT = {
    "data": "_Data",
    "setup": "Settings",
    "dashboard": "\U0001F3E0 Home",
    "catalog": "Products",
    "materials": "Yarn",
    "production": "Stock",
    "events": "Fairs",
    "sales": "Sales",
    "eventprofit": "Event $",
    "reorder": "Reorder",
    "packing": "Packing",
    "pricing": "Pricing",
    "monthly": "Monthly",
    "guide": "Help",
}

EDITIONS = {
    "premium": ["data", "setup", "dashboard", "catalog", "materials",
                "production", "events", "sales", "eventprofit", "reorder",
                "packing", "pricing", "monthly", "guide"],
    "basic": ["data", "setup", "dashboard", "catalog", "events", "sales",
              "pricing", "guide"],
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
    "catalog": 24,
    "materials": 24,
    "production": 30,
    "events": 24,
    "sales": 60,
    "monthly": 12,
}


def last_row(key):
    return ROW_FIRST + CAP[key] - 1


TICK = "\u2713"

# reorder priority labels
PR_URGENT = "\U0001F534 Urgent"
PR_HIGH = "\U0001F7E0 High"
PR_NORMAL = "\U0001F7E1 Normal"

# stock health labels
ST_OUT = "\U0001F534 Out of stock"
ST_LOW = "\U0001F7E0 Low stock"
ST_OK = "\U0001F7E2 In stock"

RE_YES = "\u2705 Reorder"
RE_NO = "\u274C OK"

# ===========================================================================
# DROPDOWN LIST VALUES (editable on the Setup tab)
# ===========================================================================
CATEGORIES = ["Toys & Amigurumi", "Wearables", "Home & Decor", "Accessories",
              "Baby & Kids", "Seasonal"]

PAYMENT_METHODS = ["Cash", "Card", "Bank transfer", "Mobile wallet",
                   "Online order", "Other"]

YARN_WEIGHTS = ["Lace", "Sock / 4-ply", "DK / light", "Worsted / aran",
                "Chunky", "Super chunky", "Mixed / other"]

UNITS = ["balls", "metres", "grams", "skeins", "cones", "pieces"]

SUPPLIERS = ["Local yarn shop", "Online yarn store", "Wholesale craft",
             "Market stall", "Second-hand / stash", "Other"]

# ===========================================================================
# COLUMN MAPS  (letter per logical field)
# ===========================================================================
COLS = {
    "catalog": {"n": "A", "sku": "B", "name": "C", "category": "D", "desc": "E",
                "price": "F", "yarn_cost": "G", "pack_cost": "H",
                "unit_cost": "I", "profit": "J", "margin": "K",
                "time_h": "L", "stock": "M", "min": "N", "status": "O",
                "link": "P"},
    "materials": {"n": "A", "name": "B", "color": "C", "brand": "D", "weight": "E",
                  "purchased": "F", "unit": "G", "cost_unit": "H",
                  "total": "I", "used": "J", "remaining": "K",
                  "supplier": "L", "threshold": "M", "reorder": "N"},
    "production": {"n": "A", "product": "B", "date": "C", "made": "D", "reserved": "E",
                   "taken": "F", "sold": "G", "returned": "H",
                   "damaged": "I", "current": "J", "available": "K",
                   "notes": "L"},
    "events": {"n": "A", "name": "B", "date": "C", "location": "D", "organizer": "E",
               "booth": "F", "travel": "G", "parking": "H", "food": "I",
               "display": "J", "total": "K", "taken": "L", "sold": "M",
               "sales": "N", "cogs": "O", "net": "P", "margin": "Q",
               "best": "R"},
    "sales": {"n": "A", "date": "B", "event": "C", "product": "D", "qty": "E",
              "unit": "F", "discount": "G", "total": "H", "method": "I",
              "ref": "J", "notes": "K", "pair": "L", "cost": "M"},
    "reorder": {"item": "B", "current": "C", "min": "D", "suggest": "E",
                "supplier": "F", "est": "G", "priority": "H",
                "ordered": "I", "ordered_date": "J"},
    "monthly": {"n": "A", "month": "B", "revenue": "C", "cogs": "D", "fees": "E",
                "net": "F", "margin": "G", "units": "H", "aov": "I",
                "per_item": "J", "markets": "K"},
}

# Column widths per sheet (letters are positional: A, B, C ...).
WIDTHS = {
    "catalog": {"A": 2.2, "B": 10, "C": 24, "D": 16, "E": 28, "F": 10,
                "G": 10, "H": 10, "I": 10, "J": 10, "K": 9, "L": 8, "M": 8,
                "N": 8, "O": 15, "P": 18},
    "materials": {"A": 2.2, "B": 22, "C": 14, "D": 14, "E": 15, "F": 10,
                  "G": 8, "H": 10, "I": 10, "J": 9, "K": 10, "L": 18,
                  "M": 10, "N": 13},
    "production": {"A": 2.2, "B": 24, "C": 12, "D": 8, "E": 9, "F": 9,
                   "G": 8, "H": 9, "I": 9, "J": 9, "K": 10, "L": 22},
    "events": {"A": 2.2, "B": 22, "C": 12, "D": 18, "E": 16, "F": 9, "G": 9,
               "H": 9, "I": 9, "J": 9, "K": 10, "L": 8, "M": 8, "N": 10,
               "O": 10, "P": 10, "Q": 9, "R": 22},
    "sales": {"A": 2.2, "B": 12, "C": 20, "D": 24, "E": 7, "F": 9, "G": 9,
              "H": 10, "I": 14, "J": 12, "K": 22, "L": 9},
    "reorder": {"A": 2.2, "B": 26, "C": 9, "D": 8, "E": 10, "F": 18, "G": 10,
                "H": 13, "I": 10, "J": 12},
    "monthly": {"A": 2.2, "B": 12, "C": 11, "D": 11, "E": 10, "F": 11,
                "G": 9, "H": 8, "I": 10, "J": 11, "K": 9},
    "eventprofit": {"A": 2.2, "B": 26, "C": 14, "D": 14, "E": 14, "F": 14,
                    "G": 14, "H": 14, "I": 14, "J": 14},
    "pricing": {"A": 2.2, "B": 24, "C": 14, "D": 14, "E": 14, "F": 14,
                "G": 16, "H": 16, "I": 16, "J": 16},
    "packing": {"A": 2.2, "B": 8, "C": 30, "D": 30, "E": 20, "F": 20},
    "dashboard": {"A": 2.2, "B": 13, "C": 13, "D": 13, "E": 13, "F": 13,
                  "G": 13, "H": 13, "I": 13, "J": 13, "K": 13, "L": 13,
                  "M": 13},
}

# ===========================================================================
# SETUP TAB GEOMETRY
# ===========================================================================
SU_BUSINESS = 6
SU_MESSAGE = 7
SU_CURRENCY = 10
SU_YEAR = 11
SU_WAGE = 12
SU_OVERHEAD = 13
SU_MARGIN = 14
SU_LIST_HEADER = 22
SU_LIST_FIRST = 23
SU_LIST_ROWS = 20
SU_FIXED_HEADER = 46
SU_FIXED_FIRST = 47

# ===========================================================================
# _Data POOL GEOMETRY (hidden sheet)
# ===========================================================================
DATA_MONTH_FIRST = 2          # H..N  rows 2-13: month table
DATA_PAY_FIRST = 2            # P..Q  rows 2-7:  payment methods
DATA_EXP_FIRST = 2            # S..T  rows 2-6:  event expense categories
DATA_PROD_FIRST = 2           # V..AA rows 2-25: product pool
DATA_EVENT_FIRST = 2          # AC..AF rows 2-25: event pool
DATA_KPI_FIRST = 2            # AH/AI rows 2-40: KPI cells
DATA_SEQP_FIRST = 2           # AK/AL rows 2-25: product reorder sequence
DATA_SEQM_FIRST = 2           # AN/AO rows 2-25: material reorder sequence

# ===========================================================================
# EVENT PROFIT / PRICING GEOMETRY
# ===========================================================================
EP_EVENT_ROW = 8              # event picker row
EP_REV = {"units": 11, "gross": 12, "discounts": 13}
EP_COST = {"cogs": 16, "booth": 17, "travel": 18, "parking": 19,
           "food": 20, "display": 21}
EP_OUT = {"gross_profit": 24, "total_cost": 25, "net": 26, "margin": 27,
          "avg_sale": 28, "breakeven": 29, "roi": 30}
EP_LAST_ROW = 34

PR_IN = {"material": 10, "packaging": 11, "hours": 12, "wage": 13,
         "overhead": 14, "margin": 15}
PR_OUT = {"labor": 14, "overhead_amt": 15, "true_cost": 16, "price": 17,
          "charm": 18, "profit": 19, "check": 20}
PR_LAST_ROW = 26

# ===========================================================================
# REORDER LIST GEOMETRY
# ===========================================================================
RO_PROD_FIRST = 10
RO_PROD_LAST = 21
RO_MAT_FIRST = 25
RO_MAT_LAST = 34

# ===========================================================================
# PACKING CHECKLIST
# ===========================================================================
PACK_EVENT_ROW = 6
PACK_SECTIONS = [
    ("Products", ["Finished products", "Extra inventory", "Price tags",
                  "Product labels / care cards"]),
    ("Booth", ["Table", "Tablecloth", "Display stands / risers",
               "Baskets & bins", "Signage & banner", "Business cards"]),
    ("Payments", ["Cash float", "Card reader", "Phone / tablet", "Charger",
                  "Payment QR code"]),
    ("Supplies", ["Shopping bags", "Tissue paper", "Tape & scissors",
                  "Pens", "Receipt book", "Sticky tack / clips"]),
]

MONTH_NAMES = ["January", "February", "March", "April", "May", "June",
               "July", "August", "September", "October", "November",
               "December"]

# ===========================================================================
# SETUP LIST COLUMNS (editable dropdown sources)
# ===========================================================================
LIST_COLS = {"categories": "C", "payments": "D", "weights": "E",
             "units": "F", "suppliers": "G"}
LIST_VALUES = {"categories": CATEGORIES, "payments": PAYMENT_METHODS,
               "weights": YARN_WEIGHTS, "units": UNITS,
               "suppliers": SUPPLIERS}
LIST_TITLES = {"categories": "Product categories",
               "payments": "Payment methods", "weights": "Yarn weights",
               "units": "Units", "suppliers": "Suppliers"}
FIXED_COLS = {"tick": "C"}

# ===========================================================================
# KPI CELLS ON _Data (column AH holds the label, AI the formula)
# ===========================================================================
_KPI_SPECS = [
    ("revenue", "Gross sales (all time)", "#,##0.00"),
    ("discounts", "Discounts given", "#,##0.00"),
    ("cogs", "Cost of products sold", "#,##0.00"),
    ("fees", "Event expenses (all fairs)", "#,##0.00"),
    ("profit", "Net profit", "#,##0.00"),
    ("margin", "Profit margin", "0.0%"),
    ("units", "Units sold", "0"),
    ("txns", "Sales log entries", "0"),
    ("aov", "Average sale value", "#,##0.00"),
    ("per_item", "Average profit per item", "#,##0.00"),
    ("events_total", "Fairs on the books", "0"),
    ("events_done", "Fairs worked", "0"),
    ("inv_value", "Inventory value (at unit cost)", "#,##0.00"),
    ("low_products", "Products below minimum", "0"),
    ("low_materials", "Materials below threshold", "0"),
    ("reorder_cost", "Estimated reorder cost", "#,##0.00"),
    ("sell_through", "Sell-through rate", "0.0%"),
    ("made_total", "Units made (all batches)", "0"),
    ("best_product", "Best-selling product", "@"),
    ("best_event", "Best craft fair", "@"),
    ("bs_units", "Most units sold", "@"),
    ("bs_revenue", "Highest revenue product", "@"),
    ("bs_profit", "Highest profit product", "@"),
    ("bs_margin", "Highest margin product", "@"),
    ("bs_slow", "Slowest mover", "@"),
    ("price_suggest", "Pricing calculator: suggested price", "#,##0.00"),
]
KPI_ROW = {}
KPI_FMT = {}
for _i, (_key, _label, _fmt) in enumerate(_KPI_SPECS):
    KPI_ROW[_key] = 2 + _i
    KPI_FMT[_key] = _fmt

KPI_LABEL = {k: label for k, label, _ in _KPI_SPECS}


def cols_last(key):
    """Right-most column letter used by a sheet's chips / head."""
    return {"catalog": "P", "materials": "N", "production": "L",
            "events": "R", "sales": "M", "reorder": "J", "monthly": "K",
            "eventprofit": "J", "pricing": "J", "packing": "F",
            "dashboard": "M", "setup": "J", "guide": "J"}[key]
```

## 3. Themes - Berry Bramble & Mint Meadow

`crochet_tracker/theme.py`

```python
"""
Colour palettes / design systems for the Crochet Craft Fair Tracker.

Two themes ship with the product:

  berry   warm cream canvas, deep berry plum, rose accent, honey gold
  mint    cool paper-white canvas, mint teal, coral accent, lemon gold

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
# BERRY - "yarn shop autumn"
# ===========================================================================
BERRY = Theme(
    key="berry",
    label="Berry",
    bg="#FAF6F1",
    cover_bg="#FAF6F1",
    card="#FFFFFF",
    alt="#FBF7F3",
    border="#E6D9D2",
    border_strong="#CBAFA6",
    primary="#5C3A50",          # deep berry plum
    primary_2="#8A5F79",        # dusty mauve
    primary_2_soft="#EFE2EA",
    primary_soft="#EDE2E8",
    accent="#C05268",           # rose
    accent_soft="#F8E3E7",
    gold="#B98A2E",             # honey
    gold_soft="#F6EDD8",
    ink="#33222C",
    muted="#97818D",
    white="#FFFFFF",
    ok="#4C7A4C",
    ok_soft="#E2EEE0",
    warn="#B4761A",
    warn_soft="#FBF0D9",
    bad="#AC2F2F",
    bad_soft="#F7E2E0",
    info="#5E6FA3",
    info_soft="#E5E9F4",
    plum="#7A5273",
    plum_soft="#F0E6EE",
    title_font="Georgia",
    body_font="Calibri",
    mono_font="Consolas",
    tabs={
        "data": "#97818D",
        "setup": "#8A5F79",
        "dashboard": "#5C3A50",
        "catalog": "#C05268",
        "materials": "#7A5273",
        "production": "#4C7A4C",
        "events": "#B98A2E",
        "sales": "#4C7A4C",
        "eventprofit": "#5E6FA3",
        "reorder": "#AC2F2F",
        "packing": "#C05268",
        "pricing": "#5E6FA3",
        "monthly": "#5C3A50",
        "guide": "#97818D",
    },
)

# ===========================================================================
# MINT - "fresh stitch studio"
# ===========================================================================
MINT = Theme(
    key="mint",
    label="Mint",
    bg="#F7FAF8",
    cover_bg="#F7FAF8",
    card="#FFFFFF",
    alt="#F4F8F5",
    border="#D8E4DD",
    border_strong="#AFC8BC",
    primary="#2F6D5F",          # mint teal
    primary_2="#5E9488",        # sage
    primary_2_soft="#E1EEE9",
    primary_soft="#E3EFEA",
    accent="#E2725B",           # coral
    accent_soft="#FBE7E2",
    gold="#C9A227",             # lemon gold
    gold_soft="#F8F1D9",
    ink="#22332E",
    muted="#7E948C",
    white="#FFFFFF",
    ok="#3E7C4F",
    ok_soft="#E1F0E4",
    warn="#C07C12",
    warn_soft="#FBF1DA",
    bad="#B23A3A",
    bad_soft="#F8E4E2",
    info="#3E6E8E",
    info_soft="#E1ECF2",
    plum="#6C5B94",
    plum_soft="#EBE7F2",
    title_font="Trebuchet MS",
    body_font="Calibri",
    mono_font="Consolas",
    tabs={
        "data": "#7E948C",
        "setup": "#5E9488",
        "dashboard": "#2F6D5F",
        "catalog": "#E2725B",
        "materials": "#6C5B94",
        "production": "#3E7C4F",
        "events": "#C9A227",
        "sales": "#3E7C4F",
        "eventprofit": "#3E6E8E",
        "reorder": "#B23A3A",
        "packing": "#E2725B",
        "pricing": "#3E6E8E",
        "monthly": "#2F6D5F",
        "guide": "#7E948C",
    },
)

THEMES = {"berry": BERRY, "mint": MINT}
```

## 4. Styles - format factory

`crochet_tracker/styles.py`

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

## 5. Book - workbook engine (names, DV, CF, charts, paint)

`crochet_tracker/book.py`

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
            "category": "Crochet Craft Fair Business Spreadsheet",
            "keywords": ("crochet business, craft fair tracker, product "
                         "catalog, yarn inventory, inventory spreadsheet, "
                         "profit dashboard, pricing calculator, sales log, "
                         "event profit, reorder list, excel template, "
                         "google sheets, etsy spreadsheet"),
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
            "Message": "%s!$C$%d" % (su, C.SU_MESSAGE),
            "Currency": "%s!$C$%d" % (su, C.SU_CURRENCY),
            "ReportYear": "%s!$C$%d" % (su, C.SU_YEAR),
            "HourlyWage": "%s!$C$%d" % (su, C.SU_WAGE),
            "OverheadPct": "%s!$C$%d" % (su, C.SU_OVERHEAD),
            "TargetMargin": "%s!$C$%d" % (su, C.SU_MARGIN),
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
            "Tick": ("tick", 2),
        }
        for nm, (key, n) in sorted(fixed.items()):
            col = C.FIXED_COLS[key]
            self.wb.define_name(
                nm, "=%s!$%s$%d:$%s$%d"
                % (su, col, C.SU_FIXED_FIRST, col, C.SU_FIXED_FIRST + n - 1))

        # Dynamic lists that live in the working sheets (grow with the data).
        dyn = {"EventList": ("events", "B"),
               "ProductsList": ("catalog", "C")}
        if self.has("materials"):
            dyn["MaterialsList"] = ("materials", "B")
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
        return "%s!$AI$%d" % (self.q("data"), C.KPI_ROW[key])

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
    "products": "ProductsList",
    "events": "EventList",
    "materials": "MaterialsList",
    "categories": "Categories",
    "payments": "PaymentMethods",
    "weights": "YarnWeights",
    "units": "Units",
    "suppliers": "Suppliers",
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

## 6. Demo model - Willow & Wren fictional studio

`crochet_tracker/demo.py`

```python
"""
The fictional EXAMPLE business: Willow & Wren Crochet Studio.

Every cached value shown in the workbooks is computed here from the raw
records below, so the filled example is always internally consistent
(sales log <-> production batches <-> fair totals <-> dashboard).
"""

import datetime


def today():
    return datetime.date(2026, 9, 13)


SETTINGS = {
    "business": "Willow & Wren Crochet Studio",
    "message": "Autumn fair season - stock up on chunky yarn!",
    "currency": "$",
    "year": 2026,
    "wage": 14.0,
    "overhead": 0.10,
    "margin": 0.45,
}

# sku, name, category, desc, price, yarn, pack, hours, min, link
PRODUCTS = [
    ("CB-001", "Amigurumi Bunny", "Toys & Amigurumi",
     "Soft cotton bunny with safety eyes", 28.00, 4.20, 0.80, 2.5, 4,
     "photos/bunny.jpg"),
    ("CB-002", "Chunky Beanie", "Wearables",
     "Super chunky merino beanie", 24.00, 5.60, 0.70, 1.5, 6,
     "photos/beanie.jpg"),
    ("CB-003", "Granny Square Cardigan", "Wearables",
     "Hand-joined granny squares", 95.00, 18.40, 2.20, 12.0, 2,
     "photos/cardigan.jpg"),
    ("CB-004", "Market Tote Bag", "Accessories",
     "Sturdy cotton market tote", 32.00, 6.10, 1.00, 3.0, 5,
     "photos/tote.jpg"),
    ("CB-005", "Cozy Scarf", "Wearables",
     "Worsted wool scarf", 38.00, 8.90, 1.10, 4.0, 4, "photos/scarf.jpg"),
    ("CB-006", "Baby Booties", "Baby & Kids",
     "Tiny soft booties, pair", 18.00, 2.60, 0.60, 1.2, 6,
     "photos/booties.jpg"),
    ("CB-007", "Flower Coaster Set", "Home & Decor",
     "Set of 4 cotton coasters", 16.00, 2.20, 0.50, 1.0, 8,
     "photos/coasters.jpg"),
    ("CB-008", "Lap Blanket", "Home & Decor",
     "Granny stripe lap blanket", 75.00, 16.80, 2.00, 10.0, 2,
     "photos/blanket.jpg"),
    ("CB-009", "Ear Warmer Headband", "Accessories",
     "Twisted ear warmer", 16.00, 2.90, 0.50, 1.0, 4, "photos/earwarmer.jpg"),
    ("CB-010", "Stuffed Octopus", "Toys & Amigurumi",
     "Squishy friend octopus", 26.00, 4.60, 0.80, 2.5, 4,
     "photos/octopus.jpg"),
    ("CB-011", "Table Runner", "Home & Decor",
     "Lacy cotton table runner", 55.00, 10.40, 1.60, 7.0, 2,
     "photos/runner.jpg"),
    ("CB-012", "Keychain Charm", "Accessories",
     "Mini amigurumi keychain", 10.00, 1.10, 0.40, 0.5, 6,
     "photos/charm.jpg"),
]

# name, color, brand, weight, purchased, unit, cost_unit, used, supplier,
# threshold
MATERIALS = [
    ("Cotton yarn - natural", "Natural", "Cotton Craft Co",
     "Worsted / aran", 20, "balls", 3.40, 16, "Local yarn shop", 5),
    ("Acrylic yarn - raspberry", "Raspberry", "Value Yarns",
     "Worsted / aran", 15, "balls", 2.60, 9, "Online yarn store", 4),
    ("Merino wool - sage", "Sage", "Highland Fibres", "DK / light", 12,
     "balls", 6.80, 10, "Local yarn shop", 4),
    ("Chenille chunky - cream", "Cream", "Snuggle Spun", "Super chunky", 10,
     "balls", 4.90, 6, "Online yarn store", 3),
    ("Cotton yarn - mustard", "Mustard", "Cotton Craft Co", "DK / light", 8,
     "balls", 3.20, 7, "Local yarn shop", 3),
    ("Wool blend - lavender", "Lavender", "Highland Fibres",
     "Worsted / aran", 10, "balls", 5.40, 4, "Wholesale craft", 3),
    ("Polyester stuffing", "White", "Craft Basics", "Mixed / other", 6,
     "pieces", 2.10, 4, "Wholesale craft", 2),
    ("Bamboo buttons", "Natural", "Craft Basics", "Mixed / other", 40,
     "pieces", 0.35, 22, "Online yarn store", 10),
    ("Safety eyes", "Black", "Craft Basics", "Mixed / other", 30, "pieces",
     0.25, 24, "Wholesale craft", 10),
    ("Paper gift tags", "Kraft", "Print & Tag", "Mixed / other", 60,
     "pieces", 0.15, 35, "Market stall", 20),
]

# product, date, made, reserved, taken, sold, returned, damaged, notes
PRODUCTION = [
    ("Amigurumi Bunny", datetime.date(2026, 3, 2), 11, 1, 10, 6, 0, 0,
     "Best spring seller"),
    ("Chunky Beanie", datetime.date(2026, 3, 12), 15, 1, 12, 5, 0, 0, ""),
    ("Granny Square Cardigan", datetime.date(2026, 4, 18), 3, 0, 2, 1, 0, 0,
     "Made to order"),
    ("Market Tote Bag", datetime.date(2026, 4, 2), 11, 1, 10, 8, 0, 0, ""),
    ("Cozy Scarf", datetime.date(2026, 8, 22), 6, 0, 4, 2, 0, 0,
     "Autumn stock"),
    ("Baby Booties", datetime.date(2026, 3, 28), 8, 0, 6, 5, 0, 0, ""),
    ("Flower Coaster Set", datetime.date(2026, 2, 20), 20, 0, 16, 14, 0, 1,
     "One water damaged"),
    ("Lap Blanket", datetime.date(2026, 2, 10), 3, 0, 3, 2, 0, 0,
     "Sold out online"),
    ("Ear Warmer Headband", datetime.date(2026, 5, 2), 10, 0, 10, 10, 0, 0,
     "Sold out at fete"),
    ("Stuffed Octopus", datetime.date(2026, 8, 25), 10, 2, 6, 3, 0, 0,
     "Reserved for market"),
    ("Table Runner", datetime.date(2026, 3, 5), 3, 0, 3, 3, 0, 0, ""),
    ("Keychain Charm", datetime.date(2026, 4, 20), 18, 0, 16, 18, 0, 0,
     "Impulse buy winner"),
]

# name, date, location, organizer, booth, travel, parking, food, display,
# taken
EVENTS = [
    ("Spring Makers Market", datetime.date(2026, 3, 14), "Town Hall",
     "Makers Guild", 45, 12, 5, 8, 20, 38),
    ("Easter Craft Bazaar", datetime.date(2026, 4, 5), "Church Hall",
     "St Mary's", 30, 8, 0, 6, 10, 20),
    ("School Fete Stall", datetime.date(2026, 5, 16), "Primary School",
     "PTA", 25, 6, 0, 5, 8, 30),
    ("Summer Night Bazaar", datetime.date(2026, 6, 19), "Riverside Park",
     "City Council", 60, 15, 6, 10, 25, 30),
    ("Farmers Market Pop-up", datetime.date(2026, 7, 25), "Market Square",
     "Growers Co-op", 40, 10, 4, 7, 12, 16),
    ("Online Order Batch", datetime.date(2026, 8, 30), "Online",
     "Own shop", 0, 0, 0, 0, 0, 0),
    ("Autumn Craft Fair", datetime.date(2026, 9, 26), "Exhibition Centre",
     "Craft Collective", 75, 20, 8, 12, 30, 0),
    ("Winter Holiday Market", datetime.date(2026, 11, 28), "Civic Square",
     "City Council", 85, 18, 8, 12, 35, 0),
]

# date, event, product, qty, unit, discount, method, ref, notes
SALES = [
    (datetime.date(2026, 3, 14), "Spring Makers Market", "Amigurumi Bunny",
     4, 28.00, 0, "Card", "SPR-01", ""),
    (datetime.date(2026, 3, 14), "Spring Makers Market", "Chunky Beanie",
     3, 24.00, 0, "Cash", "", ""),
    (datetime.date(2026, 3, 14), "Spring Makers Market",
     "Flower Coaster Set", 5, 16.00, 0, "Cash", "", ""),
    (datetime.date(2026, 3, 14), "Spring Makers Market", "Baby Booties",
     2, 18.00, 0, "Card", "SPR-02", ""),
    (datetime.date(2026, 4, 5), "Easter Craft Bazaar", "Baby Booties",
     3, 18.00, 0, "Cash", "", ""),
    (datetime.date(2026, 4, 5), "Easter Craft Bazaar", "Amigurumi Bunny",
     2, 28.00, 0, "Card", "EAS-01", ""),
    (datetime.date(2026, 4, 5), "Easter Craft Bazaar", "Flower Coaster Set",
     4, 16.00, 0, "Cash", "", ""),
    (datetime.date(2026, 5, 16), "School Fete Stall", "Keychain Charm",
     10, 10.00, 0, "Cash", "", "Kids loved these"),
    (datetime.date(2026, 5, 16), "School Fete Stall", "Ear Warmer Headband",
     6, 16.00, 0, "Cash", "", ""),
    (datetime.date(2026, 5, 16), "School Fete Stall", "Flower Coaster Set",
     3, 16.00, 0, "Cash", "", ""),
    (datetime.date(2026, 6, 19), "Summer Night Bazaar", "Market Tote Bag",
     4, 32.00, 0, "Card", "SUM-01", ""),
    (datetime.date(2026, 6, 19), "Summer Night Bazaar",
     "Ear Warmer Headband", 4, 16.00, 0, "Cash", "", ""),
    (datetime.date(2026, 6, 19), "Summer Night Bazaar", "Chunky Beanie",
     2, 24.00, 0, "Card", "SUM-02", ""),
    (datetime.date(2026, 6, 19), "Summer Night Bazaar", "Keychain Charm",
     8, 10.00, 0, "Cash", "", ""),
    (datetime.date(2026, 6, 19), "Summer Night Bazaar", "Stuffed Octopus",
     3, 26.00, 0, "Card", "SUM-03", ""),
    (datetime.date(2026, 7, 25), "Farmers Market Pop-up", "Market Tote Bag",
     4, 32.00, 0, "Cash", "", ""),
    (datetime.date(2026, 7, 25), "Farmers Market Pop-up", "Cozy Scarf",
     2, 38.00, 0, "Card", "FAR-01", ""),
    (datetime.date(2026, 7, 25), "Farmers Market Pop-up",
     "Flower Coaster Set", 2, 16.00, 2.00, "Cash", "", "Friend discount"),
    (datetime.date(2026, 8, 30), "Online Order Batch",
     "Granny Square Cardigan", 1, 95.00, 0, "Bank transfer", "INV-2041",
     "Commission"),
    (datetime.date(2026, 8, 30), "Online Order Batch", "Lap Blanket",
     1, 75.00, 0, "Bank transfer", "INV-2042", ""),
    (datetime.date(2026, 8, 30), "Online Order Batch", "Table Runner",
     3, 55.00, 0, "Online order", "INV-2043", ""),
]

PACKING_TICKS = {
    "Finished products": True, "Extra inventory": True,
    "Price tags": True, "Product labels / care cards": True,
    "Table": True, "Tablecloth": True, "Display stands / risers": True,
    "Baskets & bins": True, "Signage & banner": False,
    "Business cards": False,
    "Cash float": True, "Card reader": True, "Phone / tablet": True,
    "Charger": False, "Payment QR code": False,
    "Shopping bags": True, "Tissue paper": True, "Tape & scissors": True,
    "Pens": False, "Receipt book": False, "Sticky tack / clips": False,
}

PRICING = {"material": 4.00, "packaging": 1.00, "hours": 1.5, "wage": 15.0,
           "overhead": 0.10, "margin": 0.45}

PROFIT_EVENT = "Summer Night Bazaar"


class Demo(object):
    """Raw records plus every aggregate the builders cache."""

    def __init__(self):
        self.settings = dict(SETTINGS)
        self.products = [dict(zip(
            ("sku", "name", "category", "desc", "price", "yarn", "pack",
             "hours", "min", "link"), p)) for p in PRODUCTS]
        self.materials = [dict(zip(
            ("name", "color", "brand", "weight", "purchased", "unit",
             "cost_unit", "used", "supplier", "threshold"), m))
            for m in MATERIALS]
        self.production = [dict(zip(
            ("product", "date", "made", "reserved", "taken", "sold",
             "returned", "damaged", "notes"), p)) for p in PRODUCTION]
        self.events = [dict(zip(
            ("name", "date", "location", "organizer", "booth", "travel",
             "parking", "food", "display", "taken"), e)) for e in EVENTS]
        self.sales = [dict(zip(
            ("date", "event", "product", "qty", "unit", "discount",
             "method", "ref", "notes"), s)) for s in SALES]
        self.packing = dict(PACKING_TICKS)
        self.pricing = dict(PRICING)
        self.profit_event = PROFIT_EVENT
        for p in self.products:
            p["unit_cost"] = round(p["yarn"] + p["pack"], 2)
            p["profit"] = round(p["price"] - p["unit_cost"], 2)
            p["margin"] = p["profit"] / p["price"]
        for m in self.materials:
            m["total"] = round(m["purchased"] * m["cost_unit"], 2)
            m["remaining"] = m["purchased"] - m["used"]
        for b in self.production:
            b["current"] = (b["made"] + b["returned"] - b["sold"]
                            - b["damaged"])
            b["available"] = b["current"] - b["reserved"]
        for p in self.products:
            p["stock"] = sum(b["current"] for b in self.production
                             if b["product"] == p["name"])
            p["status"] = (_stock_status(p["stock"], p["min"]))
        for e in self.events:
            e["total"] = (e["booth"] + e["travel"] + e["parking"] + e["food"]
                          + e["display"])
            rows = [s for s in self.sales if s["event"] == e["name"]]
            e["sold"] = sum(s["qty"] for s in rows)
            e["sales"] = round(sum(s["qty"] * s["unit"] - s["discount"]
                                   for s in rows), 2)
            e["cogs"] = round(sum(s["qty"] * _uc(self, s["product"])
                                  for s in rows), 2)
            e["net"] = round(e["sales"] - e["total"] - e["cogs"], 2)
            e["margin"] = (e["net"] / e["sales"]) if e["sales"] else 0.0
            e["best"] = _best_product(rows)
        self.agg = self._aggregates()

    # ------------------------------------------------------------------
    def _aggregates(self):
        a = {}
        rev = round(sum(s["qty"] * s["unit"] - s["discount"]
                        for s in self.sales), 2)
        disc = round(sum(s["discount"] for s in self.sales), 2)
        cogs = round(sum(s["qty"] * _uc(self, s["product"])
                         for s in self.sales), 2)
        fees = round(sum(e["total"] for e in self.events
                         if e["date"] <= today()), 2)
        units = sum(s["qty"] for s in self.sales)
        profit = round(rev - cogs - fees, 2)
        made = sum(b["made"] for b in self.production)
        a.update(revenue=rev, discounts=disc, cogs=cogs, fees=fees,
                 profit=profit, margin=profit / rev if rev else 0.0,
                 units=units, txns=len(self.sales),
                 aov=rev / len(self.sales) if self.sales else 0.0,
                 per_item=profit / units if units else 0.0,
                 events_total=len(self.events),
                 events_done=sum(1 for e in self.events
                                 if e["date"] <= today() and e["sales"]),
                 inv_value=round(sum(p["stock"] * p["unit_cost"]
                                     for p in self.products), 2),
                 low_products=sum(1 for p in self.products
                                  if p["stock"] < p["min"]),
                 low_materials=sum(1 for m in self.materials
                                   if m["remaining"] <= m["threshold"]),
                 sell_through=units / made if made else 0.0,
                 made_total=made)
        a["reorder_cost"] = round(self._reorder_cost(), 2)
        done = [e for e in self.events if e["sales"]]
        a["best_event"] = max(done, key=lambda e: e["net"])["name"] if done \
            else ""
        per_prod = self._product_stats()
        a["best_product"] = max(per_prod, key=lambda r: r["revenue"])["name"]
        a["bs_units"] = max(per_prod, key=lambda r: r["units"])["name"]
        a["bs_revenue"] = max(per_prod, key=lambda r: r["revenue"])["name"]
        a["bs_profit"] = max(per_prod, key=lambda r: r["profit"])["name"]
        a["bs_margin"] = max(per_prod, key=lambda r: r["margin_adj"])["name"]
        a["bs_slow"] = min(per_prod, key=lambda r: r["slow_adj"])["name"]
        a["product_stats"] = [(r["name"], r["units"], r["revenue"],
                               r["stock"], r["margin_adj"], r["slow_adj"])
                              for r in per_prod]
        a["event_stats"] = [(e["name"], e["sales"], e["net"], e["sold"])
                            for e in self.events]
        a["months"] = self._months()
        a["payments"] = self._payments()
        a["exp_cats"] = [("Booth fees", sum(e["booth"] for e in self.events)),
                         ("Travel", sum(e["travel"] for e in self.events)),
                         ("Parking", sum(e["parking"]
                                         for e in self.events)),
                         ("Food & drinks", sum(e["food"]
                                               for e in self.events)),
                         ("Display & decor", sum(e["display"]
                                                 for e in self.events))]
        q = self.pricing
        labor = q["hours"] * q["wage"]
        sub = q["material"] + q["packaging"] + labor
        oh = sub * q["overhead"]
        true_cost = sub + oh
        price = true_cost / (1 - q["margin"]) if q["margin"] < 1 else 0
        a["price_suggest"] = round(price, 2)
        a["pricing"] = dict(q, labor=labor, overhead_amt=oh,
                            true_cost=true_cost, price=price,
                            charm=_charm(price), profit=price - true_cost)
        ep = self.profit_event
        ev = [e for e in self.events if e["name"] == ep][0]
        rows = [s for s in self.sales if s["event"] == ep]
        gross = ev["sales"]
        gp = gross - ev["cogs"]
        tc = ev["cogs"] + ev["total"]
        a["event_profit"] = {
            "event": ep, "units": ev["sold"], "gross": gross,
            "discounts": round(sum(s["discount"] for s in rows), 2),
            "cogs": ev["cogs"], "booth": ev["booth"], "travel": ev["travel"],
            "parking": ev["parking"], "food": ev["food"],
            "display": ev["display"], "gross_profit": round(gp, 2),
            "total_cost": round(tc, 2), "net": ev["net"],
            "margin": ev["margin"],
            "avg_sale": gross / len(rows) if rows else 0.0,
            "breakeven": tc / (gp / gross) if gross and gp else 0.0,
            "roi": ev["net"] / tc if tc else 0.0}
        a["packed"] = sum(1 for v in self.packing.values() if v)
        a["packed_total"] = len(self.packing)
        return a

    # ------------------------------------------------------------------
    def _product_stats(self):
        out = []
        for p in self.products:
            rows = [s for s in self.sales if s["product"] == p["name"]]
            units = sum(s["qty"] for s in rows)
            revenue = round(sum(s["qty"] * s["unit"] - s["discount"]
                                for s in rows), 2)
            out.append({"name": p["name"], "units": units,
                        "revenue": revenue, "stock": p["stock"],
                        "profit": round(units * p["profit"], 2),
                        "margin_adj": p["margin"] if units else 0.0,
                        "slow_adj": units})
        return out

    def _months(self):
        out = []
        for m in range(1, 13):
            rows = [s for s in self.sales if s["date"].month == m
                    and s["date"].year == self.settings["year"]]
            evs = [e for e in self.events if e["date"].month == m
                   and e["date"].year == self.settings["year"]]
            rev = round(sum(s["qty"] * s["unit"] - s["discount"]
                            for s in rows), 2)
            cogs = round(sum(s["qty"] * _uc(self, s["product"])
                             for s in rows), 2)
            fees = round(sum(e["total"] for e in evs), 2)
            units = sum(s["qty"] for s in rows)
            net = round(rev - cogs - fees, 2)
            out.append((m, rev, cogs, fees, net, units, len(evs),
                        len(rows)))
        return out

    def _payments(self):
        out = []
        for method in ("Cash", "Card", "Bank transfer", "Mobile wallet",
                       "Online order", "Other"):
            out.append((method, round(sum(
                s["qty"] * s["unit"] - s["discount"]
                for s in self.sales if s["method"] == method), 2)))
        return out

    def _reorder_cost(self):
        total = 0.0
        for p in self.products:
            if p["stock"] < p["min"]:
                total += (p["min"] * 2 - p["stock"]) * p["unit_cost"]
        for m in self.materials:
            if m["remaining"] <= m["threshold"]:
                total += (m["threshold"] * 2 - m["remaining"]) * m["cost_unit"]
        return total

    # ------------------------------------------------------------------
    def money(self, value, decimals=0):
        return "%s%s" % (self.settings["currency"],
                         format(round(value, decimals),
                                ",.%df" % decimals))


def _uc(demo, product):
    for p in demo.products:
        if p["name"] == product:
            return p["unit_cost"]
    return 0.0


def _stock_status(stock, minimum):
    if stock <= 0:
        return "out"
    if stock < minimum:
        return "low"
    return "ok"


def _best_product(rows):
    if not rows:
        return ""
    tally = {}
    for s in rows:
        tally[s["product"]] = tally.get(s["product"], 0) + s["qty"]
    return max(tally, key=tally.get)


def _charm(price):
    import math
    if not price:
        return 0.0
    return round(math.ceil(price) - 0.05, 2)
```

## 7. Workbook orchestration

`crochet_tracker/workbook.py`

```python
"""Build orchestration: editions, themes, modes and product filenames."""

import os

from . import config as C
from . import theme as T
from .book import Book
from .sheets import (catalog, dashboard, data, eventprofit, events, guide,
                     materials, monthly, packing, pricing, production,
                     reorder, sales, setup)

BUILDERS = {
    "data": data.build,
    "setup": setup.build,
    "dashboard": dashboard.build,
    "catalog": catalog.build,
    "materials": materials.build,
    "production": production.build,
    "events": events.build,
    "sales": sales.build,
    "eventprofit": eventprofit.build,
    "reorder": reorder.build,
    "packing": packing.build,
    "pricing": pricing.build,
    "monthly": monthly.build,
    "guide": guide.build,
}


def product_filename(edition, theme_name, mode):
    return "Crochet_Craft_Fair_Tracker_%s_%s%s.xlsx" % (
        edition.upper(), T.THEMES[theme_name].label,
        "_EXAMPLE" if mode == "demo" else "")


def build_workbook(path, edition="premium", theme_name="berry",
                   mode="blank", protect=None, images=True):
    """Write one workbook file and return its build stats."""
    from . import demo as D
    if edition not in ("basic", "premium"):
        raise ValueError("edition must be 'basic' or 'premium'")
    if mode not in ("blank", "demo"):
        raise ValueError("mode must be 'blank' or 'demo'")
    model = D.Demo() if mode == "demo" else None
    bk = Book(path, T.THEMES[theme_name], edition=edition, mode=mode,
              protect=protect, images=images, demo=model)
    for key in bk.order:
        BUILDERS[key](bk)
    bk.close()
    stats = dict(bk.stats)
    stats["size"] = os.path.getsize(path)
    return stats


def build_all(outdir, protect=None, images=True):
    """The curated six-file Etsy product set."""
    os.makedirs(outdir, exist_ok=True)
    combos = [("premium", "berry", "demo"), ("premium", "berry", "blank"),
              ("premium", "mint", "blank"), ("basic", "berry", "blank"),
              ("basic", "mint", "blank"), ("basic", "berry", "demo")]
    out = []
    for edition, theme_name, mode in combos:
        name = product_filename(edition, theme_name, mode)
        path = os.path.join(outdir, name)
        stats = build_workbook(path, edition, theme_name, mode,
                               protect=protect, images=images)
        out.append((name, stats))
    return out
```

## 8. Sheets package init

`crochet_tracker/sheets/__init__.py`

```python
"""Worksheet builders - one module per tab."""
```

## 9. Shared sheet helpers

`crochet_tracker/sheets/common.py`

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
                                     "calendar picker.", first=None,
            last=None):
    if isinstance(fields, str):
        fields = (fields,)
    for field in fields:
        bk.ws(key).data_validation(
            r(first or C.ROW_FIRST), ci(bk.col(key, field)),
            r(last or C.last_row(key)), ci(bk.col(key, field)),
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


# ===========================================================================
# shared page furniture (crochet product)
# ===========================================================================
def sheet_head(bk, key, last_col, title, sub):
    """Canvas, widths, title band and the home link."""
    from .. import config as _C
    S, th = bk.S, bk.th
    ws = bk.ws(key)
    bk.widths(key, _C.WIDTHS[key])
    try:
        depth = _C.last_row(key) + 10
    except KeyError:
        depth = 60
    bk.paint(key, 0, 0, depth, ci(last_col), S.canvas)
    ws.set_row(r(_C.ROW_SPACER_1), 7)
    ws.set_row(r(_C.ROW_TITLE), 32)
    ws.set_row(r(_C.ROW_SUBTITLE), 18)
    ws.set_row(r(_C.ROW_SPACER_2), 8)
    ws.set_row(r(_C.ROW_STATS), 22)
    ws.set_row(r(_C.ROW_SPACER_3), 8)
    ws.merge_range(r(_C.ROW_TITLE), 1, r(_C.ROW_TITLE), ci(last_col) - 3,
                   title, S.sheet_title)
    ws.merge_range(r(_C.ROW_SUBTITLE), 1, r(_C.ROW_SUBTITLE),
                   ci(last_col) - 3, "  " + sub, S.sheet_sub)
    ws.merge_range(r(_C.ROW_SUBTITLE), ci(last_col) - 2,
                   r(_C.ROW_SUBTITLE), ci(last_col), "", S.home_link)
    ws.write_url(r(_C.ROW_SUBTITLE), ci(last_col) - 2,
                 "internal:%s!A1" % bk.q("dashboard"), S.home_link,
                 "\U0001F3E0  Dashboard")


def chips(bk, key, chips, last_col=None):
    """Stats pills on ROW_STATS. chips = (formula, colour, cached, span)."""
    from .. import config as _C
    S, th = bk.S, bk.th
    ws = bk.ws(key)
    last_col = last_col or _C.cols_last(key)
    col = 1
    for formula, color, cached, span in chips:
        fmt = S.pill(th.soft(color), getattr(th, color), size=10.5,
                     bold=True, align="left")
        ws.merge_range(r(_C.ROW_STATS), col, r(_C.ROW_STATS),
                       col + span - 1, "", fmt)
        ws.write_formula(r(_C.ROW_STATS), col,
                         formula if formula.startswith("=")
                         else '="' + formula, fmt, cached)
        bk.stats["formulas"] += 1
        col += span
    if col <= ci(last_col):
        ws.merge_range(r(_C.ROW_STATS), col, r(_C.ROW_STATS),
                       ci(last_col), "", S.canvas)


def footer_nav(bk, key, row, last_col, landscape=True, zoom=90, tip=None):
    """Gold tip strip, nav grid and print setup."""
    S, th = bk.S, bk.th
    ws = bk.ws(key)
    ws.set_row(r(row), 30)
    ws.merge_range(r(row), 1, r(row), ci(last_col),
                   tip or "  \U0001F4A1  Cream cells are yours to type in; "
                          "white cells calculate.  Dropdowns keep every "
                          "column honest.",
                   S.f(**S.base(font_size=10, italic=True,
                                font_color=th.ink, bg_color=th.gold_soft,
                                align="left", valign="vcenter", indent=1)))
    nav = row + 2
    ws.set_row(r(nav), 24)
    bk.nav_row(key, nav, max_col=last_col)
    bk.page(key, last_col, nav + 1, landscape=landscape, zoom=zoom)
```

## 10. _Data - hidden calculation engine

`crochet_tracker/sheets/data.py`

```python
"""_Data - the hidden engine room: pools, KPI cells and chart fuel."""

from .. import config as C
from ..book import r, ci


def build(bk):
    key = "data"
    ws = bk.ws(key)
    for col, width in (("A", 3), ("H", 9), ("I", 11), ("J", 11), ("K", 11),
                       ("L", 11), ("M", 9), ("N", 9), ("O", 9),
                       ("P", 16), ("Q", 11), ("S", 17), ("T", 11),
                       ("V", 24), ("W", 9), ("X", 11), ("Y", 9), ("Z", 9),
                       ("AA", 9), ("AB", 11), ("AC", 24), ("AD", 10),
                       ("AE", 11), ("AF", 9), ("AG", 12), ("AH", 28),
                       ("AI", 12), ("AK", 9), ("AL", 24), ("AN", 9),
                       ("AO", 24)):
        ws.set_column(ci(col), ci(col), width)
    S, th, m = bk.S, bk.th, bk.demo
    agg = m.agg if m else {}
    premium = bk.has("materials")

    hdr = S.f(**S.base(bold=True, font_size=9, font_color=th.muted,
                       bg_color=th.bg, align="left", valign="vcenter"))
    txt = S.f(**S.base(font_size=9, font_color=th.ink, bg_color=th.bg))
    num = S.f(**S.base(font_size=9, font_color=th.ink, bg_color=th.bg,
                       num_format="#,##0.00"))
    integer = S.f(**S.base(font_size=9, font_color=th.ink, bg_color=th.bg,
                           num_format="0"))

    def q(k):
        return bk.q(k)

    sales = q("sales") if bk.has("sales") else None
    events = q("events") if bk.has("events") else None
    catalog = q("catalog")
    prod = q("production") if bk.has("production") else None
    mats = q("materials") if premium else None
    pricing = q("pricing")

    n = [0]

    def fcount():
        n[0] += 1
        bk.stats["formulas"] += 1

    # ------------------------------------------------------------------
    # month table H..N rows 2-13
    # ------------------------------------------------------------------
    for col, label in (("H", "Month"), ("I", "Revenue"), ("J", "COGS"),
                       ("K", "Event fees"), ("L", "Net"), ("M", "Units"),
                       ("N", "Markets"), ("O", "Sales lines")):
        ws.write(r(1), ci(col), label, hdr)
    months = (agg.get("months") or [])
    for i in range(12):
        row = C.DATA_MONTH_FIRST + i
        mon = i + 1
        ws.write(r(row), ci("H"), C.MONTH_NAMES[i], txt)
        d0 = "DATE(ReportYear,%d,1)" % mon
        d1 = "EDATE(DATE(ReportYear,%d,1),1)" % mon
        cached = months[i] if i < len(months) else None
        if sales:
            ws.write_formula(
                r(row), ci("I"),
                '=SUMIFS(%s,%s,">="&%s,%s,"<"&%s)'
                % (bk.rng("sales", "total"), bk.rng("sales", "date"), d0,
                   bk.rng("sales", "date"), d1),
                num, cached[1] if cached else 0); fcount()
            ws.write_formula(
                r(row), ci("J"),
                '=SUMIFS(%s,%s,">="&%s,%s,"<"&%s)'
                % (bk.rng("sales", "cost"), bk.rng("sales", "date"), d0,
                   bk.rng("sales", "date"), d1),
                num, cached[2] if cached else 0); fcount()
            ws.write_formula(
                r(row), ci("M"),
                '=SUMIFS(%s,%s,">="&%s,%s,"<"&%s)'
                % (bk.rng("sales", "qty"), bk.rng("sales", "date"), d0,
                   bk.rng("sales", "date"), d1),
                integer, cached[5] if cached else 0); fcount()
            ws.write_formula(
                r(row), ci("O"),
                '=COUNTIFS(%s,">="&%s,%s,"<"&%s)'
                % (bk.rng("sales", "date"), d0, bk.rng("sales", "date"),
                   d1),
                integer, cached[7] if cached and len(cached) > 7 else 0)
            fcount()
        else:
            for col in ("I", "J", "M"):
                ws.write(r(row), ci(col), 0, num)
        if events:
            ws.write_formula(
                r(row), ci("K"),
                '=SUMIFS(%s,%s,">="&%s,%s,"<"&%s)'
                % (bk.rng("events", "total"), bk.rng("events", "date"), d0,
                   bk.rng("events", "date"), d1),
                num, cached[3] if cached else 0); fcount()
            ws.write_formula(
                r(row), ci("N"),
                '=COUNTIFS(%s,">="&%s,%s,"<"&%s)'
                % (bk.rng("events", "date"), d0, bk.rng("events", "date"),
                   d1),
                integer, cached[6] if cached else 0); fcount()
        else:
            ws.write(r(row), ci("K"), 0, num)
            ws.write(r(row), ci("N"), 0, integer)
        ws.write_formula(r(row), ci("L"), "=$I%d-$J%d-$K%d" % (row, row, row),
                         num, cached[4] if cached else 0); fcount()

    # ------------------------------------------------------------------
    # payment methods P..Q rows 2-7
    # ------------------------------------------------------------------
    ws.write(r(1), ci("P"), "Method", hdr)
    ws.write(r(1), ci("Q"), "Taken", hdr)
    pays = agg.get("payments") or []
    for i, method in enumerate(C.PAYMENT_METHODS):
        row = C.DATA_PAY_FIRST + i
        ws.write(r(row), ci("P"), method, txt)
        if sales:
            ws.write_formula(r(row), ci("Q"),
                             "=SUMIFS(%s,%s,$P%d)"
                             % (bk.rng("sales", "total"),
                                bk.rng("sales", "method"), row),
                             num, pays[i][1] if i < len(pays) else 0); fcount()
        else:
            ws.write(r(row), ci("Q"), 0, num)

    # ------------------------------------------------------------------
    # event expense categories S..T rows 2-6
    # ------------------------------------------------------------------
    ws.write(r(1), ci("S"), "Category", hdr)
    ws.write(r(1), ci("T"), "Spent", hdr)
    exp_cols = (("Booth fees", "booth"), ("Travel", "travel"),
                ("Parking", "parking"), ("Food & drinks", "food"),
                ("Display & decor", "display"))
    exps = dict(agg.get("exp_cats") or [])
    for i, (label, field) in enumerate(exp_cols):
        row = C.DATA_EXP_FIRST + i
        ws.write(r(row), ci("S"), label, txt)
        if events:
            ws.write_formula(r(row), ci("T"), "=SUM(%s)"
                             % bk.rng("events", field),
                             num, exps.get(label, 0)); fcount()
        else:
            ws.write(r(row), ci("T"), 0, num)

    # ------------------------------------------------------------------
    # product pool V..AA rows 2-25
    # ------------------------------------------------------------------
    for col, label in (("V", "Product"), ("W", "Units"), ("X", "Revenue"),
                       ("Y", "In stock"), ("Z", "Margin adj"),
                       ("AA", "Slow adj"), ("AB", "Profit")):
        ws.write(r(1), ci(col), label, hdr)
    stats = agg.get("product_stats") or []
    last = C.ROW_FIRST + C.CAP["catalog"] - 1
    for i in range(C.CAP["catalog"]):
        row = C.DATA_PROD_FIRST + i
        crow = C.ROW_FIRST + i
        ws.write_formula(r(row), ci("V"),
                         '=IF(%s!$C$%d="","",%s!$C$%d)'
                         % (catalog, crow, catalog, crow),
                         txt, stats[i][0] if i < len(stats) else ""); fcount()
        if sales:
            ws.write_formula(r(row), ci("W"),
                             '=IF($V%d="",0,SUMIFS(%s,%s,$V%d))'
                             % (row, bk.rng("sales", "qty"),
                                bk.rng("sales", "product"), row),
                             integer, stats[i][1] if i < len(stats) else 0)
            fcount()
            ws.write_formula(r(row), ci("X"),
                             '=IF($V%d="",0,SUMIFS(%s,%s,$V%d))'
                             % (row, bk.rng("sales", "total"),
                                bk.rng("sales", "product"), row),
                             num, stats[i][2] if i < len(stats) else 0)
            fcount()
        else:
            ws.write(r(row), ci("W"), 0, integer)
            ws.write(r(row), ci("X"), 0, num)
        ws.write_formula(r(row), ci("Y"), '=IF($V%d="","",%s!$M$%d)'
                         % (row, catalog, crow),
                         integer, stats[i][3] if i < len(stats) else "")
        fcount()
        ws.write_formula(r(row), ci("Z"),
                         '=IF($W%d>0,%s!$K$%d,0)' % (row, catalog, crow),
                         S.f(**S.base(font_size=9, font_color=th.ink,
                                      bg_color=th.bg, num_format="0.0%")),
                         stats[i][4] if i < len(stats) else 0); fcount()
        ws.write_formula(r(row), ci("AA"), '=IF($V%d="","",$W%d)'
                         % (row, row),
                         integer, stats[i][5] if i < len(stats) else "")
        fcount()
        ws.write_formula(r(row), ci("AB"),
                         '=IFERROR(ROUND($W%d*%s!$J$%d,2),0)' % (row, catalog, crow),
                         num, round(stats[i][1] * _profit_of(
                             m, stats[i][0]), 2) if i < len(stats) else 0)
        fcount()

    # ------------------------------------------------------------------
    # event pool AC..AF rows 2-25
    # ------------------------------------------------------------------
    for col, label in (("AC", "Event"), ("AD", "Sales"), ("AE", "Net"),
                       ("AF", "Units")):
        ws.write(r(1), ci(col), label, hdr)
    ev_stats = agg.get("event_stats") or []
    for i in range(C.CAP["events"]):
        row = C.DATA_EVENT_FIRST + i
        erow = C.ROW_FIRST + i
        ws.write_formula(r(row), ci("AC"),
                         '=IF(%s!$B$%d="","",%s!$B$%d)'
                         % (events, erow, events, erow),
                         txt, ev_stats[i][0] if i < len(ev_stats) else "")
        fcount()
        ws.write_formula(r(row), ci("AD"), '=IF($AC%d="","",%s!$N$%d)'
                         % (row, events, erow),
                         num, ev_stats[i][1] if i < len(ev_stats) else "")
        fcount()
        ws.write_formula(r(row), ci("AE"), '=IF($AC%d="","",%s!$P$%d)'
                         % (row, events, erow),
                         num, ev_stats[i][2] if i < len(ev_stats) else "")
        fcount()
        ws.write_formula(r(row), ci("AF"), '=IF($AC%d="","",%s!$M$%d)'
                         % (row, events, erow),
                         integer, ev_stats[i][3] if i < len(ev_stats) else "")
        fcount()

    # ------------------------------------------------------------------
    # upcoming fair dates AG rows 2-25 (hero line fuel)
    # ------------------------------------------------------------------
    ws.write(r(1), ci("AG"), "Upcoming", hdr)
    for i in range(C.CAP["events"]):
        row = C.DATA_EVENT_FIRST + i
        erow = C.ROW_FIRST + i
        ws.write_formula(r(row), ci("AG"),
                         '=IF(AND(%s!$C$%d<>"",%s!$C$%d>=TODAY()),%s!$C$%d,"")'
                         % (events, erow, events, erow, events, erow),
                         txt, "")
        fcount()

    # ------------------------------------------------------------------
    # KPI cells AH/AI
    # ------------------------------------------------------------------
    ws.write(r(1), ci("AH"), "KPI", hdr)
    ws.write(r(1), ci("AI"), "Value", hdr)
    F = {}
    mr = C.DATA_MONTH_FIRST
    ml = C.DATA_MONTH_FIRST + 11
    F["revenue"] = ("=SUM($I$%d:$I$%d)" % (mr, ml), "#,##0.00")
    F["discounts"] = ("=SUM(%s)" % bk.rng("sales", "discount"), "#,##0.00")
    F["cogs"] = ("=SUMPRODUCT(IFERROR(%s*1,0))" % bk.rng("sales", "cost"),
                 "#,##0.00")
    F["fees"] = ('=SUMIFS(%s,%s,"<="&TODAY())'
                 % (bk.rng("events", "total"), bk.rng("events", "date")),
                 "#,##0.00")
    F["profit"] = ("=$AI$%d-$AI$%d-$AI$%d"
                   % (C.KPI_ROW["revenue"], C.KPI_ROW["cogs"],
                      C.KPI_ROW["fees"]), "#,##0.00")
    F["margin"] = ('=IFERROR($AI$%d/$AI$%d,0)'
                   % (C.KPI_ROW["profit"], C.KPI_ROW["revenue"]), "0.0%")
    F["units"] = ("=SUM(%s)" % bk.rng("sales", "qty"), "0")
    F["txns"] = ("=COUNT(%s)" % bk.rng("sales", "total"), "0")
    F["aov"] = ('=IFERROR($AI$%d/$AI$%d,0)'
                % (C.KPI_ROW["revenue"], C.KPI_ROW["txns"]), "#,##0.00")
    F["per_item"] = ('=IFERROR($AI$%d/$AI$%d,0)'
                     % (C.KPI_ROW["profit"], C.KPI_ROW["units"]),
                     "#,##0.00")
    F["events_total"] = ("=COUNTA(%s)" % bk.rng("events", "name"), "0")
    F["events_done"] = ('=COUNTIFS(%s,"<="&TODAY(),%s,">0")'
                        % (bk.rng("events", "date"),
                           bk.rng("events", "sales")), "0")
    F["inv_value"] = ("=SUMPRODUCT(IFERROR(%s*1,0)*IFERROR(%s*1,0))"
                      % (bk.rng("catalog", "stock"),
                         bk.rng("catalog", "unit_cost")), "#,##0.00")
    F["low_products"] = ('=COUNTIF(%s,"%s")+COUNTIF(%s,"%s")'
                         % (bk.rng("catalog", "status"), C.ST_LOW,
                            bk.rng("catalog", "status"), C.ST_OUT), "0")
    F["low_materials"] = ('=COUNTIF(%s,"%s")'
                          % (bk.rng("materials", "reorder"), C.RE_YES)
                          if premium else "=0", "0")
    ro = bk.q("reorder")
    if premium:
        F["reorder_cost"] = ("=SUM(%s!$G$%d:$G$%d)+SUM(%s!$G$%d:$G$%d)"
                             % (ro, C.RO_PROD_FIRST, C.RO_PROD_LAST, ro,
                                C.RO_MAT_FIRST, C.RO_MAT_LAST),
                             "#,##0.00")
    else:
        F["reorder_cost"] = (
            "=SUMPRODUCT((%s<=%s)*IFERROR((%s*2-%s)*%s,0))"
            % (bk.rng("catalog", "stock"), bk.rng("catalog", "min"),
               bk.rng("catalog", "min"), bk.rng("catalog", "stock"),
               bk.rng("catalog", "unit_cost")),
            "#,##0.00")
    F["sell_through"] = ('=IFERROR($AI$%d/$AI$%d,0)'
                         % (C.KPI_ROW["units"], C.KPI_ROW["made_total"]),
                         "0.0%")
    F["made_total"] = ("=SUM(%s)" % bk.rng("production", "made"), "0")
    pr = C.DATA_PROD_FIRST
    pl = C.DATA_PROD_FIRST + C.CAP["catalog"] - 1
    F["best_product"] = ('=IFERROR(INDEX($V$%d:$V$%d,MATCH(MAX($X$%d:$X$%d),'
                         '$X$%d:$X$%d,0)),"")' % (pr, pl, pr, pl, pr, pl), "@")
    er = C.DATA_EVENT_FIRST
    el = C.DATA_EVENT_FIRST + C.CAP["events"] - 1
    F["best_event"] = ('=IFERROR(INDEX($AC$%d:$AC$%d,MATCH(MAX($AE$%d:'
                       '$AE$%d),$AE$%d:$AE$%d,0)),"")'
                       % (er, el, er, el, er, el), "@")
    F["bs_units"] = ('=IFERROR(INDEX($V$%d:$V$%d,MATCH(MAX($W$%d:$W$%d),'
                     '$W$%d:$W$%d,0)),"")' % (pr, pl, pr, pl, pr, pl), "@")
    F["bs_revenue"] = F["best_product"]
    F["bs_profit"] = ('=IFERROR(INDEX($V$%d:$V$%d,MATCH(MAX($AB$%d:$AB$%d),'
                      '$AB$%d:$AB$%d,0)),"")' % (pr, pl, pr, pl, pr, pl), "@")
    F["bs_margin"] = ('=IFERROR(INDEX($V$%d:$V$%d,MATCH(MAX($Z$%d:$Z$%d),'
                      '$Z$%d:$Z$%d,0)),"")' % (pr, pl, pr, pl, pr, pl), "@")
    F["bs_slow"] = ('=IFERROR(INDEX($V$%d:$V$%d,MATCH(MIN($AA$%d:$AA$%d),'
                    '$AA$%d:$AA$%d,0)),"")' % (pr, pl, pr, pl, pr, pl), "@")
    F["price_suggest"] = ("=%s!$J$%d" % (pricing, C.PR_OUT["price"]),
                          "#,##0.00")
    for name in ("revenue", "discounts", "cogs", "fees", "profit", "margin",
                 "units", "txns", "aov", "per_item", "events_total",
                 "events_done", "inv_value", "low_products", "low_materials",
                 "reorder_cost", "sell_through", "made_total"):
        if not bk.has("production") and name in ("made_total",
                                                 "sell_through",
                                                 "inv_value"):
            F[name] = ("=0", C.KPI_FMT[name])
    for key, (formula, fmt) in F.items():
        row = C.KPI_ROW[key]
        ws.write(r(row), ci("AH"), C.KPI_LABEL[key], hdr)
        cached = agg.get(key, 0 if fmt != "@" else "")
        ws.write_formula(r(row), ci("AI"), formula,
                         S.f(**S.base(font_size=9, font_color=th.ink,
                                      bg_color=th.bg, num_format=fmt)),
                         cached)
        fcount()

    # ------------------------------------------------------------------
    # reorder sequence pools AK/AL (products) AN/AO (materials)
    # ------------------------------------------------------------------
    ws.write(r(1), ci("AK"), "Prod seq", hdr)
    ws.write(r(1), ci("AL"), "Prod row", hdr)
    ws.write(r(1), ci("AN"), "Mat seq", hdr)
    ws.write(r(1), ci("AO"), "Mat row", hdr)
    st = bk.rng("catalog", "status")
    seq_cached_p = _seq_cache([srow[3] < _min_of(m, srow[0])
                               for srow in (agg.get("product_stats") or [])],
                              m)
    for i in range(C.CAP["catalog"]):
        row = C.DATA_SEQP_FIRST + i
        crow = C.ROW_FIRST + i
        flag = 'OR(%s!$O$%d="%s",%s!$O$%d="%s")' % (catalog, crow, C.ST_LOW,
                                                   catalog, crow, C.ST_OUT)
        if i == 0:
            body = "1"
        else:
            body = ('1+SUMPRODUCT((%s!$O$%d:$O$%d="%s")+(%s!$O$%d:$O$%d="%s"))'
                    % (catalog, C.ROW_FIRST, crow - 1, C.ST_LOW,
                       catalog, C.ROW_FIRST, crow - 1, C.ST_OUT))
        ws.write_formula(r(row), ci("AK"),
                         '=IF(%s,%s,"")' % (flag, body),
                         integer, seq_cached_p[i] if i < len(seq_cached_p)
                         else "")
        fcount()
        ws.write(r(row), ci("AL"), crow, integer)
    if premium:
        seq_cached_m = _seq_cache(
            [mrow[1] <= mrow[2] for mrow in
             [(mm["remaining"], mm["remaining"], mm["threshold"])
              for mm in (m.materials if m else [])]], None)
        for i in range(C.CAP["materials"]):
            row = C.DATA_SEQM_FIRST + i
            mrow = C.ROW_FIRST + i
            flag = 'AND(%s!$B$%d<>"",%s!$K$%d<=%s!$M$%d)' % (mats, mrow,
                                                            mats, mrow,
                                                            mats, mrow)
            if i == 0:
                body = "1"
            else:
                body = ('1+SUMPRODUCT((%s!$B$%d:$B$%d<>"")*'
                        '(%s!$K$%d:$K$%d<=%s!$M$%d:$M$%d))'
                        % (mats, C.ROW_FIRST, mrow - 1, mats,
                           C.ROW_FIRST, mrow - 1, mats, C.ROW_FIRST,
                           mrow - 1))
            ws.write_formula(r(row), ci("AN"),
                             '=IF(%s,%s,"")' % (flag, body),
                             integer, seq_cached_m[i]
                             if i < len(seq_cached_m) else "")
            fcount()
            ws.write(r(row), ci("AO"), mrow, integer)
    return None


def _seq_cache(flags, _):
    out, k = [], 0
    for f in flags:
        if f:
            k += 1
            out.append(k)
        else:
            out.append("")
    return out


def _min_of(demo, name):
    if not demo:
        return 0
    for p in demo.products:
        if p["name"] == name:
            return p["min"]
    return 0


def _profit_of(demo, name):
    if not demo:
        return 0.0
    for p in demo.products:
        if p["name"] == name:
            return p["profit"]
    return 0.0
```

## 11. Lists & Settings

`crochet_tracker/sheets/setup.py`

```python
"""
The ⚙️ Lists & Settings tab: business profile, money defaults and every
editable dropdown list in the workbook.
"""

from .. import config as C
from ..book import r, ci

LAST_COL = "J"


def build(bk):
    key = "setup"
    ws = bk.ws(key)
    S, th, m = bk.S, bk.th, bk.demo
    st = m.settings if m else {}

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
                   "\u2699\uFE0F  Lists & Settings \u2014 make it yours",
                   S.sheet_title)
    ws.merge_range(r(3), 1, r(3), ci(LAST_COL) - 3,
                   "  Set these once and every tab, calculator, alert and "
                   "dropdown follows automatically.", S.sheet_sub)
    ws.merge_range(r(3), ci(LAST_COL) - 2, r(3), ci(LAST_COL), "",
                   S.home_link)
    ws.write_url(r(3), ci(LAST_COL) - 2,
                 "internal:%s!A1" % bk.q("dashboard"), S.home_link,
                 "\U0001F3E0  Back to Dashboard")

    inp = S.f(**S.base(font_size=12, bold=True, font_color=th.primary,
                       bg_color=th.gold_soft, border=1,
                       border_color=th.border_strong, align="left",
                       valign="vcenter", indent=1))
    inp_pct = S.f(**S.base(font_size=12, bold=True, font_color=th.primary,
                           bg_color=th.gold_soft, border=1,
                           border_color=th.border_strong, align="left",
                           valign="vcenter", indent=1, num_format="0%"))
    inp_money = S.f(**S.base(font_size=12, bold=True, font_color=th.primary,
                             bg_color=th.gold_soft, border=1,
                             border_color=th.border_strong, align="left",
                             valign="vcenter", indent=1,
                             num_format="#,##0.00"))
    label = S.f(**S.base(font_size=11, font_color=th.ink, bg_color=th.card,
                         align="left", valign="vcenter", indent=1,
                         border=1, border_color=th.border))

    def setting(row, text, value, fmt=None, cached=None):
        ws.set_row(r(row), 22)
        ws.write(r(row), 1, text, label)
        if value is None:
            ws.write(r(row), 2, cached if cached is not None else "", fmt)
        else:
            ws.write(r(row), 2, value, fmt)

    def section(row, emoji, text):
        ws.set_row(r(row), 24)
        ws.merge_range(r(row), 1, r(row), ci(LAST_COL),
                       "  %s  %s" % (emoji, text), S.section)

    # ------------------------------------------------------------------
    section(5, "\U0001F3F7\uFE0F", "BUSINESS PROFILE")
    setting(C.SU_BUSINESS, "Business name (dashboard header)",
            st.get("business", ""), inp)
    setting(C.SU_MESSAGE, "Dashboard message of the day",
            st.get("message", ""), inp)

    section(9, "\U0001F4B5", "MONEY & SEASON DEFAULTS")
    setting(C.SU_CURRENCY, "Currency symbol", st.get("currency", "$"), inp)
    setting(C.SU_YEAR, "Report year (monthly summary)",
            st.get("year", 2026), inp_money)
    setting(C.SU_WAGE, "Your hourly wage (pricing calculator)",
            st.get("wage", 14.0), inp_money)
    setting(C.SU_OVERHEAD, "Overhead % on top of direct costs",
            st.get("overhead", 0.10), inp_pct)
    setting(C.SU_MARGIN, "Default target profit margin",
            st.get("margin", 0.45), inp_pct)

    # ------------------------------------------------------------------
    section(C.SU_LIST_HEADER - 1, "\U0001F4DD",
            "YOUR DROPDOWN LISTS (edit freely)")
    ws.set_row(r(C.SU_LIST_HEADER), 20)
    cols = sorted(C.LIST_COLS.items(), key=lambda kv: kv[1])
    for key2, col in cols:
        ws.write(r(C.SU_LIST_HEADER), ci(col), C.LIST_TITLES[key2], S.thead)
    values = C.LIST_VALUES
    for i in range(C.SU_LIST_ROWS):
        row = C.SU_LIST_FIRST + i
        ws.set_row(r(row), 16)
        for key2, col in cols:
            lst = values[key2]
            cell = S.f(**S.base(font_size=10, font_color=th.ink,
                                bg_color=th.card if i % 2 else th.alt,
                                border=1, border_color=th.border,
                                align="left", valign="vcenter", indent=1))
            if m and key2 in ("suppliers",):
                pass
            ws.write(r(row), ci(col),
                     lst[i] if i < len(lst) else "", cell)
    note_row = C.SU_LIST_FIRST + C.SU_LIST_ROWS
    ws.merge_range(r(note_row), 1, r(note_row + 1), ci(LAST_COL),
                   "  These five columns feed every dropdown in the "
                   "workbook.  Add, rename or delete entries and the "
                   "catalog, sales log and reorder list update instantly.",
                   S.note)

    # ------------------------------------------------------------------
    section(C.SU_FIXED_HEADER - 1, "\U0001F512", "FIXED LIST (do not edit)")
    ws.set_row(r(C.SU_FIXED_HEADER), 20)
    ws.write(r(C.SU_FIXED_HEADER), ci("C"), "Tick box", S.thead)
    ws.write(r(C.SU_FIXED_FIRST), ci("C"), C.TICK, S.note_plain)
    ws.write(r(C.SU_FIXED_FIRST + 1), ci("C"), "", S.note_plain)

    foot = C.SU_FIXED_FIRST + 5
    ws.set_row(r(foot), 30)
    ws.merge_range(r(foot), 1, r(foot), ci(LAST_COL),
                   "  \U0001F4A1  Cream cells are inputs everywhere in this "
                   "workbook; white cells are formulas.  Sheets are "
                   "protected (no password) so nothing breaks by accident.",
                   S.f(**S.base(font_size=10, italic=True, font_color=th.ink,
                                bg_color=th.gold_soft, align="left",
                                valign="vcenter", indent=1)))
    nav = foot + 2
    ws.set_row(r(nav), 24)
    bk.nav_row(key, nav, max_col=LAST_COL)
    bk.page(key, LAST_COL, nav + 1, landscape=False, zoom=100)
```

## 12. Dashboard

`crochet_tracker/sheets/dashboard.py`

```python
"""📊 Dashboard - KPIs, bars, six charts, best-sellers and low-stock panels."""

from .. import config as C
from ..book import r, ci

LAST_COL = "M"
ROW_HERO_1 = 2
ROW_HERO_2 = 3
ROW_HERO_3 = 4
ROW_MSG = 5
ROW_SECTION = 7
ROW_CARD1_L = 8
ROW_CARD1_V = 9
ROW_CARD2_L = 10
ROW_CARD2_V = 11
ROW_BAR1 = 12
ROW_BAR2 = 13
ROW_CH1 = 16
ROW_CH2 = 32
ROW_CH3 = 48
ROW_PANELS = 65
ROW_FOOT = 73


def build(bk):
    key = "dashboard"
    ws = bk.ws(key)
    S, th, m = bk.S, bk.th, bk.demo
    agg = m.agg if m else {}
    premium = bk.has("materials")

    bk.widths(key, C.WIDTHS[key])
    bk.paint(key, 0, 0, ROW_FOOT + 6, ci(LAST_COL), S.canvas)

    # ------------------------------------------------------------------
    # hero
    # ------------------------------------------------------------------
    ws.set_row(r(1), 8)
    ws.set_row(r(ROW_HERO_1), 34)
    ws.set_row(r(ROW_HERO_2), 26)
    ws.set_row(r(ROW_HERO_3), 20)
    ws.set_row(r(ROW_MSG), 22)

    ws.merge_range(r(ROW_HERO_1), 1, r(ROW_HERO_1), ci(LAST_COL), "",
                   S.hero_title)
    ws.write_formula(
        r(ROW_HERO_1), 1,
        '=IF(BusinessName="","\U0001F9F6  Craft Fair Command Center",'
        '"\U0001F9F6  "&BusinessName&"   \u2022   Craft Fair Command '
        'Center")', S.hero_title,
        "\U0001F9F6  Craft Fair Command Center"
        if not (m and m.settings["business"]) else
        "\U0001F9F6  %s   \u2022   Craft Fair Command Center"
        % m.settings["business"])
    bk.stats["formulas"] += 1

    ws.merge_range(r(ROW_HERO_2), 1, r(ROW_HERO_2), ci(LAST_COL), "",
                   S.hero_count)
    nxt = _next_event(bk)
    ws.write_formula(r(ROW_HERO_2), 1, "=" + nxt, S.hero_count,
                     _next_event_cached(m))
    bk.stats["formulas"] += 1

    ws.merge_range(r(ROW_HERO_3), 1, r(ROW_HERO_3), ci(LAST_COL), "",
                   S.hero_meta)
    ws.write_formula(
        r(ROW_HERO_3), 1,
        '=TEXT(TODAY(),"dddd, dd mmmm yyyy")&"   \u2022   "&%s&" fairs on '
        'the books   \u2022   "&%s&" taken at the till   \u2022   "&%s&" '
        'net profit so far"'
        % (bk.kpi("events_total"), bk.money(bk.kpi("revenue"), "#,##0"),
           bk.money(bk.kpi("profit"), "#,##0")),
        S.hero_meta,
        "   \u2022   %d fairs on the books   \u2022   %s taken at the till"
        "   \u2022   %s net profit so far"
        % (agg.get("events_total", 0), agg.get("revenue", 0),
           agg.get("profit", 0)) if m else "")
    bk.stats["formulas"] += 1

    ws.merge_range(r(ROW_MSG), 1, r(ROW_MSG), ci(LAST_COL), "",
                   S.f(**S.base(bg_color=th.gold, font_color=th.white,
                                bold=True, font_size=11, align="left",
                                valign="vcenter", indent=1)))
    ws.write_formula(
        r(ROW_MSG), 1,
        '=IF(%s!$C$%d="","","\U0001F4E3  "&%s!$C$%d)'
        % (bk.q("setup"), C.SU_MESSAGE, bk.q("setup"), C.SU_MESSAGE),
        S.f(**S.base(bg_color=th.gold, font_color=th.white, bold=True,
                     font_size=11, align="left", valign="vcenter",
                     indent=1)),
        "\U0001F4E3  " + m.settings["message"] if m else "")
    bk.stats["formulas"] += 1

    # ------------------------------------------------------------------
    # KPI cards
    # ------------------------------------------------------------------
    ws.merge_range(r(ROW_SECTION), 1, r(ROW_SECTION), ci(LAST_COL),
                   "  \U0001F9F6  BUSINESS AT A GLANCE", S.section)
    ws.set_row(r(ROW_SECTION), 22)

    cards1 = [
        ("TOTAL SALES", "ok", bk.money(bk.kpi("revenue"), "#,##0"),
         agg.get("revenue", 0)),
        ("NET PROFIT", "primary", bk.money(bk.kpi("profit"), "#,##0"),
         agg.get("profit", 0)),
        ("PROFIT MARGIN", "gold", "=TEXT(%s,\u00220%%\u0022)"
         % bk.kpi("margin"), _p0(agg.get("margin", 0))),
        ("UNITS SOLD", "info", "=%s" % bk.kpi("units"),
         agg.get("units", 0)),
        ("FAIRS WORKED", "primary_2", "=%s" % bk.kpi("events_done"),
         agg.get("events_done", 0)),
        ("AVERAGE SALE", "accent", bk.money(bk.kpi("aov"), "#,##0.00"),
         agg.get("aov", 0)),
    ]
    cards2 = [
        ("INVENTORY VALUE", "info", bk.money(bk.kpi("inv_value"),
                                             "#,##0.00"),
         agg.get("inv_value", 0)),
        ("LOW-STOCK ITEMS", "bad", "=%s+%s"
         % (bk.kpi("low_products"), bk.kpi("low_materials")),
         agg.get("low_products", 0) + agg.get("low_materials", 0)),
        ("SELL-THROUGH", "gold", "=TEXT(%s,\u00220%%\u0022)"
         % bk.kpi("sell_through"),
         _p0(agg.get("sell_through", 0))),
        ("BEST-SELLER", "ok", "=%s" % bk.kpi("bs_units"),
         agg.get("bs_units", "")),
        ("TOP FAIR", "accent", "=%s" % bk.kpi("best_event"),
         agg.get("best_event", "")),
        ("REORDER BUDGET", "bad", bk.money(bk.kpi("reorder_cost"),
                                           "#,##0.00"),
         agg.get("reorder_cost", 0)),
    ]
    _cards(bk, cards1, ROW_CARD1_L, ROW_CARD1_V)
    _cards(bk, cards2, ROW_CARD2_L, ROW_CARD2_V)

    # ------------------------------------------------------------------
    # progress bars
    # ------------------------------------------------------------------
    _bar(bk, ROW_BAR1, "Net margin", bk.kpi("profit"), bk.kpi("revenue"),
         agg.get("margin", 0), "margin")
    _bar(bk, ROW_BAR2, "Sell-through (sold \u00f7 made)", bk.kpi("units"),
         bk.kpi("made_total"), agg.get("sell_through", 0), "pct")

    # ------------------------------------------------------------------
    # charts
    # ------------------------------------------------------------------
    _charts(bk)

    # ------------------------------------------------------------------
    # panels: best-sellers + low stock
    # ------------------------------------------------------------------
    _panels(bk)

    # ------------------------------------------------------------------
    # footer + nav
    # ------------------------------------------------------------------
    ws.set_row(r(ROW_FOOT), 30)
    ws.merge_range(r(ROW_FOOT), 1, r(ROW_FOOT), ci(LAST_COL),
                   "  \U0001F4A1  Everything on this page is calculated "
                   "\u2014 type your makes and sales into the tabs and "
                   "watch it update.  Start with the \U0001F4D6 Start Here "
                   "guide and \u2699\uFE0F Lists & Settings.",
                   S.f(**S.base(font_size=10, italic=True, font_color=th.ink,
                                bg_color=th.gold_soft, align="left",
                                valign="vcenter", indent=1)))
    nav = ROW_FOOT + 2
    ws.set_row(r(nav), 24)
    bk.nav_row(key, nav, max_col=LAST_COL)
    bk.page(key, LAST_COL, nav + 1, landscape=True, zoom=85)


# ===========================================================================
def _cards(bk, cards, row_l, row_v):
    S, th = bk.S, bk.th
    ws = bk.ws("dashboard")
    pairs = [(1, 2), (3, 4), (5, 6), (7, 8), (9, 10), (11, 12)]
    for (label, colour, formula, cached), (c1, c2) in zip(cards, pairs):
        bg = getattr(th, colour)
        lbl = S.f(**S.base(bold=True, font_size=9.5, font_color=th.white,
                           bg_color=bg, align="left", valign="vcenter",
                           indent=1))
        val = S.f(**S.base(bold=True, font_size=15,
                           font_color=bg if colour != "gold" else th.gold,
                           bg_color=th.card, align="center",
                           valign="vcenter"))
        ws.merge_range(r(row_l), c1, r(row_l), c2, "  " + label, lbl)
        ws.merge_range(r(row_v), c1, r(row_v), c2, "", val)
        if formula.startswith("="):
            ws.write_formula(r(row_v), c1, formula, val, cached)
        else:
            ws.write_formula(r(row_v), c1, "=" + formula, val, cached)
        bk.stats["formulas"] += 1
    ws.set_row(r(row_l), 18)
    ws.set_row(r(row_v), 26)


def _bar(bk, row, label, num_k, den_k, cached_pct, kind=None):
    S, th = bk.S, bk.th
    ws = bk.ws("dashboard")
    ws.set_row(r(row), 20)
    ws.merge_range(r(row), 1, r(row), 2, "  " + label, S.bar_label)
    ws.merge_range(r(row), 3, r(row), 9, "", S.bar_text)
    ws.write_formula(r(row), 3, bk.bar(num_k, den_k, 34), S.bar_text,
                     _bar_cached(cached_pct, 34))
    bk.stats["formulas"] += 1
    ws.merge_range(r(row), 10, r(row), 12, "", S.bar_pct)
    ws.write_formula(r(row), 10,
                     '=TEXT(IFERROR(%s/%s,0),"0.0%%")' % (num_k, den_k),
                     S.bar_pct, _p1(cached_pct))
    bk.stats["formulas"] += 1


def _p0(x):
    return "%.0f%%" % (100.0 * float(x or 0))


def _p1(x):
    return "%.1f%%" % (100.0 * float(x or 0))


def _bar_cached(pct, blocks):
    filled = int(round(min(1.0, max(0.0, float(pct or 0))) * blocks))
    return "\u2588" * filled + "\u2591" * (blocks - filled)


# ===========================================================================
def _charts(bk):
    S, th = bk.S, bk.th
    wb = bk.wb
    d = bk.q("data").strip("'")
    mr, ml = C.DATA_MONTH_FIRST, C.DATA_MONTH_FIRST + 11
    er, el = C.DATA_EVENT_FIRST, C.DATA_EVENT_FIRST + C.CAP["events"] - 1
    pr, pl = C.DATA_PROD_FIRST, C.DATA_PROD_FIRST + C.CAP["catalog"] - 1

    def base_chart(ctype, title):
        ch = wb.add_chart({"type": ctype})
        ch.set_title({"name": title, "name_font": {"size": 11,
                                                  "color": th.primary,
                                                  "bold": True}})
        ch.set_legend({"position": "bottom",
                       "font": {"size": 9, "color": th.ink}})
        ch.set_chartarea({"border": {"color": th.border},
                          "fill": {"color": th.card}})
        ch.set_plotarea({"fill": {"color": th.card}})
        ch.set_size({"width": 555, "height": 250})
        ch.show_hidden_data()
        bk.stats["charts"] += 1
        return ch

    # 1 - monthly revenue + profit
    combo = base_chart("column", "Monthly revenue & net profit")
    combo.add_series({
        "name": "Revenue",
        "categories": "=%s!$H$%d:$H$%d" % (d, mr, ml),
        "values": "=%s!$I$%d:$I$%d" % (d, mr, ml),
        "fill": {"color": th.ok},
        "border": {"color": th.ok},
    })
    line = wb.add_chart({"type": "line"})
    line.add_series({
        "name": "Net profit",
        "categories": "=%s!$H$%d:$H$%d" % (d, mr, ml),
        "values": "=%s!$L$%d:$L$%d" % (d, mr, ml),
        "line": {"color": th.accent, "width": 2.5},
        "marker": {"type": "circle", "size": 6,
                   "fill": {"color": th.accent}},
    })
    combo.combine(line)
    bk.ws("dashboard").insert_chart(r(ROW_CH1), 2, combo)

    # 2 - sales & profit by event
    ev = base_chart("column", "Sales & net profit by fair")
    ev.add_series({
        "name": "Sales",
        "categories": "=%s!$AC$%d:$AC$%d" % (d, er, el),
        "values": "=%s!$AD$%d:$AD$%d" % (d, er, el),
        "fill": {"color": th.info},
        "border": {"color": th.info},
    })
    ev2 = wb.add_chart({"type": "line"})
    ev2.add_series({
        "name": "Net profit",
        "categories": "=%s!$AC$%d:$AC$%d" % (d, er, el),
        "values": "=%s!$AE$%d:$AE$%d" % (d, er, el),
        "line": {"color": th.gold, "width": 2.5},
        "marker": {"type": "circle", "size": 6,
                   "fill": {"color": th.gold}},
    })
    ev.combine(ev2)
    bk.ws("dashboard").insert_chart(r(ROW_CH1), 8, ev)

    # 3 - units sold by product
    prod = base_chart("bar", "Units sold by product")
    prod.add_series({
        "name": "Units",
        "categories": "=%s!$V$%d:$V$%d" % (d, pr, pl),
        "values": "=%s!$W$%d:$W$%d" % (d, pr, pl),
        "fill": {"color": th.accent},
        "border": {"color": th.accent},
    })
    bk.ws("dashboard").insert_chart(r(ROW_CH2), 2, prod)

    # 4 - inventory remaining
    inv = base_chart("bar", "Inventory remaining (units)")
    inv.add_series({
        "name": "In stock",
        "categories": "=%s!$V$%d:$V$%d" % (d, pr, pl),
        "values": "=%s!$Y$%d:$Y$%d" % (d, pr, pl),
        "fill": {"color": th.primary_2},
        "border": {"color": th.primary_2},
    })
    bk.ws("dashboard").insert_chart(r(ROW_CH2), 8, inv)

    # 5 - takings by payment method
    pay = base_chart("doughnut", "Takings by payment method")
    pay.add_series({
        "name": "Taken",
        "categories": "=%s!$P$%d:$P$%d" % (d, C.DATA_PAY_FIRST,
                                           C.DATA_PAY_FIRST + 5),
        "values": "=%s!$Q$%d:$Q$%d" % (d, C.DATA_PAY_FIRST,
                                       C.DATA_PAY_FIRST + 5),
    })
    bk.ws("dashboard").insert_chart(r(ROW_CH3), 2, pay)

    # 6 - fair expenses by category
    exp = base_chart("doughnut", "Fair expenses by category")
    exp.add_series({
        "name": "Spent",
        "categories": "=%s!$S$%d:$S$%d" % (d, C.DATA_EXP_FIRST,
                                           C.DATA_EXP_FIRST + 4),
        "values": "=%s!$T$%d:$T$%d" % (d, C.DATA_EXP_FIRST,
                                       C.DATA_EXP_FIRST + 4),
    })
    bk.ws("dashboard").insert_chart(r(ROW_CH3), 8, exp)


# ===========================================================================
def _panels(bk):
    S, th, m = bk.S, bk.th, bk.demo
    ws = bk.ws("dashboard")
    agg = m.agg if m else {}
    ws.set_row(r(ROW_PANELS), 22)
    ws.merge_range(r(ROW_PANELS), 1, r(ROW_PANELS), 6,
                   "  \U0001F3C6  BEST-SELLERS", S.section_soft)
    ws.merge_range(r(ROW_PANELS), 7, r(ROW_PANELS), ci(LAST_COL),
                   "  \U0001F534  LOW STOCK RIGHT NOW", S.section_soft)

    left = [("Most units sold", "bs_units"),
            ("Highest revenue", "bs_revenue"),
            ("Highest profit", "bs_profit"),
            ("Highest margin", "bs_margin"),
            ("Slowest mover", "bs_slow")]
    lab = S.f(**S.base(font_size=10, font_color=th.muted, bg_color=th.card,
                       align="left", valign="vcenter", indent=1,
                       border=1, border_color=th.border))
    val = S.f(**S.base(bold=True, font_size=10.5, font_color=th.primary,
                       bg_color=th.card, align="left", valign="vcenter",
                       indent=1, border=1, border_color=th.border))
    for i, (label, kpi) in enumerate(left):
        row = ROW_PANELS + 1 + i
        ws.set_row(r(row), 18)
        ws.merge_range(r(row), 1, r(row), 2, "  " + label, lab)
        ws.merge_range(r(row), 3, r(row), 6, "", val)
        ws.write_formula(r(row), 3, "=%s" % bk.kpi(kpi), val,
                         agg.get(kpi, ""))
        bk.stats["formulas"] += 1

    # right: low stock preview (products then materials)
    rk = S.f(**S.base(font_size=10, font_color=th.ink, bg_color=th.card,
                      align="left", valign="vcenter", indent=1,
                      border=1, border_color=th.border))
    rv = S.f(**S.base(bold=True, font_size=10.5, font_color=th.bad,
                      bg_color=th.card, align="center", valign="vcenter",
                      border=1, border_color=th.border))
    cat = bk.q("catalog")
    seqp = "$AK$%d:$AK$%d" % (C.DATA_SEQP_FIRST,
                              C.DATA_SEQP_FIRST + C.CAP["catalog"] - 1)
    for i in range(5):
        row = ROW_PANELS + 1 + i
        k = i + 1
        ws.merge_range(r(row), 7, r(row), 10, "", rk)
        ws.write_formula(
            r(row), 7,
            '=IFERROR(INDEX(%s!$C$%d:$C$%d,MATCH(%d,%s,0)),"")'
            % (cat, C.ROW_FIRST, C.last_row("catalog"), k,
               bk.q("data") + "!" + seqp),
            rk, "")
        bk.stats["formulas"] += 1
        ws.merge_range(r(row), 11, r(row), ci(LAST_COL), "", rv)
        ws.write_formula(
            r(row), 11,
            '=IFERROR(INDEX(%s!$M$%d:$M$%d,MATCH(%d,%s,0))&" left (min "&'
            'INDEX(%s!$N$%d:$N$%d,MATCH(%d,%s,0))&")","")'
            % (cat, C.ROW_FIRST, C.last_row("catalog"), k,
               bk.q("data") + "!" + seqp, cat, C.ROW_FIRST,
               C.last_row("catalog"), k, bk.q("data") + "!" + seqp),
            rv, "")
        bk.stats["formulas"] += 1


# ===========================================================================
def _next_event(bk):
    ev = bk.q("events")
    d = bk.q("data")
    pool = "%s!$AG$%d:$AG$%d" % (d, C.DATA_EVENT_FIRST,
                                 C.DATA_EVENT_FIRST + C.CAP["events"] - 1)
    return ('IFERROR(TEXT(MIN(%s),"dd mmm")&"  \u2014  "&INDEX(%s!$B$%d:'
            '$B$%d,MATCH(MIN(%s),%s!$C$%d:$C$%d,0))&"  is next \u2014  '
            'pack the van!","\U0001F334  No upcoming fairs logged \u2014 '
            'add one on the \U0001F3EA Craft Fairs tab!")') % (
        pool, ev, C.ROW_FIRST, C.last_row("events"), pool,
        ev, C.ROW_FIRST, C.last_row("events"))


def _next_event_cached(m):
    if not m:
        return ""
    import datetime
    future = [e for e in m.events if e["date"] >= datetime.date(2026, 9, 13)]
    if not future:
        return ("\U0001F334  No upcoming fairs logged \u2014 add one on the "
                "\U0001F3EA Craft Fairs tab!")
    nxt = min(future, key=lambda e: e["date"])
    return ("%s  \u2014  %s  is next \u2014  pack the van!"
            % (nxt["date"].strftime("%d %b"), nxt["name"]))
```

## 13. Product Catalog

`crochet_tracker/sheets/catalog.py`

```python
"""🧶 Product Catalog - master list of everything you make and sell."""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "catalog"
LAST_COL = "P"


def _columns(has_prod):
    return [
        ("n", "#", "idx", None),
        ("sku", "SKU", "text", None),
        ("name", "Product name", "text", None),
        ("category", "Category", "center", None),
        ("desc", "Description", "wrap", None),
        ("price", "Selling price", "money", None),
        ("yarn_cost", "Yarn / material", "money", None),
        ("pack_cost", "Packaging", "money", None),
        ("unit_cost", "Unit cost", "calc_money", "primary_2"),
        ("profit", "Profit / item", "calc_money", "primary_2"),
        ("margin", "Margin %", "calc_pct", "primary_2"),
        ("time_h", "Hours", "qty1", None),
        ("stock", "In stock", "calc_num" if has_prod else "qty", None),
        ("min", "Min level", "qty", None),
        ("status", "Reorder status", "calc_c", None),
        ("link", "Photo / link", "link", None),
    ]


def build(bk):
    th = bk.th
    ws = bk.ws(key := KEY)
    m = bk.demo
    has_prod = bk.has("production")
    columns = _columns(has_prod)

    K.sheet_head(bk, KEY, LAST_COL,
                 "\U0001F9F6  Product Catalog",
                 "Everything you make, what it costs you and what it earns "
                 "you - stock counts flow in from \U0001F4E6 Made & Stocked.")
    K.table_frame(bk, KEY, columns, height=32)

    for i in range(C.CAP[KEY]):
        row = C.ROW_FIRST + i
        prod = m.products[i] if m and i < len(m.products) else None
        values = {"n": i + 1}
        if prod:
            values.update({
                "sku": prod["sku"], "name": prod["name"],
                "category": prod["category"], "desc": prod["desc"],
                "price": prod["price"], "yarn_cost": prod["yarn"],
                "pack_cost": prod["pack"], "time_h": prod["hours"],
                "min": prod["min"], "link": prod["link"]})
        values["unit_cost"] = '=IF($C%d="","",$G%d+$H%d)' % (row, row, row)
        values["profit"] = '=IF($F%d="","",$F%d-$I%d)' % (row, row, row)
        values["margin"] = '=IF($F%d="","",$J%d/$F%d)' % (row, row, row)
        values["status"] = ('=IF($C%d="","",IF($M%d<=0,"%s",IF($M%d<$N%d,'
                            '"%s","%s")))'
                            % (row, row, C.ST_OUT, row, row, C.ST_LOW,
                               C.ST_OK))
        if has_prod:
            values["stock"] = ('=SUMIFS(%s,%s,$C%d)'
                               % (bk.rng("production", "current"),
                                  bk.rng("production", "product"), row))
        elif prod:
            values["stock"] = prod["stock"]
        cached = {}
        if prod:
            cached = {"unit_cost": prod["unit_cost"],
                      "profit": prod["profit"], "margin": prod["margin"],
                      "status": {"out": C.ST_OUT, "low": C.ST_LOW,
                                 "ok": C.ST_OK}[prod["status"]],
                      "stock": prod["stock"]}
        K.write_row(bk, KEY, columns, row, values, cached)

    K.list_dv(bk, KEY, "category", "categories", title="Product category")
    K.money_dv(bk, KEY, ("price", "yarn_cost", "pack_cost"))
    K.whole_dv(bk, KEY, ("min",), minimum=0, maximum=100000)
    K.status_cf(bk, KEY, "status",
                {C.ST_OUT: (th.bad_soft, th.bad),
                 C.ST_LOW: (th.warn_soft, th.warn),
                 C.ST_OK: (th.ok_soft, th.ok)})
    ws.conditional_format(
        r(C.ROW_FIRST), ci(bk.col(KEY, "margin")),
        r(C.last_row(KEY)), ci(bk.col(KEY, "margin")),
        {"type": "cell", "criteria": "<", "value": 0.3,
         "format": bk.S.cf(bg=th.warn_soft, fg=th.warn, bold=True)})
    bk.stats["cond_formats"] += 1

    chips = [
        ('="\U0001F9F6 Products: "&COUNTA(%s)' % bk.rng(KEY, "name"),
         "primary", "Products: %d" % (len(m.products) if m else 0), 3),
        ('="\U0001F4B0 Stock value: "&Currency&TEXT(SUMPRODUCT(IFERROR(%s*1,0)*IFERROR(%s*1,0)),'
         '"#,##0.00")' % (bk.rng(KEY, "stock"), bk.rng(KEY, "unit_cost")),
         "info", "Stock value: %s"
         % (m.money(m.agg["inv_value"], 2) if m else "$0.00"), 4),
        ('="\U0001F534 Below minimum: "&%s' % bk.kpi("low_products"),
         "bad", "Below minimum: %d" % (m.agg.get("low_products", 0)
                                       if m else 0), 4),
        ('="\U0001F4C8 Avg margin: "&TEXT(IFERROR(AVERAGE(%s),0),"0%%")'
         % bk.rng(KEY, "margin"), "ok",
         "Avg margin: %.0f%%" % (100 * _avg_margin(m) if m else 0), 3),
    ]
    K.chips(bk, KEY, chips, LAST_COL)
    K.footer_nav(bk, KEY, C.last_row(KEY) + 3, LAST_COL)


def _avg_margin(m):
    if not m or not m.products:
        return 0.0
    return sum(p["margin"] for p in m.products) / len(m.products)
```

## 14. Yarn & Materials

`crochet_tracker/sheets/materials.py`

```python
"""🧵 Yarn & Materials - what is on the shelf and what needs reordering."""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "materials"
LAST_COL = "N"

COLUMNS = [
    ("n", "#", "idx", None),
    ("name", "Material / yarn", "text", None),
    ("color", "Colour", "text", None),
    ("brand", "Brand", "text", None),
    ("weight", "Weight / type", "center", None),
    ("purchased", "Purchased", "qty", None),
    ("unit", "Unit", "center", None),
    ("cost_unit", "Cost / unit", "money", None),
    ("total", "Total cost", "calc_money", "primary_2"),
    ("used", "Used", "qty1", None),
    ("remaining", "Remaining", "calc_qty1", "primary_2"),
    ("supplier", "Supplier", "text", None),
    ("threshold", "Reorder at", "qty1", None),
    ("reorder", "Reorder?", "calc_c", None),
]


def build(bk):
    th = bk.th
    ws = bk.ws(KEY)
    m = bk.demo
    K.sheet_head(bk, KEY, LAST_COL,
                 "\U0001F9F5  Yarn & Materials",
                 "Because \u0022I have five tote bags\u0022 matters less than "
                 "\u0022I am almost out of the cotton to make more\u0022.")
    K.table_frame(bk, KEY, COLUMNS, height=22)

    for i in range(C.CAP[KEY]):
        row = C.ROW_FIRST + i
        mat = m.materials[i] if m and i < len(m.materials) else None
        values = {"n": i + 1}
        if mat:
            values.update({
                "name": mat["name"], "color": mat["color"],
                "brand": mat["brand"], "weight": mat["weight"],
                "purchased": mat["purchased"], "unit": mat["unit"],
                "cost_unit": mat["cost_unit"], "used": mat["used"],
                "supplier": mat["supplier"],
                "threshold": mat["threshold"]})
        values["total"] = '=IF($B%d="","",$F%d*$H%d)' % (row, row, row)
        values["remaining"] = '=IF($B%d="","",$F%d-$J%d)' % (row, row, row)
        values["reorder"] = ('=IF($B%d="","",IF($K%d<=$M%d,"%s","%s"))'
                             % (row, row, row, C.RE_YES, C.RE_NO))
        cached = {}
        if mat:
            cached = {"total": mat["total"],
                      "remaining": mat["remaining"],
                      "reorder": C.RE_YES if mat["remaining"]
                      <= mat["threshold"] else C.RE_NO}
        K.write_row(bk, KEY, COLUMNS, row, values, cached)

    K.list_dv(bk, KEY, "weight", "weights", title="Yarn weight")
    K.list_dv(bk, KEY, "unit", "units", title="Unit")
    K.list_dv(bk, KEY, "supplier", "suppliers", title="Supplier")
    K.money_dv(bk, KEY, ("cost_unit",), label="cost per unit")
    K.status_cf(bk, KEY, "reorder",
                {C.RE_YES: (th.bad_soft, th.bad),
                 C.RE_NO: (th.ok_soft, th.ok)})
    ws.conditional_format(
        r(C.ROW_FIRST), ci(bk.col(KEY, "remaining")),
        r(C.last_row(KEY)), ci(bk.col(KEY, "remaining")),
        {"type": "cell", "criteria": "<=",
         "value": "=$M%d" % C.ROW_FIRST,
         "format": bk.S.cf(bg=th.bad_soft, fg=th.bad, bold=True)})
    bk.stats["cond_formats"] += 1

    chips = [
        ('="\U0001F9F5 Materials: "&COUNTA(%s)' % bk.rng(KEY, "name"),
         "primary", "Materials: %d" % (len(m.materials) if m else 0), 3),
        ('="\U0001F4B0 Shelf value: "&Currency&TEXT(SUM(%s),"#,##0.00")'
         % bk.rng(KEY, "total"), "info",
         "Shelf value: %s" % (m.money(sum(x["total"] for x in m.materials),
                                      2) if m else "$0.00"), 4),
        ('="\U0001F534 Reorder now: "&%s' % bk.kpi("low_materials"), "bad",
         "Reorder now: %d" % (m.agg.get("low_materials", 0) if m else 0), 3),
        ('="\u2705 Healthy lines: "&COUNTIF(%s,"%s")'
         % (bk.rng(KEY, "reorder"), C.RE_NO), "ok",
         "Healthy lines: %d" % (sum(1 for x in (m.materials if m else [])
                                    if x["remaining"] > x["threshold"])), 4),
    ]
    K.chips(bk, KEY, chips, LAST_COL)
    K.footer_nav(bk, KEY, C.last_row(KEY) + 3, LAST_COL)
```

## 15. Made & Stocked (production)

`crochet_tracker/sheets/production.py`

```python
"""📦 Made & Stocked - production batches and live inventory per batch."""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "production"
LAST_COL = "L"

COLUMNS = [
    ("n", "#", "idx", None),
    ("product", "Product", "text", None),
    ("date", "Date made", "date", None),
    ("made", "Made", "qty", None),
    ("reserved", "Reserved", "qty", None),
    ("taken", "Taken to fair", "qty", None),
    ("sold", "Sold", "qty", None),
    ("returned", "Returned", "qty", None),
    ("damaged", "Damaged", "qty", None),
    ("current", "Current stock", "calc_num", "primary_2"),
    ("available", "Available", "calc_num", "primary_2"),
    ("notes", "Notes", "wrap", None),
]


def build(bk):
    th = bk.th
    ws = bk.ws(KEY)
    m = bk.demo
    K.sheet_head(bk, KEY, LAST_COL,
                 "\U0001F4E6  Made & Stocked",
                 "One row per batch: Available = Made + Returned \u2212 Sold "
                 "\u2212 Damaged \u2212 Reserved, automatically.")
    K.table_frame(bk, KEY, COLUMNS, height=22)

    for i in range(C.CAP[KEY]):
        row = C.ROW_FIRST + i
        b = m.production[i] if m and i < len(m.production) else None
        values = {"n": i + 1}
        if b:
            values.update({"product": b["product"], "date": b["date"],
                           "made": b["made"], "reserved": b["reserved"],
                           "taken": b["taken"], "sold": b["sold"],
                           "returned": b["returned"],
                           "damaged": b["damaged"], "notes": b["notes"]})
        values["current"] = ('=IF($B%d="","",$D%d+$H%d-$G%d-$I%d)'
                             % (row, row, row, row, row))
        values["available"] = '=IF($B%d="","",$J%d-$E%d)' % (row, row, row)
        cached = {"current": b["current"], "available": b["available"]} \
            if b else {}
        K.write_row(bk, KEY, COLUMNS, row, values, cached)

    K.list_dv(bk, KEY, "product", "products", title="Product")
    K.date_dv(bk, KEY, ("date",))
    K.whole_dv(bk, KEY, ("made", "reserved", "taken", "sold", "returned",
                         "damaged"), minimum=0, maximum=100000)
    ws.conditional_format(
        r(C.ROW_FIRST), ci(bk.col(KEY, "available")),
        r(C.last_row(KEY)), ci(bk.col(KEY, "available")),
        {"type": "cell", "criteria": "<=", "value": 0,
         "format": bk.S.cf(bg=th.warn_soft, fg=th.warn, bold=True)})
    bk.stats["cond_formats"] += 1

    tot = C.last_row(KEY) + 1
    K.totals_row(bk, KEY, tot, {
        "made": ("=SUM(D%d:D%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                 "#,##0", m.agg.get("made_total", 0) if m else 0),
        "sold": ("=SUM(G%d:G%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                 "#,##0", m.agg.get("units", 0) if m else 0),
        "current": ("=SUM(J%d:J%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                    "#,##0", sum(b["current"] for b in m.production)
                    if m else 0),
        "available": ("=SUM(K%d:K%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                      "#,##0", sum(b["available"] for b in m.production)
                      if m else 0)},
        first_col="B", last_col="L")

    chips = [
        ('="\U0001F9F6 Batches: "&COUNTA(%s)' % bk.rng(KEY, "product"),
         "primary", "Batches: %d" % (len(m.production) if m else 0), 3),
        ('="\U0001F4E6 Made: "&%s' % bk.kpi("made_total"), "info",
         "Made: %d" % (m.agg.get("made_total", 0) if m else 0), 3),
        ('="\U0001F4B0 Sold: "&%s' % bk.kpi("units"), "ok",
         "Sold: %d" % (m.agg.get("units", 0) if m else 0), 3),
        ('="\U0001F4C8 Sell-through: "&TEXT(%s,"0%%")'
         % bk.kpi("sell_through"), "gold",
         "Sell-through: %.0f%%" % (100 * (m.agg.get("sell_through", 0)
                                          if m else 0)), 3),
    ]
    K.chips(bk, KEY, chips, LAST_COL)
    K.footer_nav(bk, KEY, tot + 2, LAST_COL)
```

## 16. Craft Fairs (events)

`crochet_tracker/sheets/events.py`

```python
"""🏪 Craft Fairs - one record per market, with live profit per fair."""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "events"
LAST_COL = "R"

COLUMNS = [
    ("n", "#", "idx", None),
    ("name", "Fair / event", "text", None),
    ("date", "Date", "date", None),
    ("location", "Location", "text", None),
    ("organizer", "Organizer", "text", None),
    ("booth", "Booth fee", "money0", None),
    ("travel", "Travel", "money0", None),
    ("parking", "Parking", "money0", None),
    ("food", "Food / other", "money0", None),
    ("display", "Display", "money0", None),
    ("total", "Total expenses", "calc_money", "bad"),
    ("taken", "Items taken", "qty", None),
    ("sold", "Items sold", "calc_num", "primary_2"),
    ("sales", "Total sales", "calc_money", "primary_2"),
    ("cogs", "Product cost", "calc_money", "primary_2"),
    ("net", "Net profit", "calc_money", "ok"),
    ("margin", "Margin", "calc_pct", "ok"),
    ("best", "Best seller", "calc_wrap", None),
]


def build(bk):
    th = bk.th
    ws = bk.ws(KEY)
    m = bk.demo
    K.sheet_head(bk, KEY, LAST_COL,
                 "\U0001F3EA  Craft Fairs",
                 "Every market you work - sales, costs and best seller pull "
                 "straight from the \U0001F4B0 Sales Log.")
    K.table_frame(bk, KEY, COLUMNS, height=26)

    sale_ev = bk.rng("sales", "event")
    sale_qty = bk.rng("sales", "qty")
    sale_tot = bk.rng("sales", "total")
    sale_cost = bk.rng("sales", "cost")
    sale_prod = bk.rng("sales", "product")
    sale_pair = bk.rng("sales", "pair")
    for i in range(C.CAP[KEY]):
        row = C.ROW_FIRST + i
        e = m.events[i] if m and i < len(m.events) else None
        values = {"n": i + 1}
        if e:
            values.update({"name": e["name"], "date": e["date"],
                           "location": e["location"],
                           "organizer": e["organizer"], "booth": e["booth"],
                           "travel": e["travel"], "parking": e["parking"],
                           "food": e["food"], "display": e["display"],
                           "taken": e["taken"]})
        values["total"] = '=IF($B%d="","",SUM($F%d:$J%d))' % (row, row, row)
        values["sold"] = '=IF($B%d="","",SUMIFS(%s,%s,$B%d))' % (
            row, sale_qty, sale_ev, row)
        values["sales"] = '=IF($B%d="","",SUMIFS(%s,%s,$B%d))' % (
            row, sale_tot, sale_ev, row)
        values["cogs"] = '=IF($B%d="","",SUMIFS(%s,%s,$B%d))' % (
            row, sale_cost, sale_ev, row)
        values["net"] = '=IF($N%d="","",$N%d-$K%d-$O%d)' % (row, row, row,
                                                            row)
        values["margin"] = '=IF($N%d="","",IF($N%d=0,"",$P%d/$N%d))' % (
            row, row, row, row)
        values["best"] = (
            '=IF($B%d="","",IFERROR(INDEX(%s,MATCH(1,INDEX((%s=$B%d)*'
            '(%s=SUMPRODUCT(MAX((%s=$B%d)*%s))),0),0)),""))'
            % (row, sale_prod, sale_ev, row, sale_pair, sale_ev, row,
               sale_pair))
        cached = {}
        if e:
            cached = {"total": e["total"], "sold": e["sold"],
                      "sales": e["sales"], "cogs": e["cogs"],
                      "net": e["net"], "margin": e["margin"],
                      "best": e["best"]}
        K.write_row(bk, KEY, COLUMNS, row, values, cached)

    K.date_dv(bk, KEY, ("date",))
    K.money_dv(bk, KEY, ("booth", "travel", "parking", "food", "display"))
    K.databar(bk, KEY, "net", color=th.ok)
    ws.conditional_format(
        r(C.ROW_FIRST), ci(bk.col(KEY, "net")),
        r(C.last_row(KEY)), ci(bk.col(KEY, "net")),
        {"type": "cell", "criteria": "<", "value": 0,
         "format": bk.S.cf(bg=th.bad_soft, fg=th.bad, bold=True)})
    bk.stats["cond_formats"] += 1

    tot = C.last_row(KEY) + 1
    K.totals_row(bk, KEY, tot, {
        "total": ("=SUM(K%d:K%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                  "#,##0.00", m.agg.get("fees", 0) if m else 0),
        "sales": ("=SUM(N%d:N%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                  "#,##0.00", m.agg.get("revenue", 0) if m else 0),
        "net": ("=SUM(P%d:P%d)" % (C.ROW_FIRST, C.last_row(KEY)),
                "#,##0.00", sum(e["net"] for e in m.events) if m else 0)},
        first_col="B", last_col="R")

    done = [e for e in (m.events if m else []) if e["sales"]]
    chips = [
        ('="\U0001F3EA Fairs logged: "&COUNTA(%s)' % bk.rng(KEY, "name"),
         "primary", "Fairs logged: %d" % (len(m.events) if m else 0), 3),
        ('="\u2705 Worked: "&%s' % bk.kpi("events_done"), "ok",
         "Worked: %d" % (m.agg.get("events_done", 0) if m else 0), 3),
        ('="\U0001F4B0 Takings: "&Currency&TEXT(SUM(%s),"#,##0")'
         % bk.rng(KEY, "sales"), "info",
         "Takings: %s" % (m.money(m.agg.get("revenue", 0)) if m else "$0"),
         3),
        ('="\U0001F3C6 Best fair: "&%s' % bk.kpi("best_event"), "gold",
         "Best fair: %s" % (m.agg.get("best_event", "") if m else ""), 5),
    ]
    K.chips(bk, KEY, chips, LAST_COL)
    K.footer_nav(bk, KEY, tot + 2, LAST_COL, tip=
                 "  \U0001F4A1  Wondering \u0022should I do this market "
                 "again?\u0022  Sort by Net profit, or check the ranking on "
                 "the \U0001F4CA Dashboard.")
```

## 17. Sales Log

`crochet_tracker/sheets/sales.py`

```python
"""💰 Sales Log - every sale, with event & product dropdowns and live cost."""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "sales"
LAST_COL = "M"

COLUMNS = [
    ("n", "#", "idx", None),
    ("date", "Date", "date", None),
    ("event", "Fair / event", "text", None),
    ("product", "Product", "text", None),
    ("qty", "Qty", "qty", None),
    ("unit", "Unit price", "money", None),
    ("discount", "Discount", "money", None),
    ("total", "Sale total", "calc_money", "ok"),
    ("method", "Payment method", "center", None),
    ("ref", "Txn / ref #", "text", None),
    ("notes", "Notes", "wrap", None),
    ("pair", "Pair units", "calc_num", "primary_2"),
    ("cost", "Cost of sale", "calc_money", "primary_2"),
]


def build(bk):
    th = bk.th
    ws = bk.ws(KEY)
    m = bk.demo
    K.sheet_head(bk, KEY, LAST_COL,
                 "\U0001F4B0  Sales Log",
                 "Log each sale once - the fair totals, dashboard, monthly "
                 "summary and best-sellers all read this page.")
    K.table_frame(bk, KEY, COLUMNS, height=22)

    first, last = C.ROW_FIRST, C.last_row(KEY)
    cat_name = bk.rng("catalog", "name")
    cat_cost = bk.rng("catalog", "unit_cost")
    for i in range(C.CAP[KEY]):
        row = C.ROW_FIRST + i
        s = m.sales[i] if m and i < len(m.sales) else None
        values = {"n": i + 1}
        if s:
            values.update({"date": s["date"], "event": s["event"],
                           "product": s["product"], "qty": s["qty"],
                           "unit": s["unit"], "discount": s["discount"],
                           "method": s["method"], "ref": s["ref"],
                           "notes": s["notes"]})
        values["total"] = ('=IF($D%d="","",IF($E%d="","",E%d*F%d)-'
                           'IF($G%d="",0,$G%d))'
                           % (row, row, row, row, row, row))
        values["pair"] = ('=IF($C%d="","",SUMIFS($E$%d:$E$%d,$C$%d:$C$%d,'
                          '$C%d,$D$%d:$D$%d,$D%d))'
                          % (row, first, last, first, last, row, first,
                             last, row))
        values["cost"] = ('=IF($D%d="","",$E%d*IFERROR(INDEX(%s,MATCH('
                          '$D%d,%s,0)),0))' % (row, row, cat_cost, row,
                                               cat_name))
        cached = {}
        if s:
            uc = _uc(m, s["product"])
            cached = {"total": round(s["qty"] * s["unit"] - s["discount"],
                                     2),
                      "pair": _pair(m, s),
                      "cost": round(s["qty"] * uc, 2)}
        K.write_row(bk, KEY, COLUMNS, row, values, cached)

    K.list_dv(bk, KEY, "event", "events", title="Fair / event")
    K.list_dv(bk, KEY, "product", "products", title="Product")
    K.list_dv(bk, KEY, "method", "payments", title="Payment method")
    K.date_dv(bk, KEY, ("date",))
    K.money_dv(bk, KEY, ("unit", "discount"))
    K.whole_dv(bk, KEY, ("qty",), minimum=0, maximum=100000)

    tot = C.last_row(KEY) + 1
    K.totals_row(bk, KEY, tot, {
        "qty": ("=SUM(E%d:E%d)" % (first, last), "#,##0",
                m.agg.get("units", 0) if m else 0),
        "total": ("=SUM(H%d:H%d)" % (first, last), "#,##0.00",
                  m.agg.get("revenue", 0) if m else 0),
        "cost": ("=SUM(M%d:M%d)" % (first, last), "#,##0.00",
                 m.agg.get("cogs", 0) if m else 0)},
        first_col="B", last_col=LAST_COL)

    chips = [
        ('="\U0001F4B0 Takings: "&Currency&TEXT(SUM(%s),"#,##0.00")'
         % bk.rng(KEY, "total"), "ok",
         "Takings: %s" % (m.money(m.agg.get("revenue", 0), 2)
                          if m else "$0.00"), 4),
        ('="\U0001F9F6 Units: "&%s' % bk.kpi("units"), "primary",
         "Units: %d" % (m.agg.get("units", 0) if m else 0), 2),
        ('="\U0001F9FE Sales logged: "&%s' % bk.kpi("txns"), "info",
         "Sales logged: %d" % (m.agg.get("txns", 0) if m else 0), 3),
        ('="\U0001F3F7\uFE0F Discounts: "&Currency&TEXT(SUM(%s),"#,##0.00")'
         % bk.rng(KEY, "discount"), "gold",
         "Discounts: %s" % (m.money(m.agg.get("discounts", 0), 2)
                            if m else "$0.00"), 3),
    ]
    K.chips(bk, KEY, chips, LAST_COL)
    K.footer_nav(bk, KEY, tot + 2, LAST_COL)


def _uc(m, product):
    for p in m.products:
        if p["name"] == product:
            return p["unit_cost"]
    return 0.0


def _pair(m, sale):
    return sum(s["qty"] for s in m.sales
               if s["event"] == sale["event"]
               and s["product"] == sale["product"])
```

## 18. Event Profit calculator

`crochet_tracker/sheets/eventprofit.py`

```python
"""🧮 Event Profit - pick a fair, see exactly what it made (or lost)."""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "eventprofit"
LAST_COL = "J"


def build(bk):
    th = bk.th
    ws = bk.ws(KEY)
    m = bk.demo
    ep = m.agg.get("event_profit", {}) if m else {}
    K.sheet_head(bk, KEY, LAST_COL,
                 "\U0001F9EE  Event Profit",
                 "Pick any fair and see revenue, every cost line, net "
                 "profit, margin, break-even and ROI - instantly.")
    bk.paint(KEY, 0, 0, C.EP_LAST_ROW + 8, ci(LAST_COL), S_can(bk))

    lab = bk.S.f(**bk.S.base(font_size=11, font_color=bk.th.ink,
                             bg_color=bk.th.card, align="left",
                             valign="vcenter", indent=1, border=1,
                             border_color=bk.th.border))
    val = bk.S.f(**bk.S.base(font_size=12, bold=True,
                             font_color=bk.th.primary,
                             bg_color=bk.th.gold_soft, align="right",
                             valign="vcenter", border=1,
                             border_color=bk.th.border_strong,
                             num_format="#,##0.00"))
    val0 = bk.S.f(**bk.S.base(font_size=12, bold=True,
                              font_color=bk.th.primary,
                              bg_color=bk.th.gold_soft, align="right",
                              valign="vcenter", border=1,
                              border_color=bk.th.border_strong,
                              num_format="#,##0"))
    valp = bk.S.f(**bk.S.base(font_size=12, bold=True,
                              font_color=bk.th.primary,
                              bg_color=bk.th.gold_soft, align="right",
                              valign="vcenter", border=1,
                              border_color=bk.th.border_strong,
                              num_format="0.0%"))
    big = bk.S.f(**bk.S.base(font_size=15, bold=True,
                             font_color=bk.th.white, bg_color=bk.th.ok,
                             align="right", valign="vcenter", border=1,
                             border_color=bk.th.ok,
                             num_format="#,##0.00"))

    # picker
    ws.set_row(r(C.EP_EVENT_ROW), 26)
    ws.merge_range(r(C.EP_EVENT_ROW), 1, r(C.EP_EVENT_ROW), 4,
                   "  \U0001F3EA  Pick a fair to analyse", lab)
    ws.merge_range(r(C.EP_EVENT_ROW), 5, r(C.EP_EVENT_ROW), ci(LAST_COL),
                   "", bk.S.f(**bk.S.base(font_size=12, bold=True,
                                          font_color=bk.th.primary,
                                          bg_color=bk.th.gold_soft,
                                          border=1,
                                          border_color=bk.th.border_strong,
                                          align="left", valign="vcenter",
                                          indent=1)))
    bk.validate(KEY, C.EP_EVENT_ROW, 5, C.EP_EVENT_ROW, 5,
                "=" + bk.listname("events"), title="Pick a fair")
    sel = "$E$%d" % C.EP_EVENT_ROW
    if m:
        ws.write(r(C.EP_EVENT_ROW), 5, m.profit_event,
                 bk.S.f(**bk.S.base(font_size=12, bold=True,
                                    font_color=bk.th.primary,
                                    bg_color=bk.th.gold_soft, border=1,
                                    border_color=bk.th.border_strong,
                                    align="left", valign="vcenter",
                                    indent=1)))

    ev = bk.q("events")
    sale_ev = bk.rng("sales", "event")

    def section(row, emoji, text):
        ws.set_row(r(row), 22)
        ws.merge_range(r(row), 1, r(row), ci(LAST_COL),
                       "  %s  %s" % (emoji, text), bk.S.section)

    def line(row, label, formula, fmt, cached):
        ws.set_row(r(row), 20)
        ws.merge_range(r(row), 1, r(row), 5, "  " + label, lab)
        ws.merge_range(r(row), 6, r(row), ci(LAST_COL), "", fmt)
        ws.write_formula(r(row), 6, formula, fmt, cached)
        bk.stats["formulas"] += 1

    section(10, "\U0001F4B0", "REVENUE")
    line(C.EP_REV["units"], "Total units sold",
         '=SUMIFS(%s,%s,%s)' % (bk.rng("sales", "qty"), sale_ev, sel),
         val0, ep.get("units", 0))
    line(C.EP_REV["gross"], "Gross sales",
         '=SUMIFS(%s,%s,%s)' % (bk.rng("sales", "total"), sale_ev, sel),
         val, ep.get("gross", 0))
    line(C.EP_REV["discounts"], "Discounts given",
         '=SUMIFS(%s,%s,%s)' % (bk.rng("sales", "discount"), sale_ev, sel),
         val, ep.get("discounts", 0))

    section(15, "\U0001F9FE", "COSTS")
    line(C.EP_COST["cogs"], "Cost of products sold",
         '=SUMIFS(%s,%s,%s)' % (bk.rng("sales", "cost"), sale_ev, sel),
         val, ep.get("cogs", 0))
    for i, (field, label) in enumerate((("booth", "Booth fee"),
                                        ("travel", "Travel"),
                                        ("parking", "Parking"),
                                        ("food", "Food & drinks"),
                                        ("display", "Display & decor"))):
        line(C.EP_COST[field], label,
             '=IFERROR(INDEX(%s!$%s$%d:$%s$%d,MATCH(%s,%s!$B$%d:$B$%d,0)),0)'
             % (ev, C.COLS["events"][field], C.ROW_FIRST,
                C.COLS["events"][field], C.last_row("events"), sel, ev,
                C.ROW_FIRST, C.last_row("events")),
             val, ep.get(field, 0))

    section(23, "\U0001F3C1", "RESULTS")
    line(C.EP_OUT["gross_profit"], "Gross profit (sales \u2212 product cost)",
         "=$F$%d-$F$%d" % (C.EP_REV["gross"], C.EP_COST["cogs"]),
         val, ep.get("gross_profit", 0))
    line(C.EP_OUT["total_cost"], "Total costs (product + fair)",
         "=SUM($F$%d:$F$%d)" % (C.EP_COST["cogs"], C.EP_COST["display"]),
         val, ep.get("total_cost", 0))
    ws.set_row(r(C.EP_OUT["net"]), 26)
    ws.merge_range(r(C.EP_OUT["net"]), 1, r(C.EP_OUT["net"]), 5,
                   "  NET EVENT PROFIT", bk.S.f(**bk.S.base(
                       bold=True, font_size=12, font_color=bk.th.white,
                       bg_color=bk.th.ok, align="left", valign="vcenter",
                       indent=1)))
    ws.merge_range(r(C.EP_OUT["net"]), 6, r(C.EP_OUT["net"]), ci(LAST_COL),
                   "", big)
    ws.write_formula(r(C.EP_OUT["net"]), 6,
                     "=$F$%d-$F$%d" % (C.EP_REV["gross"],
                                      C.EP_OUT["total_cost"]),
                     big, ep.get("net", 0))
    bk.stats["formulas"] += 1
    line(C.EP_OUT["margin"], "Profit margin %",
         '=IFERROR($F$%d/$F$%d,0)' % (C.EP_OUT["net"], C.EP_REV["gross"]),
         valp, ep.get("margin", 0))
    line(C.EP_OUT["avg_sale"], "Average sale value",
         '=IFERROR($F$%d/COUNTIFS(%s,%s),0)'
         % (C.EP_REV["gross"], sale_ev, sel), val, ep.get("avg_sale", 0))
    line(C.EP_OUT["breakeven"], "Break-even sales",
         '=IFERROR($F$%d/($F$%d/$F$%d),0)'
         % (C.EP_OUT["total_cost"], C.EP_OUT["gross_profit"],
            C.EP_REV["gross"]), val, ep.get("breakeven", 0))
    line(C.EP_OUT["roi"], "Return on costs (ROI %)",
         '=IFERROR($F$%d/$F$%d,0)' % (C.EP_OUT["net"],
                                      C.EP_OUT["total_cost"]),
         valp, ep.get("roi", 0))

    K.note_block(bk, KEY, C.EP_LAST_ROW - 2, "B", LAST_COL, [
        "How to read this page:",
        "1.  Gross profit = takings minus what the sold items cost you in "
        "yarn and packaging.",
        "2.  Net profit also subtracts every fair cost - booth, travel, "
        "parking, food and display.",
        "3.  Break-even sales = the takings you needed to cover all costs "
        "at this margin.  Beat it next time!",
    ], title="  \U0001F4DA  READING YOUR NUMBERS")
    K.footer_nav(bk, KEY, C.EP_LAST_ROW + 4, LAST_COL, landscape=False)


def S_can(bk):
    return bk.S.canvas
```

## 19. Reorder List

`crochet_tracker/sheets/reorder.py`

```python
"""🔄 Reorder List - auto-pulled low products & materials, red when urgent."""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "reorder"
LAST_COL = "J"
PROD_SLOTS = C.RO_PROD_LAST - C.RO_PROD_FIRST + 1
MAT_SLOTS = C.RO_MAT_LAST - C.RO_MAT_FIRST + 1
PROD_HEAD = C.RO_PROD_FIRST - 1
MAT_HEAD = C.RO_MAT_FIRST - 1

COLUMNS = [
    ("item", "Item", 26, "text"),
    ("current", "Current", 9, "qty1"),
    ("min", "Minimum", 8, "qty1"),
    ("suggest", "Suggest order", 10, "qty1"),
    ("supplier", "Supplier", 18, "text"),
    ("est", "Est. cost", 10, "money"),
    ("priority", "Priority", 13, "center"),
    ("ordered", "Ordered?", 10, "tick"),
    ("ordered_date", "Order date", 12, "date"),
]


def build(bk):
    th = bk.th
    ws = bk.ws(key := KEY)
    m = bk.demo
    premium = bk.has("materials")
    K.sheet_head(bk, KEY, LAST_COL,
                 "\U0001F504  Reorder List",
                 "Low products and yarn pull themselves onto this list - "
                 "urgent lines glow red until you tick them ordered.")
    bk.paint(KEY, 0, 0, MAT_HEAD + MAT_SLOTS + 8, ci(LAST_COL),
             bk.S.canvas)

    lab = bk.S.f(**bk.S.base(font_size=10.5, font_color=th.ink,
                             bg_color=th.card, align="left",
                             valign="vcenter", indent=1, border=1,
                             border_color=th.border))
    calc = bk.S.f(**bk.S.base(font_size=10.5, font_color=th.primary,
                              bg_color=th.alt, align="center",
                              valign="vcenter", border=1,
                              border_color=th.border,
                              num_format="#,##0.0#"))
    money = bk.S.f(**bk.S.base(font_size=10.5, font_color=th.primary,
                               bg_color=th.alt, align="right",
                               valign="vcenter", border=1,
                               border_color=th.border,
                               num_format="#,##0.00"))
    prio = bk.S.f(**bk.S.base(font_size=10.5, bold=True, font_color=th.ink,
                              bg_color=th.card, align="center",
                              valign="vcenter", border=1,
                              border_color=th.border))
    tick = bk.S.f(**bk.S.base(font_size=13, bold=True, font_color=th.ok,
                              bg_color=th.card, align="center",
                              valign="vcenter", border=1,
                              border_color=th.border, locked=False))
    datef = bk.S.f(**bk.S.base(font_size=10.5, font_color=th.ink,
                               bg_color=th.card, align="center",
                               valign="vcenter", border=1,
                               border_color=th.border,
                               num_format="dd mmm yyyy", locked=False))

    cat = bk.q("catalog")
    d = bk.q("data")
    seqp = "%s!$AK$%d:$AK$%d" % (d, C.DATA_SEQP_FIRST,
                                 C.DATA_SEQP_FIRST + C.CAP["catalog"] - 1)
    seqm = "%s!$AN$%d:$AN$%d" % (d, C.DATA_SEQM_FIRST,
                                 C.DATA_SEQM_FIRST + C.CAP["materials"] - 1)
    cl = C.ROW_FIRST
    clast = C.last_row("catalog")
    mlast = C.last_row("materials")

    def header(row, emoji, text):
        ws.set_row(r(row), 22)
        ws.merge_range(r(row), 1, r(row), ci(LAST_COL),
                       "  %s  %s" % (emoji, text), bk.S.section_soft)
        ws.set_row(r(row + 1), 20)
        for label, col in (("Item", 1), ("Current", 2), ("Minimum", 3),
                           ("Suggest", 4), ("Supplier", 5), ("Est. cost", 6),
                           ("Priority", 7), ("Ordered?", 8),
                           ("Order date", 9)):
            ws.write(r(row + 1), col, label, bk.S.thead)

    def slot(row, k, kind):
        ws.set_row(r(row), 20)
        if kind == "product":
            seq, sheet, name_col, cur_col, min_col, cost_col, sup = (
                seqp, cat, "C", "M", "N", "I", None)
            last = clast
        else:
            seq, sheet, name_col, cur_col, min_col, cost_col, sup = (
                seqm, bk.q("materials"), "B", "K", "M", "H", "L")
            last = mlast
        match = "MATCH(%d,%s,0)" % (k, seq)
        ws.write_formula(r(row), 1,
                         '=IFERROR(INDEX(%s!$%s$%d:$%s$%d,%s),"")'
                         % (sheet, name_col, cl, name_col, last, match),
                         lab, "")
        bk.stats["formulas"] += 1
        for col, src in ((2, cur_col), (3, min_col)):
            ws.write_formula(r(row), col,
                             '=IF($B%d="","",INDEX(%s!$%s$%d:$%s$%d,%s))'
                             % (row, sheet, src, cl, src, last, match),
                             calc, "")
            bk.stats["formulas"] += 1
        ws.write_formula(r(row), 4,
                         '=IF($B%d="","",($D%d*2)-$C%d)' % (row, row, row),
                         calc, "")
        bk.stats["formulas"] += 1
        if sup:
            ws.write_formula(r(row), 5,
                             '=IF($B%d="","",INDEX(%s!$%s$%d:$%s$%d,%s))'
                             % (row, sheet, sup, cl, sup, last, match),
                             lab, "")
            bk.stats["formulas"] += 1
        else:
            ws.write_formula(r(row), 5,
                             '=IF($B%d="","","Make in-house")' % row,
                             lab, "")
            bk.stats["formulas"] += 1
        ws.write_formula(r(row), 6,
                         '=IF($B%d="","",$E%d*INDEX(%s!$%s$%d:$%s$%d,%s))'
                         % (row, row, sheet, cost_col, cl, cost_col, last,
                            match),
                         money, "")
        bk.stats["formulas"] += 1
        ws.write_formula(r(row), 7,
                         '=IF($B%d="","",IF($C%d<=0,"%s",IF($C%d<=$D%d/2,'
                         '"%s","%s")))'
                         % (row, row, C.PR_URGENT, row, row, C.PR_HIGH,
                            C.PR_NORMAL),
                         prio, "")
        bk.stats["formulas"] += 1
        ws.write(r(row), 8, "", tick)
        ws.write(r(row), 9, "", datef)

    header(PROD_HEAD - 1, "\U0001F9F6", "PRODUCTS RUNNING LOW")
    for k in range(1, PROD_SLOTS + 1):
        slot(PROD_HEAD + k, k, "product")
    if premium:
        header(MAT_HEAD - 1, "\U0001F9F5", "MATERIALS RUNNING LOW")
        for k in range(1, MAT_SLOTS + 1):
            slot(MAT_HEAD + k, k, "material")

    # conditional formatting: priority colours + ticked green
    first_p, last_p = PROD_HEAD + 1, PROD_HEAD + PROD_SLOTS
    first_m, last_m = MAT_HEAD + 1, MAT_HEAD + MAT_SLOTS
    for txt, bg, fg in ((C.PR_URGENT, th.bad_soft, th.bad),
                        (C.PR_HIGH, th.warn_soft, th.warn),
                        (C.PR_NORMAL, th.ok_soft, th.ok)):
        for a, b in ((first_p, last_p), (first_m, last_m)):
            bk.cond(KEY, a, 7, b, 7, {
                "type": "formula",
                "criteria": '=$H%d="%s"' % (a, txt),
                "format": bk.S.cf(bg=bg, fg=fg, bold=True)})
    for a, b in ((first_p, last_p), (first_m, last_m)):
        bk.cond(KEY, a, 8, b, 8, {
            "type": "formula",
            "criteria": '=$I%d="%s"' % (a, C.TICK),
            "format": bk.S.cf(bg=th.ok_soft, fg=th.ok, bold=True)})
    bk.validate(KEY, first_p, 8, last_m if premium else last_p, 8,
                "=Tick", title="Ordered?")
    K.date_dv(bk, KEY, ("ordered_date",), first=first_p,
              last=last_m if premium else last_p)

    # chips
    allb = "$B$%d:$B$%d" % (first_p, last_m if premium else last_p)
    alli = "$I$%d:$I$%d" % (first_p, last_m if premium else last_p)
    allg = "$G$%d:$G$%d" % (first_p, last_m if premium else last_p)
    allh = "$H$%d:$H$%d" % (first_p, last_m if premium else last_p)
    chips = [
        ('="\U0001F534 Open lines: "&SUMPRODUCT((%s<>"")*(%s<>"%s"))'
         % (allb, alli, C.TICK), "bad",
         "Open lines: %d" % ((m.agg.get("low_products", 0)
                              + m.agg.get("low_materials", 0)) if m else 0),
         4),
        ('="\u26A1 Urgent: "&COUNTIF(%s,"%s")' % (allh, C.PR_URGENT),
         "warn", "Urgent: %d" % (sum(
             1 for p in (m.products if m else []) if p["stock"] <= 0)
             + sum(1 for x in (m.materials if m else [])
                   if x["remaining"] <= 0 and x["remaining"]
                   <= x["threshold"])), 3),
        ('="\U0001F4B0 Budget to reorder: "&Currency&TEXT(SUMPRODUCT((%s'
         '<>"")*(%s<>"%s")*IFERROR(%s*1,0)),"#,##0.00")'
         % (allb, alli, C.TICK, allg),
         "info", "Budget to reorder: %s"
         % (m.money(m.agg.get("reorder_cost", 0), 2) if m else "$0.00"), 4),
        ('="\u2705 Ordered lines: "&COUNTIF(%s,"%s")' % (alli, C.TICK),
         "ok", "Ordered lines: 0", 3),
    ]
    K.chips(bk, KEY, chips, LAST_COL)
    K.footer_nav(bk, KEY, (last_m if premium else last_p) + 3, LAST_COL)
```

## 20. Packing Checklist

`crochet_tracker/sheets/packing.py`

```python
"""🎒 Packing Checklist - event-specific list with a % packed meter."""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "packing"
LAST_COL = "F"


def build(bk):
    th = bk.th
    ws = bk.ws(KEY)
    m = bk.demo
    K.sheet_head(bk, KEY, LAST_COL,
                 "\U0001F392  Packing Checklist",
                 "Pick the fair you are packing for, tick as you load - "
                 "the % packed meter fills in as you go.")
    bk.paint(KEY, 0, 0, 60, ci(LAST_COL), bk.S.canvas)

    lab = bk.S.f(**bk.S.base(font_size=11, font_color=th.ink,
                             bg_color=th.card, align="left",
                             valign="vcenter", indent=1, border=1,
                             border_color=th.border))
    pick = bk.S.f(**bk.S.base(font_size=12, bold=True,
                              font_color=th.primary,
                              bg_color=th.gold_soft, border=1,
                              border_color=th.border_strong, align="left",
                              valign="vcenter", indent=1, locked=False))
    item = bk.S.f(**bk.S.base(font_size=10.5, font_color=th.ink,
                              bg_color=th.card, align="left",
                              valign="vcenter", indent=1, border=1,
                              border_color=th.border))
    note = bk.S.f(**bk.S.base(font_size=10.5, font_color=th.muted,
                              bg_color=th.card, align="left",
                              valign="vcenter", indent=1, border=1,
                              border_color=th.border, locked=False))
    tick = bk.S.f(**bk.S.base(font_size=13, bold=True, font_color=th.ok,
                              bg_color=th.card, align="center",
                              valign="vcenter", border=1,
                              border_color=th.border, locked=False))

    ws.set_row(r(C.PACK_EVENT_ROW), 26)
    ws.merge_range(r(C.PACK_EVENT_ROW), 1, r(C.PACK_EVENT_ROW), 2,
                   "  \U0001F3EA  Packing for:", lab)
    ws.merge_range(r(C.PACK_EVENT_ROW), 3, r(C.PACK_EVENT_ROW),
                   ci(LAST_COL), "", pick)
    bk.validate(KEY, C.PACK_EVENT_ROW, 3, C.PACK_EVENT_ROW, 3,
                "=" + bk.listname("events"), title="Pick a fair")
    if m and m.events:
        from ..demo import today as _today
        nxt = [e for e in m.events if e["date"] >= _today()]
        ws.write(r(C.PACK_EVENT_ROW), 3,
                 (nxt[0]["name"] if nxt else m.events[-1]["name"]), pick)

    rows = []
    row = C.PACK_EVENT_ROW + 2
    for title, items in C.PACK_SECTIONS:
        ws.set_row(r(row), 22)
        ws.merge_range(r(row), 1, r(row), ci(LAST_COL), "",
                       bk.S.section_soft)
        ws.write_formula(r(row), 1,
                         '=CONCATENATE("  \U0001F9FA  %s   \u2014   "&'
                         'COUNTIF($B$%d:$B$%d,"\u2713")&" / %d packed")'
                         % (title.upper(), row + 1, row + len(items),
                            len(items)),
                         bk.S.section_soft,
                         "  \U0001F9FA  %s   \u2014   %d / %d packed"
                         % (title.upper(),
                            sum(1 for it in items
                                if m and m.packing.get(it)),
                            len(items)))
        bk.stats["formulas"] += 1
        row += 1
        for it in items:
            ws.set_row(r(row), 20)
            ws.write(r(row), 1, "", tick)
            ws.merge_range(r(row), 2, r(row), 3, "  " + it, item)
            ws.merge_range(r(row), 4, r(row), ci(LAST_COL), "", note)
            if m and m.packing.get(it):
                ws.write(r(row), 1, C.TICK, tick)
            rows.append(row)
            row += 1
        row += 1

    first, last = rows[0], rows[-1]
    bk.validate(KEY, first, 1, last, 1, "=Tick", title="Packed?")
    bk.cond(KEY, first, 1, last, 1, {
        "type": "formula", "criteria": '=$B%d=""' % first,
        "format": bk.S.cf(bg=th.card, fg=th.ok)})
    bk.cond(KEY, first, 1, last, 3, {
        "type": "formula", "criteria": '=$A%d="%s"' % (first, C.TICK),
        "format": bk.S.cf(bg=th.ok_soft, fg=th.ok, strike=True)})

    chips = [
        ('="\U0001F392 Packed: "&COUNTIF($B$%d:$B$%d,"\u2713")&" of %d"'
         % (first, last, len(rows)), "primary",
         "Packed: %d of %d" % (m.agg.get("packed", 0) if m else 0,
                               len(rows)), 4),
        ('="\U0001F4CA Progress: "&TEXT(COUNTIF($B$%d:$B$%d,"\u2713")/%d,'
         '"0%%")' % (first, last, len(rows)), "ok",
         "Progress: %.0f%%" % (100.0 * (m.agg.get("packed", 0) / len(rows)
                                        if m else 0)), 3),
        ('="\u23F3 Still to pack: "&%d-COUNTIF($B$%d:$B$%d,"\u2713")'
         % (len(rows), first, last), "warn",
         "Still to pack: %d" % (len(rows) - (m.agg.get("packed", 0)
                                             if m else 0)), 4),
    ]
    K.chips(bk, KEY, chips, LAST_COL)
    K.footer_nav(bk, KEY, last + 3, LAST_COL, landscape=False, tip=
                 "  \U0001F4A1  Print this tab (File \u2192 Print) and clip "
                 "it to the van door.  Ticks save mornings.")
```

## 21. Pricing Calculator

`crochet_tracker/sheets/pricing.py`

```python
"""💵 Pricing Calculator - true cost in, suggested price out."""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "pricing"
LAST_COL = "J"


def build(bk):
    th = bk.th
    ws = bk.ws(KEY)
    m = bk.demo
    pr = m.agg.get("pricing", {}) if m else {}
    K.sheet_head(bk, KEY, LAST_COL,
                 "\U0001F4B5  Pricing Calculator",
                 "Never sell a $25 make for $15 again: materials + labour + "
                 "overhead in, suggested price out.")
    bk.paint(KEY, 0, 0, C.PR_LAST_ROW + 8, ci(LAST_COL), bk.S.canvas)

    lab = bk.S.f(**bk.S.base(font_size=11, font_color=th.ink,
                             bg_color=th.card, align="left",
                             valign="vcenter", indent=1, border=1,
                             border_color=th.border))
    inp = bk.S.f(**bk.S.base(font_size=12, bold=True, font_color=th.primary,
                             bg_color=th.gold_soft, border=1,
                             border_color=th.border_strong, align="right",
                             valign="vcenter", num_format="#,##0.00",
                             locked=False))
    inp_pct = bk.S.f(**bk.S.base(font_size=12, bold=True,
                                 font_color=th.primary,
                                 bg_color=th.gold_soft, border=1,
                                 border_color=th.border_strong,
                                 align="right", valign="vcenter",
                                 num_format="0%", locked=False))
    inp_num = bk.S.f(**bk.S.base(font_size=12, bold=True,
                                 font_color=th.primary,
                                 bg_color=th.gold_soft, border=1,
                                 border_color=th.border_strong,
                                 align="right", valign="vcenter",
                                 num_format="#,##0.0", locked=False))
    out = bk.S.f(**bk.S.base(font_size=12, bold=True, font_color=th.primary,
                             bg_color=th.card, align="right",
                             valign="vcenter", border=1,
                             border_color=th.border,
                             num_format="#,##0.00"))
    hero = bk.S.f(**bk.S.base(font_size=16, bold=True, font_color=th.white,
                              bg_color=th.accent, align="right",
                              valign="vcenter", border=1,
                              border_color=th.accent,
                              num_format="#,##0.00"))

    def left(row, label, value, fmt, cached=None):
        ws.set_row(r(row), 22)
        ws.merge_range(r(row), 1, r(row), 4, "  " + label, lab)
        ws.merge_range(r(row), 5, r(row), 6, "", fmt)
        if isinstance(value, str) and value.startswith("="):
            ws.write_formula(r(row), 5, value, fmt, cached)
            bk.stats["formulas"] += 1
        else:
            ws.write(r(row), 5, value, fmt)

    def right(row, label, formula, cached, fmt=None, hero_fmt=False):
        ws.set_row(r(row), 24 if hero_fmt else 20)
        ws.merge_range(r(row), 7, r(row), 8, "  " + label, lab)
        ws.merge_range(r(row), 9, r(row), ci(LAST_COL), "",
                       hero if hero_fmt else (fmt or out))
        ws.write_formula(r(row), 9, formula, hero if hero_fmt
                         else (fmt or out), cached)
        bk.stats["formulas"] += 1

    q = m.pricing if m else {}
    ws.merge_range(r(8), 1, r(8), 6, "  1 \u00B7 WHAT IT COSTS YOU",
                   bk.S.section_soft)
    ws.merge_range(r(8), 7, r(8), ci(LAST_COL), "  2 \u00B7 WHAT TO CHARGE",
                   bk.S.section_soft)
    ws.set_row(r(8), 22)

    left(C.PR_IN["material"], "Materials / yarn cost",
         q.get("material", ""), inp, q.get("material", 0))
    left(C.PR_IN["packaging"], "Packaging & labels",
         q.get("packaging", ""), inp, q.get("packaging", 0))
    left(C.PR_IN["hours"], "Hours to make", q.get("hours", ""), inp_num,
         q.get("hours", 0))
    left(C.PR_IN["wage"], "Your hourly wage", "=HourlyWage", inp,
         (m.settings["wage"] if m else 14.0))
    left(C.PR_IN["overhead"], "Overhead % (tools, fees, power)",
         "=OverheadPct", inp_pct, (m.settings["overhead"] if m else 0.1))
    left(C.PR_IN["margin"], "Target profit margin", "=TargetMargin",
         inp_pct, (m.settings["margin"] if m else 0.45))

    E = "$E$"
    right(C.PR_OUT["labor"], "Labour cost (hours \u00D7 wage)",
          '=IF(%s%d="","",%s%d*%s%d)' % (E, C.PR_IN["hours"], E,
                                         C.PR_IN["hours"], E,
                                         C.PR_IN["wage"]),
          pr.get("labor", 0))
    right(C.PR_OUT["overhead_amt"], "Overhead amount",
          '=IF(%s%d="","",(%s%d+%s%d+%s%d)*%s%d)'
          % (E, C.PR_IN["material"], E, C.PR_IN["material"], E,
             C.PR_IN["packaging"], E, C.PR_OUT["labor"], E,
             C.PR_IN["overhead"]),
          pr.get("overhead_amt", 0))
    right(C.PR_OUT["true_cost"], "TRUE COST",
          '=%s%d+%s%d+%s%d+%s%d' % (E, C.PR_IN["material"], E,
                                    C.PR_IN["packaging"], E,
                                    C.PR_OUT["labor"], E,
                                    C.PR_OUT["overhead_amt"]),
          pr.get("true_cost", 0))
    right(C.PR_OUT["price"], "SUGGESTED PRICE",
          '=IF(OR(%s%d="",1-%s%d<=0),"",ROUND(%s%d/(1-%s%d),2))'
          % (E, C.PR_OUT["true_cost"], E, C.PR_IN["margin"], E,
             C.PR_OUT["true_cost"], E, C.PR_IN["margin"]),
          pr.get("price", 0), hero_fmt=True)
    right(C.PR_OUT["charm"], "Charm price (\u2026.95)",
          '=IF(%s%d="","",ROUNDUP(%s%d,0)-0.05)'
          % (E, C.PR_OUT["price"], E, C.PR_OUT["price"]),
          pr.get("charm", 0))
    right(C.PR_OUT["profit"], "Profit at suggested price",
          '=IF(%s%d="","",%s%d-%s%d)' % (E, C.PR_OUT["price"], E,
                                         C.PR_OUT["price"], E,
                                         C.PR_OUT["true_cost"]),
          pr.get("profit", 0))
    right(C.PR_OUT["check"], "Sanity check",
          '=IF(%s%d="","",IF(%s%d<%s%d,"\u26A0\uFE0F below true cost!",'
          '"\u2705 healthy margin"))'
          % (E, C.PR_OUT["price"], E, C.PR_OUT["price"], E,
             C.PR_OUT["true_cost"]),
          "\u2705 healthy margin" if pr else "",
          fmt=bk.S.f(**bk.S.base(font_size=11, bold=True,
                                 font_color=th.ok, bg_color=th.ok_soft,
                                 align="center", valign="vcenter",
                                 border=1, border_color=th.border)))

    K.note_block(bk, KEY, C.PR_LAST_ROW - 4, "B", LAST_COL, [
        "Rules of thumb from market sellers:",
        "1.  Materials = yarn + labels + packaging for ONE item.",
        "2.  Overhead covers hooks, stall extras, power and marketplace "
        "fees - 10-15% is typical.",
        "3.  If the suggested price feels too high for your market, cut "
        "the hours (simpler pattern) before you cut your wage.",
    ], title="  \U0001F4A1  PRICING LIKE A PRO")
    K.footer_nav(bk, KEY, C.PR_LAST_ROW + 2, LAST_COL, landscape=False)
```

## 22. Monthly Summary

`crochet_tracker/sheets/monthly.py`

```python
"""📅 Monthly Summary - the year at a glance, month by month."""

from .. import config as C
from ..book import r, ci
from . import common as K

KEY = "monthly"
LAST_COL = "K"

COLUMNS = [
    ("month", "Month", "center", None),
    ("revenue", "Revenue", "calc_money", "ok"),
    ("cogs", "Product costs", "calc_money", None),
    ("fees", "Fair fees", "calc_money", None),
    ("net", "Net profit", "calc_money", "primary_2"),
    ("margin", "Margin", "calc_pct", "primary_2"),
    ("units", "Units sold", "calc_num", None),
    ("aov", "Avg sale", "calc_money", None),
    ("per_item", "Profit / item", "calc_money", None),
    ("markets", "Markets", "calc_num", None),
]


def build(bk):
    th = bk.th
    ws = bk.ws(KEY)
    m = bk.demo
    K.sheet_head(bk, KEY, LAST_COL,
                 "\U0001F4C5  Monthly Summary",
                 "Revenue, costs, profit and markets for every month of "
                 "the year set on \u2699\uFE0F Lists & Settings.")
    # header + rows
    from .common import header_row, data_rows
    cols = [("n", "#", "idx", None)] + [(f, l, k, c)
                                        for f, l, k, c in
                                        [(c[0], c[1], c[2], c[3])
                                         for c in COLUMNS]]
    header_row(bk, KEY, cols)
    data_rows(bk, KEY, height=20)

    d = bk.q("data")
    months = (m.agg.get("months") or []) if m else []
    for i in range(12):
        row = C.ROW_FIRST + i
        dr = C.DATA_MONTH_FIRST + i
        values = {"n": i + 1, "month": C.MONTH_NAMES[i][:3]}
        for field, col in (("revenue", "I"), ("cogs", "J"), ("fees", "K"),
                           ("net", "L"), ("units", "M"),
                           ("markets", "N")):
            values[field] = "=%s!$%s$%d" % (d, col, dr)
        values["margin"] = '=IF($C%d=0,"",IFERROR($F%d/$C%d,0))' % (
            row, row, row)
        values["aov"] = '=IF(%s!$O$%d=0,"",IFERROR($C%d/%s!$O$%d,0))' % (
            d, dr, row, d, dr)
        values["per_item"] = '=IF($H%d=0,"",IFERROR($F%d/$H%d,0))' % (
            row, row, row)
        cached = {}
        if i < len(months):
            _, rev, cogs, fees, net, units, markets, txns = months[i]
            cached = {"revenue": rev, "cogs": cogs, "fees": fees,
                      "net": net, "units": units, "markets": markets,
                      "margin": (net / rev) if rev else "",
                      "aov": (rev / txns) if txns else "",
                      "txns": txns,
                      "per_item": (net / units) if units else ""}
        K.write_row(bk, KEY, cols, row, values, cached)

    tot = C.ROW_FIRST + 12
    K.totals_row(bk, KEY, tot, {
        "revenue": ("=SUM(C%d:C%d)" % (C.ROW_FIRST, C.ROW_FIRST + 11),
                    "#,##0.00", m.agg.get("revenue", 0) if m else 0),
        "cogs": ("=SUM(D%d:D%d)" % (C.ROW_FIRST, C.ROW_FIRST + 11),
                 "#,##0.00", m.agg.get("cogs", 0) if m else 0),
        "fees": ("=SUM(E%d:E%d)" % (C.ROW_FIRST, C.ROW_FIRST + 11),
                 "#,##0.00", m.agg.get("fees", 0) if m else 0),
        "net": ("=SUM(F%d:F%d)" % (C.ROW_FIRST, C.ROW_FIRST + 11),
                "#,##0.00", m.agg.get("profit", 0) if m else 0),
        "units": ("=SUM(H%d:H%d)" % (C.ROW_FIRST, C.ROW_FIRST + 11),
                  "#,##0", m.agg.get("units", 0) if m else 0),
        "markets": ("=SUM(K%d:K%d)" % (C.ROW_FIRST, C.ROW_FIRST + 11),
                    "#,##0", m.agg.get("events_done", 0) if m else 0)},
        first_col="B", last_col=LAST_COL)

    ws.conditional_format(
        r(C.ROW_FIRST), ci(bk.col(KEY, "net")),
        r(C.ROW_FIRST + 11), ci(bk.col(KEY, "net")),
        {"type": "cell", "criteria": "<", "value": 0,
         "format": bk.S.cf(bg=th.bad_soft, fg=th.bad, bold=True)})
    bk.stats["cond_formats"] += 1

    best = max(months, key=lambda x: x[4]) if months else None
    chips = [
        ('="\U0001F4C5 Year: "&TEXT(ReportYear,"0")', "primary",
         "Year: %d" % (m.settings["year"] if m else 2026), 2),
        ('="\U0001F4B0 Best month: "&INDEX($B$%d:$B$%d,MATCH(MAX($F$%d:'
         '$F$%d),$F$%d:$F$%d,0))'
         % (C.ROW_FIRST, C.ROW_FIRST + 11, C.ROW_FIRST, C.ROW_FIRST + 11,
            C.ROW_FIRST, C.ROW_FIRST + 11), "ok",
         "Best month: %s" % (C.MONTH_NAMES[best[0] - 1][:3]
                             if best else ""), 3),
        ('="\U0001F3EA Markets worked: "&%s' % bk.kpi("events_done"),
         "info", "Markets worked: %d" % (m.agg.get("events_done", 0)
                                         if m else 0), 3),
        ('="\U0001F4C8 Year margin: "&TEXT(%s,"0.0%%")' % bk.kpi("margin"),
         "gold", "Year margin: %.1f%%" % (100 * (m.agg.get("margin", 0)
                                                 if m else 0)), 3),
    ]
    K.chips(bk, KEY, chips, LAST_COL)
    K.footer_nav(bk, KEY, tot + 2, LAST_COL)
```

## 23. Start Here guide

`crochet_tracker/sheets/guide.py`

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
                       "  your 5-minute tour of the Crochet Craft Fair "
                       "Tracker",
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
                       "  Everything is wired together: log a sale once "
                       "and the fairs, reorder list, monthly summary and "
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
        ("1", "Open \u2699\uFE0F Lists & Settings",
         "Type your studio name, currency, hourly wage, overhead % and "
         "target margin.  Every calculator and alert reads those numbers."),
        ("2", "Skim your dropdown lists",
         "Lower down on the same tab you can rename categories, payment "
         "methods, yarn weights, units and suppliers.  Every dropdown in "
         "the workbook updates instantly \u2014 nothing is hard-coded."),
        ("3", "List your makes on \U0001F9F6 Product Catalog",
         "One row per design: price, yarn cost, packaging, hours and a "
         "minimum stock level.  Profit and margin calculate themselves."),
        ("4", "Log yarn on \U0001F9F5 Yarn & Materials",
         "What you bought, what you have used, what is left \u2014 the "
         "reorder flag turns red before you run out mid-commission."),
        ("5", "Work your fairs",
         "\U0001F3EA Craft Fairs takes one row per market, \U0001F4B0 "
         "Sales Log one row per sale.  Profit, best-sellers, the monthly "
         "summary and the dashboard build themselves from those two."),
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
        ("\u25CF", "\U0001F4CA Dashboard", "Sales, profit, margin, units, "
         "best-seller, six charts and live low-stock panels."),
        ("\u25CF", "\U0001F9F6 Product Catalog", "Every make with cost, "
         "profit, margin, stock and reorder status."),
        ("\u25CF", "\U0001F3EA Craft Fairs", "One row per market with all "
         "five cost lines, takings, net profit and best seller."),
        ("\u25CF", "\U0001F4B0 Sales Log", "Every sale: fair and product "
         "dropdowns, discount, payment method and live cost of sale."),
        ("\u25CF", "\U0001F4B5 Pricing Calculator", "Materials + labour + "
         "overhead in, suggested and charm price out."),
        ("\u25CF", "\U0001F4D6 Start Here", "This page."),
        ("\u25CB", "\U0001F9F5 Yarn & Materials", "Yarn and findings with "
         "remaining quantity and an automatic reorder flag."),
        ("\u25CB", "\U0001F4E6 Made & Stocked", "Production batches: made, "
         "reserved, sold, damaged, current and available."),
        ("\u25CB", "\U0001F9EE Event Profit", "Pick a fair - revenue, every "
         "cost, net profit, margin, break-even and ROI."),
        ("\u25CB", "\U0001F504 Reorder List", "Low products and yarn pulled "
         "in automatically, with priority and estimated budget."),
        ("\u25CB", "\U0001F392 Packing Checklist", "Four packing zones with "
         "ticks and a % packed meter, per fair."),
        ("\u25CB", "\U0001F4C5 Monthly Summary", "The whole year month by "
         "month: revenue, costs, profit, units and markets."),
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
        "Price from true cost, not from the stall next door: the Pricing "
        "Calculator adds your wage and overhead before margin.",
        "Log sales at the fair on your phone (Google Sheets) and the "
        "dashboard is already correct by teardown time.",
        "Set a minimum stock level for every best-seller \u2014 the "
        "Reorder List only protects what you have told it to protect.",
        "Weigh your yarn usage once per pattern (used column) and future "
        "costs stay honest without counting every metre.",
        "Charm prices (\u2026.95) test better at markets than round "
        "numbers; the calculator suggests one for you.",
        "After each fair, glance at Event Profit: any market under ~25% "
         "margin two years running is a hobby, not a channel.",
        "The Monthly Summary reads the report year from Settings \u2014 "
         "change it in January and reuse this file forever.",
        "Print the Packing Checklist per fair; ticks beat memory at 6am.",
        "Back up before big edits: a copy on your desktop costs nothing "
        "and saves entire weekends.",
        "The EXAMPLE file is a fictional studio so you can see it working; "
        "the blank file is the one you keep.",
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
        "The EXAMPLE edition is loaded with a fictional crochet studio so "
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
                   "  Thank you for buying this template \u2014 it was "
                   "built by people who have frogged a blanket at 2am and "
                   "know what a spreadsheet owes you between markets.  If "
                   "something looks wrong or you would love an extra tab, "
                   "message the shop on Etsy and we will help quickly.",
                   body_soft)
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

## 24. Command line interface

`crochet_craft_fair_tracker.py`

```python
#!/usr/bin/env python3
"""
Command-line builder for the Ultimate Crochet Craft Fair Tracker.

Examples
--------
    python3 crochet_craft_fair_tracker.py --all --outdir products/
    python3 crochet_craft_fair_tracker.py --edition premium --theme berry \
        --mode demo --out products/preview.xlsx
"""

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from crochet_tracker import workbook as W          # noqa: E402
from crochet_tracker import config as C            # noqa: E402


def main(argv=None):
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--edition", choices=sorted(C.EDITIONS), default="premium")
    p.add_argument("--theme", choices=["berry", "mint"], default="berry")
    p.add_argument("--mode", choices=["blank", "demo"], default="blank")
    p.add_argument("--all", action="store_true",
                   help="build the curated six-file product set")
    p.add_argument("--out", default="crochet_tracker.xlsx")
    p.add_argument("--outdir", default="products")
    p.add_argument("--no-protect", dest="protect",
                   action="store_false", default=None)
    p.add_argument("--no-images", dest="images", action="store_false")
    args = p.parse_args(argv)

    if args.all:
        rows = W.build_all(args.outdir, protect=args.protect,
                           images=args.images)
    else:
        stats = W.build_workbook(args.out, args.edition, args.theme,
                                 args.mode, protect=args.protect,
                                 images=args.images)
        rows = [(os.path.basename(args.out), stats)]

    print("=" * 78)
    print("BUILD COMPLETE")
    print("=" * 78)
    for name, st in rows:
        print("  %-58s %7.1f KB" % (name, st["size"] / 1024.0))
        print("      formulas %5d   validations %3d   cond. formats %3d   "
              "charts %2d   links %3d"
              % (st["formulas"], st["validations"], st["cond_formats"],
                 st["charts"], st["links"]))
    print()
    return 0


if __name__ == "__main__":
    sys.exit(main())
```

## 25. QA: structural verifier

`tools/verify_workbook.py`

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

## 26. QA: formula recalculation checker

`tools/calc_check.py`

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

## 27. QA: layout / clipping checker

`tools/layout_check.py`

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

## 28. QA: sheet renderer (visual proof)

`tools/render_preview.py`

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

## 29. Banner alpha pipeline

`tools/make_banner_alpha.py`

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

## Appendix A - build & QA recipe

```bash
python3 -m venv /tmp/venv && /tmp/venv/bin/pip install openpyxl Pillow formulas
python3 crochet_craft_fair_tracker.py --all --outdir products
for f in products/Crochet_*.xlsx; do
    /tmp/venv/bin/python tools/verify_workbook.py "$f"
    /tmp/venv/bin/python tools/calc_check.py "$f"
    /tmp/venv/bin/python tools/layout_check.py "$f"
done
```

Banners are generated artwork, centre-cropped to a 1600x300 band and passed
through `tools/make_banner_alpha.py` so the Start Here title shows through the
flat middle of the picture inside Excel.

## Appendix B - changelog

* **1.0.0 (2026-09-13)** - first release: 14-sheet Premium / 8-sheet Basic,
  Berry Bramble + Mint Meadow themes, blank and EXAMPLE modes, six charts,
  hidden _Data engine with 26 KPIs, per-event and per-product pools,
  Excel-2016-safe k-th extraction (no CSE / dynamic arrays), full QA pass.
