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

## ✅ Quality Checks

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
