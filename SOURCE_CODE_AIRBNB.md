# Ultimate Airbnb & Short-Term Rental Management Spreadsheet — source code

Complete source code of the generator that writes the `.xlsx` product.

Writes a 20 tab `.xlsx` workbook with over 138,000 formulas for an Airbnb, Vrbo or Booking.com host running more than one property: a property register with automatic IDs, rates, fees and ownership basis; a booking tracker that turns two dates into nights, looks up the nightly rate, applies the right platform commission and flags a double booking before a guest finds it; a formula-driven calendar showing one property by month or twelve properties over a fortnight; a guest CRM with repeat-stay history; the nine arrival steps; a cleaning schedule generated from the check-outs with a printable turnover checklist; expected payouts against what the platform really paid; 28 expense categories including mortgage interest and rent; supplies with low-stock flags; maintenance with warranty countdowns; reviews with seven sub-ratings; cleaners and contractors; a tax summary; occupancy, ADR and RevPAR by property and month; and monthly and annual reports. Document author is **Novality Store**; every formula cell is locked and every sheet protected with the password **`premium`**, while the cells you are meant to type in stay open.

Nothing is hard coded to one platform's pricing. The eight platforms, their commission and their payment-processing rates are a table on Settings with two spare rows, so a host on a different split, in a different market or on a negotiated deal types in their own figures and every fee, payout and profit figure in the workbook follows. The lodging tax and who remits it, the turnover costs, the targets, the alert windows and every drop-down list are settings, not assumptions.

Two things it refuses to do, both deliberately. It never claims to synchronise with a booking platform - there is no iCal, no API and no live feed, and the workbook says so on the Getting Started tab. And **ADR and RevPAR exclude the cleaning fee**, because including it is the most common error in rental spreadsheets sold online and it flatters both numbers; the validator fails the build if either formula so much as mentions cleaning.

Pure Python, one dependency. Tested on XlsxWriter 3.2.0 and on the much older 0.3.0, so it does not rely on anything recent. The workbook itself is Excel 2010 safe — no dynamic arrays, no `XLOOKUP`, no `IFS`, nothing that needs array entry — and it opens in Google Sheets, LibreOffice and Numbers.

```bash
python build_airbnb_workbook.py --help

  -o, --output PATH     where to write the file
  --blank               no example data, just the empty workbook
  --no-protect          leave the sheets unlocked
  --password TEXT       protection password (default: premium)
  --author TEXT         document author  (default: Novality Store)
  --business TEXT       your business name on the cover
  --owner TEXT          the host or owner's name
  --currency SYMBOL     currency symbol  (default: $)
  --currency-code CODE  three letter code (default: USD)
  --nightly AMOUNT      the default nightly rate
  --cleaning-fee AMT    the cleaning fee charged to a guest
  --cleaning-cost AMT   what a turnover actually costs you
  --deposit AMOUNT      the security deposit
  --min-nights N        the default minimum stay
  --tax PERCENT         lodging / occupancy tax rate
  --tax-self            you remit the lodging tax, not the platform
  --occupancy-target P  the occupancy you are aiming at
  --adr-target AMOUNT   the average daily rate you are aiming at
  --revenue-target AMT  monthly revenue target
  --properties N        properties the targets are scaled for
  --rows N              rows of capacity in the big registers
  --today YYYY-MM-DD    pretend it is this date (for the demo data)
```

---

## Contents

* [0. How it is put together](#0-how-it-is-put-together)
  * [The build pipeline](#the-build-pipeline)
  * [Describing a column](#describing-a-column)
  * [The rules every formula obeys](#the-rules-every-formula-obeys)
  * [Ranking columns instead of a hidden engine](#ranking-columns-instead-of-a-hidden-engine)
  * [Every tab at a glance](#every-tab-at-a-glance)
  * [Every file at a glance](#every-file-at-a-glance)
* [1. Run it](#1-run-it)
  * [`build_airbnb_workbook.py`](#build_airbnb_workbookpy)
* [2. The package](#2-the-package)
  * [`airbnb_rental_spreadsheet/__init__.py`](#airbnb_rental_spreadsheet__init__py)
  * [`airbnb_rental_spreadsheet/__main__.py`](#airbnb_rental_spreadsheet__main__py)
  * [`airbnb_rental_spreadsheet/cli.py`](#airbnb_rental_spreadsheetclipy)
  * [`airbnb_rental_spreadsheet/config.py`](#airbnb_rental_spreadsheetconfigpy)
  * [`airbnb_rental_spreadsheet/layout.py`](#airbnb_rental_spreadsheetlayoutpy)
  * [`airbnb_rental_spreadsheet/formats.py`](#airbnb_rental_spreadsheetformatspy)
  * [`airbnb_rental_spreadsheet/writer.py`](#airbnb_rental_spreadsheetwriterpy)
  * [`airbnb_rental_spreadsheet/names.py`](#airbnb_rental_spreadsheetnamespy)
  * [`airbnb_rental_spreadsheet/build.py`](#airbnb_rental_spreadsheetbuildpy)
  * [`airbnb_rental_spreadsheet/sample_data.py`](#airbnb_rental_spreadsheetsample_datapy)
* [3. Properties and bookings](#3-properties-and-bookings)
  * [`airbnb_rental_spreadsheet/sheets/__init__.py`](#airbnb_rental_spreadsheetsheets__init__py)
  * [`airbnb_rental_spreadsheet/sheets/common.py`](#airbnb_rental_spreadsheetsheetscommonpy)
  * [`airbnb_rental_spreadsheet/sheets/properties.py`](#airbnb_rental_spreadsheetsheetspropertiespy)
  * [`airbnb_rental_spreadsheet/sheets/bookings.py`](#airbnb_rental_spreadsheetsheetsbookingspy)
  * [`airbnb_rental_spreadsheet/sheets/guests.py`](#airbnb_rental_spreadsheetsheetsguestspy)
* [4. The calendar](#4-the-calendar)
  * [`airbnb_rental_spreadsheet/sheets/calendar.py`](#airbnb_rental_spreadsheetsheetscalendarpy)
* [5. The turnover](#5-the-turnover)
  * [`airbnb_rental_spreadsheet/sheets/arrivals.py`](#airbnb_rental_spreadsheetsheetsarrivalspy)
  * [`airbnb_rental_spreadsheet/sheets/cleaning.py`](#airbnb_rental_spreadsheetsheetscleaningpy)
  * [`airbnb_rental_spreadsheet/sheets/turnover.py`](#airbnb_rental_spreadsheetsheetsturnoverpy)
* [6. The money](#6-the-money)
  * [`airbnb_rental_spreadsheet/sheets/payouts.py`](#airbnb_rental_spreadsheetsheetspayoutspy)
  * [`airbnb_rental_spreadsheet/sheets/expenses.py`](#airbnb_rental_spreadsheetsheetsexpensespy)
  * [`airbnb_rental_spreadsheet/sheets/tax.py`](#airbnb_rental_spreadsheetsheetstaxpy)
* [7. The property itself](#7-the-property-itself)
  * [`airbnb_rental_spreadsheet/sheets/supplies.py`](#airbnb_rental_spreadsheetsheetssuppliespy)
  * [`airbnb_rental_spreadsheet/sheets/maintenance.py`](#airbnb_rental_spreadsheetsheetsmaintenancepy)
  * [`airbnb_rental_spreadsheet/sheets/reviews.py`](#airbnb_rental_spreadsheetsheetsreviewspy)
  * [`airbnb_rental_spreadsheet/sheets/vendors.py`](#airbnb_rental_spreadsheetsheetsvendorspy)
* [8. The pages](#8-the-pages)
  * [`airbnb_rental_spreadsheet/sheets/performance.py`](#airbnb_rental_spreadsheetsheetsperformancepy)
  * [`airbnb_rental_spreadsheet/sheets/dashboard.py`](#airbnb_rental_spreadsheetsheetsdashboardpy)
  * [`airbnb_rental_spreadsheet/sheets/getting_started.py`](#airbnb_rental_spreadsheetsheetsgetting_startedpy)
  * [`airbnb_rental_spreadsheet/sheets/settings.py`](#airbnb_rental_spreadsheetsheetssettingspy)
  * [`airbnb_rental_spreadsheet/sheets/monthly.py`](#airbnb_rental_spreadsheetsheetsmonthlypy)
  * [`airbnb_rental_spreadsheet/sheets/annual.py`](#airbnb_rental_spreadsheetsheetsannualpy)
* [9. The quality tools](#9-the-quality-tools)
  * [`tools/validate_airbnb_workbook.py`](#toolsvalidate_airbnb_workbookpy)
  * [`tools/check_parens.py`](#toolscheck_parenspy)
  * [`tools/check_airbnb_sample.py`](#toolscheck_airbnb_samplepy)
  * [`tools/recalc_check.py`](#toolsrecalc_checkpy)
* [Appendix A. Data dictionary](#appendix-a-data-dictionary)

---

## 0. How it is put together

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

`build_airbnb_workbook.py` parses the command line into a `Config`, then
`airbnb_rental_spreadsheet.build.build_workbook(path, cfg)` does six things in order:

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
6. **Activate the first tab** and close the file - the buyer should land on the Dashboard, not on Settings.

### Describing a column

A column is declared once, for one row, with placeholders:

```python
Col("er", "Engagement rate", 11, "percent",
    formula='=IFERROR({engagements}{r}/{reach}{r},"")',
    cf=[{"type": "cell", "criteria": ">=", "value": "=TargetEngagement",
         "fmt": "good"}])
```

`Col` carries the key, the header text, the width, the style, the
formula template, the data validation, the header comment, the
conditional formats, whether the column is hidden and which outline
group it belongs to. `Registry.render()` expands the template per row:

| Token | Becomes |
|-------|---------|
| `{r}` / `{r_prev}` | this row number / the row above |
| `{key}` | the column letter of another column on the same sheet |
| `{first}` / `{last}` | the first and last data row of the sheet |
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
`Stores.spend_rank`.


### Every tab at a glance

20 tabs in the order they appear in the finished workbook. "Rows" is the capacity of a register at the default `--rows` setting.

| # | Tab | Kind | Written by | Columns | Rows |
|---|-----|------|------------|--------:|-----:|
| 1 | **Dashboard** | page | `airbnb_rental_spreadsheet/sheets/dashboard.py` | - | - |
| 2 | **Getting Started** | page | `airbnb_rental_spreadsheet/sheets/getting_started.py` | - | - |
| 3 | **Settings & Lists** | page | `airbnb_rental_spreadsheet/sheets/settings.py` | - | - |
| 4 | **Property Database** | register | `airbnb_rental_spreadsheet/sheets/properties.py` | 63 | 30 |
| 5 | **Booking Tracker** | register | `airbnb_rental_spreadsheet/sheets/bookings.py` | 60 | 900 |
| 6 | **Booking Calendar** | page | `airbnb_rental_spreadsheet/sheets/calendar.py` | - | - |
| 7 | **Guest Database** | register | `airbnb_rental_spreadsheet/sheets/guests.py` | 29 | 400 |
| 8 | **Check-in & Check-out** | register | `airbnb_rental_spreadsheet/sheets/arrivals.py` | 32 | 900 |
| 9 | **Cleaning Schedule** | register | `airbnb_rental_spreadsheet/sheets/cleaning.py` | 45 | 900 |
| 10 | **Turnover Checklist** | page | `airbnb_rental_spreadsheet/sheets/turnover.py` | - | - |
| 11 | **Platform Revenue & Fees** | register | `airbnb_rental_spreadsheet/sheets/payouts.py` | 27 | 900 |
| 12 | **Expense Tracker** | register | `airbnb_rental_spreadsheet/sheets/expenses.py` | 26 | 700 |
| 13 | **Supplies Inventory** | register | `airbnb_rental_spreadsheet/sheets/supplies.py` | 27 | 300 |
| 14 | **Maintenance Log** | register | `airbnb_rental_spreadsheet/sheets/maintenance.py` | 30 | 300 |
| 15 | **Reviews Tracker** | register | `airbnb_rental_spreadsheet/sheets/reviews.py` | 32 | 600 |
| 16 | **Staff & Vendors** | register | `airbnb_rental_spreadsheet/sheets/vendors.py` | 31 | 60 |
| 17 | **Tax Summary** | page | `airbnb_rental_spreadsheet/sheets/tax.py` | - | - |
| 18 | **Occupancy & Performance** | register | `airbnb_rental_spreadsheet/sheets/performance.py` | 34 | 400 |
| 19 | **Monthly Reports** | page | `airbnb_rental_spreadsheet/sheets/monthly.py` | - | - |
| 20 | **Annual Reports** | page | `airbnb_rental_spreadsheet/sheets/annual.py` | - | - |

### Every file at a glance

| File | Lines | What it is |
|------|------:|------------|
| [`build_airbnb_workbook.py`](#build_airbnb_workbookpy) | 17 | Build the Ultimate Airbnb & Short-Term Rental Management Spreadsheet |
| [`airbnb_rental_spreadsheet/__init__.py`](#airbnb_rental_spreadsheet__init__py) | 19 | Ultimate Airbnb & Short-Term Rental Management Spreadsheet |
| [`airbnb_rental_spreadsheet/__main__.py`](#airbnb_rental_spreadsheet__main__py) | 5 | ``python -m airbnb_rental_spreadsheet`` |
| [`airbnb_rental_spreadsheet/cli.py`](#airbnb_rental_spreadsheetclipy) | 193 | Command line front end |
| [`airbnb_rental_spreadsheet/config.py`](#airbnb_rental_spreadsheetconfigpy) | 694 | Every default, list, rate and colour the workbook is built from |
| [`airbnb_rental_spreadsheet/layout.py`](#airbnb_rental_spreadsheetlayoutpy) | 253 | Layout engine |
| [`airbnb_rental_spreadsheet/formats.py`](#airbnb_rental_spreadsheetformatspy) | 202 | Cell format factory |
| [`airbnb_rental_spreadsheet/writer.py`](#airbnb_rental_spreadsheetwriterpy) | 473 | Sheet writing helpers shared by every tab of the workbook |
| [`airbnb_rental_spreadsheet/names.py`](#airbnb_rental_spreadsheetnamespy) | 81 | Workbook level defined names that point at the register sheets |
| [`airbnb_rental_spreadsheet/build.py`](#airbnb_rental_spreadsheetbuildpy) | 160 | Workbook builder: puts the twenty tabs together in order |
| [`airbnb_rental_spreadsheet/sample_data.py`](#airbnb_rental_spreadsheetsample_datapy) | 1,151 | A worked example: four properties, four months, every tab populated |
| [`airbnb_rental_spreadsheet/sheets/__init__.py`](#airbnb_rental_spreadsheetsheets__init__py) | 0 |  |
| [`airbnb_rental_spreadsheet/sheets/common.py`](#airbnb_rental_spreadsheetsheetscommonpy) | 483 | Small building blocks shared by the register sheets |
| [`airbnb_rental_spreadsheet/sheets/properties.py`](#airbnb_rental_spreadsheetsheetspropertiespy) | 489 | `Property Database` - one row per place you let, and what it earns |
| [`airbnb_rental_spreadsheet/sheets/bookings.py`](#airbnb_rental_spreadsheetsheetsbookingspy) | 527 | `Booking Tracker` - every reservation, block and enquiry, priced |
| [`airbnb_rental_spreadsheet/sheets/guests.py`](#airbnb_rental_spreadsheetsheetsguestspy) | 252 | `Guest Database` - who stayed, how often, and what they were worth |
| [`airbnb_rental_spreadsheet/sheets/calendar.py`](#airbnb_rental_spreadsheetsheetscalendarpy) | 427 | `Booking Calendar` - a month at a glance, and the next fortnight |
| [`airbnb_rental_spreadsheet/sheets/arrivals.py`](#airbnb_rental_spreadsheetsheetsarrivalspy) | 234 | `Check-in & Check-out` - the eight messages every guest needs |
| [`airbnb_rental_spreadsheet/sheets/cleaning.py`](#airbnb_rental_spreadsheetsheetscleaningpy) | 318 | `Cleaning Schedule` - the turnover, who does it and what it costs |
| [`airbnb_rental_spreadsheet/sheets/turnover.py`](#airbnb_rental_spreadsheetsheetsturnoverpy) | 245 | `Turnover Checklist` - one changeover, on one sheet of paper |
| [`airbnb_rental_spreadsheet/sheets/payouts.py`](#airbnb_rental_spreadsheetsheetspayoutspy) | 248 | `Platform Revenue & Fees` - what the platform said, and what arrived |
| [`airbnb_rental_spreadsheet/sheets/expenses.py`](#airbnb_rental_spreadsheetsheetsexpensespy) | 252 | `Expense Tracker` - everything the properties cost you |
| [`airbnb_rental_spreadsheet/sheets/tax.py`](#airbnb_rental_spreadsheetsheetstaxpy) | 340 | `Tax Summary` - your figures, grouped the way an accountant wants them |
| [`airbnb_rental_spreadsheet/sheets/supplies.py`](#airbnb_rental_spreadsheetsheetssuppliespy) | 234 | `Supplies Inventory` - linen, loo roll and the things that run out |
| [`airbnb_rental_spreadsheet/sheets/maintenance.py`](#airbnb_rental_spreadsheetsheetsmaintenancepy) | 243 | `Maintenance Log` - what is broken, who is fixing it, what it cost |
| [`airbnb_rental_spreadsheet/sheets/reviews.py`](#airbnb_rental_spreadsheetsheetsreviewspy) | 250 | `Reviews Tracker` - the seven scores, and which one is costing you |
| [`airbnb_rental_spreadsheet/sheets/vendors.py`](#airbnb_rental_spreadsheetsheetsvendorspy) | 240 | `Staff & Vendors` - the people who keep the properties running |
| [`airbnb_rental_spreadsheet/sheets/performance.py`](#airbnb_rental_spreadsheetsheetsperformancepy) | 348 | `Occupancy & Performance` - one row per property per month |
| [`airbnb_rental_spreadsheet/sheets/dashboard.py`](#airbnb_rental_spreadsheetsheetsdashboardpy) | 459 | `Dashboard` - the page that answers "how is the portfolio doing?" Twenty numbers across the top, the jobs... |
| [`airbnb_rental_spreadsheet/sheets/getting_started.py`](#airbnb_rental_spreadsheetsheetsgetting_startedpy) | 286 | `Getting Started` - twenty minutes from download to running |
| [`airbnb_rental_spreadsheet/sheets/settings.py`](#airbnb_rental_spreadsheetsheetssettingspy) | 513 | `Settings & Lists` - every number and every list the workbook runs on |
| [`airbnb_rental_spreadsheet/sheets/monthly.py`](#airbnb_rental_spreadsheetsheetsmonthlypy) | 446 | `Monthly Reports` - any month you choose, in full |
| [`airbnb_rental_spreadsheet/sheets/annual.py`](#airbnb_rental_spreadsheetsheetsannualpy) | 534 | `Annual Reports` - twelve months across, and the year that was |
| [`tools/validate_airbnb_workbook.py`](#toolsvalidate_airbnb_workbookpy) | 1,015 | Static checks on a generated Airbnb & Short-Term Rental workbook |
| [`tools/check_parens.py`](#toolscheck_parenspy) | 121 | Parenthesis and quote balance check for every formula in a workbook |
| [`tools/check_airbnb_sample.py`](#toolscheck_airbnb_samplepy) | 96 | Validate the Airbnb product's example data before it is built |
| [`tools/recalc_check.py`](#toolsrecalc_checkpy) | 216 | Really calculate every formula in a generated workbook and report any cell that ends up as an Excel error... |

---

## 1. Run it

Save the files below keeping these exact paths, then run the one liner. Nothing else is needed - no project file, no build step.

```bash
pip install XlsxWriter
python build_airbnb_workbook.py
```

The workbook lands in `output/Airbnb & Short-Term Rental Management Spreadsheet.xlsx`.

### `build_airbnb_workbook.py`

*17 lines*

> Build the Ultimate Airbnb & Short-Term Rental Management Spreadsheet.
>
> python3 build_airbnb_workbook.py
> python3 build_airbnb_workbook.py -o "output/Airbnb.xlsx" --blank
> python3 build_airbnb_workbook.py --currency "GBP" --rows 1200

````python
#!/usr/bin/env python3
"""Build the Ultimate Airbnb & Short-Term Rental Management Spreadsheet.

    python3 build_airbnb_workbook.py
    python3 build_airbnb_workbook.py -o "output/Airbnb.xlsx" --blank
    python3 build_airbnb_workbook.py --currency "GBP" --rows 1200
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from airbnb_rental_spreadsheet.cli import main   # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main())
````

---

## 2. The package

Configuration, the formula template engine, cell formats, the sheet writer, the cross sheet names and the demo portfolio. These are the parts every tab leans on. `config.py` is also where the platform fee table starts out, so a host on any platform, in any market, on any commission deal can be matched without touching a formula.

### `airbnb_rental_spreadsheet/__init__.py`

*19 lines*

> Ultimate Airbnb & Short-Term Rental Management Spreadsheet.
>
> Build the workbook with::
>
>     python -m airbnb_rental_spreadsheet -o "output/Airbnb.xlsx"
>
> or from Python::
>
>     from airbnb_rental_spreadsheet import build_workbook
>     build_workbook("Airbnb Rental Management Spreadsheet.xlsx")

````python
"""Ultimate Airbnb & Short-Term Rental Management Spreadsheet.

Build the workbook with::

    python -m airbnb_rental_spreadsheet -o "output/Airbnb.xlsx"

or from Python::

    from airbnb_rental_spreadsheet import build_workbook
    build_workbook("Airbnb Rental Management Spreadsheet.xlsx")
"""

from .build import build_workbook
from .config import (AUTHOR, LISTING_TITLE, PRODUCT_NAME,
                     PRODUCT_SUBTITLE, VERSION, Config)

__all__ = ["build_workbook", "Config", "AUTHOR", "PRODUCT_NAME",
           "PRODUCT_SUBTITLE", "LISTING_TITLE", "VERSION"]
__version__ = VERSION
````

### `airbnb_rental_spreadsheet/__main__.py`

*5 lines*

> ``python -m airbnb_rental_spreadsheet``.

````python
"""``python -m airbnb_rental_spreadsheet``."""

from .cli import main

raise SystemExit(main())
````

### `airbnb_rental_spreadsheet/cli.py`

*193 lines · 3 functions*

> Command line front end.
>
> python -m airbnb_rental_spreadsheet
> python -m airbnb_rental_spreadsheet -o "output/Airbnb.xlsx" --blank
> python -m airbnb_rental_spreadsheet --currency "GBP" --rows 1200

| Defines | What it does |
|---------|--------------|
| `build_parser()` |  |
| `config_from_args(args)` |  |
| `main(argv)` |  |

**Module constants:** `DEFAULT_OUTPUT` (os.path.join(...)).

````python
"""Command line front end.

    python -m airbnb_rental_spreadsheet
    python -m airbnb_rental_spreadsheet -o "output/Airbnb.xlsx" --blank
    python -m airbnb_rental_spreadsheet --currency "GBP" --rows 1200
"""

from __future__ import annotations

import argparse
import datetime as _dt
import os
import sys
from typing import List, Optional

from .build import SHEET_ORDER, build_workbook
from .config import (AUTHOR, LISTING_TITLE, MAX_PROPERTIES,
                     PRODUCT_NAME, Config)

DEFAULT_OUTPUT = os.path.join(
    "output", "Airbnb & Short-Term Rental Management Spreadsheet.xlsx")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="airbnb_rental_spreadsheet",
        description=f"Build the {PRODUCT_NAME} workbook.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument("-o", "--output", default=DEFAULT_OUTPUT,
                        help="where to write the .xlsx file")
    parser.add_argument("--blank", action="store_true",
                        help="leave the registers empty (no example "
                             "portfolio)")
    parser.add_argument("--no-protect", action="store_true",
                        help="do not protect the sheets (for "
                             "development)")
    parser.add_argument("--password", default=None,
                        help="sheet protection password")
    parser.add_argument("--author", default=AUTHOR,
                        help="document author / creator")
    parser.add_argument("--business", default=None,
                        help="pre-fill the business name on Settings")
    parser.add_argument("--owner", default=None,
                        help="pre-fill the host or owner's name")
    parser.add_argument("--currency", default=None,
                        help="currency symbol, e.g. $ £ € A$")
    parser.add_argument("--currency-code", default=None,
                        help="three letter currency code, e.g. USD")
    parser.add_argument("--year", type=int, default=None,
                        help="the year the reports open on")
    parser.add_argument("--nightly", type=float, default=None,
                        help="default nightly rate for a new property")
    parser.add_argument("--cleaning-fee", type=float, default=None,
                        help="default cleaning fee charged to guests")
    parser.add_argument("--cleaning-cost", type=float, default=None,
                        help="default cost of a turnover to you")
    parser.add_argument("--deposit", type=float, default=None,
                        help="default security deposit")
    parser.add_argument("--min-nights", type=int, default=None,
                        help="default minimum stay")
    parser.add_argument("--tax", type=float, default=None,
                        help="lodging tax rate as a percentage, e.g. "
                             "6 for 6%% (zero by default: most "
                             "platforms collect and remit it)")
    parser.add_argument("--tax-self", action="store_true",
                        help="you collect and remit lodging tax "
                             "yourself rather than the platform")
    parser.add_argument("--occupancy-target", type=float, default=None,
                        help="occupancy target as a percentage, e.g. 72")
    parser.add_argument("--adr-target", type=float, default=None,
                        help="average daily rate target")
    parser.add_argument("--revenue-target", type=float, default=None,
                        help="monthly revenue target")
    parser.add_argument("--properties", type=int, default=None,
                        help="how many properties the example "
                             "portfolio has")
    parser.add_argument("--rows", type=int, default=None,
                        help="scale every register (the default row "
                             "counts are scaled to fit this many "
                             "booking rows)")
    parser.add_argument("--today", default=None,
                        help="treat this ISO date as 'today' when "
                             "generating the example data")
    parser.add_argument("--quiet", action="store_true",
                        help="only print the output path")
    return parser


def config_from_args(args: argparse.Namespace) -> Config:
    kwargs = {}
    if args.blank:
        kwargs["include_sample_data"] = False
    if args.no_protect:
        kwargs["protect_sheets"] = False
    if args.password:
        kwargs["password"] = args.password
    if args.author:
        kwargs["author"] = args.author
    if args.currency:
        kwargs["currency_symbol"] = args.currency
    if args.currency_code:
        kwargs["currency_code"] = args.currency_code
    if args.year:
        kwargs["plan_year"] = args.year
    if args.nightly is not None:
        kwargs["default_nightly_rate"] = args.nightly
    if args.cleaning_fee is not None:
        kwargs["default_cleaning_fee"] = args.cleaning_fee
    if args.cleaning_cost is not None:
        kwargs["default_cleaning_cost"] = args.cleaning_cost
    if args.deposit is not None:
        kwargs["default_security_deposit"] = args.deposit
    if args.min_nights is not None:
        kwargs["default_min_nights"] = args.min_nights
    if args.tax is not None:
        kwargs["tax_rate"] = round(args.tax / 100.0, 5)
    if args.tax_self:
        kwargs["tax_collected_by_platform"] = "No"
    if args.occupancy_target is not None:
        kwargs["occupancy_target"] = round(
            args.occupancy_target / 100.0, 4)
    if args.adr_target is not None:
        kwargs["adr_target"] = args.adr_target
    if args.revenue_target is not None:
        kwargs["revenue_target_month"] = args.revenue_target
        kwargs["revenue_target_year"] = args.revenue_target * 12
    if args.properties is not None:
        kwargs["property_count"] = args.properties
    if args.today:
        kwargs["today"] = _dt.date.fromisoformat(args.today)
    cfg = Config(**kwargs)
    if args.business:
        cfg.host.name = args.business
    if args.owner:
        cfg.host.owner = args.owner
    if args.rows:
        # every register scales against the Booking Tracker, which is
        # what `row_count()` and `scale_rows()` both read
        cfg.rows_override = args.rows
    return cfg


def main(argv: Optional[List[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        cfg = config_from_args(args)
    except ValueError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 2

    path = build_workbook(args.output, cfg)
    size = os.path.getsize(path) / 1024.0

    if args.quiet:
        print(path)
        return 0

    platforms = ", ".join(name for name, _c, _p, counts, _n
                          in cfg.platform_table if counts)
    who = ("the platform" if cfg.tax_collected_by_platform == "Yes"
           else "you")

    print(f"{cfg.product_name}")
    print(f"  listing title : {LISTING_TITLE}")
    print(f"  author        : {cfg.author}")
    print(f"  sheets        : {len(SHEET_ORDER)}")
    print(f"  year          : {cfg.year}")
    print(f"  business      : {cfg.host.name}")
    print(f"  properties    : {cfg.property_count} in the example "
          f"portfolio (room for {MAX_PROPERTIES} on the "
          f"dashboard)")
    print(f"  platforms     : {platforms}")
    print(f"  nightly rate  : {cfg.currency_symbol}"
          f"{cfg.default_nightly_rate:,.2f}, cleaning fee "
          f"{cfg.currency_symbol}{cfg.default_cleaning_fee:,.2f}, "
          f"cost {cfg.currency_symbol}"
          f"{cfg.default_cleaning_cost:,.2f}")
    print(f"  lodging tax   : {cfg.tax_rate:.2%}, remitted by {who}")
    print(f"  targets       : {cfg.occupancy_target:.0%} occupancy, "
          f"ADR {cfg.currency_symbol}{cfg.adr_target:,.0f}, RevPAR "
          f"{cfg.currency_symbol}{cfg.revpar_target:,.0f}")
    print(f"  target        : {cfg.currency_symbol}"
          f"{cfg.revenue_target_month:,.0f} a month")
    print(f"  example data  : "
          f"{'no' if not cfg.include_sample_data else 'yes'}")
    print(f"  protection    : "
          f"{'off' if not cfg.protect_sheets else 'on, password ' + cfg.password}")
    print(f"  written to    : {os.path.abspath(path)}  ({size:,.0f} KB)")
    return 0


if __name__ == "__main__":        # pragma: no cover
    raise SystemExit(main())
````

### `airbnb_rental_spreadsheet/config.py`

*694 lines · 3 classes*

> Every default, list, rate and colour the workbook is built from.
>
> Nothing in the generated spreadsheet is hard coded to one country, one
> platform or one pricing model.  Platform commission rates, payment
> processing fees, tax rates, cleaning costs and every drop down list are
> settings the host edits on the *Settings & Lists* tab, because an Airbnb
> host in Lisbon, a Vrbo host in Colorado and a Booking.com partner in
> Manchester all pay different fees on different terms.

| Defines | What it does |
|---------|--------------|
| `class Palette` |  |
| `class HostDefaults` | Sample values written into the business profile on Settings |
| `class Config` | Everything the build needs to know, in one place |

**Module constants:** `PRODUCT_NAME` ('Airbnb & Short-Term Rental Management Spreadsheet'), `PRODUCT_SHORT` ('Short-Term Rental Manager'), `LISTING_TITLE` ('Airbnb Host Spreadsheet, Short-Term Rental Management, B...), `PRODUCT_SUBTITLE` ('Bookings, cleaning turnovers, platform fees, expenses, o...), `AUTHOR` ('Novality Store'), `VERSION` ('1.0.0'), `PASSWORD` ('premium'), `PLATFORMS` (8 items), `PLATFORM_NAMES` ([p[0] for p in PLATFORMS]), `BOOKING_STATUSES` (7 items), `HOLDING_STATUSES` (5 items), `EARNING_STATUSES` (3 items), `BLOCK_STATUSES` (2 items), `PAYMENT_STATUSES` (6 items) and 44 more.

````python
"""
Every default, list, rate and colour the workbook is built from.

Nothing in the generated spreadsheet is hard coded to one country, one
platform or one pricing model.  Platform commission rates, payment
processing fees, tax rates, cleaning costs and every drop down list are
settings the host edits on the *Settings & Lists* tab, because an Airbnb
host in Lisbon, a Vrbo host in Colorado and a Booking.com partner in
Manchester all pay different fees on different terms.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Dict, List, Optional, Tuple

PRODUCT_NAME = "Airbnb & Short-Term Rental Management Spreadsheet"
PRODUCT_SHORT = "Short-Term Rental Manager"
LISTING_TITLE = (
    "Airbnb Host Spreadsheet, Short-Term Rental Management, Booking "
    "Calendar, Cleaning, Guest Tracker, Expenses and Profit Dashboard"
)
PRODUCT_SUBTITLE = (
    "Bookings, cleaning turnovers, platform fees, expenses, "
    "occupancy and profit - across every property and every "
    "platform"
)
AUTHOR = "Novality Store"
VERSION = "1.0.0"
PASSWORD = "premium"


# ---------------------------------------------------------------------------
# Booking platforms
#
# `commission` is what the platform keeps out of the host's money, and
# `payment_fee` is the card / payout processing charge on top.  Both ship
# as editable cells: Airbnb alone runs a split fee in most markets and a
# host-only fee in others, Vrbo has two completely different models, and
# Booking.com commission is negotiated per property.
# ---------------------------------------------------------------------------

#: (platform, host commission %, payment processing %, counts as a platform
#:  booking, note shown on Settings)
PLATFORMS: List[Tuple[str, float, float, str, str]] = [
    ("Airbnb", 0.030, 0.000, "Yes",
     "host-only service fee is usually 14-16%, split fee about 3%"),
    ("Vrbo", 0.050, 0.030, "Yes",
     "pay-per-booking is around 5% plus 3% card processing"),
    ("Booking.com", 0.150, 0.000, "Yes",
     "commission is negotiated, commonly 15-18%"),
    ("Expedia", 0.150, 0.000, "Yes", "similar terms to Booking.com"),
    ("Direct booking", 0.000, 0.029, "No",
     "no commission, just whatever your card processor charges"),
    ("Returning guest", 0.000, 0.029, "No", "direct repeat booking"),
    ("Corporate / relocation", 0.000, 0.020, "No",
     "invoiced stays, often monthly"),
    ("Other", 0.100, 0.000, "Yes", "anything else - set your own rate"),
]

PLATFORM_NAMES = [p[0] for p in PLATFORMS]


# ---------------------------------------------------------------------------
# Booking states
#
# A block is just a reservation that earns nothing: it takes the dates off
# the calendar, it shows up in the availability grid, and it is counted as
# a blocked night rather than a booked night.  That keeps one register
# behind the calendar instead of two.
# ---------------------------------------------------------------------------

BOOKING_STATUSES = [
    "Inquiry",
    "Confirmed",
    "Checked in",
    "Completed",
    "Cancelled",
    "Owner block",
    "Maintenance block",
]

#: statuses that hold the dates on the calendar
HOLDING_STATUSES = ["Confirmed", "Checked in", "Completed",
                    "Owner block", "Maintenance block"]

#: statuses that earn money
EARNING_STATUSES = ["Confirmed", "Checked in", "Completed"]

#: statuses that take the night off the market without earning
BLOCK_STATUSES = ["Owner block", "Maintenance block"]

PAYMENT_STATUSES = [
    "Not due yet",
    "Awaiting payout",
    "Part paid",
    "Paid in full",
    "Refunded",
    "Written off",
]

PROPERTY_TYPES = [
    "Entire apartment", "Entire house", "Villa", "Cabin", "Studio",
    "Condo", "Townhouse", "Cottage", "Loft", "Guest suite",
    "Private room", "Serviced apartment", "Tiny home", "Chalet",
]

PROPERTY_STATUSES = [
    "Active", "Inactive", "Under maintenance", "Seasonal closure",
    "Being set up", "Sold / ended",
]

OWNERSHIP_TYPES = [
    "Owned outright", "Owned with mortgage", "Rental arbitrage",
    "Co-host (managed for owner)", "Family owned", "Head lease",
]

CLEANING_STATUSES = [
    "Not scheduled", "Scheduled", "In progress", "Complete",
    "Complete - issues", "Skipped", "Cancelled",
]

CLEANING_TASKS: List[Tuple[str, str]] = [
    ("task_bedroom", "Bedrooms"),
    ("task_bathroom", "Bathrooms"),
    ("task_kitchen", "Kitchen"),
    ("task_living", "Living area"),
    ("task_floors", "Floors"),
    ("task_laundry", "Laundry & linen"),
    ("task_rubbish", "Rubbish out"),
    ("task_restock", "Restock supplies"),
    ("task_inspection", "Final inspection"),
]

#: the printable room-by-room checklist on the Turnover Checklist tab
TURNOVER_CHECKLIST: List[Tuple[str, str, List[str]]] = [
    ("task_bedroom", "Bedrooms", [
        "Strip and change all bed sheets",
        "Replace pillowcases and check pillows",
        "Check mattress and protector for stains",
        "Make beds to listing-photo standard",
        "Dust surfaces, headboards and skirting",
        "Vacuum floors and under the beds",
        "Check bedside lights, lamps and sockets",
        "Empty and wipe out wardrobes and drawers",
    ]),
    ("task_bathroom", "Bathrooms", [
        "Clean and disinfect the toilet",
        "Clean shower, bath, screen and grout",
        "Clean sink, taps and mirrors",
        "Replace all towels and bath mats",
        "Refill shampoo, conditioner and soap",
        "Replace toilet roll (and leave a spare)",
        "Check drains, plumbing and extractor fan",
        "Empty the bin and replace the liner",
    ]),
    ("task_kitchen", "Kitchen", [
        "Clear, clean and disinfect all countertops",
        "Wash, dry and put away every dish",
        "Clean hob, oven, microwave and kettle",
        "Empty the fridge and freezer, wipe shelves",
        "Refill coffee, tea, sugar and condiments",
        "Restock dishwasher tablets and washing-up liquid",
        "Check cutlery, glasses and crockery counts",
        "Empty bins and replace liners",
    ]),
    ("task_living", "Living area", [
        "Dust all surfaces, shelves and skirting",
        "Vacuum or mop all floors and rugs",
        "Plump cushions and straighten furniture",
        "Check TV, remote and batteries",
        "Test the Wi-Fi and leave the code visible",
        "Tidy books, games and welcome folder",
    ]),
    ("task_inspection", "Final inspection", [
        "Check every window and door locks properly",
        "Set heating or cooling to the welcome setting",
        "Check the access code or key box works",
        "Walk through and photograph each room",
        "Log any damage or missing items",
        "Confirm the property is guest-ready",
    ]),
]

MAINTENANCE_CATEGORIES = [
    "Plumbing", "Electrical", "Heating / cooling", "Appliance",
    "Furniture", "Internet / TV", "Locks & access", "Pest control",
    "Garden / outdoor", "Pool / hot tub", "Damp & mould",
    "Paint & decorating", "Flooring", "Glass & windows",
    "Guest damage", "Deep clean", "Safety equipment", "Other",
]

MAINTENANCE_STATUSES = [
    "Reported", "Quoted", "Scheduled", "In progress", "Complete",
    "Waiting on parts", "Waiting on owner", "Cancelled",
]

PRIORITY_LEVELS = ["Urgent", "High", "Normal", "Low", "When convenient"]

EXPENSE_CATEGORIES = [
    "Cleaning", "Laundry", "Utilities - electric", "Utilities - gas",
    "Utilities - water", "Internet & TV", "Maintenance", "Repairs",
    "Furniture", "Appliances", "Guest amenities", "Welcome basket",
    "Insurance", "Property taxes", "Mortgage interest", "Rent",
    "Service charge", "Supplies", "Management fees", "Co-host fees",
    "Marketing & photography", "Software & subscriptions",
    "Licences & permits", "Accounting & legal", "Travel & mileage",
    "Bank & card fees", "Guest refunds", "Miscellaneous",
]

#: categories that are capital rather than running costs - excluded from the
#: operating profit figure but still totalled for the tax summary
CAPITAL_CATEGORIES = ["Furniture", "Appliances"]

#: categories the tax summary treats as a financing or ownership cost
OWNERSHIP_CATEGORIES = ["Mortgage interest", "Rent", "Property taxes",
                        "Service charge", "Insurance"]

TAX_CATEGORIES = [
    "Operating expense", "Capital / asset", "Financing cost",
    "Owner drawing", "Not deductible", "Review with accountant",
]

PAYMENT_METHODS = [
    "Bank transfer", "Debit card", "Credit card", "Direct debit",
    "Cash", "PayPal", "Platform deducted", "Owner paid", "Other",
]

RECURRING_FREQUENCIES = [
    "One off", "Weekly", "Fortnightly", "Monthly", "Quarterly",
    "Twice a year", "Yearly", "Per booking",
]

#: how many times a year each frequency happens - used to turn any cost
#: into a monthly equivalent
FREQUENCY_PER_YEAR: Dict[str, float] = {
    "One off": 0.0,
    "Weekly": 52.0,
    "Fortnightly": 26.0,
    "Monthly": 12.0,
    "Quarterly": 4.0,
    "Twice a year": 2.0,
    "Yearly": 1.0,
    "Per booking": 0.0,
}

SUPPLY_CATEGORIES = [
    "Towels", "Bed linen", "Pillows & duvets", "Toilet paper",
    "Shampoo & conditioner", "Soap & hand wash", "Cleaning products",
    "Coffee & tea", "Kitchen supplies", "Bin bags", "Laundry detergent",
    "Welcome treats", "Batteries & bulbs", "First aid", "Paper & stationery",
    "Pool & hot tub", "Garden", "Other",
]

STOCK_STATUSES = ["In stock", "Low", "Order now", "Out of stock",
                  "Overstocked", "Not tracked"]

STAFF_ROLES = [
    "Cleaner", "Head cleaner", "Laundry service", "Handyman",
    "Plumber", "Electrician", "Gardener", "Pest control",
    "Pool technician", "Co-host", "Property manager", "Photographer",
    "Linen hire", "Locksmith", "Window cleaner", "Accountant", "Other",
]

PAY_BASIS = ["Per turnover", "Per hour", "Per job", "Monthly retainer",
             "Per property per month", "% of revenue", "Invoice each time"]

CONTRACT_STATUSES = ["Active", "Trial", "Occasional", "Paused",
                     "Ended", "Do not use"]

AVAILABILITY = ["Any day", "Weekdays only", "Weekends only",
                "Mornings", "Afternoons", "By arrangement", "Unavailable"]

REVIEW_STATUSES = ["Not requested", "Requested", "Received",
                   "No review left", "Removed", "Disputed"]

RATING_CATEGORIES = [
    ("rate_overall", "Overall"),
    ("rate_clean", "Cleanliness"),
    ("rate_accuracy", "Accuracy"),
    ("rate_checkin", "Check-in"),
    ("rate_comm", "Communication"),
    ("rate_location", "Location"),
    ("rate_value", "Value"),
]

GUEST_TYPES = ["Couple", "Family", "Solo", "Friends", "Business",
               "Group", "Contractor", "Long stay", "Unknown"]

TRIP_PURPOSES = ["Holiday", "Business", "Family visit", "Event",
                 "Relocation", "Medical", "Contract work", "Other"]

PAYOUT_STATUSES = ["Expected", "Received", "Part received",
                   "Late", "Disputed", "Written off"]

YES_NO = ["Yes", "No"]
YES_NO_NA = ["Yes", "No", "N/A"]

COUNTRIES = [
    "United States", "United Kingdom", "Canada", "Ireland", "Australia",
    "New Zealand", "Germany", "France", "Spain", "Portugal", "Italy",
    "Netherlands", "Belgium", "Sweden", "Norway", "Denmark", "Poland",
    "Mexico", "Brazil", "Argentina", "Japan", "South Korea", "Singapore",
    "India", "United Arab Emirates", "South Africa", "Other",
]

MONTH_NAMES = ["January", "February", "March", "April", "May", "June",
               "July", "August", "September", "October", "November",
               "December"]

WEEKDAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday",
                 "Saturday", "Sunday"]

#: how the length of a stay is bucketed on the performance tab
STAY_BANDS: List[Tuple[str, int, int]] = [
    ("1 night", 1, 1),
    ("2 nights", 2, 2),
    ("3-4 nights", 3, 4),
    ("5-6 nights", 5, 6),
    ("1 week", 7, 9),
    ("10+ nights", 10, 27),
    ("Monthly stay", 28, 400),
]

#: seasons, so a host can see peak against shoulder without a pivot table
SEASONS: List[Tuple[str, int, int]] = [
    ("Winter", 12, 2),
    ("Spring", 3, 5),
    ("Summer", 6, 8),
    ("Autumn", 9, 11),
]


# ---------------------------------------------------------------------------
# Colours - a coastal / host palette: deep navy, sea green, warm sand
# ---------------------------------------------------------------------------

class Palette:
    ink = "#17262F"            # near black navy - titles
    deep = "#113A4C"           # deep navy - money tabs
    primary = "#1C6E78"        # sea green headers
    primary_mid = "#3F97A0"    # section bands
    primary_soft = "#E6F3F4"   # light band / alternating rows
    accent = "#E2703A"         # terracotta accent
    accent_soft = "#FCEFE7"
    input_bg = "#FFFFFF"
    input_border = "#D2E6E8"
    calc_bg = "#F4FAFA"
    calc_border = "#E2F0F1"
    white = "#FFFFFF"
    grey_text = "#5A6B70"
    line = "#C4DEE1"
    good_bg = "#D8EDDD"
    good_text = "#23583A"
    warn_bg = "#FCEBD4"
    warn_text = "#8A5300"
    bad_bg = "#F9DAD6"
    bad_text = "#9B2D20"
    info_bg = "#E1E8F2"
    info_text = "#2F4576"
    note_bg = "#F6FBFB"


# ---------------------------------------------------------------------------
# Row capacity of every register
# ---------------------------------------------------------------------------

DEFAULT_ROWS = {
    "Dashboard": 0,
    "Getting Started": 0,
    "Settings & Lists": 0,
    "Property Database": 30,
    "Booking Tracker": 900,
    "Booking Calendar": 0,
    "Guest Database": 400,
    "Check-in & Check-out": 900,
    "Cleaning Schedule": 900,
    "Turnover Checklist": 0,
    "Platform Revenue & Fees": 900,
    "Expense Tracker": 700,
    "Supplies Inventory": 300,
    "Maintenance Log": 300,
    "Reviews Tracker": 600,
    "Staff & Vendors": 60,
    "Tax Summary": 0,
    "Occupancy & Performance": 400,
    "Monthly Reports": 0,
    "Annual Reports": 0,
}

#: the register the `--rows` switch scales everything else against
SCALE_ANCHOR = "Booking Tracker"

#: how many properties the calendar grid and the comparison tables cover
MAX_PROPERTIES = 12

#: days shown on the multi-property availability strip
CALENDAR_STRIP_DAYS = 14

#: rows on the printable turnover checklist's extra-notes block
TURNOVER_NOTE_LINES = 4


# ---------------------------------------------------------------------------
# Automatic ID prefixes
# ---------------------------------------------------------------------------

ID_PREFIXES: Dict[str, Tuple[str, int]] = {
    "property": ("PRP-", 2),
    "booking": ("BKG-", 4),
    "guest": ("GST-", 4),
    "arrival": ("ARR-", 4),
    "clean": ("CLN-", 4),
    "payout": ("PAY-", 4),
    "expense": ("EXP-", 4),
    "supply": ("SUP-", 3),
    "maintenance": ("MNT-", 3),
    "review": ("REV-", 4),
    "vendor": ("VND-", 3),
    "period": ("PER-", 3),
}


@dataclass
class HostDefaults:
    """Sample values written into the business profile on Settings."""

    name: str = "Harbour & Pine Stays"
    owner: str = "Nadia Whitlock"
    brand: str = "Harbour & Pine"
    email: str = "stay@harbourandpine.example"
    phone: str = "(503) 555 0188"
    address: str = "18 Fenwick Street, Portland, OR 97209"
    website: str = "harbourandpine.example"
    tax_id: str = "EIN 87-4410923"
    licence: str = "STR-PDX-2021-00714"
    tagline: str = "Four coastal and city stays, run properly"


# ---------------------------------------------------------------------------
# The configuration object every sheet module reads
# ---------------------------------------------------------------------------

@dataclass
class Config:
    """Everything the build needs to know, in one place."""

    # --------------------------------------------------------- identity ---
    author: str = AUTHOR
    version: str = VERSION
    password: Optional[str] = PASSWORD
    protect_sheets: bool = True
    include_sample_data: bool = True

    host: HostDefaults = field(default_factory=HostDefaults)

    # ------------------------------------------------------------ money ---
    currency_symbol: str = "$"
    currency_code: str = "USD"
    date_format: str = "dd mmm yyyy"
    time_format: str = "hh:mm"

    #: sales / occupancy / lodging tax on the accommodation total
    tax_rate: float = 0.0
    tax_label: str = "Lodging tax"
    #: does the platform collect and remit the lodging tax for you
    tax_collected_by_platform: str = "Yes"

    # ------------------------------------------------- portfolio shape ---
    property_count: int = 4
    default_nightly_rate: float = 185.0
    default_cleaning_fee: float = 95.0
    default_cleaning_cost: float = 65.0
    default_security_deposit: float = 300.0
    default_min_nights: int = 2
    default_check_in: str = "15:00"
    default_check_out: str = "11:00"
    default_max_guests: int = 4
    extra_guest_fee: float = 25.0
    extra_guest_after: int = 2
    pet_fee: float = 40.0

    #: nights a year the host expects a property to be lettable at all
    lettable_nights_year: int = 365
    #: nights the owner expects to block for their own use
    owner_nights_year: int = 14

    # ------------------------------------------------------- platforms ---
    default_platform: str = "Airbnb"
    airbnb_commission: float = 0.030
    vrbo_commission: float = 0.050
    booking_commission: float = 0.150
    direct_card_fee: float = 0.029
    payout_days: int = 1

    # -------------------------------------------------------- operating ---
    laundry_cost_turnover: float = 18.0
    supplies_cost_turnover: float = 12.0
    welcome_cost_turnover: float = 8.0
    cleaner_default_rate: float = 65.0
    management_fee_percent: float = 0.0

    # ---------------------------------------------------------- targets ---
    occupancy_target: float = 0.72
    adr_target: float = 195.0
    revpar_target: float = 140.0
    margin_target: float = 0.38
    rating_target: float = 4.80
    revenue_target_month: float = 14000.0
    revenue_target_year: float = 168000.0
    profit_target_month: float = 5200.0

    # ----------------------------------------------------------- alerts ---
    turnover_warning_hours: int = 6
    review_chase_days: int = 3
    payout_chase_days: int = 7
    maintenance_warning_days: int = 7
    restock_warning_weeks: float = 2.0
    upcoming_window_days: int = 7
    long_stay_nights: int = 28
    lead_time_warning_days: int = 2

    # ------------------------------------------------------------ build ---
    year: int = 0
    today_override: Optional[date] = None
    rows_override: Optional[int] = None
    output: str = ""

    # ----------------------------------------------------------- lookups ---
    def __post_init__(self) -> None:
        if not self.year:
            self.year = self.today.year

    # ------------------------------------------------------------- dates ---
    @property
    def today(self) -> date:
        return self.today_override or date.today()

    # ------------------------------------------------------------ formats ---
    @property
    def money_format(self) -> str:
        sign = self.currency_symbol
        return f'{sign}#,##0.00;[Red]-{sign}#,##0.00'

    @property
    def money_format_0(self) -> str:
        sign = self.currency_symbol
        return f'{sign}#,##0;[Red]-{sign}#,##0'

    @property
    def money_format_4(self) -> str:
        sign = self.currency_symbol
        return f'{sign}#,##0.0000'

    # ------------------------------------------------------------- labels ---
    @property
    def product_name(self) -> str:
        return PRODUCT_NAME

    @property
    def listing_title(self) -> str:
        return LISTING_TITLE

    @property
    def tax_name(self) -> str:
        return self.tax_label

    @property
    def nights_per_year(self) -> int:
        return self.lettable_nights_year

    @property
    def platform_table(self) -> List[Tuple[str, float, float, str, str]]:
        """Platform fee table with the CLI overrides applied."""
        table: List[Tuple[str, float, float, str, str]] = []
        overrides = {
            "Airbnb": self.airbnb_commission,
            "Vrbo": self.vrbo_commission,
            "Booking.com": self.booking_commission,
        }
        for name, commission, payment, counts, note in PLATFORMS:
            commission = overrides.get(name, commission)
            if name in ("Direct booking", "Returning guest"):
                payment = self.direct_card_fee
            table.append((name, commission, payment, counts, note))
        return table

    # -------------------------------------------------------------- rows ---
    def row_count(self, sheet: str) -> int:
        base = DEFAULT_ROWS.get(sheet, 0)
        if not base or not self.rows_override:
            return base
        anchor = DEFAULT_ROWS[SCALE_ANCHOR]
        scaled = int(round(base * self.rows_override / anchor))
        return max(12, scaled)

    def scale_rows(self, count: int) -> int:
        """Scale a hand written row count the same way a register scales."""
        if not self.rows_override:
            return count
        anchor = DEFAULT_ROWS[SCALE_ANCHOR]
        return max(6, int(round(count * self.rows_override / anchor)))

    @property
    def scaled(self) -> bool:
        return bool(self.rows_override)

    # ---------------------------------------------------------------- ids ---
    def rid(self, kind: str, row: int) -> str:
        """The ID a given record row will show, as plain text."""
        prefix, digits = ID_PREFIXES[kind]
        return f"{prefix}{row:0{digits}d}"

    def prefix(self, kind: str) -> str:
        return ID_PREFIXES[kind][0]

    def digits(self, kind: str) -> int:
        return ID_PREFIXES[kind][1]


# ---------------------------------------------------------------------------
# Drop down list names, so a typo is a build error and not a broken workbook
# ---------------------------------------------------------------------------

LIST_NAMES: List[str] = [
    "YesNoList", "YesNoNaList", "BookingStatusList",
    "PaymentStatusList", "PropertyTypeList", "PropertyStatusList",
    "OwnershipList", "CleaningStatusList", "MaintenanceCategoryList",
    "MaintenanceStatusList", "PriorityList", "ExpenseCategoryList",
    "TaxCategoryList", "PaymentMethodList", "FrequencyList",
    "SupplyCategoryList", "StockStatusList", "StaffRoleList",
    "PayBasisList", "ContractStatusList", "AvailabilityList",
    "ReviewStatusList", "GuestTypeList", "TripPurposeList",
    "PayoutStatusList", "CountryList", "MonthNameList", "SeasonList",
    "StayBandList", "RatingList", "PlatformList", "PlatformCommission",
    "PlatformPayment", "PlatformIsOTA", "FrequencyPerYear",
    "MonthKeyList", "MonthKeys", "PropertyList", "PropertyNameList",
    "BookingList", "GuestList", "GuestNameList", "VendorList",
    "VendorNameList", "SupplyList", "SupplyNameList", "CleanTaskList",
    "MaintenanceList",
]


# ---------------------------------------------------------------------------
# The ID prefix table as it appears on Settings
# ---------------------------------------------------------------------------

#: (label on Settings, default prefix, defined name, digits)
ID_PREFIX_FIELDS: List[Tuple[str, str, str, int]] = [
    ("Property", "PRP-", "PropertyPrefix", 2),
    ("Booking", "BKG-", "BookingPrefix", 4),
    ("Guest", "GST-", "GuestPrefix", 4),
    ("Arrival", "ARR-", "ArrivalPrefix", 4),
    ("Cleaning job", "CLN-", "CleanPrefix", 4),
    ("Payout", "PAY-", "PayoutPrefix", 4),
    ("Expense", "EXP-", "ExpensePrefix", 4),
    ("Supply line", "SUP-", "SupplyPrefix", 3),
    ("Maintenance job", "MNT-", "MaintenancePrefix", 3),
    ("Review", "REV-", "ReviewPrefix", 4),
    ("Vendor", "VND-", "VendorPrefix", 3),
    ("Performance row", "PER-", "PeriodPrefix", 3),
]


PRIVACY_NOTE = (
    "A word about guest data. This workbook asks for a name, a way to "
    "contact someone and what they asked for, because that is what it "
    "takes to run the business. It deliberately has nowhere to put a "
    "passport number, a card number or a scan of an ID document, and "
    "you should not add one. Keep the file somewhere only you can "
    "reach, only mail the guests whose consent column says Yes, and "
    "delete the rows you no longer need."
)

DISCLAIMER = (
    "What this workbook is, and is not. It is a management spreadsheet "
    "for running short-term lets: bookings, turnovers, costs and "
    "performance. It is not connected to Airbnb, Vrbo, Booking.com or "
    "any other platform, it is not affiliated with or endorsed by any "
    "of them, and it cannot read or write their calendars - every "
    "figure in it is one you typed or pasted. It does not take "
    "payments and stores no card details. It is not accounting, legal "
    "or tax software, and nothing in it is accounting, legal or tax "
    "advice: the Tax Summary tab groups your own figures under your "
    "own headings so your accountant has less work to do, but what is "
    "deductible, how rental income is classified, whether lodging tax "
    "applies and who must remit it all depend on where the property "
    "is, where you live and how you hold it. It is also not a "
    "licensing or compliance system - short-term let registration, "
    "safety certificates, insurance and local caps on letting nights "
    "are your responsibility. Check your own rules, and check them "
    "with someone qualified."
)
````

### `airbnb_rental_spreadsheet/layout.py`

*253 lines · 3 classes · 10 functions*

> Layout engine.
>
> The workbook is described declaratively: every register sheet (Clients,
> Properties, Jobs, Invoices, ...) is a :class:`SheetSpec` of :class:`Col`
> objects.  Formulas are written as templates that never contain hard coded
> column letters, so columns can be reordered or inserted without breaking a
> single formula:
>
>     "=IF({client}{r}="","",SUMIFS([[Payments.amount]],[[Payments.invoice]],{id}{r}))"
>
>     {r}            -> the current Excel row number (1 based)
>     {key}          -> column letter of the column `key` on the *same* sheet
>     [[Sheet.key]]  -> absolute range of that column's data rows, e.g.
>                       'Job Costing'!$K$5:$K$504
>     [[Sheet.key#r]]-> absolute single cell of that column on row r

| Defines | What it does |
|---------|--------------|
| `col_letter(index)` | 0 based column index -> Excel column letter |
| `quote_sheet(name)` | Quote a sheet name for use inside a formula when it needs it |
| `cell_ref(col, row, absolute)` |  |
| `class Col` | One column of a register sheet |
| `class SheetSpec` | A register sheet: title band, header row and N pre-built data rows |
| `class Registry` | Lookup table of every registered :class:`SheetSpec` |
| `max_if(value_range, *conditions)` | MAX of `value_range` where all conditions hold (no CSE required) |
| `min_if(value_range, *conditions, *, sentinel)` | MIN of `value_range` where all conditions hold (no CSE required) |
| `blank_if(test, result)` |  |
| `safe(expr, fallback)` |  |
| `lookup(value, table_range, index_range, fallback)` | INDEX/MATCH lookup (works left or right, unlike VLOOKUP) |
| `dynamic_name(sheet, letter, first, last, count_text)` | OFFSET based defined name that grows with the populated rows |
| `join_nonblank(parts)` | Build `a & b & c` style concatenation |

**Module constants:** `TITLE_ROW` (1), `SUBTITLE_ROW` (2), `NAV_ROW` (3), `HEADER_ROW` (4), `DATA_ROW` (5).

````python
"""
Layout engine.

The workbook is described declaratively: every register sheet (Clients,
Properties, Jobs, Invoices, ...) is a :class:`SheetSpec` of :class:`Col`
objects.  Formulas are written as templates that never contain hard coded
column letters, so columns can be reordered or inserted without breaking a
single formula:

    "=IF({client}{r}="","",SUMIFS([[Payments.amount]],[[Payments.invoice]],{id}{r}))"

    {r}            -> the current Excel row number (1 based)
    {key}          -> column letter of the column `key` on the *same* sheet
    [[Sheet.key]]  -> absolute range of that column's data rows, e.g.
                      'Job Costing'!$K$5:$K$504
    [[Sheet.key#r]]-> absolute single cell of that column on row r
"""

from __future__ import annotations

import re
import string
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Sequence

# Row 1 title, row 2 strapline, row 3 navigation, row 4 headers, row 5+ data.
TITLE_ROW = 1
SUBTITLE_ROW = 2
NAV_ROW = 3
HEADER_ROW = 4
DATA_ROW = 5


# ---------------------------------------------------------------------------
# A1 helpers
# ---------------------------------------------------------------------------

def col_letter(index: int) -> str:
    """0 based column index -> Excel column letter."""
    letters = ""
    index += 1
    while index:
        index, rem = divmod(index - 1, 26)
        letters = string.ascii_uppercase[rem] + letters
    return letters


def quote_sheet(name: str) -> str:
    """Quote a sheet name for use inside a formula when it needs it."""
    if re.match(r"^[A-Za-z_][A-Za-z0-9_.]*$", name):
        return name
    return "'" + name.replace("'", "''") + "'"


def cell_ref(col: int, row: int, absolute: bool = True) -> str:
    dollar = "$" if absolute else ""
    return f"{dollar}{col_letter(col)}{dollar}{row}"


# ---------------------------------------------------------------------------
# Column / sheet specifications
# ---------------------------------------------------------------------------

@dataclass
class Col:
    """One column of a register sheet."""

    key: str
    header: str
    width: float = 14.0
    style: str = "text"              # key into the format book (see formats.py)
    formula: Optional[str] = None    # template -> locked, calculated column
    validate: Optional[Dict[str, Any]] = None
    note: Optional[str] = None       # help text shown when the cell is selected
    cf: List[Dict[str, Any]] = field(default_factory=list)
    hidden: bool = False
    group: str = ""                  # section label (documentation only)

    @property
    def is_formula(self) -> bool:
        return self.formula is not None


@dataclass
class SheetSpec:
    """A register sheet: title band, header row and N pre-built data rows."""

    name: str
    title: str
    subtitle: str = ""
    columns: List[Col] = field(default_factory=list)
    rows: int = 100
    header_row: int = HEADER_ROW        # 1 based Excel row of the header
    tab_color: Optional[str] = None
    freeze_cols: int = 2
    freeze_at: Optional[int] = None   # 1 based row to freeze below
    samples: List[Dict[str, Any]] = field(default_factory=list)
    help_text: str = ""
    extra_notes: List[str] = field(default_factory=list)

    # -------------------------------------------------------------- geometry
    @property
    def first_data_row(self) -> int:
        return self.header_row + 1

    @property
    def last_data_row(self) -> int:
        return self.header_row + self.rows

    def index(self, key: str) -> int:
        for i, col in enumerate(self.columns):
            if col.key == key:
                return i
        raise KeyError(f"{self.name}: unknown column '{key}'")

    def letter(self, key: str) -> str:
        return col_letter(self.index(key))

    def col(self, key: str) -> Col:
        return self.columns[self.index(key)]

    @property
    def last_col_index(self) -> int:
        return len(self.columns) - 1


class Registry:
    """Lookup table of every registered :class:`SheetSpec`."""

    def __init__(self) -> None:
        self._specs: Dict[str, SheetSpec] = {}

    def add(self, spec: SheetSpec) -> SheetSpec:
        self._specs[spec.name] = spec
        return spec

    def __contains__(self, name: str) -> bool:
        return name in self._specs

    def __getitem__(self, name: str) -> SheetSpec:
        return self._specs[name]

    def names(self) -> Sequence[str]:
        return list(self._specs)

    # ------------------------------------------------------------ references
    def rng(self, sheet: str, key: str, first: Optional[int] = None,
            last: Optional[int] = None) -> str:
        spec = self._specs[sheet]
        letter = spec.letter(key)
        first = first or spec.first_data_row
        last = last or spec.last_data_row
        return f"{quote_sheet(sheet)}!${letter}${first}:${letter}${last}"

    def cell(self, sheet: str, key: str, row: int) -> str:
        spec = self._specs[sheet]
        return f"{quote_sheet(sheet)}!${spec.letter(key)}${row}"

    def letter(self, sheet: str, key: str) -> str:
        return self._specs[sheet].letter(key)

    # -------------------------------------------------------------- renderer
    _REF = re.compile(r"\[\[([^\].#]+)\.([A-Za-z0-9_]+)(?:#(\d*))?\]\]")

    def render(self, template: str, row: Optional[int] = None,
               spec: Optional[SheetSpec] = None) -> str:
        """Expand [[Sheet.key]] and {key}/{r} tokens into a real formula."""

        def _ref(match: "re.Match[str]") -> str:
            sheet, key, single = match.group(1), match.group(2), match.group(3)
            if single:
                return self.cell(sheet, key, int(single))
            if single == "":
                # [[Sheet.key#]] -> same row as the formula being written
                if row is None:
                    raise ValueError(
                        f"[[{sheet}.{key}#]] needs a row context")
                return self.cell(sheet, key, row)
            return self.rng(sheet, key)

        text = self._REF.sub(_ref, template)

        if spec is not None:
            mapping: Dict[str, Any] = {c.key: spec.letter(c.key)
                                       for c in spec.columns}
            # `{key_first}` / `{key_last}` give the absolute top and bottom
            # of a column on the same sheet, so an in-sheet COUNTIFS reads
            # `{key_first}:{key_last}` instead of a hand built $D$5:$D$104.
            for col in spec.columns:
                letter = spec.letter(col.key)
                mapping[f"{col.key}_first"] = (
                    f"${letter}${spec.first_data_row}")
                mapping[f"{col.key}_last"] = (
                    f"${letter}${spec.last_data_row}")
            if row is not None:
                mapping["r"] = row
                mapping["r_prev"] = row - 1
            mapping["first"] = spec.first_data_row
            mapping["last"] = spec.last_data_row
            text = text.format(**mapping)
        elif row is not None:
            text = text.format(r=row)
        return text


# ---------------------------------------------------------------------------
# Reusable formula fragments
# ---------------------------------------------------------------------------

def max_if(value_range: str, *conditions: str) -> str:
    """MAX of `value_range` where all conditions hold (no CSE required)."""
    cond = "*".join(f"({c})" for c in conditions)
    return f"SUMPRODUCT(MAX({cond}*{value_range}))"


def min_if(value_range: str, *conditions: str, sentinel: int = 2958465) -> str:
    """MIN of `value_range` where all conditions hold (no CSE required).

    Non matching rows are swapped for a sentinel (31-Dec-9999 as a serial) so
    that MIN ignores them.  Returns the sentinel when nothing matches, so the
    caller should wrap the result in an IF test.
    """
    cond = "*".join(f"({c})" for c in conditions)
    return (f"SUMPRODUCT(MIN({cond}*{value_range}"
            f"+(1-{cond})*{sentinel}))")


def blank_if(test: str, result: str) -> str:
    return f'IF({test},"",{result})'


def safe(expr: str, fallback: str = '""') -> str:
    return f"IFERROR({expr},{fallback})"


def lookup(value: str, table_range: str, index_range: str,
           fallback: str = '""') -> str:
    """INDEX/MATCH lookup (works left or right, unlike VLOOKUP)."""
    return f"IFERROR(INDEX({index_range},MATCH({value},{table_range},0)),{fallback})"


def dynamic_name(sheet: str, letter: str, first: int, last: int,
                 count_text: bool = True) -> str:
    """OFFSET based defined name that grows with the populated rows."""
    rng = f"{quote_sheet(sheet)}!${letter}${first}:${letter}${last}"
    counter = f'COUNTIF({rng},"?*")' if count_text else f"COUNTA({rng})"
    return (f"=OFFSET({quote_sheet(sheet)}!${letter}${first},0,0,"
            f"MAX(1,{counter}),1)")


def join_nonblank(parts: Sequence[str]) -> str:
    """Build `a & b & c` style concatenation."""
    return "&".join(parts)
````

### `airbnb_rental_spreadsheet/formats.py`

*202 lines · 1 class*

> Cell format factory.
>
> Two things matter here:
>
> * **Input cells** are unlocked (``locked: False``) and painted cream, so the
>   user can type in them while the sheet is protected.
> * **Formula cells** are locked (``locked: True``) and painted cool grey.  With
>   the sheet protection password in place ("premium" by default) they cannot be
>   overwritten by accident.

| Defines | What it does |
|---------|--------------|
| `class FormatBook` | Creates (and caches) every :class:`xlsxwriter.format.Format` used |

**Module constants:** `P` (Palette), `INPUT` ('input'), `CALC` ('calc'), `PLAIN` ('plain').

````python
"""
Cell format factory.

Two things matter here:

* **Input cells** are unlocked (``locked: False``) and painted cream, so the
  user can type in them while the sheet is protected.
* **Formula cells** are locked (``locked: True``) and painted cool grey.  With
  the sheet protection password in place ("premium" by default) they cannot be
  overwritten by accident.
"""

from __future__ import annotations

from typing import Any, Dict

from .config import Config, Palette

P = Palette

INPUT = "input"
CALC = "calc"
PLAIN = "plain"


class FormatBook:
    """Creates (and caches) every :class:`xlsxwriter.format.Format` used."""

    def __init__(self, book, config: Config) -> None:
        self.book = book
        self.cfg = config
        self._cache: Dict[Any, Any] = {}

        money = config.money_format
        money0 = config.money_format_0
        money4 = config.money_format_4
        self._styles: Dict[str, Dict[str, Any]] = {
            "text": {"align": "left"},
            "id": {"align": "left", "bold": True, "font_color": P.primary},
            "wrap": {"align": "left", "valign": "top", "text_wrap": True},
            "date": {"num_format": config.date_format, "align": "center"},
            "time": {"num_format": config.time_format, "align": "center"},
            "money": {"num_format": money, "align": "right"},
            "money0": {"num_format": money0, "align": "right"},
            "money4": {"num_format": money4, "align": "right"},
            "number": {"num_format": "#,##0.00", "align": "right"},
            "number1": {"num_format": "#,##0.0", "align": "right"},
            "int": {"num_format": "#,##0", "align": "right"},
            "percent": {"num_format": "0.0%", "align": "center"},
            "percent0": {"num_format": "0%", "align": "center"},
            "rating": {"num_format": "0", "align": "center"},
            "flag": {"align": "center"},
            "status": {"align": "center", "bold": True},
            "link": {"align": "left", "font_color": P.info_text,
                     "underline": 1},
        }

        self._modes: Dict[str, Dict[str, Any]] = {
            INPUT: {"locked": False, "bg_color": P.input_bg,
                    "border": 1, "border_color": P.input_border},
            CALC: {"locked": True, "bg_color": P.calc_bg,
                   "border": 1, "border_color": P.calc_border,
                   "font_color": P.ink},
            PLAIN: {"locked": True},
        }

    # ------------------------------------------------------------------ core
    def _make(self, props: Dict[str, Any]):
        key = tuple(sorted((k, str(v)) for k, v in props.items()))
        if key not in self._cache:
            self._cache[key] = self.book.add_format(props)
        return self._cache[key]

    def cell(self, style: str = "text", mode: str = INPUT, **extra):
        """Data cell format: `style` = value type, `mode` = input/calc/plain."""
        props: Dict[str, Any] = {"font_name": "Calibri", "font_size": 10,
                                 "valign": "vcenter"}
        props.update(self._modes[mode])
        props.update(self._styles.get(style, {}))
        props.update(extra)
        return self._make(props)

    def input(self, style: str = "text", **extra):
        return self.cell(style, INPUT, **extra)

    def calc(self, style: str = "text", **extra):
        return self.cell(style, CALC, **extra)

    # ---------------------------------------------------------------- chrome
    def chrome(self, name: str, **extra):
        base: Dict[str, Any] = {"font_name": "Calibri", "font_size": 10,
                                "locked": True, "valign": "vcenter"}
        styles: Dict[str, Dict[str, Any]] = {
            "title": {"font_size": 18, "bold": True, "font_color": P.white,
                      "bg_color": P.primary, "align": "left", "indent": 1},
            "title_right": {"font_size": 10, "bold": True,
                            "font_color": P.white, "bg_color": P.primary,
                            "align": "right", "indent": 1},
            "subtitle": {"font_size": 10, "italic": True,
                         "font_color": P.grey_text, "bg_color": P.note_bg,
                         "align": "left", "indent": 1},
            "nav": {"font_size": 9, "bold": True, "font_color": P.info_text,
                    "bg_color": P.note_bg, "underline": 1, "align": "left",
                    "indent": 1},
            "nav_plain": {"font_size": 9, "font_color": P.grey_text,
                          "bg_color": P.note_bg, "align": "left"},
            "section": {"font_size": 11, "bold": True, "font_color": P.white,
                        "bg_color": P.primary_mid, "align": "left",
                        "indent": 1},
            "subsection": {"font_size": 10, "bold": True,
                           "font_color": P.primary, "bg_color": P.primary_soft,
                           "align": "left", "indent": 1},
            "header": {"font_size": 9, "bold": True, "font_color": P.white,
                       "bg_color": P.primary, "align": "center",
                       "valign": "vcenter", "text_wrap": True, "border": 1,
                       "border_color": P.primary},
            "header_left": {"font_size": 9, "bold": True,
                            "font_color": P.white, "bg_color": P.primary,
                            "align": "left", "valign": "vcenter",
                            "text_wrap": True, "border": 1,
                            "border_color": P.primary, "indent": 1},
            "label": {"font_size": 10, "align": "left", "indent": 1,
                      "font_color": P.ink},
            "label_bold": {"font_size": 10, "bold": True, "align": "left",
                           "indent": 1, "font_color": P.ink},
            "label_right": {"font_size": 10, "align": "right",
                            "font_color": P.ink},
            "note": {"font_size": 9, "italic": True,
                     "font_color": P.grey_text, "align": "left",
                     "text_wrap": True, "valign": "top"},
            "note_box": {"font_size": 9, "font_color": P.ink,
                         "bg_color": P.accent_soft, "align": "left",
                         "valign": "top", "text_wrap": True, "border": 1,
                         "border_color": P.accent, "indent": 1},
            "kpi_label": {"font_size": 9, "bold": True,
                          "font_color": P.white, "bg_color": P.primary_mid,
                          "align": "left", "indent": 1},
            "kpi_value": {"font_size": 16, "bold": True,
                          "font_color": P.ink, "bg_color": P.white,
                          "align": "center", "valign": "vcenter",
                          "border": 1, "border_color": P.line},
            "kpi_note": {"font_size": 8, "italic": True,
                         "font_color": P.grey_text, "bg_color": P.white,
                         "align": "center", "border": 1,
                         "border_color": P.line},
            "total_label": {"font_size": 10, "bold": True,
                            "font_color": P.white, "bg_color": P.primary,
                            "align": "left", "indent": 1},
            "total_value": {"font_size": 10, "bold": True,
                            "font_color": P.white, "bg_color": P.primary,
                            "align": "right"},
            "form_label": {"font_size": 10, "bold": True, "align": "left",
                           "font_color": P.primary, "bottom": 1,
                           "bottom_color": P.line, "indent": 1},
            "form_line": {"font_size": 10, "align": "left", "bottom": 1,
                          "bottom_color": P.ink, "locked": False},
            "form_box": {"font_size": 10, "align": "left", "valign": "top",
                         "text_wrap": True, "border": 1,
                         "border_color": P.ink, "locked": False},
            "print_title": {"font_size": 16, "bold": True,
                            "font_color": P.primary, "align": "left"},
            "print_sub": {"font_size": 10, "italic": True,
                          "font_color": P.grey_text, "align": "left"},
            "legend_input": {"font_size": 9, "bg_color": P.input_bg,
                             "border": 1, "border_color": P.input_border,
                             "align": "center"},
            "legend_calc": {"font_size": 9, "bg_color": P.calc_bg,
                            "border": 1, "border_color": P.calc_border,
                            "align": "center"},
            "band": {"bg_color": P.primary_soft},
            "rule": {"bottom": 2, "bottom_color": P.primary},
            "legend_label": {"font_size": 9, "bold": True,
                             "font_color": P.grey_text, "align": "right"},
            "day_number": {"font_size": 11, "bold": True,
                           "font_color": P.primary,
                           "bg_color": P.primary_soft, "align": "left",
                           "indent": 1, "num_format": "d ddd",
                           "border": 1, "border_color": P.line},
            "day_item": {"font_size": 8, "font_color": P.ink,
                         "align": "left", "indent": 1, "border": 1,
                         "border_color": P.line, "bg_color": P.white},
        }
        props = dict(base)
        props.update(styles.get(name, {}))
        props.update(extra)
        return self._make(props)

    # -------------------------------------------------- conditional formats
    def cf(self, kind: str):
        """Formats used by conditional formatting rules."""
        palettes = {
            "good": {"bg_color": P.good_bg, "font_color": P.good_text},
            "warn": {"bg_color": P.warn_bg, "font_color": P.warn_text},
            "bad": {"bg_color": P.bad_bg, "font_color": P.bad_text},
            "info": {"bg_color": P.info_bg, "font_color": P.info_text},
            "muted": {"font_color": P.grey_text, "italic": True},
            "bold_bad": {"bg_color": P.bad_bg, "font_color": P.bad_text,
                         "bold": True},
            "bold_good": {"bg_color": P.good_bg, "font_color": P.good_text,
                          "bold": True},
        }
        return self._make(palettes[kind])
````

### `airbnb_rental_spreadsheet/writer.py`

*473 lines · 1 class · 8 functions*

> Sheet writing helpers shared by every tab of the workbook.
>
> :class:`SheetWriter` wraps an XlsxWriter worksheet and adds the house style:
> title band, navigation strip, section headers, KPI tiles, label/value settings
> rows and - most importantly - :meth:`SheetWriter.grid`, which turns a
> :class:`~daycare_business_spreadsheet.layout.SheetSpec` into a finished register
> with headers, drop downs, conditional formats, sample rows, frozen panes and
> locked formula cells.

| Defines | What it does |
|---------|--------------|
| `class SheetWriter` | Convenience wrapper around a worksheet |
| `yes_no(validate_note)` |  |
| `list_from(name)` | Drop down fed by a defined name (see lists.py) |
| `date_check()` |  |
| `time_check()` | Any time of day |
| `number_check(minimum, maximum, decimal)` |  |
| `cf_text(text, fmt, criteria)` |  |
| `cf_formula(formula, fmt)` |  |
| `cf_bar(color)` |  |

**Module constants:** `P` (Palette), `NAV_LINKS` (10 items), `PROTECT_OPTIONS` (15 entries).

````python
"""
Sheet writing helpers shared by every tab of the workbook.

:class:`SheetWriter` wraps an XlsxWriter worksheet and adds the house style:
title band, navigation strip, section headers, KPI tiles, label/value settings
rows and - most importantly - :meth:`SheetWriter.grid`, which turns a
:class:`~daycare_business_spreadsheet.layout.SheetSpec` into a finished register
with headers, drop downs, conditional formats, sample rows, frozen panes and
locked formula cells.
"""

from __future__ import annotations

import datetime as _dt
from typing import Any, Dict, List, Optional, Sequence, Tuple

from .config import Config, Palette
from .formats import CALC, FormatBook, INPUT, PLAIN
from .layout import (NAV_ROW, SUBTITLE_ROW, TITLE_ROW, Registry,
                     SheetSpec, col_letter, quote_sheet)

P = Palette

#: Sheets linked from the navigation strip at the top of every tab.
NAV_LINKS: List[Tuple[str, str]] = [
    ("Dashboard", "Dashboard"),
    ("Start", "Getting Started"),
    ("Properties", "Property Database"),
    ("Bookings", "Booking Tracker"),
    ("Calendar", "Booking Calendar"),
    ("Guests", "Guest Database"),
    ("Cleaning", "Cleaning Schedule"),
    ("Expenses", "Expense Tracker"),
    ("Reports", "Monthly Reports"),
    ("Settings", "Settings & Lists"),
]

#: Worksheet protection: formulas are locked, everything else stays usable.
PROTECT_OPTIONS: Dict[str, bool] = {
    "objects": False,
    "scenarios": False,
    "format_cells": True,
    "format_columns": True,
    "format_rows": True,
    "insert_columns": False,
    "insert_rows": True,
    "insert_hyperlinks": True,
    "delete_columns": False,
    "delete_rows": True,
    "select_locked_cells": True,
    "sort": False,          # stops rows being re-sorted out of sync with IDs
    "autofilter": True,     # filtering stays available
    "pivot_tables": False,
    "select_unlocked_cells": True,
}


class SheetWriter:
    """Convenience wrapper around a worksheet."""

    def __init__(self, book, fmts: FormatBook, cfg: Config, reg: Registry,
                 name: str, tab_color: Optional[str] = None,
                 landscape: bool = True) -> None:
        self.book = book
        self.f = fmts
        self.cfg = cfg
        self.reg = reg
        self.name = name
        self.ws = book.add_worksheet(name)
        self._protected = False
        if tab_color:
            self.ws.set_tab_color(tab_color)
        if landscape:
            self.ws.set_landscape()
        self.ws.set_paper(9)  # A4
        self.ws.set_margins(0.4, 0.4, 0.6, 0.5)
        self.ws.set_header(
            f"&L&\"Calibri,Bold\"&12{''}"
            f"&C&\"Calibri,Bold\"&11{name}"
            f"&R&\"Calibri,Italic\"&9{cfg.product_name}")
        self.ws.set_footer(
            f"&L&\"Calibri,Italic\"&8{cfg.author} - v{cfg.version}"
            "&C&\"Calibri,Italic\"&8Page &P of &N"
            "&R&\"Calibri,Italic\"&8&D")

    # ------------------------------------------------------------- plumbing
    def write(self, row: int, col: int, value: Any, fmt=None):
        """Type aware write (0 based row/col, like XlsxWriter)."""
        if value is None:
            if fmt is not None:
                self.ws.write_blank(row, col, None, fmt)
            return
        if isinstance(value, str) and value.startswith("="):
            self.ws.write_formula(row, col, value.strip(), fmt)
        elif isinstance(value, (_dt.datetime, _dt.date, _dt.time)):
            self.ws.write_datetime(row, col, value, fmt)
        elif isinstance(value, bool):
            self.ws.write_string(row, col, "Yes" if value else "No", fmt)
        elif isinstance(value, (int, float)):
            self.ws.write_number(row, col, value, fmt)
        else:
            self.ws.write_string(row, col, str(value), fmt)

    def formula(self, row: int, col: int, text: str, fmt=None):
        """Write a formula, resolving any ``[[Sheet.key]]`` tokens.

        The rendered text is stripped: Excel tolerates a stray space
        on the end of a formula but several calculation engines -
        including the one the recalculation check uses - refuse to
        parse it, and a space is the easiest thing in the world to
        leave behind when a formula is concatenated over ten lines of
        source.
        """
        self.ws.write_formula(row, col, self.reg.render(text).strip(), fmt)

    def merge(self, row: int, first_col: int, last_col: int, value: Any,
              fmt=None, row2: Optional[int] = None):
        row2 = row if row2 is None else row2
        if first_col == last_col and row == row2:
            self.write(row, first_col, value, fmt)
        else:
            text = "" if value is None else value
            if isinstance(text, (_dt.datetime, _dt.date, _dt.time)):
                self.ws.merge_range(row, first_col, row2, last_col, "", fmt)
                self.ws.write_datetime(row, first_col, text, fmt)
            else:
                self.ws.merge_range(row, first_col, row2, last_col, text, fmt)

    def set_col(self, first: int, last: int, width: float, fmt=None,
                hidden: bool = False):
        options = {"hidden": 1} if hidden else {}
        self.ws.set_column(first, last, width, fmt, options)

    def row_height(self, row: int, height: float):
        self.ws.set_row(row, height)

    # ----------------------------------------------------------- page chrome
    def banner(self, title: str, subtitle: str = "", last_col: int = 9,
               nav: bool = True, badge: str = "",
               tail_text: str = "") -> None:
        """Title row + strapline + navigation strip (Excel rows 1-3).

        `tail_text` is written in the empty space to the right of the
        navigation links - used for the register help line on sheets whose
        header row leaves no spare row for it.
        """
        r = TITLE_ROW - 1
        self.row_height(r, 34)
        split = max(last_col - 3, 1)
        self.merge(r, 0, split, title, self.f.chrome("title"))
        self.merge(r, split + 1, last_col,
                   badge or f"{self.cfg.author}  |  v{self.cfg.version}",
                   self.f.chrome("title_right"))

        r = SUBTITLE_ROW - 1
        self.row_height(r, 18)
        self.merge(r, 0, last_col, subtitle, self.f.chrome("subtitle"))

        if nav:
            self.nav_strip(NAV_ROW - 1, last_col, tail_text=tail_text)

    def nav_strip(self, row: int, last_col: int,
                  links: Optional[Sequence[Tuple[str, str]]] = None,
                  tail_text: str = "") -> int:
        """Row of internal links. Returns the first column left over."""
        links = list(links or NAV_LINKS)
        self.row_height(row, 16)
        col = 0
        fmt = self.f.chrome("nav")
        for label, sheet in links:
            if col + 1 > last_col:
                break
            self.ws.merge_range(row, col, row, min(col + 1, last_col), "", fmt)
            self.ws.write_url(row, col, f"internal:{quote_sheet(sheet)}!A1",
                              fmt, f"> {label}")
            col += 2
        if col <= last_col:
            room = last_col - col
            if tail_text and room >= 4:
                self.merge(row, col, last_col, tail_text,
                           self.f.chrome("note", bg_color=P.note_bg))
            else:
                self.merge(row, col, last_col, "",
                           self.f.chrome("nav_plain"))
        return col

    def section(self, row: int, text: str, last_col: int, first_col: int = 0):
        self.row_height(row, 20)
        self.merge(row, first_col, last_col, text, self.f.chrome("section"))
        return row + 1

    def subsection(self, row: int, text: str, last_col: int,
                   first_col: int = 0):
        self.row_height(row, 17)
        self.merge(row, first_col, last_col, text,
                   self.f.chrome("subsection"))
        return row + 1

    def note(self, row: int, text: str, last_col: int, first_col: int = 0,
             height: float = 28, boxed: bool = True):
        self.row_height(row, height)
        self.merge(row, first_col, last_col, text,
                   self.f.chrome("note_box" if boxed else "note"))
        return row + 1

    def bullets(self, row: int, lines: Sequence[str], last_col: int,
                first_col: int = 0, height: float = 15):
        for line in lines:
            self.row_height(row, height)
            self.merge(row, first_col, last_col, line, self.f.chrome("label"))
            row += 1
        return row

    # -------------------------------------------------------- settings rows
    def field(self, row: int, label: str, value: Any = None, *,
              label_col: int = 1, value_col: int = 3,
              value_end: Optional[int] = None, style: str = "text",
              mode: str = INPUT, name: Optional[str] = None,
              validate: Optional[Dict[str, Any]] = None,
              help_text: str = "", help_col: Optional[int] = None,
              formula: Optional[str] = None):
        """A `Label | value | help` settings row, optionally named."""
        self.merge(row, label_col, value_col - 1, label,
                   self.f.chrome("label_bold"))
        fmt = self.f.cell(style, mode)
        end = value_end if value_end is not None else value_col
        if formula is not None:
            rendered = self.reg.render(formula)
            if end > value_col:
                self.ws.merge_range(row, value_col, row, end, "", fmt)
            self.ws.write_formula(row, value_col, rendered, fmt)
        else:
            if end > value_col:
                self.merge(row, value_col, end, None, fmt)
                self.write(row, value_col, value, fmt)
            else:
                self.write(row, value_col, value, fmt)
        if validate:
            self.ws.data_validation(row, value_col, row, value_col, validate)
        if help_text:
            hc = help_col if help_col is not None else end + 1
            self.merge(row, hc, hc + 3, help_text, self.f.chrome("note"))
        if name:
            self.book.define_name(
                name, f"={quote_sheet(self.name)}!"
                      f"${col_letter(value_col)}${row + 1}")
        return row + 1

    # ------------------------------------------------------------ KPI tiles
    def kpi(self, row: int, col: int, label: str, formula: str,
            style: str = "int", span: int = 2, note: str = ""):
        """Three row KPI tile: caption, big number, small note."""
        self.merge(row, col, col + span - 1, label, self.f.chrome("kpi_label"))
        self.row_height(row + 1, 26)
        fmt = self.f.chrome("kpi_value",
                            num_format=self._num_format(style))
        if span > 1:
            self.ws.merge_range(row + 1, col, row + 1, col + span - 1,
                                "", fmt)
        self.ws.write_formula(row + 1, col, self.reg.render(formula), fmt)
        self.merge(row + 2, col, col + span - 1, note,
                   self.f.chrome("kpi_note"))
        return row + 3

    def num_format(self, style: str) -> str:
        """Number format string for a style key (public helper)."""
        return self._num_format(style)

    def _num_format(self, style: str) -> str:
        return {
            "money": self.cfg.money_format,
            "money0": self.cfg.money_format_0,
            "money4": self.cfg.money_format_4,
            "int": "#,##0",
            "number": "#,##0.00",
            "number1": "#,##0.0",
            "percent": "0.0%",
            "percent0": "0%",
            "date": self.cfg.date_format,
            "text": "General",
        }.get(style, "General")

    # ------------------------------------------------------- generic tables
    def table_head(self, row: int, col: int, headers: Sequence[str],
                   widths: Optional[Sequence[float]] = None,
                   align_left_first: bool = True):
        for i, head in enumerate(headers):
            fmt = self.f.chrome(
                "header_left" if (i == 0 and align_left_first) else "header")
            self.write(row, col + i, head, fmt)
            if widths:
                self.set_col(col + i, col + i, widths[i])
        self.row_height(row, 26)
        return row + 1

    def calc_row(self, row: int, col: int, values: Sequence[Any],
                 styles: Sequence[str], mode: str = CALC,
                 first_style: str = "text"):
        for i, value in enumerate(values):
            style = styles[i] if i < len(styles) else "text"
            fmt = self.f.cell(style, PLAIN if i == 0 and mode == PLAIN
                              else mode)
            if isinstance(value, str) and value.startswith("="):
                self.ws.write_formula(row, col + i, self.reg.render(value),
                                      fmt)
            else:
                self.write(row, col + i, value, fmt)
        return row + 1

    # ----------------------------------------------------------- the grid --
    def grid(self, spec: SheetSpec, extra_help: str = "") -> None:
        """Render a whole register sheet from its specification."""
        ws = self.ws
        last_col = spec.last_col_index
        header_row0 = spec.header_row - 1
        first = spec.first_data_row
        last = spec.last_data_row

        help_text = extra_help or spec.help_text
        spare_row = spec.header_row - 1 > NAV_ROW   # room above the headers?
        self.banner(spec.title, spec.subtitle, last_col,
                    tail_text="" if spare_row else help_text)

        if help_text and spare_row:
            self.note(header_row0 - 1, help_text, last_col, height=26)

        # Column widths, default formats (so blank cells keep the right lock
        # state) and the header row itself.
        for idx, col in enumerate(spec.columns):
            mode = CALC if col.is_formula else INPUT
            self.set_col(idx, idx, col.width, self.f.cell(col.style, mode),
                         hidden=col.hidden)
            self.write(header_row0, idx, col.header,
                       self.f.chrome("header_left" if idx <= 1 else "header"))
        self.row_height(header_row0, 30)

        # Data rows: formulas everywhere, sample values in the input columns.
        samples = spec.samples if self.cfg.include_sample_data else []
        for offset in range(spec.rows):
            excel_row = first + offset
            row0 = excel_row - 1
            sample = samples[offset] if offset < len(samples) else {}
            for idx, col in enumerate(spec.columns):
                if col.is_formula:
                    fmt = self.f.cell(col.style, CALC)
                    ws.write_formula(
                        row0, idx,
                        self.reg.render(col.formula, excel_row, spec), fmt)
                elif col.key in sample and sample[col.key] is not None:
                    self.write(row0, idx, sample[col.key],
                               self.f.cell(col.style, INPUT))

        # Drop downs, help bubbles and conditional formats.
        for idx, col in enumerate(spec.columns):
            if col.validate:
                options = dict(col.validate)
                if "source" in options and isinstance(options["source"], str):
                    options["source"] = self.reg.render(options["source"])
                if col.note:
                    options.setdefault("input_title", col.header[:32])
                    options.setdefault("input_message", col.note[:255])
                ws.data_validation(first - 1, idx, last - 1, idx, options)
            if col.note:
                # Header comment: works in every Excel version and survives
                # the trip through Google Sheets.
                ws.write_comment(header_row0, idx, col.note, {
                    "author": self.cfg.author,
                    "width": 220,
                    "height": 100,
                    "x_scale": 1.1,
                })
            for rule in col.cf:
                options = dict(rule)
                fmt_key = options.pop("fmt", None)
                if fmt_key:
                    options["format"] = self.f.cf(fmt_key)
                for field_name in ("criteria", "value", "minimum", "maximum"):
                    val = options.get(field_name)
                    if isinstance(val, str) and ("[[" in val or "{" in val):
                        options[field_name] = self.reg.render(
                            val, first, spec)
                ws.conditional_format(first - 1, idx, last - 1, idx, options)

        ws.autofilter(header_row0, 0, last - 1, last_col)
        freeze_row = (spec.freeze_at if spec.freeze_at is not None
                      else spec.header_row)
        ws.freeze_panes(freeze_row, spec.freeze_cols)
        ws.print_area(0, 0, last - 1, last_col)
        ws.repeat_rows(0, header_row0 if spec.header_row <= 6 else 2)
        ws.fit_to_pages(1, 0)

    # ------------------------------------------------------------- finishing
    def finish(self, freeze_row: Optional[int] = None,
               freeze_col: int = 0, print_area: Optional[str] = None,
               fit_width: bool = True, gridlines: bool = True):
        if freeze_row is not None:
            self.ws.freeze_panes(freeze_row, freeze_col)
        if print_area:
            self.ws.print_area(print_area)
        if fit_width:
            self.ws.fit_to_pages(1, 0)
        if not gridlines:
            self.ws.hide_gridlines(2)
        self.protect()

    def protect(self):
        if self.cfg.protect_sheets and not self._protected:
            self.ws.protect(self.cfg.password, dict(PROTECT_OPTIONS))
            self._protected = True


# ---------------------------------------------------------------------------
# Small shared helpers
# ---------------------------------------------------------------------------

def yes_no(validate_note: str = "") -> Dict[str, Any]:
    opts: Dict[str, Any] = {"validate": "list", "source": ["Yes", "No"]}
    if validate_note:
        opts["input_message"] = validate_note
    return opts


def list_from(name: str) -> Dict[str, Any]:
    """Drop down fed by a defined name (see lists.py)."""
    return {"validate": "list", "source": f"={name}"}


def date_check() -> Dict[str, Any]:
    return {
        "validate": "date",
        "criteria": "between",
        "minimum": _dt.date(2000, 1, 1),
        "maximum": _dt.date(2100, 12, 31),
        "error_title": "Check the date",
        "error_message": "Enter a real date between 2000 and 2100.",
    }


def time_check() -> Dict[str, Any]:
    """Any time of day. Typing 9:00 AM or 14:30 both work."""
    return {
        "validate": "time",
        "criteria": "between",
        "minimum": _dt.time(0, 0),
        "maximum": _dt.time(23, 59),
        "error_title": "Check the time",
        "error_message": "Enter a time of day, for example 9:00 AM "
                         "or 14:30.",
    }


def number_check(minimum: float = 0, maximum: float = 1000000,
                 decimal: bool = True) -> Dict[str, Any]:
    return {
        "validate": "decimal" if decimal else "integer",
        "criteria": "between",
        "minimum": minimum,
        "maximum": maximum,
        "error_title": "Check the value",
        "error_message": f"Enter a number between {minimum} and {maximum}.",
    }


def cf_text(text: str, fmt: str, criteria: str = "containing") -> Dict[str, Any]:
    return {"type": "text", "criteria": criteria, "value": text, "fmt": fmt}


def cf_formula(formula: str, fmt: str) -> Dict[str, Any]:
    return {"type": "formula", "criteria": formula, "fmt": fmt}


def cf_bar(color: str = "#2E7D8F") -> Dict[str, Any]:
    return {"type": "data_bar", "bar_color": color}
````

### `airbnb_rental_spreadsheet/names.py`

*81 lines · 1 function*

> Workbook level defined names that point at the register sheets.
>
> The drop-down lists on the `Settings & Lists` tab define their own
> names.  These are the ones that have to follow live data - property IDs,
> booking references, guests, cleaners and contractors, the linen cupboard
> - so they are built from the registry once every register exists.  Each
> one is an ``OFFSET`` range that grows and shrinks with the number of
> filled rows, which is why a property you add this morning turns up in
> the booking drop-down this afternoon.

| Defines | What it does |
|---------|--------------|
| `define_cross_sheet_names(book, reg)` | Create the live lookup names |

**Module constants:** `CROSS_SHEET_NAMES` (45 entries).

````python
"""Workbook level defined names that point at the register sheets.

The drop-down lists on the `Settings & Lists` tab define their own
names.  These are the ones that have to follow live data - property IDs,
booking references, guests, cleaners and contractors, the linen cupboard
- so they are built from the registry once every register exists.  Each
one is an ``OFFSET`` range that grows and shrinks with the number of
filled rows, which is why a property you add this morning turns up in
the booking drop-down this afternoon.
"""

from __future__ import annotations

from typing import Dict, Tuple

from .layout import Registry, dynamic_name

#: defined name -> (sheet, column key)
CROSS_SHEET_NAMES: Dict[str, Tuple[str, str]] = {
    # the spine of the workbook - every drop-down that must follow data
    "PropertyList": ("Property Database", "id"),
    "PropertyNameList": ("Property Database", "name"),
    "BookingList": ("Booking Tracker", "id"),
    "GuestList": ("Guest Database", "id"),
    "GuestNameList": ("Guest Database", "name"),
    "VendorList": ("Staff & Vendors", "id"),
    "VendorNameList": ("Staff & Vendors", "name"),
    "SupplyList": ("Supplies Inventory", "id"),
    "SupplyNameList": ("Supplies Inventory", "name"),
    "CleanTaskList": ("Cleaning Schedule", "id"),
    "MaintenanceList": ("Maintenance Log", "id"),
    # property columns the dashboard and the comparison tables scan
    "PropertyStatusCol": ("Property Database", "status"),
    "PropertyNightlyCol": ("Property Database", "nightly_rate"),
    "PropertyRankCol": ("Property Database", "profit_rank"),
    # booking columns - the busiest ranges in the workbook
    "BookingProperty": ("Booking Tracker", "property"),
    "BookingPlatform": ("Booking Tracker", "platform"),
    "BookingStatusCol": ("Booking Tracker", "status"),
    "BookingCheckIn": ("Booking Tracker", "check_in"),
    "BookingCheckOut": ("Booking Tracker", "check_out"),
    "BookingCheckInNum": ("Booking Tracker", "ci_num"),
    "BookingCheckOutNum": ("Booking Tracker", "co_num"),
    "BookingNights": ("Booking Tracker", "nights"),
    "BookingMonths": ("Booking Tracker", "month"),
    "BookingAccommodation": ("Booking Tracker", "accommodation"),
    "BookingGross": ("Booking Tracker", "gross"),
    "BookingNet": ("Booking Tracker", "net_revenue"),
    "BookingHolds": ("Booking Tracker", "holds"),
    "BookingEarns": ("Booking Tracker", "earns"),
    "BookingRowNum": ("Booking Tracker", "seq_num"),
    # the money columns the reports and the tax summary add up
    "ExpenseMonths": ("Expense Tracker", "month"),
    "ExpenseCategoryCol": ("Expense Tracker", "category"),
    "ExpenseAmounts": ("Expense Tracker", "amount"),
    "ExpensePL": ("Expense Tracker", "pl_amount"),
    "ExpenseProperty": ("Expense Tracker", "property"),
    "CleanMonths": ("Cleaning Schedule", "month"),
    "CleanCost": ("Cleaning Schedule", "total_cost"),
    "CleanProperty": ("Cleaning Schedule", "property"),
    "MaintMonths": ("Maintenance Log", "month"),
    "MaintCost": ("Maintenance Log", "cost_used"),
    "MaintProperty": ("Maintenance Log", "property"),
    "PayoutMonths": ("Platform Revenue & Fees", "month"),
    "PayoutReceived": ("Platform Revenue & Fees", "received"),
    "ReviewMonths": ("Reviews Tracker", "month"),
    "ReviewOverall": ("Reviews Tracker", "rate_overall"),
    "ReviewProperty": ("Reviews Tracker", "property"),
}


def define_cross_sheet_names(book, reg: Registry) -> None:
    """Create the live lookup names. Call after every register is built."""
    for name, (sheet, key) in CROSS_SHEET_NAMES.items():
        if sheet not in reg:
            continue
        spec = reg[sheet]
        book.define_name(
            name,
            dynamic_name(sheet, spec.letter(key), spec.first_data_row,
                         spec.last_data_row))
````

### `airbnb_rental_spreadsheet/build.py`

*160 lines · 3 functions*

> Workbook builder: puts the twenty tabs together in order.
>
> Typical use::
>
>     from airbnb_rental_spreadsheet import build_workbook
>     build_workbook("Airbnb Rental Management Spreadsheet.xlsx")

| Defines | What it does |
|---------|--------------|
| `build_register(book, fmts, cfg, reg, module)` | Create one register sheet from its :class:`SheetSpec` |
| `build_workbook(path, cfg)` | Write the whole product to *path* and return the path |
| `_set_properties(book, cfg)` |  |

**Module constants:** `REGISTER_MODULES` (12 items), `SHEET_ORDER` (20 items), `PAGE_BUILDERS` (8 entries), `REGISTER_BY_NAME` ({module.NAME: module for module in REGISTER_MODULES}).

````python
"""Workbook builder: puts the twenty tabs together in order.

Typical use::

    from airbnb_rental_spreadsheet import build_workbook
    build_workbook("Airbnb Rental Management Spreadsheet.xlsx")
"""

from __future__ import annotations

import datetime as _dt
import inspect
import os
from typing import Callable, Dict, List, Optional

from xlsxwriter.workbook import Workbook

from .config import Config
from .formats import FormatBook
from .layout import Registry, SheetSpec
from .names import define_cross_sheet_names
from .sheets import (annual, arrivals, bookings, calendar, cleaning,
                     dashboard, expenses, getting_started, guests,
                     maintenance, monthly, payouts, performance,
                     properties, reviews, settings, supplies, tax,
                     turnover, vendors)
from .writer import SheetWriter

#: register modules, in the order their specs are registered
REGISTER_MODULES = [
    properties, bookings, guests, arrivals, cleaning, payouts,
    expenses, supplies, maintenance, reviews, vendors, performance,
]

#: tab order of the finished workbook
SHEET_ORDER: List[str] = [
    "Dashboard",
    "Getting Started",
    "Settings & Lists",
    "Property Database",
    "Booking Tracker",
    "Booking Calendar",
    "Guest Database",
    "Check-in & Check-out",
    "Cleaning Schedule",
    "Turnover Checklist",
    "Platform Revenue & Fees",
    "Expense Tracker",
    "Supplies Inventory",
    "Maintenance Log",
    "Reviews Tracker",
    "Staff & Vendors",
    "Tax Summary",
    "Occupancy & Performance",
    "Monthly Reports",
    "Annual Reports",
]

#: sheets that are not registers - name -> builder(book, fmts, cfg, reg)
PAGE_BUILDERS: Dict[str, Callable] = {
    "Dashboard": dashboard.build,
    "Getting Started": getting_started.build,
    "Settings & Lists": settings.build,
    "Booking Calendar": calendar.build,
    "Turnover Checklist": turnover.build,
    "Tax Summary": tax.build,
    "Monthly Reports": monthly.build,
    "Annual Reports": annual.build,
}

#: register modules by sheet name
REGISTER_BY_NAME = {module.NAME: module for module in REGISTER_MODULES}


# ---------------------------------------------------------------------------

def build_register(book, fmts: FormatBook, cfg: Config, reg: Registry,
                   module) -> SheetWriter:
    """Create one register sheet from its :class:`SheetSpec`."""
    spec: SheetSpec = reg[module.NAME]
    sheet = SheetWriter(book, fmts, cfg, reg, spec.name,
                        tab_color=spec.tab_color)
    sheet.grid(spec)
    decorate = getattr(module, "decorate", None)
    if decorate is not None:
        wants_registry = len(inspect.signature(decorate).parameters) >= 4
        if wants_registry:
            decorate(sheet, fmts, cfg, reg)
        else:
            decorate(sheet, fmts, cfg)
    sheet.protect()
    return sheet


def build_workbook(path: str = "Airbnb Rental Spreadsheet.xlsx",
                   cfg: Optional[Config] = None) -> str:
    """Write the whole product to *path* and return the path."""
    cfg = cfg or Config()
    folder = os.path.dirname(os.path.abspath(path))
    if folder and not os.path.isdir(folder):
        os.makedirs(folder)

    book = Workbook(path)
    _set_properties(book, cfg)

    fmts = FormatBook(book, cfg)
    reg = Registry()

    # 1. Register every spec first so cross sheet formulas resolve no
    #    matter which tab happens to be written first.
    for module in REGISTER_MODULES:
        reg.add(module.spec(cfg))

    # 2. Write the tabs in their final order.
    sheets: Dict[str, SheetWriter] = {}
    for name in SHEET_ORDER:
        if name in REGISTER_BY_NAME:
            sheets[name] = build_register(book, fmts, cfg, reg,
                                          REGISTER_BY_NAME[name])
        elif name in PAGE_BUILDERS:
            sheets[name] = PAGE_BUILDERS[name](book, fmts, cfg, reg)
        else:                                   # pragma: no cover
            raise KeyError(f"No builder for sheet '{name}'")

    # 3. Live lookup names (property IDs, booking IDs, guests, vendors).
    define_cross_sheet_names(book, reg)

    # 4. Final touches: open on the dashboard.
    first = sheets[SHEET_ORDER[0]].ws
    first.activate()
    first.set_first_sheet()

    book.close()
    return path


def _set_properties(book, cfg: Config) -> None:
    book.set_properties({
        "title": cfg.product_name,
        "subject": "Short-term rental bookings, cleaning turnovers, "
                   "platform fees, expenses, occupancy and profit",
        "author": cfg.author,
        "manager": cfg.author,
        "company": cfg.author,
        "category": "Small business spreadsheet template",
        "keywords": ("airbnb spreadsheet, short term rental tracker, "
                     "vacation rental management, booking calendar, "
                     "cleaning schedule, guest tracker, occupancy "
                     "calculator, ADR RevPAR, str investor "
                     "spreadsheet, vrbo tracker, property management "
                     "template, rental income expense tracker"),
        "comments": (f"{cfg.product_name}. Created by {cfg.author}. "
                     f"Version {cfg.version}. Formula cells are locked "
                     f"and every sheet is protected with the password "
                     f"'{cfg.password}'. Not affiliated with or "
                     f"connected to Airbnb, Vrbo or Booking.com."),
        "status": "Final",
        "created": _dt.datetime(cfg.today.year, cfg.today.month,
                                cfg.today.day),
    })
````

### `airbnb_rental_spreadsheet/sample_data.py`

*1,151 lines · 15 functions*

> A worked example: four properties, four months, every tab populated.
>
> The numbers are deliberately not flattering.  One property is a
> rent-to-rent that barely washes its face, one is a one-bedroom that
> takes short stays and gets eaten alive by cleaning costs, and there is
> a boiler failure in February that blocks a week.  A sample portfolio
> where everything works teaches nobody anything.
>
> Every row here is checked twice before it ships: the keys must all be
> *input* columns on their register, and every value that feeds a
> drop-down must exist in that drop-down's list.

| Defines | What it does |
|---------|--------------|
| `_day(cfg, offset)` | A date `offset` days from the build date |
| `_time(hour, minute)` |  |
| `properties(cfg)` |  |
| `vendors(cfg)` |  |
| `guests(cfg)` |  |
| `bookings(cfg)` |  |
| `arrivals(cfg)` | One row per recent and upcoming stay, in varying states |
| `cleaning(cfg)` | A turnover for every completed stay, plus the ones coming up |
| `payouts(cfg)` | Reconciled payouts, including two where the fee was not as assumed - which is the entire point of the tab |
| `expenses(cfg)` |  |
| `supplies(cfg)` |  |
| `maintenance(cfg)` |  |
| `reviews(cfg)` |  |
| `performance(cfg)` |  |
| `get(name, cfg)` | Sample rows for one register, or `[]` when there are none |

````python
"""A worked example: four properties, four months, every tab populated.

The numbers are deliberately not flattering.  One property is a
rent-to-rent that barely washes its face, one is a one-bedroom that
takes short stays and gets eaten alive by cleaning costs, and there is
a boiler failure in February that blocks a week.  A sample portfolio
where everything works teaches nobody anything.

Every row here is checked twice before it ships: the keys must all be
*input* columns on their register, and every value that feeds a
drop-down must exist in that drop-down's list.
"""

from __future__ import annotations

import datetime as _dt
from typing import Any, Callable, Dict, List

from .config import Config

Rows = List[Dict[str, Any]]


def _day(cfg: Config, offset: int) -> _dt.date:
    """A date `offset` days from the build date."""
    return cfg.today + _dt.timedelta(days=offset)


def _time(hour: int, minute: int = 0) -> _dt.time:
    return _dt.time(hour, minute)


# ---------------------------------------------------------------------------
# 1. Property Database
# ---------------------------------------------------------------------------

def properties(cfg: Config) -> Rows:
    return [
        {
            "name": "Harbour View Loft", "nickname": "Harbour",
            "address": "14 Quayside Walk, Apt 3B", "city": "Portland",
            "country": "United States", "type": "Loft",
            "bedrooms": 1, "bathrooms": 1, "beds": 2, "max_guests": 3,
            "status": "Active", "listed_from": _day(cfg, -690),
            "owner": "Nadia Whitlock", "manager": "Nadia Whitlock",
            "ownership": "Owned outright",
            "main_platform": "Airbnb",
            "listing_url": "https://example.com/listing/harbour-loft",
            "licence": "STR-2023-04417",
            "check_in_time": _time(15), "check_out_time": _time(11),
            "min_nights": 2, "wifi_name": "HarbourLoft-5G",
            "access": "Lockbox by the lift, code 4417",
            "nightly_rate": 165, "cleaning_fee": 85,
            "cleaning_cost": 60, "deposit": 250,
            "monthly_rent": 0, "monthly_mortgage": 0,
            "purchase_price": 289000,
            "notes": "The reliable one. Short stays, high turnover, "
                     "cleaning is the whole game here.",
        },
        {
            "name": "Pine Ridge Cabin", "nickname": "Cabin",
            "address": "2 Pine Ridge Road", "city": "Mount Hood",
            "country": "United States", "type": "Cabin",
            "bedrooms": 3, "bathrooms": 2, "beds": 4, "max_guests": 8,
            "status": "Active", "listed_from": _day(cfg, -520),
            "owner": "Nadia Whitlock", "manager": "Nadia Whitlock",
            "ownership": "Owned with mortgage",
            "main_platform": "Vrbo",
            "listing_url": "https://example.com/listing/pine-ridge",
            "licence": "STR-2023-09112",
            "check_in_time": _time(16), "check_out_time": _time(10),
            "min_nights": 3, "wifi_name": "PineRidge",
            "access": "Key safe on the porch post, code 2291",
            "nightly_rate": 395, "cleaning_fee": 180,
            "cleaning_cost": 135, "deposit": 600,
            "monthly_rent": 0, "monthly_mortgage": 1840,
            "purchase_price": 612000,
            "notes": "Long stays, families, strong summer. The "
                     "mortgage makes winter look thin.",
        },
        {
            "name": "Old Mill Townhouse", "nickname": "Mill",
            "address": "71 Old Mill Lane", "city": "Portland",
            "country": "United States", "type": "Townhouse",
            "bedrooms": 2, "bathrooms": 2, "beds": 3, "max_guests": 5,
            "status": "Active", "listed_from": _day(cfg, -400),
            "owner": "Marcus Oyelaran", "manager": "Nadia Whitlock",
            "ownership": "Rental arbitrage",
            "main_platform": "Booking.com",
            "listing_url": "https://example.com/listing/old-mill",
            "licence": "STR-2024-01188",
            "check_in_time": _time(15), "check_out_time": _time(11),
            "min_nights": 2, "wifi_name": "OldMill-Guest",
            "access": "Smart lock, code sent the morning of arrival",
            "nightly_rate": 210, "cleaning_fee": 110,
            "cleaning_cost": 80, "deposit": 300,
            "monthly_rent": 1950, "monthly_mortgage": 0,
            "purchase_price": 0,
            "notes": "Rent-to-rent. The rent goes out whether anyone "
                     "stays or not, which is the whole risk.",
        },
        {
            "name": "Creekside Bungalow", "nickname": "Creek",
            "address": "9 Creekside Avenue", "city": "Hood River",
            "country": "United States", "type": "Cottage",
            "bedrooms": 2, "bathrooms": 1, "beds": 3, "max_guests": 4,
            "status": "Active", "listed_from": _day(cfg, -210),
            "owner": "Nadia Whitlock", "manager": "Priya Raman",
            "ownership": "Owned with mortgage",
            "main_platform": "Airbnb",
            "listing_url": "https://example.com/listing/creekside",
            "licence": "STR-2025-02204",
            "check_in_time": _time(15), "check_out_time": _time(11),
            "min_nights": 2, "wifi_name": "Creekside",
            "access": "Lockbox on the gate, code 8830",
            "nightly_rate": 185, "cleaning_fee": 95,
            "cleaning_cost": 70, "deposit": 300,
            "monthly_rent": 0, "monthly_mortgage": 980,
            "purchase_price": 338000,
            "notes": "Newest. Still building a review count, which is "
                     "why the rate is where it is.",
        },
        {
            "name": "Fir Street Apartment", "nickname": "Fir St",
            "address": "140 Fir Street, Unit 7", "city": "Portland",
            "country": "United States", "type": "Entire apartment",
            "bedrooms": 1, "bathrooms": 1, "beds": 1, "max_guests": 2,
            "status": "Being set up", "listed_from": _day(cfg, 21),
            "owner": "Nadia Whitlock", "manager": "Nadia Whitlock",
            "ownership": "Owned with mortgage",
            "main_platform": "Airbnb",
            "listing_url": "",
            "licence": "Applied for",
            "check_in_time": _time(15), "check_out_time": _time(11),
            "min_nights": 2, "wifi_name": "FirSt-Guest",
            "access": "To be fitted",
            "nightly_rate": 145, "cleaning_fee": 75,
            "cleaning_cost": 55, "deposit": 200,
            "monthly_rent": 0, "monthly_mortgage": 720,
            "purchase_price": 241000,
            "notes": "Opens next month. Here to show what a property "
                     "looks like before it earns anything.",
        },
    ]


# ---------------------------------------------------------------------------
# 2. Staff & Vendors
# ---------------------------------------------------------------------------

def vendors(cfg: Config) -> Rows:
    return [
        {"name": "Priya Raman", "role": "Head cleaner",
         "company": "Raman Turnovers", "phone": "+1 503 555 0142",
         "email": "priya@example.com", "properties": "All Portland",
         "availability": "Any day",
         "pay_basis": "Per turnover", "rate": 65,
         "contract": "Active", "started": _day(cfg, -640),
         "insured": "Yes", "insurance_expiry": _day(cfg, 142),
         "keyholder": "Yes",
         "notes": "Does the bulk of the work and is worth every "
                  "penny. Also co-manages Creekside."},
        {"name": "Tomas Beric", "role": "Cleaner",
         "company": "", "phone": "+1 503 555 0188",
         "email": "tomas@example.com", "properties": "Cabin, Creek",
         "availability": "Weekends only",
         "pay_basis": "Per turnover", "rate": 120,
         "contract": "Active", "started": _day(cfg, -300),
         "insured": "No", "insurance_expiry": "",
         "keyholder": "Yes",
         "notes": "Covers the cabin, which is a four-hour turnover "
                  "and priced accordingly. Chase him for insurance."},
        {"name": "Glenbrook Laundry", "role": "Laundry service",
         "company": "Glenbrook Laundry Co.",
         "phone": "+1 503 555 0310", "email": "orders@example.com",
         "properties": "All", "availability": "Weekdays only",
         "pay_basis": "Invoice each time", "rate": 3.40,
         "contract": "Active", "started": _day(cfg, -610),
         "insured": "Yes", "insurance_expiry": _day(cfg, 310),
         "keyholder": "No",
         "notes": "Collects Tuesday and Friday. Cheaper than "
                  "in-house once you count the machine."},
        {"name": "Dale Kirkbride", "role": "Handyman",
         "company": "Kirkbride Property Services",
         "phone": "+1 503 555 0266", "email": "dale@example.com",
         "properties": "All", "availability": "By arrangement",
         "pay_basis": "Per hour", "rate": 72,
         "contract": "Active", "started": _day(cfg, -580),
         "insured": "Yes", "insurance_expiry": _day(cfg, 24),
         "keyholder": "Yes",
         "notes": "Insurance is about to lapse - the Action column "
                  "has already noticed."},
        {"name": "Hollis Plumbing", "role": "Plumber",
         "company": "Hollis & Sons", "phone": "+1 503 555 0401",
         "email": "jobs@example.com", "properties": "All",
         "availability": "By arrangement", "pay_basis": "Per job",
         "rate": 240, "contract": "Active",
         "started": _day(cfg, -420), "insured": "Yes",
         "insurance_expiry": _day(cfg, 201), "keyholder": "No",
         "notes": "Emergency rate after six. Did the cabin boiler."},
        {"name": "Ines Fuentes", "role": "Co-host",
         "company": "", "phone": "+1 503 555 0529",
         "email": "ines@example.com", "properties": "Harbour, Mill",
         "availability": "Weekdays only", "pay_basis": "% of revenue",
         "rate": 0, "contract": "Active",
         "started": _day(cfg, -120), "insured": "No",
         "insurance_expiry": "", "keyholder": "Yes",
         "notes": "Takes 12% of net on the two Portland units for "
                  "guest messaging. Logged as a co-host fee."},
        {"name": "Evergreen Gardens", "role": "Gardener",
         "company": "Evergreen Gardens LLC",
         "phone": "+1 503 555 0633", "email": "hello@example.com",
         "properties": "Cabin, Creek", "availability": "Weekdays only",
         "pay_basis": "Monthly retainer", "rate": 180,
         "contract": "Active", "started": _day(cfg, -500),
         "insured": "Yes", "insurance_expiry": _day(cfg, 95),
         "keyholder": "No", "notes": "Fortnightly in summer."},
        {"name": "Pacific Pest Control", "role": "Pest control",
         "company": "Pacific Pest Control",
         "phone": "+1 503 555 0744", "email": "book@example.com",
         "properties": "All", "availability": "By arrangement",
         "pay_basis": "Per job", "rate": 150,
         "contract": "Occasional", "started": _day(cfg, -360),
         "insured": "Yes", "insurance_expiry": _day(cfg, 180),
         "keyholder": "No",
         "notes": "Annual preventative visit. Not used since."},
    ]


# ---------------------------------------------------------------------------
# 3. Guest Database
# ---------------------------------------------------------------------------

_GUESTS = [
    ("Rebecca Thorne", "rebecca.t@example.com", "+1 206 555 0118",
     "United States", "Couple", "Holiday", "Airbnb", "Yes", "Received",
     "Ground floor if possible", "Late check-in, arriving 10pm",
     "Loved it. Said the bed was the best part.",
     "Excellent guest. Would have back tomorrow."),
    ("Dai Nguyen", "dai.nguyen@example.com", "+1 503 555 0227",
     "United States", "Business", "Business", "Direct booking",
     "Yes", "Received", "Desk and fast wifi",
     "Needs a receipt addressed to the company",
     "Quiet, tidy, books the same week every quarter.",
     "Repeat corporate. Give him the direct rate."),
    ("Marion Esseltine", "marion.e@example.com", "+44 7700 900412",
     "United Kingdom", "Family", "Holiday", "Vrbo", "No", "Received",
     "Travel cot and high chair",
     "Flying in late, needs flexible check-in",
     "Said the cabin was perfect for the children.",
     "Family of five. Books the cabin in August."),
    ("Grant Oyelaran", "grant.o@example.com", "+1 971 555 0340",
     "United States", "Friends", "Holiday", "Booking.com", "No",
     "Not requested", "", "", "", "One night, left a towel behind."),
    ("Suzanna Pell", "suzanna.pell@example.com", "+1 415 555 0455",
     "United States", "Couple", "Holiday", "Airbnb", "Yes", "Received",
     "Prefers a firm mattress", "Early check-in if free",
     "Marked us down on cleanliness. Fair - that was the "
     "Tuesday Tomas could not cover.",
     "The review that cost us. Worth winning back."),
    ("Hector Bellweather", "hector.b@example.com", "+1 503 555 0562",
     "United States", "Couple", "Holiday", "Airbnb", "Yes", "Received",
     "Dog friendly", "Bringing one small dog",
     "Great stay, will return with the dog.",
     "Pet fee paid without being asked. Good sign."),
    ("Amira Sadiq", "amira.sadiq@example.com", "+1 646 555 0677",
     "United States", "Long stay", "Relocation", "Corporate / "
     "relocation", "No", "Not requested", "Long stay, needs laundry",
     "Six weeks while her flat is found",
     "", "Longest booking of the year. Worth a discount."),
    ("Niall Fergusson", "niall.f@example.com", "+353 87 555 0781",
     "Ireland", "Couple", "Holiday", "Expedia", "No", "Received", "",
     "", "Nothing but five stars.",
     "Booked through Expedia - the expensive channel."),
    ("Petra Lindqvist", "petra.l@example.com", "+46 70 555 0894",
     "Sweden", "Couple", "Event", "Airbnb", "Yes", "Received",
     "Quiet room, late check-out",
     "Celebrating an anniversary - a bottle would be welcome",
     "Said the welcome basket made the trip.",
     "Welcome basket costs eleven dollars and bought a "
     "five-star review."),
    ("Owen Castellanos", "owen.c@example.com", "+1 503 555 0902",
     "United States", "Group", "Event", "Vrbo", "No",
     "Received", "Eight people, two cars",
     "Wedding party - needs the full cabin",
     "Great space, slightly tired bathroom.",
     "Flagged the bathroom. Booked Dale to look at it."),
    ("Fenella Adebayo", "fenella.a@example.com", "+1 212 555 1014",
     "United States", "Couple", "Holiday", "Direct booking", "Yes",
     "Not requested", "", "Returning - third stay", "",
     "Our best direct guest. Zero commission three times over."),
    ("Jurgen Haas", "jurgen.haas@example.com", "+49 151 555 1127",
     "Germany", "Couple", "Holiday", "Booking.com", "No",
     "Not requested", "", "Arriving by train, no car", "",
     "Cancelled late. Booking.com terms meant no fee."),
]


def guests(cfg: Config) -> Rows:
    keys = ("name", "email", "phone", "country", "guest_type",
            "purpose", "first_platform", "marketing_ok",
            "review_status", "preferences", "requests", "feedback",
            "internal")
    return [dict(zip(keys, row)) for row in _GUESTS]


# ---------------------------------------------------------------------------
# 4. Booking Tracker
# ---------------------------------------------------------------------------

#: (property, guest, platform, status, booked_offset, in_offset,
#:  nights, adults, children, pets, rate, payment, requests)
_BOOKINGS = [
    # ---- finished stays, three months back ----------------------------
    ("PRP-01", "GST-01", "Airbnb", "Completed", -112, -96, 3, 2, 0, 0,
     165, "Paid in full", "Late check-in, arriving 10pm"),
    ("PRP-02", "GST-03", "Vrbo", "Completed", -120, -93, 7, 2, 3, 0,
     410, "Paid in full", "Travel cot and high chair please"),
    ("PRP-03", "GST-04", "Booking.com", "Completed", -95, -90, 1, 2, 0,
     0, 210, "Paid in full", ""),
    ("PRP-01", "GST-02", "Direct booking", "Completed", -100, -88, 2,
     1, 0, 0, 155, "Paid in full", "Receipt to the company please"),
    ("PRP-04", "GST-06", "Airbnb", "Completed", -99, -85, 4, 2, 0, 1,
     185, "Paid in full", "One small dog"),
    ("PRP-01", "GST-05", "Airbnb", "Completed", -90, -82, 2, 2, 0, 0,
     165, "Paid in full", "Early check-in if free"),
    ("PRP-02", "GST-10", "Vrbo", "Completed", -102, -78, 3, 8, 0, 0,
     450, "Paid in full", "Wedding party, eight people"),
    ("PRP-03", "GST-08", "Expedia", "Completed", -84, -74, 4, 2, 0, 0,
     215, "Paid in full", ""),
    ("PRP-01", "GST-09", "Airbnb", "Completed", -80, -70, 3, 2, 0, 0,
     175, "Paid in full", "Anniversary - quiet room"),
    ("PRP-04", "GST-01", "Airbnb", "Completed", -75, -66, 2, 2, 0, 0,
     185, "Paid in full", "Second stay"),

    # ---- two months back ----------------------------------------------
    ("PRP-02", "GST-03", "Vrbo", "Completed", -78, -62, 5, 2, 3, 0,
     395, "Paid in full", "Returning family"),
    ("PRP-03", "GST-11", "Direct booking", "Completed", -70, -58, 3, 2,
     0, 0, 195, "Paid in full", "Third stay - direct"),
    ("PRP-01", "GST-04", "Booking.com", "Completed", -62, -55, 1, 2, 0,
     0, 170, "Paid in full", ""),
    ("PRP-04", "GST-05", "Airbnb", "Completed", -66, -52, 3, 2, 0, 0,
     185, "Paid in full", "Firm mattress if possible"),
    ("PRP-01", "GST-06", "Airbnb", "Completed", -58, -48, 2, 2, 0, 1,
     165, "Paid in full", "Dog again"),
    ("PRP-02", "GST-10", "Vrbo", "Cancelled", -64, -46, 4, 6, 0, 0,
     395, "Refunded", "Cancelled eleven days out"),
    ("PRP-03", "GST-12", "Booking.com", "Cancelled", -55, -44, 2, 2, 0,
     0, 210, "Refunded", "Cancelled, no fee under their terms"),
    ("PRP-01", "GST-02", "Direct booking", "Completed", -50, -41, 2, 1,
     0, 0, 155, "Paid in full", "Quarterly work trip"),
    ("PRP-04", "GST-08", "Expedia", "Completed", -52, -38, 3, 2, 0, 0,
     190, "Paid in full", ""),
    ("PRP-02", "GST-07", "Corporate / relocation", "Completed", -60,
     -36, 14, 1, 0, 0, 340, "Paid in full",
     "Relocation - six weeks split over two bookings"),

    # ---- last month, including a maintenance block ---------------------
    ("PRP-02", "", "Direct booking", "Maintenance block", -34, -32, 6,
     0, 0, 0, 0, "Not due yet", "Boiler failed - closed the cabin"),
    ("PRP-01", "GST-09", "Airbnb", "Completed", -40, -29, 3, 2, 0, 0,
     175, "Paid in full", ""),
    ("PRP-03", "GST-01", "Booking.com", "Completed", -36, -26, 2, 2, 0,
     0, 210, "Paid in full", "Third property for this guest"),
    ("PRP-04", "GST-11", "Direct booking", "Completed", -38, -24, 4, 2,
     0, 0, 185, "Paid in full", ""),
    ("PRP-01", "GST-05", "Airbnb", "Completed", -30, -21, 1, 2, 0, 0,
     165, "Paid in full", "One night - the expensive kind"),
    ("PRP-02", "GST-07", "Corporate / relocation", "Completed", -45,
     -20, 14, 1, 0, 0, 340, "Paid in full", "Second half"),
    ("PRP-03", "GST-06", "Airbnb", "Completed", -28, -18, 3, 2, 0, 1,
     215, "Paid in full", "Dog"),
    ("PRP-01", "GST-08", "Expedia", "Completed", -26, -14, 2, 2, 0, 0,
     170, "Paid in full", ""),
    ("PRP-04", "GST-09", "Airbnb", "Completed", -24, -11, 3, 2, 0, 0,
     190, "Paid in full", ""),

    # ---- this month, in progress ---------------------------------------
    ("PRP-01", "GST-11", "Direct booking", "Completed", -20, -8, 2, 2,
     0, 0, 160, "Paid in full", "Direct again - no commission"),
    ("PRP-03", "GST-04", "Booking.com", "Completed", -18, -6, 2, 2, 0,
     0, 210, "Awaiting payout", ""),
    ("PRP-02", "GST-03", "Vrbo", "Completed", -30, -5, 4, 2, 3, 0,
     420, "Awaiting payout", "Half term"),
    ("PRP-04", "GST-05", "Airbnb", "Checked in", -16, -2, 5, 2, 0, 0,
     195, "Awaiting payout", "Currently in the property"),
    ("PRP-01", "GST-06", "Airbnb", "Checked in", -14, -1, 3, 2, 0, 1,
     175, "Awaiting payout", "Dog, as usual"),

    # ---- today and the days around it ----------------------------------
    ("PRP-03", "GST-09", "Airbnb", "Confirmed", -10, 0, 3, 2, 0, 0,
     220, "Awaiting payout", "Arriving today"),
    ("PRP-02", "GST-10", "Vrbo", "Confirmed", -12, 1, 4, 7, 0, 0, 430,
     "Part paid", "Arriving tomorrow - seven guests"),
    ("PRP-01", "GST-01", "Airbnb", "Confirmed", -8, 2, 2, 2, 0, 0, 180,
     "Part paid", ""),
    ("PRP-04", "GST-08", "Expedia", "Confirmed", -6, 4, 3, 2, 0, 0,
     195, "Not due yet", ""),
    ("PRP-03", "GST-11", "Direct booking", "Confirmed", -5, 6, 4, 2, 0,
     0, 200, "Part paid", "Fourth stay"),
    ("PRP-01", "GST-02", "Direct booking", "Confirmed", -4, 9, 2, 1, 0,
     0, 155, "Part paid", "Quarterly work trip"),
    ("PRP-02", "GST-03", "Vrbo", "Confirmed", -22, 12, 7, 2, 3, 0, 440,
     "Part paid", "Booked well ahead"),
    ("PRP-04", "GST-06", "Airbnb", "Confirmed", -3, 15, 3, 2, 0, 1,
     195, "Not due yet", "Dog"),
    ("PRP-01", "", "Airbnb", "Owner block", -2, 18, 4, 0, 0, 0, 0,
     "Not due yet", "Keeping it for family"),
    ("PRP-03", "GST-12", "Booking.com", "Inquiry", 0, 22, 3, 2, 0, 0,
     0, "Not due yet", "Asked about the train station"),
    ("PRP-02", "GST-10", "Vrbo", "Confirmed", -1, 26, 5, 8, 0, 0, 460,
     "Part paid", "Full house again"),
    ("PRP-04", "GST-05", "Airbnb", "Inquiry", 0, 31, 2, 2, 0, 0, 0,
     "Not due yet", "Asked whether the creek is swimmable"),
]


def bookings(cfg: Config) -> Rows:
    rows: Rows = []
    for (prop, guest, platform, status, booked, start, nights, adults,
         children, pets, rate, payment, requests) in _BOOKINGS:
        row: Dict[str, Any] = {
            "property": prop,
            "platform": platform,
            "status": status,
            "booked_date": _day(cfg, booked),
            "check_in": _day(cfg, start),
            "check_out": _day(cfg, start + nights),
            "adults": adults,
            "payment_status": payment,
        }
        if guest:
            row["guest"] = guest
        if children:
            row["children"] = children
        if pets:
            row["pets"] = pets
        if rate:
            row["rate_in"] = rate
        if requests:
            row["requests"] = requests
        if platform in ("Airbnb", "Vrbo", "Booking.com", "Expedia"):
            row["reference"] = f"HM{abs(start):04d}{adults}{nights}"
        rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# 5. Check-in & Check-out
# ---------------------------------------------------------------------------

def arrivals(cfg: Config) -> Rows:
    """One row per recent and upcoming stay, in varying states."""
    rows: Rows = []
    finished = [
        (22, "Smart lock code", True, "Guest left a towel behind."),
        (23, "Lockbox", True, ""),
        (24, "Lockbox", True, "Shower pressure was mentioned again."),
        (25, "Lockbox", True, "One night. Barely worth the turnover."),
        (26, "Key safe", True, "Long stay - laundry done mid-way."),
        (27, "Smart lock code", True, ""),
        (28, "Lockbox", True, "Late departure, cleaner had to wait."),
        (29, "Lockbox", True, ""),
        (30, "Lockbox", True, ""),
        (31, "Smart lock code", True, ""),
        (32, "Key safe", True, ""),
    ]
    for index, method, complete, issues in finished:
        row: Dict[str, Any] = {
            "booking": f"BKG-{index:04d}",
            "arrival_time": _time(16),
            "departure_time": _time(10, 30),
            "late_checkout": "No", "early_checkin": "No",
            "key_method": method,
            "issues": issues,
        }
        for key in ("sent_directions", "sent_access", "sent_wifi",
                    "sent_rules", "sent_welcome"):
            row[key] = "Yes"
        if complete:
            for key in ("sent_checkout", "checked_out", "inspected",
                        "photographed"):
                row[key] = "Yes"
        rows.append(row)

    # in-house now: pre-arrival done, nothing after departure yet
    for index, method in ((33, "Lockbox"), (34, "Smart lock code")):
        rows.append({
            "booking": f"BKG-{index:04d}",
            "arrival_time": _time(15, 30),
            "departure_time": _time(11),
            "late_checkout": "No", "early_checkin": "Yes",
            "key_method": method,
            "sent_directions": "Yes", "sent_access": "Yes",
            "sent_wifi": "Yes", "sent_rules": "Yes",
            "sent_welcome": "Yes",
            "notes": "Guest is in the property now.",
        })

    # arriving today: everything sent
    rows.append({
        "booking": "BKG-0035", "arrival_time": _time(15),
        "departure_time": _time(11), "late_checkout": "No",
        "early_checkin": "No", "key_method": "Lockbox",
        "sent_directions": "Yes", "sent_access": "Yes",
        "sent_wifi": "Yes", "sent_rules": "Yes", "sent_welcome": "Yes",
        "notes": "Arriving today and fully briefed.",
    })
    # arriving tomorrow: access code not sent - the one to chase
    rows.append({
        "booking": "BKG-0036", "arrival_time": _time(16),
        "departure_time": _time(10), "late_checkout": "No",
        "early_checkin": "No", "key_method": "Key safe",
        "sent_directions": "Yes", "sent_access": "No",
        "sent_wifi": "Yes", "sent_rules": "Yes", "sent_welcome": "No",
        "notes": "Seven guests arriving tomorrow and the access code "
                 "has not gone out. This is what the Action column is "
                 "for.",
    })
    # arriving in a few days: barely started
    for index in (37, 38, 39):
        rows.append({
            "booking": f"BKG-{index:04d}",
            "arrival_time": _time(15),
            "departure_time": _time(11),
            "late_checkout": "No", "early_checkin": "No",
            "key_method": "Lockbox",
            "sent_directions": "Yes", "sent_access": "No",
            "sent_wifi": "No", "sent_rules": "No",
            "sent_welcome": "No",
        })
    return rows


# ---------------------------------------------------------------------------
# 6. Cleaning Schedule
# ---------------------------------------------------------------------------

def cleaning(cfg: Config) -> Rows:
    """A turnover for every completed stay, plus the ones coming up."""
    every_task = ("task_bedroom", "task_bathroom", "task_kitchen",
                  "task_living", "task_floors", "task_laundry",
                  "task_rubbish", "task_restock", "task_inspection")
    # booking number -> (cleaner, hours, status, skipped tasks, note)
    jobs = [
        (1, "VND-01", 2.0, "Complete", (), ""),
        (2, "VND-02", 4.0, "Complete", (), "Cabin after a week - "
                                            "four hours is normal."),
        (3, "VND-01", 2.5, "Complete", (), ""),
        (4, "VND-01", 2.0, "Complete", (), ""),
        (5, "VND-01", 2.5, "Complete", (), ""),
        (6, "VND-01", 2.0, "Complete", ("task_restock",),
         "Ran out of coffee pods - restock not done."),
        (7, "VND-02", 4.5, "Complete", (), "Eight guests. Enough "
                                            "said."),
        (8, "VND-01", 2.5, "Complete", (), ""),
        (9, "VND-01", 2.0, "Complete", (), ""),
        (10, "VND-01", 2.5, "Complete", (), ""),
        (11, "VND-02", 4.0, "Complete", (), ""),
        (12, "VND-01", 2.5, "Complete", (), ""),
        (13, "VND-01", 2.0, "Complete", (), ""),
        (14, "VND-01", 2.5, "Complete", ("task_floors",),
         "Vacuum broke mid-clean. This is the turnover that cost us "
         "a cleanliness score."),
        (15, "VND-01", 2.0, "Complete", (), ""),
        (18, "VND-01", 2.0, "Complete", (), ""),
        (19, "VND-01", 2.5, "Complete", (), ""),
        (20, "VND-02", 3.5, "Complete", (), "Mid-stay clean on the "
                                             "relocation booking."),
        (22, "VND-01", 2.0, "Complete", (), ""),
        (23, "VND-01", 2.5, "Complete", (), ""),
        (24, "VND-01", 2.5, "Complete", (), ""),
        (25, "VND-01", 2.0, "Complete", (), "One-night stay. Same "
                                             "cost as a week."),
        (26, "VND-02", 3.5, "Complete", (), ""),
        (27, "VND-01", 2.5, "Complete", (), ""),
        (28, "VND-01", 2.0, "Complete", (), ""),
        (29, "VND-01", 2.5, "Complete", (), ""),
        (30, "VND-01", 2.0, "Complete", (), ""),
        (31, "VND-01", 2.5, "Complete", (), ""),
        (32, "VND-02", 4.0, "Complete", (), ""),
    ]
    rows: Rows = []
    for number, cleaner, hours, status, skipped, note in jobs:
        start_hour = 11
        row: Dict[str, Any] = {
            "booking": f"BKG-{number:04d}",
            "cleaner": cleaner,
            "start_time": _time(start_hour),
            "end_time": _time(start_hour + int(hours),
                              30 if hours % 1 else 0),
            "status": status,
            "laundry_cost": 28 if cleaner == "VND-02" else 16,
            "supplies_cost": 14 if cleaner == "VND-02" else 9,
            "paid": "Yes",
        }
        for key in every_task:
            row[key] = "No" if key in skipped else "Yes"
        if note:
            row["notes"] = note
        rows.append(row)

    # the in-flight ones
    rows.append({
        "booking": "BKG-0033", "cleaner": "VND-01",
        "status": "Scheduled", "laundry_cost": 16, "supplies_cost": 9,
        "paid": "No",
        "notes": "Guest still in. Clean is booked for check-out day.",
    })
    rows.append({
        "booking": "BKG-0034", "cleaner": "VND-01",
        "start_time": _time(11), "end_time": _time(13),
        "status": "In progress", "laundry_cost": 16,
        "supplies_cost": 9, "paid": "No",
        "task_bedroom": "Yes", "task_bathroom": "Yes",
        "task_kitchen": "Yes", "task_living": "No",
        "task_floors": "No", "task_laundry": "Yes",
        "task_rubbish": "No", "task_restock": "No",
        "task_inspection": "No",
        "notes": "Same-day turnover happening right now - a guest "
                 "arrives this afternoon.",
    })
    rows.append({
        "booking": "BKG-0035", "cleaner": "VND-01",
        "status": "Scheduled", "laundry_cost": 16, "supplies_cost": 9,
        "paid": "No", "notes": "Booked ahead for check-out day.",
    })
    rows.append({
        "booking": "BKG-0036", "cleaner": "VND-02",
        "status": "Scheduled", "laundry_cost": 28, "supplies_cost": 14,
        "paid": "No",
        "notes": "Cabin, eight guests. Tomas, four hours.",
    })
    return rows


# ---------------------------------------------------------------------------
# 7. Platform Revenue & Fees
# ---------------------------------------------------------------------------

def payouts(cfg: Config) -> Rows:
    """Reconciled payouts, including two where the fee was not as
    assumed - which is the entire point of the tab."""
    settled = [
        (1, 0.0, 0.0, 0.0), (2, 0.0, 0.0, 0.0), (3, 0.0, 0.0, 0.0),
        (4, 0.0, 0.0, 0.0), (5, 0.0, 0.0, 0.0), (6, 0.0, 0.0, 0.0),
        (7, 0.0, 0.0, 0.0), (8, -18.40, 0.0, 0.0),
        (9, 0.0, 0.0, 0.0), (10, 0.0, 0.0, 0.0),
        (11, 0.0, 0.0, 0.0), (12, 0.0, 0.0, 0.0),
        (13, 0.0, 0.0, 0.0), (14, 0.0, 0.0, -35.00),
        (15, 0.0, 0.0, 0.0), (18, 0.0, 0.0, 0.0),
        (19, -22.10, 0.0, 0.0), (20, 0.0, 0.0, 0.0),
        (22, 0.0, 0.0, 0.0), (23, 0.0, 0.0, 0.0),
        (24, 0.0, 0.0, 0.0), (25, 0.0, 0.0, 0.0),
        (26, 0.0, 0.0, 0.0), (27, 0.0, 0.0, 0.0),
        (28, 0.0, 0.0, 0.0), (29, 0.0, 0.0, 0.0),
        (30, 0.0, 0.0, 0.0),
    ]
    rows: Rows = []
    for number, variance, refund, adjustment in settled:
        rows.append({
            "booking": f"BKG-{number:04d}",
            "status": "Received",
            "refund": refund or "",
            "adjustment": adjustment or "",
            "reference": f"PO-{number:05d}",
            "notes": ("Expedia kept more than the Settings rate - "
                      "worth checking the contract."
                      if variance else ""),
        })
    # awaiting
    for number in (31, 32):
        rows.append({
            "booking": f"BKG-{number:04d}",
            "status": "Expected",
            "reference": f"PO-{number:05d}",
            "notes": "Due in the next payout run.",
        })
    # overdue - the one the dashboard chases
    rows.append({
        "booking": "BKG-0016",
        "status": "Late",
        "reference": "PO-00016",
        "notes": "Cancelled stay. Platform still owes the "
                 "non-refundable portion and it is now well past "
                 "their usual terms.",
    })
    return rows


# ---------------------------------------------------------------------------
# 8. Expense Tracker
# ---------------------------------------------------------------------------

def expenses(cfg: Config) -> Rows:
    base: Rows = []
    recurring = [
        ("PRP-03", "Rent", "Fulcrum Lettings",
         "Monthly rent to the landlord", 1950.00, "Monthly",
         "Operating expense", "Bank transfer"),
        ("PRP-02", "Mortgage interest", "Cascade Mutual",
         "Interest portion of the cabin mortgage", 1180.00, "Monthly",
         "Financing cost", "Direct debit"),
        ("PRP-04", "Mortgage interest", "Cascade Mutual",
         "Interest portion of the bungalow mortgage", 604.00,
         "Monthly", "Financing cost", "Direct debit"),
        ("PRP-01", "Utilities - electric", "Northwest Power",
         "Electricity", 86.40, "Monthly", "Operating expense",
         "Direct debit"),
        ("PRP-02", "Utilities - electric", "Northwest Power",
         "Electricity - the cabin is not cheap to heat", 214.80,
         "Monthly", "Operating expense", "Direct debit"),
        ("PRP-03", "Utilities - electric", "Northwest Power",
         "Electricity", 104.20, "Monthly", "Operating expense",
         "Direct debit"),
        ("PRP-04", "Utilities - electric", "Northwest Power",
         "Electricity", 92.10, "Monthly", "Operating expense",
         "Direct debit"),
        ("PRP-01", "Internet & TV", "Stream Fibre",
         "Guest wifi", 55.00, "Monthly", "Operating expense",
         "Direct debit"),
        ("PRP-02", "Internet & TV", "Stream Fibre",
         "Guest wifi", 62.00, "Monthly", "Operating expense",
         "Direct debit"),
        ("PRP-03", "Internet & TV", "Stream Fibre",
         "Guest wifi", 55.00, "Monthly", "Operating expense",
         "Direct debit"),
        ("PRP-04", "Internet & TV", "Stream Fibre",
         "Guest wifi", 55.00, "Monthly", "Operating expense",
         "Direct debit"),
        ("", "Insurance", "Harbourline Insurance",
         "Short-term let policy, all properties", 268.00, "Monthly",
         "Operating expense", "Direct debit"),
        ("", "Software & subscriptions", "PriceLabs",
         "Dynamic pricing tool", 39.00, "Monthly",
         "Operating expense", "Credit card"),
        ("", "Software & subscriptions", "Hospitable",
         "Guest messaging", 48.00, "Monthly", "Operating expense",
         "Credit card"),
        ("PRP-02", "Service charge", "Pine Ridge Association",
         "Road and gate maintenance levy", 95.00, "Monthly",
         "Operating expense", "Bank transfer"),
    ]
    for offset_months in range(4):
        day = _day(cfg, -offset_months * 30 - 4)
        for (prop, category, vendor, description, amount, frequency,
             tax_category, method) in recurring:
            row: Dict[str, Any] = {
                "date": day, "category": category, "vendor": vendor,
                "description": description, "amount": amount,
                "method": method, "receipt": "Yes",
                "frequency": frequency, "tax_category": tax_category,
            }
            if prop:
                row["property"] = prop
            base.append(row)

    one_offs = [
        (-104, "PRP-02", "Furniture", "Rowan & Field",
         "Replacement sofa bed for the cabin", 1240.00,
         "Capital / asset", "Credit card",
         "Capital, not a repair. It will not touch this year's "
         "profit and your accountant will be pleased."),
        (-92, "", "Marketing & photography", "Lumen Studio",
         "Re-shoot of the Portland listings", 780.00,
         "Operating expense", "Bank transfer",
         "Paid for itself in six weeks of better conversion."),
        (-76, "PRP-01", "Appliances", "Harbour Appliance Co.",
         "New washer-dryer", 689.00, "Capital / asset", "Credit card", ""),
        (-64, "PRP-04", "Guest amenities", "Cedar & Salt",
         "Welcome baskets, twelve", 132.00, "Operating expense",
         "Credit card", "Eleven dollars a basket bought a five-star "
                 "review. Best money spent all quarter."),
        (-58, "", "Accounting & legal", "Mercer & Vance",
         "Quarterly bookkeeping", 420.00, "Operating expense",
         "Bank transfer", ""),
        (-46, "PRP-02", "Repairs", "Hollis & Sons",
         "Emergency boiler repair - cabin closed six nights",
         1180.00, "Operating expense", "Bank transfer",
         "The week that cost the cabin its February."),
        (-44, "PRP-02", "Guest refunds", "Marion Esseltine",
         "Partial refund, boiler outage", 395.00,
         "Operating expense", "Bank transfer", ""),
        (-38, "", "Licences & permits", "City of Portland",
         "Short-term let registration renewal", 310.00,
         "Operating expense", "Credit card", ""),
        (-31, "PRP-03", "Co-host fees", "Ines Fuentes",
         "12% of net on the Portland units", 288.00,
         "Operating expense", "Bank transfer", ""),
        (-24, "PRP-01", "Supplies", "Restock Wholesale",
         "Linen, towels, consumables", 246.00, "Operating expense",
         "Credit card", ""),
        (-18, "", "Travel & mileage", "",
         "Driving between properties, 340 miles", 221.00,
         "Operating expense", "Cash", ""),
        (-12, "PRP-04", "Repairs", "Kirkbride Property Services",
         "Re-seal the bathroom, re-grout tiles", 340.00,
         "Operating expense", "Bank transfer",
         "Raised by a guest in a review. Fixed within the week."),
        (-9, "", "Bank & card fees", "Northline Bank",
         "Account and card fees", 42.00, "Operating expense",
         "Direct debit", ""),
        (-6, "PRP-05", "Furniture", "Rowan & Field",
         "Furnishing the new Fir Street flat", 2840.00,
         "Capital / asset", "Credit card",
         "Setting up a property that has not earned a night yet."),
        (-3, "PRP-05", "Appliances", "Harbour Appliance Co.",
         "Fridge, washer and oven for Fir Street", 1690.00,
         "Capital / asset", "Credit card", ""),
    ]
    for (offset, prop, category, vendor, description, amount,
         tax_category, method, note) in one_offs:
        row = {
            "date": _day(cfg, offset), "category": category,
            "vendor": vendor, "description": description,
            "amount": amount, "method": method, "receipt": "Yes",
            "frequency": "One off", "tax_category": tax_category,
        }
        if prop:
            row["property"] = prop
        if note:
            row["notes"] = note
        base.append(row)
    return base


# ---------------------------------------------------------------------------
# 9. Supplies Inventory
# ---------------------------------------------------------------------------

def supplies(cfg: Config) -> Rows:
    items = [
        ("Toilet rolls", "Soap & hand wash", "", "Pack of 9",
         "Restock Wholesale", "RW-TR09", 14, 6, 3, 8.40, 12,
         "Across the whole portfolio."),
        ("Kitchen roll", "Kitchen supplies", "", "Pack of 4",
         "Restock Wholesale", "RW-KR04", 9, 4, 1.5, 6.20, 10, ""),
        ("Hand soap refill", "Soap & hand wash", "", "500ml",
         "Restock Wholesale", "RW-HS50", 7, 6, 0.5, 3.10, 12,
         "At the minimum. The Action column has noticed."),
        ("Shampoo refill", "Soap & hand wash", "", "500ml",
         "Restock Wholesale", "RW-SH50", 11, 6, 0.6, 3.80, 12, ""),
        ("Dishwasher tablets", "Kitchen supplies", "", "Box of 60",
         "Restock Wholesale", "RW-DT60", 3, 3, 4, 11.50, 6,
         "Exactly at minimum with four turnovers a month."),
        ("Coffee pods", "Welcome treats", "", "Box of 50",
         "Cedar & Salt", "CS-CP50", 0, 2, 6, 18.90, 6,
         "Out of stock. This is why one turnover shows restock "
         "not done."),
        ("Tea bags", "Welcome treats", "", "Box of 80",
         "Cedar & Salt", "CS-TB80", 4, 2, 4, 6.40, 4, ""),
        ("Bin bags", "Cleaning products", "", "Roll of 50",
         "Restock Wholesale", "RW-BB50", 6, 3, 3, 7.20, 6, ""),
        ("Multi-surface cleaner", "Cleaning products", "", "750ml",
         "Restock Wholesale", "RW-MS75", 8, 4, 1, 4.30, 8, ""),
        ("Bleach", "Cleaning products", "", "1 litre",
         "Restock Wholesale", "RW-BL10", 5, 3, 0.5, 2.90, 6, ""),
        ("King duvet covers", "Bed linen", "PRP-02", "Each",
         "Glenbrook Laundry", "GL-DK", 6, 4, 0, 42.00, 4,
         "Cabin only. Counted, not consumed."),
        ("Double duvet covers", "Bed linen", "", "Each",
         "Glenbrook Laundry", "GL-DD", 14, 8, 0, 34.00, 6, ""),
        ("Pillowcases", "Bed linen", "", "Each", "Glenbrook Laundry",
         "GL-PC", 38, 24, 0, 7.50, 12, ""),
        ("Bath towels", "Towels", "", "Each", "Glenbrook Laundry",
         "GL-BT", 31, 24, 0, 11.00, 12,
         "Guests take them. Budget for shrinkage."),
        ("Hand towels", "Towels", "", "Each", "Glenbrook Laundry",
         "GL-HT", 22, 16, 0, 5.50, 10, ""),
        ("Tea towels", "Kitchen supplies", "", "Each",
         "Glenbrook Laundry", "GL-TT", 12, 8, 0, 3.20, 8, ""),
        ("Mattress protectors", "Bed linen", "", "Each",
         "Glenbrook Laundry", "GL-MP", 7, 6, 0, 24.00, 4, ""),
        ("Light bulbs", "Batteries & bulbs", "", "Each",
         "Kirkbride Property Services", "KP-LB", 9, 6, 0.3, 3.60, 10,
         ""),
        ("Spare smart lock batteries", "Batteries & bulbs", "",
         "Pack of 4", "Kirkbride Property Services", "KP-BT", 2, 2,
         0.1, 9.80, 4, "One pack per property is the rule."),
        ("Welcome basket - local goods", "Welcome treats", "",
         "Each", "Cedar & Salt", "CS-WB", 5, 4, 1, 11.00, 10,
         "The eleven dollars that keeps buying five-star reviews."),
    ]
    rows: Rows = []
    for (name, category, prop, unit, supplier, sku, on_hand, min_level,
         per_turn, unit_cost, reorder, note) in items:
        row: Dict[str, Any] = {
            "name": name, "category": category, "unit": unit,
            "supplier": supplier, "sku": sku, "on_hand": on_hand,
            "min_level": min_level, "last_count": _day(cfg, -9),
            "last_restock": _day(cfg, -24),
            "per_turn": per_turn, "unit_cost": unit_cost,
            "reorder_qty": reorder,
        }
        if prop:
            row["property"] = prop
        if note:
            row["notes"] = note
        rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# 10. Maintenance Log
# ---------------------------------------------------------------------------

def maintenance(cfg: Config) -> Rows:
    jobs = [
        (-108, "PRP-01", "Plumbing", "Dripping kitchen tap", "Low",
         "Cleaner", "", "VND-04", -104, -103, "Complete", "No",
         80, 94, "KP-2291", "Yes", 365,
         "Small job, slight overrun. Normal."),
        (-96, "PRP-03", "Appliance", "Dishwasher not draining",
         "Normal", "Guest", "BKG-0003", "VND-04", -94, -93,
         "Complete", "No", 150, 150, "KP-2314", "Yes", 180, ""),
        (-72, "PRP-02", "Heating / cooling", "Thermostat replaced", "Normal",
         "Owner inspection", "", "VND-04", -70, -69, "Complete",
         "No", 220, 265, "KP-2402", "Yes", 730, ""),
        (-47, "PRP-02", "Heating / cooling",
         "Boiler failure - no hot water or heating", "Urgent",
         "Guest", "BKG-0011", "VND-05", -46, -41, "Complete", "Yes",
         900, 1180, "HS-88410", "Yes", 365,
         "Six nights blocked, one partial refund and a long "
         "weekend of phone calls. The single most expensive thing "
         "that happened this quarter."),
        (-44, "PRP-01", "Electrical", "Hallway light flickering",
         "Low", "Cleaner", "", "VND-04", -40, -39, "Complete", "No",
         60, 72, "KP-2477", "Yes", 180, ""),
        (-33, "PRP-04", "Furniture",
         "Bed slat broken in the second bedroom", "Normal", "Guest",
         "BKG-0024", "VND-04", -30, -29, "Complete", "No", 90, 118,
         "KP-2511", "Yes", 90, ""),
        (-26, "PRP-04", "Plumbing",
         "Re-seal bath, re-grout tiles - raised in a review",
         "Normal", "Review", "BKG-0010", "VND-04", -14, -12,
         "Complete", "No", 300, 340, "KP-2566", "Yes", 365,
         "Guest mentioned a tired bathroom. Fixed inside a "
         "fortnight, which is the right answer to a review."),
        (-16, "PRP-03", "Locks & access",
         "Smart lock battery warning", "Normal", "Owner inspection",
         "", "VND-04", -6, "", "Scheduled", "No", 40, "", "", "No",
         "", "Batteries on order. Not urgent until it is."),
        (-11, "PRP-02", "Garden / outdoor",
         "Decking boards lifting at the north end", "Normal",
         "Cleaner", "", "VND-04", 4, "", "Scheduled", "No", 480, "",
         "", "No", "",
         "Quoted. Booked for the gap between bookings."),
        (-5, "PRP-01", "Appliance",
         "Washer-dryer making a grinding noise", "High", "Cleaner",
         "", "VND-04", 2, "", "Reported", "No", 150, "", "", "No",
         "", "If it fails, every Harbour turnover needs the "
             "laundry service instead."),
        (-2, "PRP-03", "Plumbing",
         "Shower pressure poor - mentioned twice now", "High",
         "Guest", "BKG-0031", "VND-05", 1, "", "Scheduled", "Yes",
         260, "", "", "No", "",
         "Blocking the calendar on the day of the visit. Two "
         "guests have now mentioned it, so it is a rating "
         "problem rather than a plumbing one."),
        (-1, "PRP-05", "Other", "Snagging list for the new flat",
         "Low", "Owner inspection", "", "VND-04", 7, "", "Reported",
         "No", 400, "", "", "No", "",
         "Pre-launch. Nothing earning here yet."),
    ]
    keys = ("report_date", "property", "category", "description",
            "priority", "reported_by", "booking", "vendor",
            "scheduled", "completed", "status", "blocks_calendar",
            "est_cost", "act_cost", "invoice_ref", "paid",
            "warranty_until", "notes")
    rows: Rows = []
    for job in jobs:
        row: Dict[str, Any] = {}
        for key, value in zip(keys, job):
            if value == "":
                continue
            if key in ("report_date", "scheduled", "completed"):
                value = _day(cfg, value)
            elif key == "warranty_until":
                value = _day(cfg, value)
            row[key] = value
        rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# 11. Reviews Tracker
# ---------------------------------------------------------------------------

def reviews(cfg: Config) -> Rows:
    #  booking, overall, clean, accuracy, checkin, comm, location,
    #  value, responded, public, action
    entries = [
        (1, 5, 5, 5, 5, 5, 5, 5, "Yes",
         "Spotless, easy check-in, great spot by the water.", ""),
        (2, 5, 5, 5, 5, 5, 4, 5, "Yes",
         "Perfect for the children. We will be back.", ""),
        (3, 4, 4, 4, 5, 5, 4, 3, "Yes",
         "Fine for one night, but pricey once the cleaning fee "
         "goes on.",
         "Looked at the one-night pricing - raised the minimum "
         "stay to two."),
        (4, 5, 5, 5, 5, 5, 5, 4, "No",
         "Everything I need for a work trip.", ""),
        (5, 5, 5, 5, 5, 5, 5, 5, "Yes",
         "The dog was made as welcome as we were.", ""),
        (6, 3, 2, 4, 4, 5, 4, 3, "Yes",
         "The floors had not been vacuumed and there was no "
         "coffee. Everything else was lovely, which made it more "
         "annoying.",
         "Replaced the vacuum, added a coffee minimum to the "
         "inventory, and rang Suzanna to apologise."),
        (7, 4, 4, 5, 4, 5, 5, 4, "Yes",
         "Great space for a wedding party. Bathroom is tired.",
         "Booked the re-seal and re-grout the same week."),
        (8, 5, 5, 5, 5, 5, 5, 5, "Yes", "Nothing but five stars.",
         ""),
        (9, 5, 5, 5, 5, 5, 5, 5, "Yes",
         "The welcome basket made our anniversary.", ""),
        (11, 5, 5, 5, 5, 5, 4, 5, "Yes",
         "Second stay and just as good.", ""),
        (12, 5, 5, 5, 5, 5, 5, 5, "No",
         "Booked direct, which was easy. Lovely house.", ""),
        (14, 4, 3, 5, 5, 5, 5, 4, "Yes",
         "Comfortable but not quite as clean as the photos "
         "suggest.",
         "Same week as the broken vacuum. Two reviews, one "
         "cause."),
        (18, 5, 5, 5, 5, 5, 5, 5, "Yes", "Faultless.", ""),
        (19, 4, 5, 4, 4, 4, 5, 4, "No",
         "Good stay. Check-in instructions arrived late.",
         "Moved the access message to two days before arrival."),
        (22, 5, 5, 5, 5, 5, 5, 5, "Yes", "Home from home.", ""),
        (23, 4, 4, 4, 5, 5, 5, 4, "Yes", "Very good value.", ""),
        (24, 5, 5, 5, 5, 5, 5, 4, "Yes",
         "Third stay. Says it all.", ""),
        (26, 5, 5, 5, 5, 5, 4, 5, "Yes",
         "Six weeks and it felt like a home.", ""),
        (27, 5, 5, 5, 5, 5, 5, 5, "Yes", "The dog loved it too.",
         ""),
        (28, 4, 4, 4, 4, 5, 5, 4, "No", "Comfortable and quiet.",
         ""),
        (29, 5, 5, 5, 5, 5, 5, 5, "Yes", "Beautiful spot.", ""),
        (30, 5, 5, 5, 5, 5, 5, 5, "Yes",
         "Booking direct was simple and the house was perfect.",
         ""),
    ]
    rows: Rows = []
    for (number, overall, clean, accuracy, checkin, comm, location,
         value, responded, public, action) in entries:
        row: Dict[str, Any] = {
            "booking": f"BKG-{number:04d}",
            "requested": _day(cfg, -70),
            "status": "Received",
            "review_date": _day(cfg, -68),
            "rate_overall": overall, "rate_clean": clean,
            "rate_accuracy": accuracy, "rate_checkin": checkin,
            "rate_comm": comm, "rate_location": location,
            "rate_value": value,
            "public_review": public,
            "responded": responded,
        }
        if responded == "Yes":
            row["response_date"] = _day(cfg, -66)
            row["response"] = ("Thank you - we are glad you enjoyed "
                               "it." if overall >= 4 else
                               "You are right and I am sorry. Here "
                               "is what we have changed.")
        if action:
            row["action_taken"] = action
        rows.append(row)

    # asked for but not yet left - the chase case
    rows.append({
        "booking": "BKG-0025", "requested": _day(cfg, -19),
        "status": "Requested", "responded": "No",
        "public_review": "",
        "action_taken": "One-night stay, no review. The window is "
                        "closing.",
    })
    rows.append({
        "booking": "BKG-0032", "requested": _day(cfg, -1),
        "status": "Requested", "responded": "No",
        "public_review": "",
    })
    rows.append({
        "booking": "BKG-0013", "status": "No review left",
        "responded": "No", "public_review": "",
        "action_taken": "Guest never reviews. Marked so it stops "
                        "being chased.",
    })
    return rows


# ---------------------------------------------------------------------------
# 12. Occupancy & Performance
# ---------------------------------------------------------------------------

def performance(cfg: Config) -> Rows:
    rows: Rows = []
    notes = {
        (1, 3): "Boiler failure blocked six nights and cost a "
                "partial refund. The month the cabin lost money.",
        (0, 0): "Strongest month so far. Direct bookings doing the "
                "heavy lifting.",
        (2, 2): "Two one-night stays. Cleaning ate the margin.",
    }
    for back in range(4):
        first = cfg.today.replace(day=1)
        month = first.month - back
        year = first.year
        while month <= 0:
            month += 12
            year -= 1
        period = _dt.date(year, month, 1)
        for index in range(4):
            row: Dict[str, Any] = {
                "property": f"PRP-{index + 1:02d}",
                "period": period,
            }
            note = notes.get((back, index + 1))
            if note:
                row["notes"] = note
            rows.append(row)
    return rows


# ---------------------------------------------------------------------------

_BUILDERS: Dict[str, Callable[[Config], Rows]] = {
    "properties": properties,
    "bookings": bookings,
    "guests": guests,
    "arrivals": arrivals,
    "cleaning": cleaning,
    "payouts": payouts,
    "expenses": expenses,
    "supplies": supplies,
    "maintenance": maintenance,
    "reviews": reviews,
    "vendors": vendors,
    "performance": performance,
}


def get(name: str, cfg: Config) -> Rows:
    """Sample rows for one register, or `[]` when there are none."""
    builder = _BUILDERS.get(name)
    if builder is None:
        return []
    return builder(cfg)
````

---

## 3. Properties and bookings

The two registers everything else joins to. A property carries its own rate, fees, times and ownership basis, and collects twenty-odd calculated columns back from the rest of the workbook - occupancy, ADR, RevPAR, profit, margin and its rank in the portfolio. A booking turns two dates into nights, looks up the nightly rate, applies the platform's commission from the editable fee table, and flags a double booking by reading every other row on the same property. The shared helpers in `common.py` are where the alert colours, the ID lookups and the never-blank numeric mirror columns live.

### `airbnb_rental_spreadsheet/sheets/__init__.py`

*0 lines*

````python

````

### `airbnb_rental_spreadsheet/sheets/common.py`

*483 lines · 46 functions*

> Small building blocks shared by the register sheets.

| Defines | What it does |
|---------|--------------|
| `id_col(trigger, prefix, digits, width, header, note)` | Auto numbered ID column (blank until the trigger column is filled) |
| `month_col(date_key, key, header, width, hidden)` | `yyyy-mm` text column - the key every monthly roll-up groups on |
| `year_col(date_key, key, header, width)` |  |
| `weekday_col(date_key, key, header, width, hidden)` |  |
| `week_start_col(date_key, key, header, width)` | The Monday of that week - a far safer grouping key than WEEKNUM |
| `date_num_col(date_key, key, header, width)` | Hidden numeric mirror of a date column that is **never blank** |
| `seq_num_col(trigger, key, header)` | Hidden never-blank row counter, so INDEX/MATCH can find a row |
| `alert_cf()` | Traffic light colouring for the 'Action' / 'Check' columns |
| `status_cf()` | Colouring for booking, cleaning, payout and maintenance statuses |
| `priority_cf()` |  |
| `bar_cf(color)` |  |
| `scale_cf()` |  |
| `money_cf()` |  |
| `loss_cf()` | Red when a profit column goes negative |
| `owing_cf()` | Amber on anything still owed, blue on an overpayment |
| `flag_yes_no_cf()` |  |
| `flag_no_bad_cf()` | For the ones where 'No' is a real problem, not a preference |
| `target_cf(target_name)` | Green at or above a named target, amber below it |
| `rating_cf()` |  |
| `yesno(note)` |  |
| `score_validation(maximum)` |  |
| `blank(trigger, body)` | `=IF(trigger="","",body)` with the standard row placeholders |
| `zero(trigger, body)` | `=IF(trigger="",0,body)` - for columns that must never return "" |
| `lookup(sheet, target, key_col, key, missing)` | Generic `INDEX/MATCH` against any register, blank row safe |
| `lookup_num(sheet, target, key_col, key)` | Numeric `INDEX/MATCH` - 0 when the ID is missing or the cell empty |
| `lookup_num_zero(sheet, target, key_col, key)` | As :func:`lookup_num`, but returns 0 rather than "" on a blank row |
| `lookup_property(target, key, missing)` |  |
| `lookup_property_num(target, key)` |  |
| `lookup_booking(target, key, missing)` |  |
| `lookup_booking_num(target, key)` |  |
| `lookup_guest(target, key, missing)` |  |
| `lookup_vendor(target, key, missing)` |  |
| `lookup_vendor_num(target, key)` |  |
| `lookup_supply(target, key, missing)` |  |
| `in_list(cell, values)` | `TRUE` when `cell` matches any of `values` |
| `is_holding(cell)` | The booking holds the dates on the calendar (booked or blocked) |
| `is_earning(cell)` | The booking earns money (confirmed, checked in or completed) |
| `is_block(cell)` | An owner or maintenance block - takes the night off the market |
| `is_capital(cell)` | A capital purchase rather than a running cost |
| `platform_rate(column, platform_cell)` | The configured commission or processing rate for a platform |
| `nights_between(check_in, check_out)` | Nights of a stay: the arrival night counts, the departure one does not |
| `rank_desc(key, guard, header, width, note)` | Hidden dense rank, 1 = the biggest value in `key` |
| `rank_asc(key, guard, header, width, note)` | Hidden dense rank, 1 = the smallest value in `key` |
| `seq_key(parent, trigger)` | Hidden `1, 2, 3 ...` within a parent, plus a `parent|n` key |
| `lodging_tax(amount)` | Lodging / occupancy tax on an accommodation total |
| `monthly_equivalent(amount, frequency_cell)` | Turn any recurring cost into what it really costs each month |

**Module constants:** `ROW_OFFSET` (HEADER_ROW).

````python
"""Small building blocks shared by the register sheets."""

from __future__ import annotations

from typing import Any, Dict, List

from ..config import (BLOCK_STATUSES, CAPITAL_CATEGORIES, EARNING_STATUSES,
                      HOLDING_STATUSES)
from ..layout import HEADER_ROW, Col
from ..writer import cf_text, yes_no

# Everything below assumes the standard geometry: headers on row 4, data from
# row 5 - ROW()-4 therefore gives a 1 based record number.
ROW_OFFSET = HEADER_ROW


# ---------------------------------------------------------------------------
# Standard columns
# ---------------------------------------------------------------------------

def id_col(trigger: str, prefix: str, digits: int = 4, width: float = 11,
           header: str = "ID", note: str = "") -> Col:
    """Auto numbered ID column (blank until the trigger column is filled).

    `prefix` is a formula fragment - either a defined name such as
    ``BookingPrefix`` or a quoted literal such as ``'"BKG-"'``.
    """
    zeros = "0" * digits
    return Col(
        "id", header, width, "id",
        formula=(f'=IF({{{trigger}}}{{r}}="","",'
                 f'{prefix}&TEXT(ROW()-{ROW_OFFSET},"{zeros}"))'),
        note=note or "Generated automatically from the row number. "
                     "Filter the list rather than re-sorting it so the IDs "
                     "keep matching your other tabs.")


def month_col(date_key: str, key: str = "month", header: str = "Month",
              width: float = 10, hidden: bool = True) -> Col:
    """`yyyy-mm` text column - the key every monthly roll-up groups on."""
    return Col(key, header, width, "text",
               formula=(f'=IF({{{date_key}}}{{r}}="","",'
                        f'TEXT({{{date_key}}}{{r}},"yyyy-mm"))'),
               hidden=hidden,
               note="Automatic. Every monthly report groups on this.")


def year_col(date_key: str, key: str = "year", header: str = "Year",
             width: float = 8) -> Col:
    return Col(key, header, width, "int",
               formula=(f'=IF({{{date_key}}}{{r}}="","",'
                        f'YEAR({{{date_key}}}{{r}}))'),
               hidden=True)


def weekday_col(date_key: str, key: str = "weekday", header: str = "Day",
                width: float = 8, hidden: bool = False) -> Col:
    return Col(key, header, width, "text",
               formula=(f'=IF({{{date_key}}}{{r}}="","",'
                        f'TEXT({{{date_key}}}{{r}},"ddd"))'),
               hidden=hidden)


def week_start_col(date_key: str, key: str = "week_start",
                   header: str = "Week beginning", width: float = 12) -> Col:
    """The Monday of that week - a far safer grouping key than WEEKNUM."""
    return Col(key, header, width, "date",
               formula=(f'=IF({{{date_key}}}{{r}}="","",'
                        f'{{{date_key}}}{{r}}-WEEKDAY({{{date_key}}}{{r}},3))'),
               hidden=True,
               note="The Monday of that week. Weekly totals group on it.")


def date_num_col(date_key: str, key: str, header: str = "",
                 width: float = 10) -> Col:
    """Hidden numeric mirror of a date column that is **never blank**.

    A blank cell inside a `SUMPRODUCT(cond * range)` is fine, but a
    formula column that returns `""` is not: Excel answers `#VALUE!` the
    moment it multiplies a zero by an empty string.  The calendar, the
    double-booking test and the availability strip all multiply date
    columns, so each one gets a mirror that returns a real 0 instead.
    """
    return Col(key, header or f"{date_key} (number)", width, "number",
               formula=(f'=IF({{{date_key}}}{{r}}="",0,'
                        f'{{{date_key}}}{{r}}*1)'),
               hidden=True,
               note="Automatic. The calendar and the clash checks read this "
                    "instead of the date column, because a blank date would "
                    "break the arithmetic.")


def seq_num_col(trigger: str, key: str = "seq_num",
                header: str = "Row number") -> Col:
    """Hidden never-blank row counter, so INDEX/MATCH can find a row."""
    return Col(key, header, 9, "int",
               formula=(f'=IF({{{trigger}}}{{r}}="",0,ROW()-{ROW_OFFSET})'),
               hidden=True,
               note="Automatic. The calendar uses this to find which "
                    "booking covers a date.")


# ---------------------------------------------------------------------------
# Conditional formatting sets
# ---------------------------------------------------------------------------

def alert_cf() -> List[Dict[str, Any]]:
    """Traffic light colouring for the 'Action' / 'Check' columns."""
    return [
        cf_text("OK", "good", "begins with"),
        cf_text("Ready", "good"),
        cf_text("Double booked", "bad"),
        cf_text("Clash", "bad"),
        cf_text("Overdue", "bad"),
        cf_text("Late", "bad"),
        cf_text("Missing", "bad"),
        cf_text("Not cleaned", "bad"),
        cf_text("Urgent", "bad"),
        cf_text("Out of stock", "bad"),
        cf_text("Loss", "bad"),
        cf_text("Unpaid", "bad"),
        cf_text("Cancelled", "bad"),
        cf_text("Chase", "warn"),
        cf_text("Check", "warn"),
        cf_text("Action", "warn"),
        cf_text("Order", "warn"),
        cf_text("Low stock", "warn"),
        cf_text("Add", "warn"),
        cf_text("Send", "warn"),
        cf_text("Confirm", "warn"),
        cf_text("Follow up", "warn"),
        cf_text("Schedule", "warn"),
        cf_text("Same day", "warn"),
        cf_text("Tight turn", "warn"),
        cf_text("Due", "warn"),
        cf_text("Today", "info"),
        cf_text("Tomorrow", "info"),
        cf_text("Arriving", "info"),
        cf_text("Leaving", "info"),
        cf_text("Upcoming", "info"),
        cf_text("Staying", "info"),
        cf_text("Blocked", "muted"),
    ]


def status_cf() -> List[Dict[str, Any]]:
    """Colouring for booking, cleaning, payout and maintenance statuses."""
    return [
        cf_text("Active", "good"),
        cf_text("Completed", "good"),
        cf_text("Complete", "good"),
        cf_text("Confirmed", "good"),
        cf_text("Checked in", "good"),
        cf_text("Received", "good"),
        cf_text("Paid in full", "good"),
        cf_text("In stock", "good"),
        cf_text("Ready", "good"),
        cf_text("Cancelled", "bad"),
        cf_text("Out of stock", "bad"),
        cf_text("Disputed", "bad"),
        cf_text("Written off", "bad"),
        cf_text("Do not use", "bad"),
        cf_text("Late", "bad"),
        cf_text("Urgent", "bad"),
        cf_text("Order now", "bad"),
        cf_text("Inquiry", "warn"),
        cf_text("Scheduled", "warn"),
        cf_text("In progress", "warn"),
        cf_text("Part paid", "warn"),
        cf_text("Part received", "warn"),
        cf_text("Awaiting payout", "warn"),
        cf_text("Expected", "warn"),
        cf_text("Reported", "warn"),
        cf_text("Quoted", "warn"),
        cf_text("Waiting", "warn"),
        cf_text("Low", "warn"),
        cf_text("Requested", "warn"),
        cf_text("Trial", "warn"),
        cf_text("Under maintenance", "warn"),
        cf_text("Owner block", "info"),
        cf_text("Maintenance block", "info"),
        cf_text("Not due yet", "muted"),
        cf_text("Not scheduled", "muted"),
        cf_text("Not requested", "muted"),
        cf_text("Not tracked", "muted"),
        cf_text("Inactive", "muted"),
        cf_text("Ended", "muted"),
        cf_text("Skipped", "muted"),
    ]


def priority_cf() -> List[Dict[str, Any]]:
    return [
        cf_text("Urgent", "bold_bad"),
        cf_text("High", "bad"),
        cf_text("Normal", "info"),
        cf_text("Low", "muted"),
        cf_text("When convenient", "muted"),
    ]


def bar_cf(color: str = "#3F97A0") -> List[Dict[str, Any]]:
    return [{"type": "data_bar", "bar_color": color}]


def scale_cf() -> List[Dict[str, Any]]:
    return [{"type": "3_color_scale", "min_color": "#F9DAD6",
             "mid_color": "#FCEBD4", "max_color": "#D8EDDD"}]


def money_cf() -> List[Dict[str, Any]]:
    return [{"type": "cell", "criteria": ">", "value": 0, "fmt": "bold_good"}]


def loss_cf() -> List[Dict[str, Any]]:
    """Red when a profit column goes negative."""
    return [{"type": "cell", "criteria": "<", "value": 0, "fmt": "bold_bad"},
            {"type": "cell", "criteria": ">", "value": 0, "fmt": "good"}]


def owing_cf() -> List[Dict[str, Any]]:
    """Amber on anything still owed, blue on an overpayment."""
    return [{"type": "cell", "criteria": ">", "value": 0.004, "fmt": "warn"},
            {"type": "cell", "criteria": "<", "value": -0.004, "fmt": "info"}]


def flag_yes_no_cf() -> List[Dict[str, Any]]:
    return [cf_text("No", "warn"), cf_text("Yes", "good")]


def flag_no_bad_cf() -> List[Dict[str, Any]]:
    """For the ones where 'No' is a real problem, not a preference."""
    return [cf_text("No", "bad"), cf_text("Yes", "good")]


def target_cf(target_name: str) -> List[Dict[str, Any]]:
    """Green at or above a named target, amber below it."""
    return [
        {"type": "cell", "criteria": ">=", "value": f"={target_name}",
         "fmt": "good"},
        {"type": "cell", "criteria": "<", "value": f"={target_name}",
         "fmt": "warn"},
    ]


def rating_cf() -> List[Dict[str, Any]]:
    return [
        {"type": "cell", "criteria": ">=", "value": 4.8, "fmt": "bold_good"},
        {"type": "cell", "criteria": ">=", "value": 4.5, "fmt": "good"},
        {"type": "cell", "criteria": ">=", "value": 4.0, "fmt": "warn"},
        {"type": "cell", "criteria": "<", "value": 4.0, "fmt": "bad"},
    ]


def yesno(note: str = "") -> Dict[str, Any]:
    return yes_no(note)


def score_validation(maximum: int = 5) -> Dict[str, Any]:
    return {
        "validate": "decimal",
        "criteria": "between",
        "minimum": 1,
        "maximum": maximum,
        "error_title": f"Rating 1-{maximum}",
        "error_message": f"Enter a rating from 1 (poor) to {maximum} "
                         f"(excellent). Half stars are fine.",
    }


def blank(trigger: str, body: str) -> str:
    """`=IF(trigger="","",body)` with the standard row placeholders."""
    return f'=IF({{{trigger}}}{{r}}="","",{body})'


def zero(trigger: str, body: str) -> str:
    """`=IF(trigger="",0,body)` - for columns that must never return "".

    Any column a `SUMPRODUCT`, `MAX(cond*range)` or calendar lookup reads
    has to be numeric on every single row, including the ones that do not
    apply, or Excel multiplies a number by an empty string and answers
    `#VALUE!`.
    """
    return f'=IF({{{trigger}}}{{r}}="",0,{body})'


# ---------------------------------------------------------------------------
# Cross sheet lookups.  Every one is blank-row safe and says what to do
# when the ID does not exist, because a typo in an ID is the single most
# common thing that goes wrong in a workbook like this.
# ---------------------------------------------------------------------------

def lookup(sheet: str, target: str, key_col: str, key: str,
           missing: str = "") -> str:
    """Generic `INDEX/MATCH` against any register, blank row safe."""
    return (f'=IF({{{key}}}{{r}}="","",IFERROR(INDEX([[{sheet}.'
            f'{target}]],MATCH({{{key}}}{{r}},'
            f'[[{sheet}.{key_col}]],0)),"{missing}"))')


def lookup_num(sheet: str, target: str, key_col: str, key: str) -> str:
    """Numeric `INDEX/MATCH` - 0 when the ID is missing or the cell empty.

    `INDEX` on an empty cell returns 0, not "", so numeric lookups are
    coerced with `*1` inside the `IFERROR` rather than compared to "".
    """
    return (f'=IF({{{key}}}{{r}}="","",IFERROR(INDEX([[{sheet}.'
            f'{target}]],MATCH({{{key}}}{{r}},'
            f'[[{sheet}.{key_col}]],0))*1,0))')


def lookup_num_zero(sheet: str, target: str, key_col: str, key: str) -> str:
    """As :func:`lookup_num`, but returns 0 rather than "" on a blank row."""
    return (f'=IF({{{key}}}{{r}}="",0,IFERROR(INDEX([[{sheet}.'
            f'{target}]],MATCH({{{key}}}{{r}},'
            f'[[{sheet}.{key_col}]],0))*1,0))')


def lookup_property(target: str, key: str = "property",
                    missing: str = "Not in Property Database") -> str:
    return lookup("Property Database", target, "id", key, missing)


def lookup_property_num(target: str, key: str = "property") -> str:
    return lookup_num("Property Database", target, "id", key)


def lookup_booking(target: str, key: str = "booking",
                   missing: str = "Not in Booking Tracker") -> str:
    return lookup("Booking Tracker", target, "id", key, missing)


def lookup_booking_num(target: str, key: str = "booking") -> str:
    return lookup_num("Booking Tracker", target, "id", key)


def lookup_guest(target: str, key: str = "guest",
                 missing: str = "Not in Guest Database") -> str:
    return lookup("Guest Database", target, "id", key, missing)


def lookup_vendor(target: str, key: str = "vendor",
                  missing: str = "Not in Staff & Vendors") -> str:
    return lookup("Staff & Vendors", target, "id", key, missing)


def lookup_vendor_num(target: str, key: str = "vendor") -> str:
    return lookup_num("Staff & Vendors", target, "id", key)


def lookup_supply(target: str, key: str = "supply",
                  missing: str = "Not in Supplies Inventory") -> str:
    return lookup("Supplies Inventory", target, "id", key, missing)


# ---------------------------------------------------------------------------
# Shared formula fragments
# ---------------------------------------------------------------------------

def in_list(cell: str, values: List[str]) -> str:
    """`TRUE` when `cell` matches any of `values`.

    `OR()` evaluates every argument whatever happens, which is how a
    perfectly good formula ends up returning `#VALUE!`, so membership is
    tested by adding the comparisons instead.
    """
    tests = "+".join(f'({cell}="{v}")' for v in values)
    return f"({tests}>0)"


def is_holding(cell: str) -> str:
    """The booking holds the dates on the calendar (booked or blocked)."""
    return in_list(cell, HOLDING_STATUSES)


def is_earning(cell: str) -> str:
    """The booking earns money (confirmed, checked in or completed)."""
    return in_list(cell, EARNING_STATUSES)


def is_block(cell: str) -> str:
    """An owner or maintenance block - takes the night off the market."""
    return in_list(cell, BLOCK_STATUSES)


def is_capital(cell: str) -> str:
    """A capital purchase rather than a running cost."""
    return in_list(cell, CAPITAL_CATEGORIES)


def platform_rate(column: str, platform_cell: str) -> str:
    """The configured commission or processing rate for a platform.

    `column` is one of the Settings range names - `PlatformCommission`
    or `PlatformPayment`.  An unknown platform costs nothing rather than
    breaking the row, because a host who types a new platform name should
    get a number they can correct, not an error.
    """
    return (f'IFERROR(INDEX({column},MATCH({platform_cell},'
            f'PlatformList,0))*1,0)')


def nights_between(check_in: str, check_out: str) -> str:
    """Nights of a stay: the arrival night counts, the departure one does not."""
    return (f'IF({check_in}="","",IF({check_out}="","",'
            f'MAX(0,{check_out}-{check_in})))')


def rank_desc(key: str, guard: str = '{name}{r}=""',
              header: str = "Rank", width: float = 9,
              note: str = "") -> Col:
    """Hidden dense rank, 1 = the biggest value in `key`.

    Ties are broken by row order, so the ranks come out 1, 2, 3 ... with
    no gaps and no duplicates.  Pages then read the nth row with an
    ordinary `MATCH(n, rank, 0)` - no array formula, so nothing in the
    workbook needs Ctrl+Shift+Enter.
    """
    return Col(f"{key}_rank", header, width, "int",
               formula=(f'=IF({guard},"",COUNTIFS({{{key}_first}}:'
                        f'{{{key}_last}},">"&{{{key}}}{{r}})+'
                        f'COUNTIFS({{{key}_first}}:{{{key}}}{{r}},'
                        f'{{{key}}}{{r}}))'),
               hidden=True, note=note or f"1 = the largest {key}.")


def rank_asc(key: str, guard: str = '{name}{r}=""',
             header: str = "Rank", width: float = 9,
             note: str = "") -> Col:
    """Hidden dense rank, 1 = the smallest value in `key`."""
    return Col(f"{key}_rank", header, width, "int",
               formula=(f'=IF({guard},"",COUNTIFS({{{key}_first}}:'
                        f'{{{key}_last}},"<"&{{{key}}}{{r}})+'
                        f'COUNTIFS({{{key}_first}}:{{{key}}}{{r}},'
                        f'{{{key}}}{{r}}))'),
               hidden=True, note=note or f"1 = the smallest {key}.")


def seq_key(parent: str, trigger: str = "") -> List[Col]:
    """Hidden `1, 2, 3 ...` within a parent, plus a `parent|n` key.

    Printables and per-parent summaries use `MATCH(parent & "|" & n, key, 0)`
    to pull the nth line of one record without an array formula.
    """
    trig = trigger or parent
    return [
        Col("seq", "Line", 6, "int",
            formula=blank(
                trig,
                f'COUNTIFS({{{parent}_first}}:{{{parent}}}{{r}},'
                f'{{{parent}}}{{r}})'),
            hidden=True),
        Col("key", "Key", 14, "text",
            formula=blank(trig, f'{{{parent}}}{{r}}&"|"&{{seq}}{{r}}'),
            hidden=True),
    ]


# ---------------------------------------------------------------------------
# Money rules used in more than one place
# ---------------------------------------------------------------------------

def lodging_tax(amount: str) -> str:
    """Lodging / occupancy tax on an accommodation total.

    Most platforms collect and remit this for the host, which is why the
    switch is on Settings and the default is 'the platform does it'.
    """
    return (f'IF(TaxCollectedByPlatform="Yes",0,'
            f'ROUND({amount}*TaxRate,2))')


def monthly_equivalent(amount: str, frequency_cell: str) -> str:
    """Turn any recurring cost into what it really costs each month.

    A yearly insurance premium and a weekly laundry bill are not
    comparable until they are both expressed per month, and a host
    looking at an expense breakdown is always asking the per-month
    question.
    """
    return (f'IF({frequency_cell}="","",'
            f'IFERROR(INDEX(FrequencyPerYear,MATCH({frequency_cell},'
            f'FrequencyList,0))*{amount}/12,0))')
````

### `airbnb_rental_spreadsheet/sheets/properties.py`

*489 lines · 2 functions*

> `Property Database` - one row per place you let, and what it earns.
>
> This is the spine of the workbook.  Every other tab points at a property
> by its ID, so a portfolio of one studio and a portfolio of twelve
> apartments behave exactly the same way: add a row here, and the property
> appears in every drop-down, every calendar column, every per-property
> total and every comparison table in the file.
>
> The commercial side of a property lives here too - the standard nightly
> rate, the cleaning fee you charge, the cleaning cost you pay, the
> mortgage or the rent if you are letting someone else's building - so a
> blank booking row can price itself from the property and a host can see
> which of their places actually makes money.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Portfolio totals under the register |

**Module constants:** `NAME` ('Property Database').

````python
"""`Property Database` - one row per place you let, and what it earns.

This is the spine of the workbook.  Every other tab points at a property
by its ID, so a portfolio of one studio and a portfolio of twelve
apartments behave exactly the same way: add a row here, and the property
appears in every drop-down, every calendar column, every per-property
total and every comparison table in the file.

The commercial side of a property lives here too - the standard nightly
rate, the cleaning fee you charge, the cleaning cost you pay, the
mortgage or the rent if you are letting someone else's building - so a
blank booking row can price itself from the property and a host can see
which of their places actually makes money.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check, time_check
from .common import (alert_cf, bar_cf, blank, flag_yes_no_cf, id_col, 
                    loss_cf, rating_cf, status_cf, target_cf)

NAME = "Property Database"


def spec(cfg: Config) -> SheetSpec:
    bk_prop = "[[Booking Tracker.property]]"
    bk_earns = "[[Booking Tracker.earns]]"
    bk_holds = "[[Booking Tracker.holds]]"
    bk_nights = "[[Booking Tracker.nights]]"
    bk_accom = "[[Booking Tracker.accommodation]]"
    bk_gross = "[[Booking Tracker.gross]]"
    bk_fees = "[[Booking Tracker.platform_fee_total]]"
    bk_net = "[[Booking Tracker.net_revenue]]"
    bk_guests = "[[Booking Tracker.guests_total]]"
    bk_block = "[[Booking Tracker.is_block]]"
    bk_cancel = "[[Booking Tracker.is_cancelled]]"
    ex_prop = "[[Expense Tracker.property]]"
    ex_pl = "[[Expense Tracker.pl_amount]]"
    cl_prop = "[[Cleaning Schedule.property]]"
    cl_cost = "[[Cleaning Schedule.total_cost]]"
    mt_prop = "[[Maintenance Log.property]]"
    mt_cost = "[[Maintenance Log.cost_used]]"
    mt_open = "[[Maintenance Log.is_open]]"
    rv_prop = "[[Reviews Tracker.property]]"
    rv_score = "[[Reviews Tracker.rate_overall]]"
    rv_has = "[[Reviews Tracker.has_score]]"
    sp_prop = "[[Supplies Inventory.property]]"
    sp_low = "[[Supplies Inventory.needs_order]]"

    def by_property(value: str, prop_range: str, *extra: str) -> str:
        """SUMIFS of a value column for this property."""
        tail = "".join(f",{c}" for c in extra)
        return f'SUMIFS({value},{prop_range},{{id}}{{r}}{tail})'

    columns = [
        # ------------------------------------------------------ the place --
        id_col("name", "PropertyPrefix", digits=cfg.digits("property"),
               width=10, header="Property ID",
               note="Automatic. This is the code every other tab uses to "
                    "point at this property, so keep the rows in place "
                    "and filter rather than re-sort."),
        Col("name", "Property name", 24, "text",
            note="What you call it day to day - 'Harbour Loft', "
                 "'Pine Cabin'. Short names read better in the calendar "
                 "and the reports.",
            group="The place"),
        Col("nickname", "Short code", 9, "text",
            note="Two or three letters for the calendar strip and the "
                 "tight comparison tables. 'HL', 'PC'.",
            group="The place"),
        Col("address", "Address", 32, "text",
            note="Street address. Used on the printable turnover sheet "
                 "so a cleaner knows where they are going.",
            group="The place"),
        Col("city", "City / area", 16, "text", group="The place"),
        Col("country", "Country", 14, "text",
            validate=list_from("CountryList"), group="The place"),
        Col("type", "Property type", 16, "text",
            validate=list_from("PropertyTypeList"), group="The place"),
        Col("bedrooms", "Bedrooms", 9, "int",
            validate=number_check(0, 30, "Bedrooms"), group="The place"),
        Col("bathrooms", "Bathrooms", 9, "number1",
            validate=number_check(0, 30, "Bathrooms"),
            note="Half a bathroom is fine - type 1.5.",
            group="The place"),
        Col("beds", "Beds", 7, "int",
            validate=number_check(0, 60, "Beds"), group="The place"),
        Col("max_guests", "Sleeps", 8, "int",
            validate=number_check(1, 60, "Maximum guests"),
            note="Maximum guests allowed. The booking tab warns you when "
                 "a reservation goes over it.",
            group="The place"),
        Col("status", "Status", 15, "status",
            validate=list_from("PropertyStatusList"),
            cf=status_cf(),
            note="Only Active properties are counted in the portfolio "
                 "occupancy and the dashboard targets.",
            group="The place"),
        Col("listed_from", "Listed since", 12, "date",
            validate=date_check(),
            note="The date it first went live. Used to work out how many "
                 "nights it has genuinely been available this year.",
            group="The place"),

        # --------------------------------------------------- who runs it --
        Col("owner", "Owner", 18, "text", group="Who runs it"),
        Col("manager", "Managed by", 18, "text",
            note="You, a co-host, or an agency. Free text.",
            group="Who runs it"),
        Col("ownership", "Ownership", 20, "text",
            validate=list_from("OwnershipList"),
            note="Rental arbitrage means you pay rent on it and sublet "
                 "it nightly - the monthly rent below is then your "
                 "biggest cost.",
            group="Who runs it"),
        Col("main_platform", "Main platform", 15, "text",
            validate=list_from("PlatformList"), group="Who runs it"),
        Col("listing_url", "Listing link", 26, "link",
            note="Paste the public listing URL. Handy when you are "
                 "checking your own calendar against the platform.",
            group="Who runs it"),
        Col("licence", "Licence / permit", 16, "text",
            note="Short-term let registration number, if your city "
                 "requires one.",
            group="Who runs it"),

        # ------------------------------------------------- house settings --
        Col("check_in_time", "Check-in", 9, "time",
            validate=time_check(),
            note="The standard arrival time. Copied onto new "
                 "arrivals so you only type an exception.",
            group="House settings"),
        Col("check_out_time", "Check-out", 9, "time",
            validate=time_check(), group="House settings"),
        Col("min_nights", "Min nights", 10, "int",
            validate=number_check(1, 365, "Minimum nights"),
            group="House settings"),
        Col("wifi_name", "Wi-Fi network", 16, "text",
            group="House settings"),
        Col("access", "Access method", 18, "text",
            note="Key box, smart lock, meet and greet. Printed on the "
                 "turnover sheet.",
            group="House settings"),

        # ----------------------------------------------------- the money --
        Col("nightly_rate", "Standard rate", 13, "money",
            validate=number_check(0, 100000, "Nightly rate"),
            note="Your usual nightly price. A booking with no rate typed "
                 "falls back to this, so seasonal pricing is an override "
                 "and not a chore.",
            group="Money"),
        Col("cleaning_fee", "Cleaning fee charged", 15, "money",
            validate=number_check(0, 100000, "Cleaning fee"),
            note="What the guest pays. Kept out of ADR and RevPAR, "
                 "because it is a cost recovery and not room revenue.",
            group="Money"),
        Col("cleaning_cost", "Cleaning cost paid", 15, "money",
            validate=number_check(0, 100000, "Cleaning cost"),
            note="What the turnover actually costs you. The default for "
                 "new cleaning jobs on this property.",
            group="Money"),
        Col("deposit", "Security deposit", 14, "money",
            validate=number_check(0, 1000000, "Deposit"), group="Money"),
        Col("monthly_rent", "Monthly rent", 13, "money",
            validate=number_check(0, 1000000, "Monthly rent"),
            note="Rental arbitrage only - what you pay the landlord. "
                 "Leave blank if you own it.",
            group="Money"),
        Col("monthly_mortgage", "Monthly mortgage", 14, "money",
            validate=number_check(0, 1000000, "Mortgage"),
            note="Optional. The full payment; the interest part belongs "
                 "in the expense tab if you track it for tax.",
            group="Money"),
        Col("purchase_price", "Purchase price", 14, "money0",
            validate=number_check(0, 100000000, "Purchase price"),
            note="Optional. Only used for the yield figure.",
            group="Money"),
        Col("fixed_monthly", "Fixed cost a month", 14, "money",
            formula=blank("name", "ROUND({monthly_rent}{r}+"
                                  "{monthly_mortgage}{r},2)"),
            note="Rent plus mortgage. This is what the property costs "
                 "you before a single guest arrives.",
            group="Money"),

        # -------------------------------------------------- what it does --
        Col("bookings", "Bookings", 10, "int",
            formula=blank("name", f'COUNTIFS({bk_prop},{{id}}{{r}},'
                                  f'{bk_earns},"Yes")'),
            note="Confirmed, checked in and completed. Inquiries, "
                 "cancellations and blocks are not counted.",
            group="Performance"),
        Col("nights_booked", "Nights booked", 12, "int",
            formula=blank("name", by_property(bk_nights, bk_prop,
                                              f'{bk_earns},"Yes"')),
            cf=bar_cf(), group="Performance"),
        Col("nights_blocked", "Nights blocked", 12, "int",
            formula=blank("name", by_property(bk_nights, bk_prop,
                                              f'{bk_block},"Yes"')),
            note="Owner stays and maintenance closures. These come off "
                 "the available nights, so blocking a week does not "
                 "damage your occupancy rate.",
            group="Performance"),
        Col("guests_hosted", "Guests hosted", 12, "int",
            formula=blank("name", by_property(bk_guests, bk_prop,
                                              f'{bk_earns},"Yes"')),
            group="Performance"),
        Col("cancellations", "Cancellations", 12, "int",
            formula=blank("name", f'COUNTIFS({bk_prop},{{id}}{{r}},'
                                  f'{bk_cancel},"Yes")'),
            group="Performance"),
        Col("nights_available", "Nights available", 13, "int",
            formula=blank(
                "name",
                'MAX(0,MIN(TodayDate,DATE(YEAR(TodayDate),12,31))-'
                'MAX(DATE(YEAR(TodayDate),1,1),'
                'IF({listed_from}{r}="",DATE(YEAR(TodayDate),1,1),'
                '{listed_from}{r}))+1-{nights_blocked}{r})'),
            note="This calendar year so far, from the day it was listed, "
                 "less the nights you blocked. This is the denominator "
                 "of the occupancy figure.",
            group="Performance"),
        Col("occupancy", "Occupancy", 10, "percent",
            formula=blank(
                "name",
                'IF({nights_available}{r}<=0,"",'
                '{nights_booked}{r}/{nights_available}{r})'),
            cf=target_cf("TargetOccupancy"),
            note="Booked nights divided by the nights the place was "
                 "actually on the market this year.",
            group="Performance"),
        Col("adr", "ADR", 11, "money",
            formula=blank(
                "name",
                'IF({nights_booked}{r}<=0,"",'
                'ROUND({accommodation}{r}/{nights_booked}{r},2))'),
            cf=target_cf("TargetADR"),
            note="Average daily rate: accommodation revenue per booked "
                 "night. Cleaning fees are deliberately excluded.",
            group="Performance"),
        Col("revpar", "RevPAR", 11, "money",
            formula=blank(
                "name",
                'IF({nights_available}{r}<=0,"",'
                'ROUND({accommodation}{r}/{nights_available}{r},2))'),
            cf=target_cf("TargetRevPAR"),
            note="Revenue per available night - occupancy and rate in "
                 "one number. The figure to compare two properties on.",
            group="Performance"),
        Col("alos", "Avg stay", 9, "number1",
            formula=blank(
                "name",
                'IF({bookings}{r}<=0,"",'
                'ROUND({nights_booked}{r}/{bookings}{r},1))'),
            note="Average length of stay, in nights.",
            group="Performance"),
        Col("rating", "Rating", 9, "number",
            formula=blank(
                "name",
                f'IF(COUNTIFS({rv_prop},{{id}}{{r}},{rv_has},"Yes")<=0,"",'
                f'ROUND(SUMIFS({rv_score},{rv_prop},{{id}}{{r}},'
                f'{rv_has},"Yes")/COUNTIFS({rv_prop},{{id}}{{r}},'
                f'{rv_has},"Yes"),2))'),
            cf=rating_cf(), group="Performance"),
        Col("reviews", "Reviews", 9, "int",
            formula=blank("name", f'COUNTIFS({rv_prop},{{id}}{{r}},'
                                  f'{rv_has},"Yes")'),
            hidden=True, group="Performance"),

        # ---------------------------------------------------- the result --
        Col("accommodation", "Room revenue", 13, "money",
            formula=blank("name", by_property(bk_accom, bk_prop,
                                              f'{bk_earns},"Yes"')),
            note="Nightly accommodation only - no cleaning fee, no extra "
                 "guest charges. ADR and RevPAR are built on this.",
            group="Result"),
        Col("gross", "Gross revenue", 13, "money",
            formula=blank("name", by_property(bk_gross, bk_prop,
                                              f'{bk_earns},"Yes"')),
            note="Everything the guest paid, before platform fees.",
            group="Result"),
        Col("platform_fees", "Platform fees", 12, "money",
            formula=blank("name", by_property(bk_fees, bk_prop,
                                              f'{bk_earns},"Yes"')),
            group="Result"),
        Col("net_revenue", "Net revenue", 13, "money",
            formula=blank("name", by_property(bk_net, bk_prop,
                                              f'{bk_earns},"Yes"')),
            note="What lands in your account after the platform takes "
                 "its cut.",
            group="Result"),
        Col("cleaning_spend", "Cleaning paid out", 13, "money",
            formula=blank("name", by_property(cl_cost, cl_prop)),
            group="Result"),
        Col("maintenance_spend", "Maintenance", 12, "money",
            formula=blank("name", by_property(mt_cost, mt_prop)),
            group="Result"),
        Col("other_expenses", "Other costs", 12, "money",
            formula=blank("name", by_property(ex_pl, ex_prop)),
            note="Everything on the expense tab for this property that "
                 "counts against profit. Cleaning typed in as an expense "
                 "is excluded automatically so it is never counted twice.",
            group="Result"),
        Col("total_costs", "Total costs", 13, "money",
            formula=blank("name",
                          "ROUND({cleaning_spend}{r}+"
                          "{maintenance_spend}{r}+{other_expenses}{r},2)"),
            group="Result"),
        Col("profit", "Operating profit", 14, "money",
            formula=blank("name",
                          "ROUND({net_revenue}{r}-{total_costs}{r},2)"),
            cf=loss_cf(),
            note="Net revenue less everything it cost to run. Rent and "
                 "mortgage are shown separately so an arbitrage host and "
                 "an owner can read the same column.",
            group="Result"),
        Col("margin", "Margin", 9, "percent",
            formula=blank("name",
                          'IF({net_revenue}{r}<=0,"",'
                          '{profit}{r}/{net_revenue}{r})'),
            cf=target_cf("TargetMargin"), group="Result"),
        Col("profit_night", "Profit a night", 12, "money",
            formula=blank("name",
                          'IF({nights_booked}{r}<=0,"",'
                          'ROUND({profit}{r}/{nights_booked}{r},2))'),
            group="Result"),
        Col("yield_pct", "Gross yield", 10, "percent",
            formula=blank("name",
                          'IF({purchase_price}{r}<=0,"",'
                          '{gross}{r}/{purchase_price}{r})'),
            note="Gross revenue against what you paid for the place. "
                 "Blank unless you filled the purchase price in.",
            hidden=True, group="Result"),
        Col("profit_rank", "Rank", 7, "int",
            formula=('=IF({name}{r}="","",COUNTIFS({profit_first}:'
                     '{profit_last},">"&{profit}{r})+'
                     'COUNTIFS({profit_first}:{profit}{r},{profit}{r}))'),
            note="1 = your best earning property.",
            group="Result"),

        # ------------------------------------------------- what to do now --
        Col("open_jobs", "Open jobs", 10, "int",
            formula=blank("name", f'COUNTIFS({mt_prop},{{id}}{{r}},'
                                  f'{mt_open},"Yes")'),
            note="Maintenance reported and not finished.",
            group="Right now"),
        Col("low_stock", "Low stock", 10, "int",
            formula=blank("name", f'COUNTIFS({sp_prop},{{id}}{{r}},'
                                  f'{sp_low},"Yes")'),
            group="Right now"),
        Col("guest_in", "Guest in tonight", 13, "flag",
            formula=blank(
                "name",
                f'IF(SUMPRODUCT(({bk_prop}={{id}}{{r}})*'
                f'({bk_holds}="Yes")*'
                f'([[Booking Tracker.ci_num]]<=TodayDate)*'
                f'([[Booking Tracker.co_num]]>TodayDate))>0,"Yes","No")'),
            cf=flag_yes_no_cf(),
            note="Is someone staying tonight? Reads the booking tab, so "
                 "it is right the moment you type a reservation.",
            group="Right now"),
        Col("next_arrival", "Next arrival", 12, "date",
            formula=blank(
                "name",
                f'IF(SUMPRODUCT(MIN(([[Booking Tracker.ci_num]]>TodayDate)*'
                f'({bk_prop}={{id}}{{r}})*({bk_holds}="Yes")*'
                f'[[Booking Tracker.ci_num]]+'
                f'(1-([[Booking Tracker.ci_num]]>TodayDate)*'
                f'({bk_prop}={{id}}{{r}})*({bk_holds}="Yes"))*2958465))'
                f'>=2958465,"",'
                f'SUMPRODUCT(MIN(([[Booking Tracker.ci_num]]>TodayDate)*'
                f'({bk_prop}={{id}}{{r}})*({bk_holds}="Yes")*'
                f'[[Booking Tracker.ci_num]]+'
                f'(1-([[Booking Tracker.ci_num]]>TodayDate)*'
                f'({bk_prop}={{id}}{{r}})*({bk_holds}="Yes"))*2958465)))'),
            note="The next check-in date after today.",
            group="Right now"),
        Col("alert", "Action", 26, "status",
            formula=blank(
                "name",
                'IF({status}{r}="Under maintenance","Check - '
                'marked under maintenance",'
                'IF({open_jobs}{r}>0,"Action - "&{open_jobs}{r}&'
                '" open job(s)",'
                'IF({low_stock}{r}>0,"Order - "&{low_stock}{r}&'
                '" supply line(s) low",'
                'IF({nightly_rate}{r}<=0,"Add - no standard nightly rate",'
                'IF({status}{r}<>"Active","",'
                'IF({nights_available}{r}<=0,"",'
                'IF({occupancy}{r}<TargetOccupancy*0.7,'
                '"Check - occupancy well under target",'
                'IF({margin}{r}<0,"Loss - costs are above net revenue",'
                '"OK"))))))))'),
            cf=alert_cf(), group="Right now"),
        Col("notes", "Notes", 34, "wrap",
            note="Anything that does not fit a column - parking, quirks, "
                 "the neighbour to warn before a party.",
            group="Right now"),
    ]

    return SheetSpec(
        name=NAME,
        title="Property Database",
        subtitle="Every place you let, what it costs to run and what it "
                 "actually earns",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=None,
        freeze_cols=2,
        samples=sample("properties", cfg),
        help_text="Add one row per property. The Property ID on the left "
                  "is what every other tab points at, so add a property "
                  "here first and it appears in every drop-down in the "
                  "workbook. White cells are yours to type in; the "
                  "shaded ones work themselves out.",
        extra_notes=[
            "Nights available counts this calendar year from the day you "
            "listed the property, less any nights you blocked - so "
            "blocking a fortnight for yourself does not dent the "
            "occupancy rate.",
            "Cleaning fees charged to guests are deliberately kept out "
            "of ADR and RevPAR. They are a cost recovery, not room "
            "revenue, and mixing them in flatters both figures.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Portfolio totals under the register."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index

    row = sheet.section(row, "The whole portfolio", last)
    row = sheet.table_head(row, 0, [
        "Properties", "Active", "Bedrooms", "Sleeps", "Bookings",
        "Nights booked", "Nights blocked", "Occupancy", "ADR", "RevPAR"],
        [11, 9, 10, 9, 10, 13, 13, 11, 11, 11])
    sheet.calc_row(row, 0, [
        '=COUNTIF([[Property Database.name]],"?*")',
        '=COUNTIFS([[Property Database.status]],"Active")',
        '=SUM([[Property Database.bedrooms]])',
        '=SUM([[Property Database.max_guests]])',
        '=SUM([[Property Database.bookings]])',
        '=SUM([[Property Database.nights_booked]])',
        '=SUM([[Property Database.nights_blocked]])',
        '=IF(SUM([[Property Database.nights_available]])<=0,"",'
        'SUM([[Property Database.nights_booked]])/'
        'SUM([[Property Database.nights_available]]))',
        '=IF(SUM([[Property Database.nights_booked]])<=0,"",'
        'ROUND(SUM([[Property Database.accommodation]])/'
        'SUM([[Property Database.nights_booked]]),2))',
        '=IF(SUM([[Property Database.nights_available]])<=0,"",'
        'ROUND(SUM([[Property Database.accommodation]])/'
        'SUM([[Property Database.nights_available]]),2))',
    ], ["int", "int", "int", "int", "int", "int", "int", "percent",
        "money", "money"])
    row += 2

    row = sheet.table_head(row, 0, [
        "Gross revenue", "Platform fees", "Net revenue", "Cleaning",
        "Maintenance", "Other costs", "Operating profit", "Margin",
        "Rent & mortgage", "After finance"],
        [13, 12, 13, 12, 12, 12, 14, 10, 14, 13])
    sheet.calc_row(row, 0, [
        '=ROUND(SUM([[Property Database.gross]]),2)',
        '=ROUND(SUM([[Property Database.platform_fees]]),2)',
        '=ROUND(SUM([[Property Database.net_revenue]]),2)',
        '=ROUND(SUM([[Property Database.cleaning_spend]]),2)',
        '=ROUND(SUM([[Property Database.maintenance_spend]]),2)',
        '=ROUND(SUM([[Property Database.other_expenses]]),2)',
        '=ROUND(SUM([[Property Database.profit]]),2)',
        '=IF(SUM([[Property Database.net_revenue]])<=0,"",'
        'SUM([[Property Database.profit]])/'
        'SUM([[Property Database.net_revenue]]))',
        '=ROUND(SUM([[Property Database.fixed_monthly]])*12,2)',
        '=ROUND(SUM([[Property Database.profit]])-'
        'SUM([[Property Database.fixed_monthly]])*12,2)',
    ], ["money", "money", "money", "money", "money", "money", "money",
        "percent", "money", "money"])
    row += 2

    sheet.note(row, "Rent and mortgage are shown as a yearly figure "
                    "(monthly x 12) and kept out of operating profit, so "
                    "the profit column compares a property you own with "
                    "one you rent on equal terms. 'After finance' is the "
                    "number an arbitrage host cares about.", last)
````

### `airbnb_rental_spreadsheet/sheets/bookings.py`

*527 lines · 2 functions*

> `Booking Tracker` - every reservation, block and enquiry, priced.
>
> One register holds the lot: confirmed stays, cancellations, enquiries
> that never happened, owner stays and maintenance closures.  A block is
> simply a reservation that earns nothing - it takes the dates off the
> calendar and comes out of the available nights, which is why blocking a
> fortnight for yourself does not wreck your occupancy figure.
>
> Three things are worked out here that a host would otherwise do in their
> head at eleven o'clock at night: the nights in a stay, what the platform
> keeps, and whether this reservation overlaps one you already took.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Booking totals and a platform breakdown under the register |

**Module constants:** `NAME` ('Booking Tracker').

````python
"""`Booking Tracker` - every reservation, block and enquiry, priced.

One register holds the lot: confirmed stays, cancellations, enquiries
that never happened, owner stays and maintenance closures.  A block is
simply a reservation that earns nothing - it takes the dates off the
calendar and comes out of the available nights, which is why blocking a
fortnight for yourself does not wreck your occupancy figure.

Three things are worked out here that a host would otherwise do in their
head at eleven o'clock at night: the nights in a stay, what the platform
keeps, and whether this reservation overlaps one you already took.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, date_num_col, flag_yes_no_cf,
                     id_col, is_block, is_earning, is_holding, lodging_tax,
                     lookup_property,
                     money_cf, month_col, platform_rate, seq_num_col,
                     status_cf, weekday_col, year_col, zero)

NAME = "Booking Tracker"


def spec(cfg: Config) -> SheetSpec:
    cl_booking = "[[Cleaning Schedule.booking]]"
    ar_booking = "[[Check-in & Check-out.booking]]"
    rv_booking = "[[Reviews Tracker.booking]]"

    columns = [
        # ------------------------------------------------- the booking --
        id_col("check_in", "BookingPrefix", digits=cfg.digits("booking"),
               width=11, header="Booking ID",
               note="Automatic, from the row number. Cleaning jobs, "
                    "arrivals, payouts and reviews all point back at "
                    "this reference."),
        Col("property", "Property", 11, "text",
            validate=list_from("PropertyList"),
            note="Pick the property ID. Add the property on the Property "
                 "Database tab first and it appears in this list.",
            group="The booking"),
        Col("property_name", "Property name", 20, "text",
            formula=lookup_property("name"), group="The booking"),
        Col("guest", "Guest", 11, "text",
            validate=list_from("GuestList"),
            note="Pick the guest ID from the Guest Database. Leave it "
                 "blank for a block - the Action column will remind you "
                 "if a real stay has no guest record.",
            group="The booking"),
        Col("guest_name", "Guest name", 20, "text",
            formula=('=IF({check_in}{r}="","",IF({guest}{r}="",'
                     'IF({is_block}{r}="Yes","(blocked)",""),'
                     'IFERROR(INDEX([[Guest Database.name]],'
                     'MATCH({guest}{r},[[Guest Database.id]],0)),'
                     '"Not in Guest Database")))'),
            group="The booking"),
        Col("platform", "Platform", 15, "text",
            validate=list_from("PlatformList"),
            note="Sets the commission and card fee for this booking. "
                 "Every rate is editable on Settings - platform fees "
                 "differ by country, account type and plan.",
            group="The booking"),
        Col("reference", "Platform ref", 14, "text",
            note="The confirmation code the platform gave you, e.g. "
                 "HM4XY29ABC. Useful when you reconcile a payout.",
            group="The booking"),
        Col("status", "Status", 16, "status",
            validate=list_from("BookingStatusList"),
            cf=status_cf(),
            note="Owner block and Maintenance block hold the dates "
                 "without earning anything. Cancelled frees them up.",
            group="The booking"),

        # ---------------------------------------------------- the dates --
        Col("booked_date", "Booked on", 12, "date",
            validate=date_check(),
            note="When the reservation came in - gives you the lead time.",
            group="Dates"),
        Col("check_in", "Check-in", 12, "date",
            validate=date_check(),
            note="Arrival date. Every row on this tab starts here: no "
                 "check-in date, no booking.",
            group="Dates"),
        Col("check_out", "Check-out", 12, "date",
            validate=date_check(), group="Dates"),
        weekday_col("check_in", "ci_day", "In", 6, hidden=True),
        weekday_col("check_out", "co_day", "Out", 6, hidden=True),
        Col("nights", "Nights", 8, "int",
            formula=('=IF({check_in}{r}="","",IF({check_out}{r}="","",'
                     'MAX(0,{check_out}{r}-{check_in}{r})))'),
            cf=bar_cf(),
            note="Worked out for you. The arrival night counts, the "
                 "departure night does not - which is how every "
                 "platform counts it too.",
            group="Dates"),
        date_num_col("check_in", "ci_num", "Check-in (number)"),
        date_num_col("check_out", "co_num", "Check-out (number)"),
        seq_num_col("check_in"),
        month_col("check_in"),
        year_col("check_in"),
        Col("season", "Season", 9, "text",
            formula=blank(
                "check_in",
                'IF(MONTH({check_in}{r})<=2,"Winter",'
                'IF(MONTH({check_in}{r})<=5,"Spring",'
                'IF(MONTH({check_in}{r})<=8,"Summer",'
                'IF(MONTH({check_in}{r})<=11,"Autumn","Winter"))))'),
            hidden=True, group="Dates"),
        Col("lead_time", "Lead time", 10, "int",
            formula=('=IF({check_in}{r}="","",IF({booked_date}{r}="","",'
                     'MAX(0,{check_in}{r}-{booked_date}{r})))'),
            note="Days between the booking coming in and the guest "
                 "arriving. A string of one-day lead times means your "
                 "price is too low.",
            hidden=True, group="Dates"),

        # --------------------------------------------------- the guests --
        Col("adults", "Adults", 8, "int",
            validate=number_check(0, 40, "Adults"), group="Guests"),
        Col("children", "Children", 9, "int",
            validate=number_check(0, 40, "Children"), group="Guests"),
        Col("pets", "Pets", 6, "int",
            validate=number_check(0, 20, "Pets"), group="Guests"),
        Col("guests_total", "Guests", 8, "int",
            formula=blank("check_in", "{adults}{r}+{children}{r}"),
            group="Guests"),
        Col("over_capacity", "Over capacity", 12, "flag",
            formula=blank(
                "check_in",
                'IF({property}{r}="","",'
                'IF({guests_total}{r}>'
                'IFERROR(INDEX([[Property Database.max_guests]],'
                'MATCH({property}{r},[[Property Database.id]],0))*1,'
                '999),"Yes","No"))'),
            cf=[{"type": "text", "criteria": "containing", "value": "Yes",
                 "fmt": "bad"}],
            hidden=True, group="Guests"),

        # ----------------------------------------------------- the money --
        Col("rate_in", "Nightly rate", 12, "money",
            validate=number_check(0, 100000, "Nightly rate"),
            note="Leave blank to use the property's standard rate. Type "
                 "a number here to override it for this stay only.",
            group="Money"),
        Col("rate_used", "Rate used", 11, "money",
            formula=blank(
                "check_in",
                'IF({rate_in}{r}<>"",{rate_in}{r},'
                'IF({property}{r}="",0,'
                'IFERROR(INDEX([[Property Database.nightly_rate]],'
                'MATCH({property}{r},[[Property Database.id]],0))*1,0)))'),
            note="The rate this booking is actually priced at: what you "
                 "typed, or the property's standard rate when you left "
                 "it blank.",
            group="Money"),
        Col("accommodation", "Room revenue", 13, "money",
            formula=blank(
                "check_in",
                'IF({earns}{r}<>"Yes",0,'
                'ROUND({nights}{r}*{rate_used}{r},2))'),
            note="Nights times the rate. This is the only figure ADR and "
                 "RevPAR are built from - cleaning and extras stay out "
                 "of it on purpose.",
            group="Money"),
        Col("cleaning_in", "Cleaning fee", 12, "money",
            validate=number_check(0, 100000, "Cleaning fee"),
            note="Blank uses the property's standard cleaning fee.",
            group="Money"),
        Col("cleaning_fee", "Cleaning charged", 13, "money",
            formula=blank(
                "check_in",
                'IF({earns}{r}<>"Yes",0,IF({cleaning_in}{r}<>"",'
                '{cleaning_in}{r},'
                'IFERROR(INDEX([[Property Database.cleaning_fee]],'
                'MATCH({property}{r},[[Property Database.id]],0))*1,0)))'),
            group="Money"),
        Col("extra_in", "Extra fees", 11, "money",
            validate=number_check(0, 100000, "Extra fees"),
            note="Blank works out the extra guest charge for you from "
                 "the Settings rate. Type a number to override, or 0 to "
                 "waive it.",
            group="Money"),
        Col("extra_fee", "Extra charged", 12, "money",
            formula=blank(
                "check_in",
                'IF({earns}{r}<>"Yes",0,IF({extra_in}{r}<>"",'
                '{extra_in}{r},'
                'ROUND(MAX(0,{guests_total}{r}-ExtraGuestAfter)*'
                'ExtraGuestFee*{nights}{r}+{pets}{r}*PetFee,2)))'),
            note="Extra guests beyond the free allowance, per night, "
                 "plus the pet fee. All three numbers live on Settings.",
            group="Money"),
        Col("discount", "Discount", 11, "money",
            validate=number_check(0, 100000, "Discount"),
            note="Weekly or monthly discount, or the goodwill you gave "
                 "back. Enter it as a positive number.",
            group="Money"),
        Col("gross", "Gross", 12, "money",
            formula=blank(
                "check_in",
                'ROUND(MAX(0,{accommodation}{r}+{cleaning_fee}{r}+'
                '{extra_fee}{r}-{discount}{r}),2)'),
            cf=money_cf(),
            note="Everything the guest paid, before the platform takes "
                 "its cut.",
            group="Money"),
        Col("tax", "Lodging tax", 11, "money",
            formula=blank("check_in", lodging_tax("{accommodation}{r}")),
            note="Zero when the platform collects and remits it for you, "
                 "which is the default. Switch it on Settings.",
            hidden=True, group="Money"),
        Col("fee_pct", "Commission %", 11, "percent",
            formula=blank("check_in", platform_rate("PlatformCommission",
                                                    "{platform}{r}")),
            hidden=True, group="Money"),
        Col("pay_pct", "Card fee %", 10, "percent",
            formula=blank("check_in", platform_rate("PlatformPayment",
                                                    "{platform}{r}")),
            hidden=True, group="Money"),
        Col("platform_fee", "Platform fee", 12, "money",
            formula=blank(
                "check_in",
                'IF({earns}{r}<>"Yes",0,'
                'ROUND({gross}{r}*{fee_pct}{r},2))'),
            group="Money"),
        Col("payment_fee", "Card fee", 10, "money",
            formula=blank(
                "check_in",
                'IF({earns}{r}<>"Yes",0,'
                'ROUND({gross}{r}*{pay_pct}{r},2))'),
            group="Money"),
        Col("platform_fee_total", "Fees total", 11, "money",
            formula=blank("check_in",
                          "ROUND({platform_fee}{r}+{payment_fee}{r},2)"),
            group="Money"),
        Col("net_revenue", "Net revenue", 13, "money",
            formula=blank(
                "check_in",
                'ROUND({gross}{r}-{platform_fee_total}{r}-{tax}{r},2)'),
            cf=money_cf(),
            note="What you actually keep: gross, less the platform fee, "
                 "less the card fee, less any lodging tax you have to "
                 "remit yourself.",
            group="Money"),
        Col("net_night", "Net a night", 11, "money",
            formula=blank("check_in",
                          'IF({nights}{r}<=0,"",'
                          'ROUND({net_revenue}{r}/{nights}{r},2))'),
            hidden=True, group="Money"),
        Col("payment_status", "Payment", 14, "status",
            validate=list_from("PaymentStatusList"),
            cf=status_cf(), group="Money"),

        # ------------------------------------------------------- flags --
        Col("holds", "Holds dates", 10, "flag",
            formula=('=IF({check_in}{r}="","No",'
                     f'IF({is_holding("{status}{r}")},"Yes","No"))'),
            hidden=True,
            note="Yes for anything that takes the dates off the market - "
                 "a real stay or a block. The calendar and the clash "
                 "check read this."),
        Col("earns", "Earns", 8, "flag",
            formula=('=IF({check_in}{r}="","No",'
                     f'IF({is_earning("{status}{r}")},"Yes","No"))'),
            hidden=True),
        Col("is_block", "Is a block", 10, "flag",
            formula=('=IF({check_in}{r}="","No",'
                     f'IF({is_block("{status}{r}")},"Yes","No"))'),
            hidden=True),
        Col("is_cancelled", "Cancelled", 10, "flag",
            formula=('=IF({check_in}{r}="","No",'
                     'IF({status}{r}="Cancelled","Yes","No"))'),
            hidden=True),
        Col("stay_band", "Stay length", 12, "text",
            formula=blank(
                "check_in",
                'IF({nights}{r}<=0,"",'
                'IF({nights}{r}=1,"1 night",'
                'IF({nights}{r}=2,"2 nights",'
                'IF({nights}{r}<=4,"3-4 nights",'
                'IF({nights}{r}<=6,"5-6 nights",'
                'IF({nights}{r}<=9,"1 week",'
                'IF({nights}{r}<=27,"10+ nights","Monthly stay")))))))'),
            hidden=True, group="Flags"),

        # ------------------------------------------------- the warnings --
        Col("clash", "Clash", 9, "int",
            formula=zero(
                "check_in",
                'IF({holds}{r}<>"Yes",0,'
                'COUNTIFS({property_first}:{property_last},{property}{r},'
                '{ci_num_first}:{ci_num_last},"<"&{co_num}{r},'
                '{co_num_first}:{co_num_last},">"&{ci_num}{r},'
                '{holds_first}:{holds_last},"Yes")-1)'),
            hidden=True,
            note="How many other reservations on this property overlap "
                 "these dates. Anything above zero is a double booking."),
        Col("same_day", "Same-day turn", 12, "flag",
            formula=blank(
                "check_in",
                'IF({holds}{r}<>"Yes","No",'
                'IF({nights}{r}<=0,"No",'
                'IF(COUNTIFS({property_first}:{property_last},'
                '{property}{r},{ci_num_first}:{ci_num_last},{co_num}{r},'
                '{holds_first}:{holds_last},"Yes")>0,"Yes","No")))'),
            cf=[{"type": "text", "criteria": "containing", "value": "Yes",
                 "fmt": "warn"}],
            note="Somebody else checks in on the day this guest leaves. "
                 "The cleaner has hours, not a day.",
            group="Warnings"),
        Col("has_cleaning", "Cleaning booked", 13, "flag",
            formula=blank(
                "check_in",
                f'IF({{holds}}{{r}}<>"Yes","No",'
                f'IF(COUNTIFS({cl_booking},{{id}}{{r}})>0,"Yes","No"))'),
            cf=flag_yes_no_cf(), hidden=True, group="Warnings"),
        Col("has_arrival", "Arrival logged", 13, "flag",
            formula=blank(
                "check_in",
                f'IF({{earns}}{{r}}<>"Yes","No",'
                f'IF(COUNTIFS({ar_booking},{{id}}{{r}})>0,"Yes","No"))'),
            hidden=True, group="Warnings"),
        Col("has_review", "Review logged", 13, "flag",
            formula=blank(
                "check_in",
                f'IF({{earns}}{{r}}<>"Yes","No",'
                f'IF(COUNTIFS({rv_booking},{{id}}{{r}})>0,"Yes","No"))'),
            hidden=True, group="Warnings"),
        Col("stay_state", "Where it is", 13, "status",
            formula=blank(
                "check_in",
                'IF({holds}{r}<>"Yes","",'
                'IF({check_out}{r}="","",'
                'IF(TodayDate<{check_in}{r},'
                'IF({check_in}{r}-TodayDate<=UpcomingDays,"Arriving soon",'
                '"Upcoming"),'
                'IF(TodayDate>={check_out}{r},"Past",'
                'IF(TodayDate={check_in}{r},"Arriving today","Staying")))))'),
            cf=alert_cf(), group="Warnings"),
        Col("arr_key", "Arrival order key", 10, "number",
            formula=('=IF({check_in}{r}="",9999999999,'
                     'IF({earns}{r}<>"Yes",9999999999,'
                     'IF({check_in}{r}<TodayDate,9999999999,'
                     '{ci_num}{r}*10000+ROW()-4)))'),
            hidden=True, group="Flags"),
        Col("arrival_rank", "Arriving next", 10, "int",
            formula=('=IF({arr_key}{r}>=9999999999,"",'
                     'COUNTIFS({arr_key_first}:{arr_key_last},"<"&'
                     '{arr_key}{r})+1)'),
            note="1 is the next guest through the door. The dashboard "
                 "reads this to list who is arriving.",
            hidden=True, group="Flags"),
        Col("alert", "Action", 28, "status",
            formula=blank(
                "check_in",
                'IF({property}{r}="","Add - no property on this row",'
                'IF({clash}{r}>0,"Double booked - '
                'overlapping reservation",'
                'IF({check_out}{r}="","Add - no check-out date",'
                'IF({nights}{r}<=0,"Check - check-out is not after '
                'check-in",'
                'IF({status}{r}="Cancelled","Cancelled",'
                'IF({is_block}{r}="Yes","Blocked - '
                'no guest, no revenue",'
                'IF({over_capacity}{r}="Yes","Check - over the sleeps '
                'limit",'
                'IF({status}{r}="Inquiry",'
                'IF(TodayDate>{check_in}{r},"Chase - enquiry went past '
                'the dates","Follow up - still an enquiry"),'
                'IF({guest}{r}="","Add - no guest record",'
                'IF({has_cleaning}{r}="No","Schedule - no cleaning job '
                'for the check-out",'
                'IF({same_day}{r}="Yes","Same day - tight turnaround",'
                'IF({rate_used}{r}<=0,"Add - no nightly rate",'
                'IF(TodayDate>={check_out}{r},'
                'IF({payment_status}{r}="Paid in full",'
                'IF({has_review}{r}="No","Chase - no review logged","OK"),'
                '"Chase - payout not marked paid"),'
                '"OK")))))))))))))'),
            cf=alert_cf(), group="Warnings"),

        # ------------------------------------------------------- notes --
        Col("requests", "Guest requests", 30, "wrap",
            note="Early check-in, travel cot, late flight. Printed on "
                 "the arrival sheet.",
            group="Notes"),
        Col("notes", "Notes", 30, "wrap", group="Notes"),
    ]

    return SheetSpec(
        name=NAME,
        title="Booking Tracker",
        subtitle="Every reservation, enquiry and block - priced, counted "
                 "and checked for clashes",
        columns=columns,
        rows=cfg.row_count(NAME),
        freeze_cols=3,
        samples=sample("bookings", cfg),
        help_text="One row per reservation. Type the property, the guest, "
                  "the platform and the two dates and the rest works "
                  "itself out: nights, revenue, the platform's cut, what "
                  "you keep, and whether the dates clash with a booking "
                  "you already took. Use the Owner block and Maintenance "
                  "block statuses to take dates off the market.",
        extra_notes=[
            "Leave the nightly rate blank to use the property's standard "
            "rate - type one only when this stay is priced differently.",
            "A Clash warning means two reservations that hold the dates "
            "overlap on the same property. Cancel one or change the "
            "dates; the calendar will keep shouting until you do.",
            "Cleaning fees are charged to the guest and recovered as a "
            "cost on the Cleaning Schedule tab. They are never counted "
            "as room revenue, so ADR and RevPAR stay honest.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Booking totals and a platform breakdown under the register."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index

    row = sheet.section(row, "Every booking on this tab", last)
    row = sheet.table_head(row, 0, [
        "Bookings", "Enquiries", "Cancelled", "Blocks", "Nights booked",
        "Nights blocked", "Guests", "Avg stay", "Clashes", "Same-day turns"],
        [11, 11, 10, 9, 13, 13, 9, 10, 9, 14])
    sheet.calc_row(row, 0, [
        '=COUNTIFS([[Booking Tracker.earns]],"Yes")',
        '=COUNTIFS([[Booking Tracker.status]],"Inquiry")',
        '=COUNTIFS([[Booking Tracker.is_cancelled]],"Yes")',
        '=COUNTIFS([[Booking Tracker.is_block]],"Yes")',
        '=SUMIFS([[Booking Tracker.nights]],'
        '[[Booking Tracker.earns]],"Yes")',
        '=SUMIFS([[Booking Tracker.nights]],'
        '[[Booking Tracker.is_block]],"Yes")',
        '=SUMIFS([[Booking Tracker.guests_total]],'
        '[[Booking Tracker.earns]],"Yes")',
        '=IF(COUNTIFS([[Booking Tracker.earns]],"Yes")<=0,"",'
        'ROUND(SUMIFS([[Booking Tracker.nights]],'
        '[[Booking Tracker.earns]],"Yes")/'
        'COUNTIFS([[Booking Tracker.earns]],"Yes"),1))',
        '=COUNTIFS([[Booking Tracker.clash]],">0")',
        '=COUNTIFS([[Booking Tracker.same_day]],"Yes")',
    ], ["int", "int", "int", "int", "int", "int", "int", "number1",
        "int", "int"])
    row += 2

    row = sheet.table_head(row, 0, [
        "Room revenue", "Cleaning charged", "Extras", "Discounts",
        "Gross", "Platform fees", "Card fees", "Lodging tax",
        "Net revenue", "Net a night"],
        [13, 14, 11, 11, 12, 13, 11, 11, 13, 12])
    sheet.calc_row(row, 0, [
        '=ROUND(SUMIFS([[Booking Tracker.accommodation]],'
        '[[Booking Tracker.earns]],"Yes"),2)',
        '=ROUND(SUMIFS([[Booking Tracker.cleaning_fee]],'
        '[[Booking Tracker.earns]],"Yes"),2)',
        '=ROUND(SUMIFS([[Booking Tracker.extra_fee]],'
        '[[Booking Tracker.earns]],"Yes"),2)',
        '=ROUND(SUMIFS([[Booking Tracker.discount]],'
        '[[Booking Tracker.earns]],"Yes"),2)',
        '=ROUND(SUMIFS([[Booking Tracker.gross]],'
        '[[Booking Tracker.earns]],"Yes"),2)',
        '=ROUND(SUMIFS([[Booking Tracker.platform_fee]],'
        '[[Booking Tracker.earns]],"Yes"),2)',
        '=ROUND(SUMIFS([[Booking Tracker.payment_fee]],'
        '[[Booking Tracker.earns]],"Yes"),2)',
        '=ROUND(SUMIFS([[Booking Tracker.tax]],'
        '[[Booking Tracker.earns]],"Yes"),2)',
        '=ROUND(SUMIFS([[Booking Tracker.net_revenue]],'
        '[[Booking Tracker.earns]],"Yes"),2)',
        '=IF(SUMIFS([[Booking Tracker.nights]],'
        '[[Booking Tracker.earns]],"Yes")<=0,"",'
        'ROUND(SUMIFS([[Booking Tracker.net_revenue]],'
        '[[Booking Tracker.earns]],"Yes")/'
        'SUMIFS([[Booking Tracker.nights]],'
        '[[Booking Tracker.earns]],"Yes"),2))',
    ], ["money", "money", "money", "money", "money", "money", "money",
        "money", "money", "money"])
    row += 2

    row = sheet.section(row, "Which platform is worth having", last)
    row = sheet.table_head(row, 0, [
        "Platform", "Bookings", "Nights", "Gross", "Fees",
        "Net revenue", "Fee rate", "Net a night", "Avg stay",
        "Share of net"],
        [18, 10, 9, 12, 11, 13, 10, 12, 10, 12])
    for name, _commission, _payment, _counts, _note in cfg.platform_table:
        key = f'"{name}"'
        earns = '[[Booking Tracker.earns]],"Yes"'
        plat = f'[[Booking Tracker.platform]],{key}'
        gross = f'SUMIFS([[Booking Tracker.gross]],{plat},{earns})'
        fees = (f'SUMIFS([[Booking Tracker.platform_fee_total]],'
                f'{plat},{earns})')
        net = f'SUMIFS([[Booking Tracker.net_revenue]],{plat},{earns})'
        nights = f'SUMIFS([[Booking Tracker.nights]],{plat},{earns})'
        count = f'COUNTIFS({plat},{earns})'
        total_net = (f'SUMIFS([[Booking Tracker.net_revenue]],{earns})')
        sheet.calc_row(row, 0, [
            name,
            f"={count}",
            f"={nights}",
            f"=ROUND({gross},2)",
            f"=ROUND({fees},2)",
            f"=ROUND({net},2)",
            f'=IF({gross}<=0,"",{fees}/{gross})',
            f'=IF({nights}<=0,"",ROUND({net}/{nights},2))',
            f'=IF({count}<=0,"",ROUND({nights}/{count},1))',
            f'=IF({total_net}<=0,"",{net}/{total_net})',
        ], ["text", "int", "int", "money", "money", "money", "percent",
            "money", "number1", "percent"])
        row += 1

    row += 1
    sheet.note(row, "Fee rate is what the platform actually kept out of "
                    "the money the guest handed over, including card "
                    "processing. Compare it against the net-a-night "
                    "column before you decide a channel is not worth "
                    "the commission - a platform that fills midweek "
                    "nights at a lower rate can still beat one that "
                    "only sends you weekends.", last)
````

### `airbnb_rental_spreadsheet/sheets/guests.py`

*252 lines · 2 functions*

> `Guest Database` - who stayed, how often, and what they were worth.
>
> A short-term rental business that only ever talks to strangers is paying
> full commission on every single night.  This tab keeps the history: how
> many times someone has stayed, which properties they liked, what they
> have spent and how they reviewed you - so the message you send before
> next summer goes to the right twenty people.
>
> Deliberately modest about what it stores.  A name, a way to contact
> them, where they came from and what they asked for is enough to run the
> business.  Passport numbers, card details and copies of ID do not
> belong in a spreadsheet.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Who is actually worth marketing to |

**Module constants:** `NAME` ('Guest Database').

````python
"""`Guest Database` - who stayed, how often, and what they were worth.

A short-term rental business that only ever talks to strangers is paying
full commission on every single night.  This tab keeps the history: how
many times someone has stayed, which properties they liked, what they
have spent and how they reviewed you - so the message you send before
next summer goes to the right twenty people.

Deliberately modest about what it stores.  A name, a way to contact
them, where they came from and what they asked for is enough to run the
business.  Passport numbers, card details and copies of ID do not
belong in a spreadsheet.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec, max_if, min_if
from ..sample_data import get as sample
from ..writer import list_from
from .common import (alert_cf, bar_cf, blank, flag_yes_no_cf, id_col,
                     money_cf, rating_cf, status_cf)

NAME = "Guest Database"


def spec(cfg: Config) -> SheetSpec:
    bk_guest = "[[Booking Tracker.guest]]"
    bk_earns = "[[Booking Tracker.earns]]"
    bk_nights = "[[Booking Tracker.nights]]"
    bk_gross = "[[Booking Tracker.gross]]"
    bk_net = "[[Booking Tracker.net_revenue]]"
    bk_ci = "[[Booking Tracker.ci_num]]"
    bk_cancel = "[[Booking Tracker.is_cancelled]]"
    rv_guest = "[[Reviews Tracker.guest]]"
    rv_score = "[[Reviews Tracker.rate_overall]]"
    rv_has = "[[Reviews Tracker.has_score]]"

    # MIN and MAX over a condition, without an array formula.  Both read
    # the hidden numeric mirror of the check-in date rather than the date
    # column itself: a formula column that returns "" would be multiplied
    # by a zero somewhere in the SUMPRODUCT and answer #VALUE!.
    _guest_is = bk_guest + "={id}{r}"
    _earning = bk_earns + '="Yes"'
    _earliest = min_if(bk_ci, _guest_is, _earning)
    _latest = max_if(bk_ci, _guest_is, _earning)

    columns = [
        # ---------------------------------------------------- the guest --
        id_col("name", "GuestPrefix", digits=cfg.digits("guest"), width=11,
               header="Guest ID",
               note="Automatic. Bookings point at this, which is how the "
                    "history below builds itself."),
        Col("name", "Guest name", 22, "text",
            note="One row per person, not per stay. If they come back, "
                 "point the new booking at the same guest ID.",
            group="The guest"),
        Col("email", "Email", 26, "text",
            note="Platforms mask this until a booking is confirmed. Add "
                 "it when you get it - it is what makes a direct "
                 "re-booking possible.",
            group="The guest"),
        Col("phone", "Phone", 16, "text", group="The guest"),
        Col("country", "Country", 15, "text",
            validate=list_from("CountryList"), group="The guest"),
        Col("guest_type", "Guest type", 13, "text",
            validate=list_from("GuestTypeList"), group="The guest"),
        Col("purpose", "Usual trip", 14, "text",
            validate=list_from("TripPurposeList"), group="The guest"),
        Col("first_platform", "Found you on", 15, "text",
            validate=list_from("PlatformList"),
            note="Where the first booking came from. A guest who first "
                 "arrived via Airbnb and now books direct is the whole "
                 "point of keeping this tab.",
            group="The guest"),
        Col("marketing_ok", "Happy to be contacted", 16, "flag",
            validate=list_from("YesNoList"),
            cf=flag_yes_no_cf(),
            note="Only mail the people who said yes. Keep this honest - "
                 "it is the column that keeps you out of trouble.",
            group="The guest"),

        # -------------------------------------------------- the history --
        Col("stays", "Stays", 8, "int",
            formula=blank("name", f'COUNTIFS({bk_guest},{{id}}{{r}},'
                                  f'{bk_earns},"Yes")'),
            cf=bar_cf(), group="History"),
        Col("cancellations", "Cancelled", 10, "int",
            formula=blank("name", f'COUNTIFS({bk_guest},{{id}}{{r}},'
                                  f'{bk_cancel},"Yes")'),
            hidden=True, group="History"),
        Col("nights", "Nights", 9, "int",
            formula=blank("name", f'SUMIFS({bk_nights},{bk_guest},'
                                  f'{{id}}{{r}},{bk_earns},"Yes")'),
            group="History"),
        Col("first_stay", "First stay", 12, "date",
            formula=blank(
                "name",
                'IF({stays}{r}<=0,"",' + _earliest + ")"),
            note="The check-in date of their very first stay.",
            group="History"),
        Col("last_stay", "Last stay", 12, "date",
            formula=blank(
                "name",
                'IF({stays}{r}<=0,"",' + _latest + ")"),
            group="History"),
        Col("days_since", "Days since", 11, "int",
            formula=blank("name",
                          'IF({last_stay}{r}="","",'
                          'MAX(0,TodayDate-{last_stay}{r}))'),
            group="History"),
        Col("repeat", "Repeat guest", 12, "flag",
            formula=blank("name", 'IF({stays}{r}>1,"Yes","No")'),
            cf=flag_yes_no_cf(), group="History"),

        # ---------------------------------------------------- the money --
        Col("gross", "Lifetime gross", 13, "money",
            formula=blank("name", f'ROUND(SUMIFS({bk_gross},{bk_guest},'
                                  f'{{id}}{{r}},{bk_earns},"Yes"),2)'),
            cf=money_cf(), group="Value"),
        Col("net", "Lifetime net", 13, "money",
            formula=blank("name", f'ROUND(SUMIFS({bk_net},{bk_guest},'
                                  f'{{id}}{{r}},{bk_earns},"Yes"),2)'),
            note="What you kept after the platform's cut. A direct "
                 "repeat guest is worth noticeably more than the gross "
                 "figure suggests.",
            group="Value"),
        Col("avg_stay_value", "Avg per stay", 12, "money",
            formula=blank("name",
                          'IF({stays}{r}<=0,"",'
                          'ROUND({net}{r}/{stays}{r},2))'),
            group="Value"),
        Col("nights_avg", "Avg nights", 11, "number1",
            formula=blank("name",
                          'IF({stays}{r}<=0,"",'
                          'ROUND({nights}{r}/{stays}{r},1))'),
            hidden=True, group="Value"),
        Col("value_rank", "Rank", 7, "int",
            formula=('=IF({name}{r}="","",COUNTIFS({net_first}:{net_last},'
                     '">"&{net}{r})+COUNTIFS({net_first}:{net}{r},'
                     '{net}{r}))'),
            note="1 = your most valuable guest by net revenue.",
            group="Value"),

        # -------------------------------------------------- the reviews --
        Col("reviews", "Reviews left", 12, "int",
            formula=blank("name", f'COUNTIFS({rv_guest},{{id}}{{r}},'
                                  f'{rv_has},"Yes")'),
            hidden=True, group="Reviews"),
        Col("rating_given", "Rating given", 12, "number",
            formula=blank(
                "name",
                f'IF({{reviews}}{{r}}<=0,"",'
                f'ROUND(SUMIFS({rv_score},{rv_guest},{{id}}{{r}},'
                f'{rv_has},"Yes")/{{reviews}}{{r}},2))'),
            cf=rating_cf(),
            note="The average score this guest has given you.",
            group="Reviews"),
        Col("review_status", "Review status", 14, "status",
            validate=list_from("ReviewStatusList"),
            cf=status_cf(), group="Reviews"),

        # ----------------------------------------------- what they like --
        Col("preferences", "Preferences", 28, "wrap",
            note="Ground floor, early check-in, extra pillows, allergic "
                 "to the cat. The reason they book you again.",
            group="Notes"),
        Col("requests", "Standing requests", 26, "wrap", group="Notes"),
        Col("feedback", "What they said", 30, "wrap", group="Notes"),
        Col("internal", "Internal notes", 28, "wrap",
            note="For you, not for them. Keep it factual - a guest can "
                 "ask to see what you hold about them.",
            group="Notes"),

        # ----------------------------------------------- what to do now --
        Col("alert", "Action", 26, "status",
            formula=blank(
                "name",
                'IF({stays}{r}<=0,"Add - no stays linked to this guest",'
                'IF({email}{r}="","Add - no email, no re-booking",'
                'IF({repeat}{r}="Yes",'
                'IF({marketing_ok}{r}="Yes","OK - repeat guest, '
                'contactable","Check - repeat guest, no consent"),'
                'IF({days_since}{r}="","",'
                'IF({days_since}{r}>330,"Follow up - a year since they '
                'stayed",'
                'IF({review_status}{r}="Requested","Chase - review '
                'requested, none received","OK"))))))'),
            cf=alert_cf(), group="Notes"),
    ]

    return SheetSpec(
        name=NAME,
        title="Guest Database",
        subtitle="Who has stayed with you, how often, and what they were "
                 "worth after commission",
        columns=columns,
        rows=cfg.row_count(NAME),
        freeze_cols=2,
        samples=sample("guests", cfg),
        help_text="One row per person, not per stay. Point every booking "
                  "at a guest ID and the stay count, the nights, the "
                  "lifetime value and the review average all build "
                  "themselves. The repeat-guest column is the one worth "
                  "sorting on.",
        extra_notes=[
            "Keep this tab modest. A name, a contact, where they came "
            "from and what they asked for is enough. Identity documents "
            "and card details do not belong in a spreadsheet.",
            "Only contact the guests whose consent column says Yes.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Who is actually worth marketing to."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index

    row = sheet.section(row, "Your guest book", last)
    row = sheet.table_head(row, 0, [
        "Guests on file", "Repeat guests", "Repeat rate", "Contactable",
        "Total stays", "Total nights", "Lifetime net", "Net per guest",
        "Avg rating given"],
        [13, 13, 11, 12, 11, 12, 13, 13, 14])
    guests = '=COUNTIF([[Guest Database.name]],"?*")'
    sheet.calc_row(row, 0, [
        guests,
        '=COUNTIFS([[Guest Database.repeat]],"Yes")',
        '=IF(COUNTIF([[Guest Database.name]],"?*")<=0,"",'
        'COUNTIFS([[Guest Database.repeat]],"Yes")/'
        'COUNTIF([[Guest Database.name]],"?*"))',
        '=COUNTIFS([[Guest Database.marketing_ok]],"Yes")',
        '=SUM([[Guest Database.stays]])',
        '=SUM([[Guest Database.nights]])',
        '=ROUND(SUM([[Guest Database.net]]),2)',
        '=IF(COUNTIF([[Guest Database.name]],"?*")<=0,"",'
        'ROUND(SUM([[Guest Database.net]])/'
        'COUNTIF([[Guest Database.name]],"?*"),2))',
        '=IF(SUM([[Guest Database.reviews]])<=0,"",'
        'ROUND(SUMPRODUCT([[Guest Database.reviews]],'
        'IFERROR([[Guest Database.rating_given]]*1,0))/'
        'SUM([[Guest Database.reviews]]),2))',
    ], ["int", "int", "percent", "int", "int", "int", "money", "money",
        "number"])
    row += 2

    sheet.note(row, "A repeat guest costs you nothing in commission if "
                    "they book direct. Sort by Rank, take the top twenty "
                    "with consent, and that is your whole marketing "
                    "list.", last)
````

---

## 4. The calendar

A month grid and a fourteen day strip across twelve properties, both drawn entirely from the Booking Tracker. No second calendar is kept, so nothing can fall out of step: an owner stay or a maintenance closure is a booking row with a block status, and the grid reads it like any other. The same page carries the Turnover Command Centre.

### `airbnb_rental_spreadsheet/sheets/calendar.py`

*427 lines · 7 functions*

> `Booking Calendar` - a month at a glance, and the next fortnight.
>
> Two views on one page.
>
> The **month grid** is one property at a time: pick the property and
> the month at the top and every day fills itself in - who is staying,
> which platform, and whether the day is free, occupied, an arrival, a
> departure, blocked by you or closed for maintenance.  Colour does the
> rest.
>
> The **fortnight strip** is every property at once for the next
> fourteen days, which is the view you actually want on a Friday
> morning.  Underneath it sits the Turnover Command Center: today's
> departures, tomorrow's arrivals, which flats are clean and which are
> not.
>
> One thing this is not, and the workbook never pretends otherwise: it
> does not talk to Airbnb.  There is no sync, no iCal, no API.  It draws
> what is in your Booking Tracker, which means it is right exactly as
> often as the Booking Tracker is.

| Defines | What it does |
|---------|--------------|
| `_occupied(day, prop)` | COUNTIFS for a paying stay covering `day` at `prop` |
| `_blocked(day, status, prop)` |  |
| `_arrives(day, prop)` |  |
| `_departs(day, prop)` |  |
| `_row_on(day, prop)` | Row number of the stay covering `day`, or 0 |
| `_guest_on(day, prop)` | Name of whoever is in on `day` |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Booking Calendar'), `LAST_COL` (7), `WEEK_ROWS` (6), `DAY_LINES` (3), `STRIP_PROPERTIES` (12).

````python
"""`Booking Calendar` - a month at a glance, and the next fortnight.

Two views on one page.

The **month grid** is one property at a time: pick the property and
the month at the top and every day fills itself in - who is staying,
which platform, and whether the day is free, occupied, an arrival, a
departure, blocked by you or closed for maintenance.  Colour does the
rest.

The **fortnight strip** is every property at once for the next
fourteen days, which is the view you actually want on a Friday
morning.  Underneath it sits the Turnover Command Center: today's
departures, tomorrow's arrivals, which flats are clean and which are
not.

One thing this is not, and the workbook never pretends otherwise: it
does not talk to Airbnb.  There is no sync, no iCal, no API.  It draws
what is in your Booking Tracker, which means it is right exactly as
often as the Booking Tracker is.
"""

from __future__ import annotations

from ..config import CALENDAR_STRIP_DAYS, Config, Palette
from ..layout import Registry, col_letter
from ..writer import SheetWriter, list_from

NAME = "Booking Calendar"
LAST_COL = 7
WEEK_ROWS = 6
DAY_LINES = 3
STRIP_PROPERTIES = 12


def _occupied(day: str, prop: str = "CalProperty") -> str:
    """COUNTIFS for a paying stay covering `day` at `prop`."""
    return (f'COUNTIFS([[Booking Tracker.property]],{prop},'
            f'[[Booking Tracker.check_in]],"<="&{day},'
            f'[[Booking Tracker.check_out]],">"&{day},'
            f'[[Booking Tracker.earns]],"Yes")')


def _blocked(day: str, status: str, prop: str = "CalProperty") -> str:
    return (f'COUNTIFS([[Booking Tracker.property]],{prop},'
            f'[[Booking Tracker.check_in]],"<="&{day},'
            f'[[Booking Tracker.check_out]],">"&{day},'
            f'[[Booking Tracker.status]],"{status}")')


def _arrives(day: str, prop: str = "CalProperty") -> str:
    return (f'COUNTIFS([[Booking Tracker.property]],{prop},'
            f'[[Booking Tracker.check_in]],{day},'
            f'[[Booking Tracker.earns]],"Yes")')


def _departs(day: str, prop: str = "CalProperty") -> str:
    return (f'COUNTIFS([[Booking Tracker.property]],{prop},'
            f'[[Booking Tracker.check_out]],{day},'
            f'[[Booking Tracker.earns]],"Yes")')


def _row_on(day: str, prop: str = "CalProperty") -> str:
    """Row number of the stay covering `day`, or 0. Numeric ranges only."""
    return (f'SUMPRODUCT(MAX(([[Booking Tracker.property]]={prop})*'
            f'([[Booking Tracker.ci_num]]<={day})*'
            f'([[Booking Tracker.co_num]]>{day})*'
            f'([[Booking Tracker.earns]]="Yes")*'
            f'[[Booking Tracker.seq_num]]))')


def _guest_on(day: str, prop: str = "CalProperty") -> str:
    """Name of whoever is in on `day`. No array formula, no CSE."""
    found = _row_on(day, prop)
    return (f'IF({found}<=0,"",IFERROR(INDEX('
            f'[[Booking Tracker.guest_name]],{found})&"",""))')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_mid)
    ws = s.ws

    s.banner("Booking Calendar",
             "One property a month, and every property for the next "
             "fortnight. Drawn from the Booking Tracker, not from "
             "Airbnb.", LAST_COL)
    s.set_col(0, 0, 20)
    for column in range(1, 8):
        s.set_col(column, column, 19)
    s.set_col(8, 20, 13)

    # ------------------------------------------------------- the picker --
    row = 4
    row = s.note(
        row,
        "This calendar is not connected to Airbnb, Vrbo or any other "
        "site and cannot be. It draws whatever is in your Booking "
        "Tracker, so keep that tab honest and this one looks after "
        "itself.",
        LAST_COL, height=26)
    row += 1

    s.merge(row, 0, 0, "Property", fmts.chrome("label_bold"))
    s.merge(row, 1, 2, None, fmts.chrome("form_box", locked=False))
    ws.data_validation(row, 1, row, 1, list_from("PropertyList"))
    book.define_name("CalProperty", f"='{NAME}'!$B${row + 1}")
    s.merge(row, 3, 4, None, fmts.calc("text"))
    s.formula(
        row, 3,
        '=IFERROR(INDEX([[Property Database.name]],MATCH(CalProperty,'
        '[[Property Database.id]],0)),"Pick a property ID")',
        fmts.calc("text"))
    s.merge(row, 5, LAST_COL, "Type a property ID, e.g. PRP-01",
            fmts.chrome("note"))
    picker_row = row
    row += 1

    s.merge(row, 0, 0, "Month", fmts.chrome("label_bold"))
    s.write(row, 1, cfg.today.replace(day=1),
            fmts.chrome("form_box", locked=False,
                        num_format=cfg.date_format))
    book.define_name("CalMonth", f"='{NAME}'!$B${row + 1}")
    s.merge(row, 2, 2, None, fmts.calc("text"))
    s.formula(row, 2, '=TEXT(CalStart,"mmmm yyyy")',
                     fmts.calc("text"))
    s.merge(row, 3, LAST_COL,
            "Any date in the month you want to see - it snaps to the "
            "first.", fmts.chrome("note"))
    row += 1

    # Hidden helpers: first of month, and the Monday the grid starts on.
    helper = row
    s.write(helper, 20, None, fmts.calc("date"))
    s.formula(
        helper, 20,
        '=IF(CalMonth="",DATE(YEAR(TodayDate),MONTH(TodayDate),1),'
        'DATE(YEAR(CalMonth),MONTH(CalMonth),1))', fmts.calc("date"))
    book.define_name("CalStart", f"='{NAME}'!$U${helper + 1}")
    s.write(helper + 1, 20, None, fmts.calc("date"))
    s.formula(helper + 1, 20,
                     "=CalStart-WEEKDAY(CalStart,3)", fmts.calc("date"))
    book.define_name("CalGridStart", f"='{NAME}'!$U${helper + 2}")
    ws.set_column(20, 20, None, None, {"hidden": True})
    row += 1

    # --------------------------------------------------- month summary --
    days_in = ('DAY(DATE(YEAR(CalStart),MONTH(CalStart)+1,1)-1)')
    nights_sold = ('SUMIFS([[Booking Tracker.nights]],'
                   '[[Booking Tracker.property]],CalProperty,'
                   '[[Booking Tracker.month]],TEXT(CalStart,"yyyy-mm"),'
                   '[[Booking Tracker.earns]],"Yes")')
    accom = ('SUMIFS([[Booking Tracker.accommodation]],'
             '[[Booking Tracker.property]],CalProperty,'
             '[[Booking Tracker.month]],TEXT(CalStart,"yyyy-mm"),'
             '[[Booking Tracker.earns]],"Yes")')
    blocked = ('SUMIFS([[Booking Tracker.nights]],'
               '[[Booking Tracker.property]],CalProperty,'
               '[[Booking Tracker.month]],TEXT(CalStart,"yyyy-mm"),'
               '[[Booking Tracker.is_block]],"Yes")')

    s.kpi(row, 0, "Nights sold", f'=IF(CalProperty="","",{nights_sold})',
          "int")
    s.kpi(row, 1, "Nights free",
          f'=IF(CalProperty="","",MAX(0,{days_in}-{nights_sold}-'
          f'{blocked}))', "int", span=1)
    s.kpi(row, 2, "Blocked", f'=IF(CalProperty="","",{blocked})', "int",
          span=1)
    s.kpi(row, 3, "Occupancy",
          f'=IF(CalProperty="","",IF({days_in}-{blocked}<=0,"",'
          f'{nights_sold}/({days_in}-{blocked})))', "percent", span=1)
    s.kpi(row, 4, "Arrivals",
          '=IF(CalProperty="","",COUNTIFS('
          '[[Booking Tracker.property]],CalProperty,'
          '[[Booking Tracker.month]],TEXT(CalStart,"yyyy-mm"),'
          '[[Booking Tracker.earns]],"Yes"))', "int", span=1)
    s.kpi(row, 5, "Room revenue",
          f'=IF(CalProperty="","",ROUND({accom},2))', "money0", span=1)
    s.kpi(row, 6, "ADR",
          f'=IF(CalProperty="","",IF({nights_sold}<=0,"",'
          f'ROUND({accom}/{nights_sold},2)))', "money", span=1)
    row = s.kpi(row, 7, "Same-day turns",
                '=IF(CalProperty="","",COUNTIFS('
                '[[Cleaning Schedule.property]],CalProperty,'
                '[[Cleaning Schedule.month]],TEXT(CalStart,"yyyy-mm"),'
                '[[Cleaning Schedule.same_day]],"Yes"))', "int", span=1)
    row += 1

    # ------------------------------------------------------ month grid --
    row = s.section(row, "The month", LAST_COL)
    for index, day in enumerate(("Monday", "Tuesday", "Wednesday",
                                 "Thursday", "Friday", "Saturday",
                                 "Sunday")):
        s.write(row, index, day, fmts.chrome("header"))
    s.row_height(row, 22)
    row += 1

    grid_top = row
    for week in range(WEEK_ROWS):
        base = row + week * DAY_LINES
        s.row_height(base, 18)
        s.row_height(base + 1, 16)
        s.row_height(base + 2, 16)
        for weekday in range(7):
            offset = week * 7 + weekday
            day = f"(CalGridStart+{offset})"
            in_month = f'MONTH({day})=MONTH(CalStart)'
            # line 1 - the date
            s.formula(
                base, weekday,
                f'=IF({in_month},DAY({day}),"")',
                fmts.chrome("day_number"))
            # line 2 - who is in
            s.formula(
                base + 1, weekday,
                f'=IF(NOT({in_month}),"",IF(CalProperty="","",'
                f'IF({_blocked(day, "Maintenance block")}>0,'
                f'"Maintenance",'
                f'IF({_blocked(day, "Owner block")}>0,"Owner block",'
                f'IF({_occupied(day)}>0,{_guest_on(day)},"Available")'
                f'))))',
                fmts.chrome("day_item"))
            # line 3 - what happens that day
            s.formula(
                base + 2, weekday,
                f'=IF(NOT({in_month}),"",IF(CalProperty="","",'
                f'IF({_departs(day)}>0,'
                f'IF({_arrives(day)}>0,"Turnover - out and in",'
                f'"Check-out"),'
                f'IF({_arrives(day)}>0,"Check-in",'
                f'IF({_occupied(day)}>1,"DOUBLE BOOKED",'
                f'IF({_occupied(day)}>0,"Staying",""))))))',
                fmts.chrome("day_item"))
    grid_bottom = grid_top + WEEK_ROWS * DAY_LINES - 1

    for text, style in (("DOUBLE BOOKED", "bold_bad"),
                        ("Turnover", "warn"),
                        ("Check-out", "info"),
                        ("Check-in", "good"),
                        ("Available", "muted"),
                        ("Owner block", "muted"),
                        ("Maintenance", "bad")):
        ws.conditional_format(grid_top, 0, grid_bottom, 6, {
            "type": "text", "criteria": "containing", "value": text,
            "format": fmts.cf(style)})
    row = grid_bottom + 2

    row = s.note(
        row,
        "A day shows the guest who is sleeping there that night, so a "
        "check-out day belongs to whoever arrives, not to whoever "
        "left. 'DOUBLE BOOKED' means two paying stays claim the same "
        "night in the same property - fix it on the Booking Tracker, "
        "where the clash column will tell you which rows.",
        LAST_COL, height=36)
    row += 1

    # --------------------------------------- the fortnight strip -------
    row = s.section(row, "The next fortnight, every property",
                    STRIP_PROPERTIES + 2)
    row = s.note(
        row,
        "Fourteen days across, every property down. This is the view "
        "worth printing on a Friday: O is occupied, A is an arrival, "
        "D is a departure, T is a same-day turnover, M is closed for "
        "maintenance, B is blocked by you, and a blank cell is a "
        "night you are not selling.",
        STRIP_PROPERTIES + 2, height=38)
    row += 1

    s.write(row, 0, "Property", fmts.chrome("header_left"))
    s.set_col(0, 0, 22)
    for offset in range(CALENDAR_STRIP_DAYS):
        s.formula(
            row, offset + 1,
            f'=TEXT(TodayDate+{offset},"ddd")&" "&'
            f'TEXT(TodayDate+{offset},"d mmm")',
            fmts.chrome("header"))
        s.set_col(offset + 1, offset + 1, 7)
    s.row_height(row, 24)
    row += 1

    strip_top = row
    for n in range(1, STRIP_PROPERTIES + 1):
        pid = f'INDEX([[Property Database.id]],{n})'
        s.formula(
            row, 0,
            f'=IFERROR(IF(INDEX([[Property Database.name]],{n})="","",'
            f'INDEX([[Property Database.name]],{n})),"")',
            fmts.calc("text"))
        for offset in range(CALENDAR_STRIP_DAYS):
            day = f"(TodayDate+{offset})"
            s.formula(
                row, offset + 1,
                f'=IFERROR(IF(INDEX([[Property Database.name]],{n})="",'
                f'"",'
                f'IF({_blocked(day, "Maintenance block", pid)}>0,"M",'
                f'IF({_blocked(day, "Owner block", pid)}>0,"B",'
                f'IF({_departs(day, pid)}+{_arrives(day, pid)}>1,"T",'
                f'IF({_arrives(day, pid)}>0,"A",'
                f'IF({_departs(day, pid)}>0,"D",'
                f'IF({_occupied(day, pid)}>0,"O","")'
                f')))))),"")',
                fmts.calc("text"))
        row += 1
    strip_bottom = row - 1
    for text, style in (("T", "bold_bad"), ("A", "good"),
                        ("D", "info"), ("O", "muted"),
                        ("M", "bad"), ("B", "warn")):
        ws.conditional_format(
            strip_top, 1, strip_bottom, CALENDAR_STRIP_DAYS, {
                "type": "cell", "criteria": "equal to",
                "value": f'"{text}"', "format": fmts.cf(style)})
    row += 1

    # ----------------------------------- turnover command center -------
    row = s.section(row, "Turnover command centre", STRIP_PROPERTIES + 2)
    row = s.note(
        row,
        "Everything happening in the next twenty-four hours, in one "
        "place. If these four numbers are all accounted for, nobody "
        "arrives to a dirty flat.",
        STRIP_PROPERTIES + 2, height=26)

    s.kpi(row, 0, "Leaving today",
          '=COUNTIFS([[Booking Tracker.check_out]],TodayDate,'
          '[[Booking Tracker.earns]],"Yes")', "int")
    s.kpi(row, 2, "Arriving today",
          '=COUNTIFS([[Booking Tracker.check_in]],TodayDate,'
          '[[Booking Tracker.earns]],"Yes")', "int")
    s.kpi(row, 4, "Arriving tomorrow",
          '=COUNTIFS([[Booking Tracker.check_in]],TodayDate+1,'
          '[[Booking Tracker.earns]],"Yes")', "int")
    s.kpi(row, 6, "Cleans booked today",
          '=COUNTIFS([[Cleaning Schedule.clean_date]],TodayDate)',
          "int")
    s.kpi(row, 8, "Cleans still open",
          '=COUNTIFS([[Cleaning Schedule.clean_date]],"<="&TodayDate,'
          '[[Cleaning Schedule.status]],"<>Completed")', "int",
          note="Due today or earlier and not finished")
    row = s.kpi(row, 10, "Ready for a guest",
                '=COUNTIFS([[Property Database.alert]],"OK*")', "int")
    row += 1

    row = s.table_head(row, 0, [
        "Property", "Out today", "In today", "Clean booked",
        "Cleaner", "Turn", "Hours", "State"],
        [22, 10, 10, 13, 18, 8, 9, 26])
    cc_top = row
    for n in range(1, STRIP_PROPERTIES + 1):
        pid = f'INDEX([[Property Database.id]],{n})'
        guard = f'INDEX([[Property Database.name]],{n})=""'
        today = "TodayDate"
        clean_here = ('COUNTIFS([[Cleaning Schedule.property]],'
                      f'{pid},[[Cleaning Schedule.clean_date]],'
                      f'{today})')
        clean_idx = ('SUMPRODUCT(MAX(([[Cleaning Schedule.property]]='
                     f'{pid})*([[Cleaning Schedule.clean_num]]='
                     f'{today})*[[Cleaning Schedule.seq_num]]))')
        s.formula(
            row, 0,
            f'=IFERROR(IF({guard},"",INDEX('
            f'[[Property Database.name]],{n})),"")', fmts.calc("text"))
        s.formula(
            row, 1,
            f'=IFERROR(IF({guard},"",{_departs(today, pid)}),"")',
            fmts.calc("int"))
        s.formula(
            row, 2,
            f'=IFERROR(IF({guard},"",{_arrives(today, pid)}),"")',
            fmts.calc("int"))
        s.formula(
            row, 3,
            f'=IFERROR(IF({guard},"",{clean_here}),"")',
            fmts.calc("int"))
        s.formula(
            row, 4,
            f'=IFERROR(IF({guard},"",IF({clean_idx}<=0,"",'
            f'INDEX([[Cleaning Schedule.cleaner_name]],{clean_idx})&"")),'
            f'"")',
            fmts.calc("text"))
        s.formula(
            row, 5,
            f'=IFERROR(IF({guard},"",IF({clean_idx}<=0,"",'
            f'INDEX([[Cleaning Schedule.same_day]],{clean_idx})&"")),'
            f'"")',
            fmts.calc("text"))
        s.formula(
            row, 6,
            f'=IFERROR(IF({guard},"",IF({clean_idx}<=0,"",'
            f'INDEX([[Cleaning Schedule.window_hours]],{clean_idx})*1)),'
            f'"")',
            fmts.calc("number1"))
        s.formula(
            row, 7,
            f'=IFERROR(IF({guard},"",'
            f'IF({_departs(today, pid)}+{_arrives(today, pid)}>1,'
            f'IF({clean_here}<=0,"Urgent - turnover with no clean '
            f'booked","Same day - out and in"),'
            f'IF({_arrives(today, pid)}>0,"Check-in - guest arriving",'
            f'IF({_departs(today, pid)}>0,'
            f'IF({clean_here}<=0,"Action - nobody booked to clean",'
            f'"Check-out - clean booked"),'
            f'IF({_occupied(today, pid)}>0,"Guest in","Empty"))))),"")',
            fmts.calc("text"))
        row += 1
    for text, style in (("Urgent", "bold_bad"), ("Action", "bad"),
                        ("Same day", "warn"), ("Check-in", "good"),
                        ("Check-out", "info"), ("Empty", "muted")):
        ws.conditional_format(cc_top, 7, row - 1, 7, {
            "type": "text", "criteria": "containing", "value": text,
            "format": fmts.cf(style)})
    row += 1

    row = s.note(
        row,
        "'Hours' is the gap between the departing guest's check-out "
        "time and the arriving guest's check-in time. Anything under "
        "the tight-turnaround setting goes red on the Cleaning "
        "Schedule, because that is the day a cleaner needs telling in "
        "advance rather than at eleven o'clock.",
        STRIP_PROPERTIES + 2, height=34)

    s.finish(freeze_row=picker_row + 3, gridlines=False, fit_width=True,
             print_area=f"A1:{col_letter(LAST_COL)}{grid_bottom + 3}")
    s.protect()
    return s
````

---

## 5. The turnover

What happens between one guest leaving and the next arriving. A cleaning job is generated from a check-out rather than typed from memory; the nine arrival steps are ticked off before a guest lands; and the printable checklist is a one page card driven by a single job ID that writes its nine report-back flags straight back into the schedule.

### `airbnb_rental_spreadsheet/sheets/arrivals.py`

*234 lines · 2 functions*

> `Check-in & Check-out` - the eight messages every guest needs.
>
> Hosting is mostly admin, and the admin is always the same eight things:
> the directions, the door code, the Wi-Fi, the house rules, the welcome
> message, the check-out instructions, the walk round afterwards and the
> photos.  Forget one and you get a one-star review about communication.
>
> So this tab is a tick list with a score.  Each row is one stay, each
> column is one of the things that has to be sent or done, and the
> readiness column tells you - in one glance down the sheet - which guest
> is still waiting to hear from you.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | How on top of the admin you actually are |

**Module constants:** `NAME` ('Check-in & Check-out'), `BEFORE_ARRIVAL` (5 items), `AFTER_DEPARTURE` (4 items), `ALL_STEPS` (BEFORE_ARRIVAL + AFTER_DEPARTURE).

````python
"""`Check-in & Check-out` - the eight messages every guest needs.

Hosting is mostly admin, and the admin is always the same eight things:
the directions, the door code, the Wi-Fi, the house rules, the welcome
message, the check-out instructions, the walk round afterwards and the
photos.  Forget one and you get a one-star review about communication.

So this tab is a tick list with a score.  Each row is one stay, each
column is one of the things that has to be sent or done, and the
readiness column tells you - in one glance down the sheet - which guest
is still waiting to hear from you.
"""

from __future__ import annotations

from typing import List, Tuple

from ..config import Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, time_check
from .common import (alert_cf, blank, flag_yes_no_cf, id_col, 
                    lookup_booking, month_col, weekday_col, yesno)

NAME = "Check-in & Check-out"

#: (column key, header, when it matters, help text)
BEFORE_ARRIVAL: List[Tuple[str, str, str]] = [
    ("sent_directions", "Directions sent",
     "How to find the place, parking, which door."),
    ("sent_access", "Access code sent",
     "Key box code, smart lock PIN or who is meeting them."),
    ("sent_wifi", "Wi-Fi details sent",
     "Network and password. The single most asked question."),
    ("sent_rules", "House rules sent",
     "Quiet hours, smoking, pets, extra visitors."),
    ("sent_welcome", "Welcome message sent",
     "The day before or the morning of arrival."),
]

AFTER_DEPARTURE: List[Tuple[str, str, str]] = [
    ("sent_checkout", "Check-out notes sent",
     "What to do with the keys, the bins and the heating."),
    ("checked_out", "Checked out",
     "They have gone and the keys are back."),
    ("inspected", "Property inspected",
     "Someone has walked round it since they left."),
    ("photographed", "Condition photographed",
     "Your evidence if a deposit claim ever comes up."),
]

ALL_STEPS = BEFORE_ARRIVAL + AFTER_DEPARTURE


def spec(cfg: Config) -> SheetSpec:
    steps_done = "+".join(
        f'({{{key}}}{{r}}="Yes")' for key, _h, _n in ALL_STEPS)
    before_done = "+".join(
        f'({{{key}}}{{r}}="Yes")' for key, _h, _n in BEFORE_ARRIVAL)

    columns = [
        id_col("booking", "ArrivalPrefix", digits=cfg.digits("arrival"),
               width=10, header="Arrival ID"),
        Col("booking", "Booking", 11, "text",
            validate=list_from("BookingList"),
            note="Pick the booking reference. Everything on the left of "
                 "this sheet is pulled from it.",
            group="The stay"),
        Col("property", "Property", 11, "text",
            formula=lookup_booking("property"), hidden=True,
            group="The stay"),
        Col("property_name", "Property", 20, "text",
            formula=lookup_booking("property_name"), group="The stay"),
        Col("guest_name", "Guest", 20, "text",
            formula=lookup_booking("guest_name"), group="The stay"),
        Col("arrival_date", "Arrives", 12, "date",
            formula=lookup_booking("check_in"), group="The stay"),
        weekday_col("arrival_date", "arr_day", "Day", 6, hidden=True),
        Col("departure_date", "Leaves", 12, "date",
            formula=lookup_booking("check_out"), group="The stay"),
        Col("nights", "Nights", 8, "int",
            formula=lookup_booking("nights"), hidden=True,
            group="The stay"),
        month_col("arrival_date"),

        # ----------------------------------------------- what happened --
        Col("arrival_time", "Arrived at", 10, "time",
            validate=time_check(),
            note="The actual arrival time, if you know it. Leave blank "
                 "for a self check-in you never saw.",
            group="On the day"),
        Col("departure_time", "Left at", 10, "time",
            validate=time_check(), group="On the day"),
        Col("late_checkout", "Late check-out agreed", 14, "flag",
            validate=yesno("Did you agree a later departure?"),
            cf=flag_yes_no_cf(), group="On the day"),
        Col("early_checkin", "Early check-in agreed", 14, "flag",
            validate=yesno(), cf=flag_yes_no_cf(), group="On the day"),
        Col("key_method", "Access used", 16, "text",
            note="Key box, smart lock, met in person. Worth recording "
                 "when a code goes wrong.",
            group="On the day"),
    ]

    for key, header, note in BEFORE_ARRIVAL:
        columns.append(
            Col(key, header, 13, "flag", validate=yesno(note),
                cf=flag_yes_no_cf(), note=note, group="Before they arrive"))

    columns.append(
        Col("pre_arrival_done", "Pre-arrival score", 13, "percent",
            formula=blank("booking",
                          f"({before_done})/{len(BEFORE_ARRIVAL)}"),
            cf=[{"type": "cell", "criteria": "<", "value": 1,
                 "fmt": "warn"},
                {"type": "cell", "criteria": ">=", "value": 1,
                 "fmt": "good"}],
            note="Everything the guest needs before they travel.",
            group="Before they arrive"))

    for key, header, note in AFTER_DEPARTURE:
        columns.append(
            Col(key, header, 13, "flag", validate=yesno(note),
                cf=flag_yes_no_cf(), note=note, group="After they leave"))

    columns += [
        Col("steps_done", "Done", 7, "int",
            formula=blank("booking", f"({steps_done})"),
            hidden=True, group="Where it stands"),
        Col("readiness", "Complete", 10, "percent",
            formula=blank("booking", f"({steps_done})/{len(ALL_STEPS)}"),
            cf=[{"type": "data_bar", "bar_color": "#3F97A0"}],
            group="Where it stands"),
        Col("outstanding", "Still to do", 11, "int",
            formula=blank("booking",
                          f"{len(ALL_STEPS)}-({steps_done})"),
            group="Where it stands"),
        Col("state", "Where it is", 14, "status",
            formula=blank(
                "booking",
                'IF({arrival_date}{r}="","",'
                'IF(TodayDate<{arrival_date}{r},"Upcoming",'
                'IF(TodayDate={arrival_date}{r},"Arriving today",'
                'IF({departure_date}{r}="","Staying",'
                'IF(TodayDate<{departure_date}{r},"Staying",'
                'IF(TodayDate={departure_date}{r},"Leaving today",'
                '"Past"))))))'),
            cf=alert_cf(), group="Where it stands"),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "booking",
                'IF({arrival_date}{r}="","Check - booking not found",'
                'IF({state}{r}="Upcoming",'
                'IF({arrival_date}{r}-TodayDate>UpcomingDays,"",'
                'IF({pre_arrival_done}{r}<1,"Send - '
                'arrival details still missing","OK - ready for them")),'
                'IF({state}{r}="Arriving today",'
                'IF({pre_arrival_done}{r}<1,"Send - '
                'they arrive today and have not been told everything",'
                '"OK - ready for them"),'
                'IF({state}{r}="Staying","OK - in residence",'
                'IF({state}{r}="Leaving today",'
                'IF({sent_checkout}{r}<>"Yes","Send - check-out notes",'
                '"OK - leaving today"),'
                'IF({checked_out}{r}<>"Yes","Check - never marked '
                'checked out",'
                'IF({inspected}{r}<>"Yes","Action - nobody has '
                'inspected it",'
                'IF({photographed}{r}<>"Yes","Action - no condition '
                'photos","OK"))))))))'),
            cf=alert_cf(), group="Where it stands"),
        Col("issues", "Anything wrong", 30, "wrap",
            note="Damage, a missing key, a complaint. If it costs money, "
                 "log it on the Maintenance Log too.",
            group="Notes"),
        Col("notes", "Notes", 28, "wrap", group="Notes"),
    ]

    return SheetSpec(
        name=NAME,
        title="Check-in & Check-out",
        subtitle="The eight things every guest needs, and which of them "
                 "you have actually done",
        columns=columns,
        rows=cfg.row_count(NAME),
        freeze_cols=2,
        samples=sample("arrivals", cfg),
        help_text="One row per stay. Pick the booking and the property, "
                  "guest and dates arrive with it. Then tick off the "
                  "nine steps. The Action column reads the dates and "
                  "the ticks together, so it only nags you about the "
                  "guests who are actually coming.",
        extra_notes=[
            "The pre-arrival score is the one to watch. A guest who has "
            "not been sent a door code the day before arrival is "
            "tomorrow's problem, and this is the column that spots it "
            "today.",
            "Condition photographs are worth the thirty seconds. They "
            "are the only evidence you will have if you ever need to "
            "claim against a deposit.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """How on top of the admin you actually are."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index

    row = sheet.section(row, "How on top of it you are", last)
    row = sheet.table_head(row, 0, [
        "Stays logged", "Upcoming", "In residence", "Leaving today",
        "Fully prepared", "Missing something", "Average complete",
        "Not inspected", "No photos"],
        [12, 11, 12, 13, 14, 15, 15, 13, 11])
    sheet.calc_row(row, 0, [
        '=COUNTIF([[Check-in & Check-out.booking]],"?*")',
        '=COUNTIFS([[Check-in & Check-out.state]],"Upcoming")',
        '=COUNTIFS([[Check-in & Check-out.state]],"Staying")',
        '=COUNTIFS([[Check-in & Check-out.state]],"Leaving today")',
        '=COUNTIFS([[Check-in & Check-out.pre_arrival_done]],1)',
        '=COUNTIFS([[Check-in & Check-out.state]],"Upcoming",'
        '[[Check-in & Check-out.pre_arrival_done]],"<1")+'
        'COUNTIFS([[Check-in & Check-out.state]],"Arriving today",'
        '[[Check-in & Check-out.pre_arrival_done]],"<1")',
        '=IF(COUNTIF([[Check-in & Check-out.booking]],"?*")<=0,"",'
        'SUM([[Check-in & Check-out.steps_done]])/'
        f'(COUNTIF([[Check-in & Check-out.booking]],"?*")*{len(ALL_STEPS)}))',
        '=COUNTIFS([[Check-in & Check-out.state]],"Past",'
        '[[Check-in & Check-out.inspected]],"<>Yes")',
        '=COUNTIFS([[Check-in & Check-out.state]],"Past",'
        '[[Check-in & Check-out.photographed]],"<>Yes")',
    ], ["int", "int", "int", "int", "int", "int", "percent", "int", "int"])
````

### `airbnb_rental_spreadsheet/sheets/cleaning.py`

*318 lines · 2 functions*

> `Cleaning Schedule` - the turnover, who does it and what it costs.
>
> The cleaning is the business.  A host can get the pricing wrong and
> survive; get a turnover wrong and the next guest walks into someone
> else's weekend.  This tab is built around the two numbers that actually
> matter on a changeover day: how long the cleaner has, and whether the
> property is finished before the next arrival.
>
> The cleaning date defaults to the booking's check-out date, so a job
> created straight off a reservation needs one click, not six.  Nine task
> flags drive a completion percentage, and the printable room-by-room
> checklist on the next tab reads the same nine flags.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Turnover totals and the cleaners' league table |

**Module constants:** `NAME` ('Cleaning Schedule').

````python
"""`Cleaning Schedule` - the turnover, who does it and what it costs.

The cleaning is the business.  A host can get the pricing wrong and
survive; get a turnover wrong and the next guest walks into someone
else's weekend.  This tab is built around the two numbers that actually
matter on a changeover day: how long the cleaner has, and whether the
property is finished before the next arrival.

The cleaning date defaults to the booking's check-out date, so a job
created straight off a reservation needs one click, not six.  Nine task
flags drive a completion percentage, and the printable room-by-room
checklist on the next tab reads the same nine flags.
"""

from __future__ import annotations

from ..config import CLEANING_TASKS, Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check, time_check
from .common import (alert_cf, blank, date_num_col, seq_num_col, flag_yes_no_cf, 
                    id_col, lookup_booking, lookup_booking_num, 
                    lookup_vendor, month_col, status_cf, weekday_col, 
                    yesno)

NAME = "Cleaning Schedule"


def spec(cfg: Config) -> SheetSpec:
    bk_prop = "[[Booking Tracker.property]]"
    bk_ci = "[[Booking Tracker.ci_num]]"
    bk_holds = "[[Booking Tracker.holds]]"

    tasks_done = "+".join(
        f'({{{key}}}{{r}}="Yes")' for key, _h in CLEANING_TASKS)

    columns = [
        id_col("booking", "CleanPrefix", digits=cfg.digits("clean"),
               width=10, header="Job ID"),
        Col("booking", "Booking", 11, "text",
            validate=list_from("BookingList"),
            note="The stay that is ending. Pick it and the property, "
                 "the guest and the check-out date come with it.",
            group="The job"),
        Col("property", "Property", 11, "text",
            formula=lookup_booking("property"), hidden=True,
            group="The job"),
        Col("property_name", "Property", 20, "text",
            formula=lookup_booking("property_name"), group="The job"),
        Col("guest_name", "Guest leaving", 18, "text",
            formula=lookup_booking("guest_name"), group="The job"),
        Col("checkout_date", "Check-out", 12, "date",
            formula=lookup_booking("check_out"), group="The job"),
        Col("checkout_num", "Check-out (number)", 10, "number",
            formula=('=IF({booking}{r}="",0,IFERROR(INDEX('
                     '[[Booking Tracker.co_num]],MATCH({booking}{r},'
                     '[[Booking Tracker.id]],0))*1,0))'),
            hidden=True, group="The job"),
        Col("date_in", "Cleaning date", 12, "date",
            validate=date_check(),
            note="Leave it blank and the job happens on the check-out "
                 "date, which is what you want nine times in ten. Type "
                 "a date only when it slips.",
            group="The job"),
        Col("clean_date", "Cleaned on", 12, "date",
            formula=('=IF({booking}{r}="","",IF({date_in}{r}<>"",'
                     '{date_in}{r},{checkout_date}{r}))'),
            group="The job"),
        weekday_col("clean_date", "clean_day", "Day", 6, hidden=True),
        date_num_col("clean_date", "clean_num", "Cleaned on (number)"),
        seq_num_col("booking"),
        month_col("clean_date"),

        # ------------------------------------------------ who and when --
        Col("cleaner", "Cleaner", 11, "text",
            validate=list_from("VendorList"),
            note="Pick them from Staff & Vendors. The rate and the "
                 "contact details come from there.",
            group="Who and when"),
        Col("cleaner_name", "Cleaner name", 18, "text",
            formula=lookup_vendor("name", "cleaner"), group="Who and when"),
        Col("start_time", "Started", 9, "time",
            validate=time_check(), group="Who and when"),
        Col("end_time", "Finished", 9, "time",
            validate=time_check(), group="Who and when"),
        Col("hours", "Hours", 8, "number1",
            formula=('=IF({booking}{r}="","",IF({start_time}{r}="","",'
                     'IF({end_time}{r}="","",'
                     'ROUND(MOD({end_time}{r}-{start_time}{r},1)*24,2))))'),
            note="Worked out for you, and it copes with a clean that "
                 "runs past midnight.",
            group="Who and when"),
        Col("status", "Status", 16, "status",
            validate=list_from("CleaningStatusList"),
            cf=status_cf(), group="Who and when"),
    ]

    for key, header in CLEANING_TASKS:
        columns.append(
            Col(key, header, 11, "flag", validate=yesno(),
                cf=flag_yes_no_cf(),
                note=f"{header} done. The printable turnover checklist "
                     f"reads this same flag.",
                group="The turnover"))

    columns += [
        Col("tasks_done", "Tasks done", 10, "int",
            formula=blank("booking", f"({tasks_done})"),
            hidden=True, group="The turnover"),
        Col("completion", "Complete", 10, "percent",
            formula=blank("booking",
                          f"({tasks_done})/{len(CLEANING_TASKS)}"),
            cf=[{"type": "cell", "criteria": ">=", "value": 1,
                 "fmt": "good"},
                {"type": "cell", "criteria": ">=", "value": 0.5,
                 "fmt": "warn"},
                {"type": "cell", "criteria": "<", "value": 0.5,
                 "fmt": "bad"}],
            note="Nine tasks, one percentage. Green only when every one "
                 "of them is ticked.",
            group="The turnover"),

        # ------------------------------------------------- the pressure --
        Col("next_arrival", "Next arrival", 12, "date",
            formula=blank(
                "booking",
                f'IF({{property}}{{r}}="","",'
                f'IF(SUMPRODUCT(MIN(({bk_prop}={{property}}{{r}})*'
                f'({bk_holds}="Yes")*({bk_ci}>={{checkout_num}}{{r}})*'
                f'{bk_ci}+(1-({bk_prop}={{property}}{{r}})*'
                f'({bk_holds}="Yes")*({bk_ci}>={{checkout_num}}{{r}}))*'
                f'2958465))>=2958465,"",'
                f'SUMPRODUCT(MIN(({bk_prop}={{property}}{{r}})*'
                f'({bk_holds}="Yes")*({bk_ci}>={{checkout_num}}{{r}})*'
                f'{bk_ci}+(1-({bk_prop}={{property}}{{r}})*'
                f'({bk_holds}="Yes")*({bk_ci}>={{checkout_num}}{{r}}))*'
                f'2958465))))'),
            note="The next check-in on this property on or after the "
                 "check-out. This is the deadline.",
            group="The pressure"),
        Col("gap_days", "Days spare", 10, "int",
            formula=blank("booking",
                          'IF({next_arrival}{r}="","",'
                          'MAX(0,{next_arrival}{r}-{checkout_date}{r}))'),
            group="The pressure"),
        Col("same_day", "Same-day turn", 12, "flag",
            formula=blank("booking",
                          'IF({next_arrival}{r}="","No",'
                          'IF({gap_days}{r}=0,"Yes","No"))'),
            cf=[{"type": "text", "criteria": "containing", "value": "Yes",
                 "fmt": "warn"}],
            note="Somebody checks in the same day. The cleaner has the "
                 "window between the two times and nothing more.",
            group="The pressure"),
        Col("window_hours", "Hours in the window", 14, "number1",
            formula=blank(
                "booking",
                'IF({same_day}{r}<>"Yes","",'
                'IF({property}{r}="","",'
                'ROUND(MAX(0,'
                'IFERROR(INDEX([[Property Database.check_in_time]],'
                'MATCH({property}{r},[[Property Database.id]],0))*1,0.625)-'
                'IFERROR(INDEX([[Property Database.check_out_time]],'
                'MATCH({property}{r},[[Property Database.id]],0))*1,'
                '0.4583))*24,1)))'),
            cf=[{"type": "cell", "criteria": "<",
                 "value": "=TightTurnHours", "fmt": "bad"}],
            note="Check-out time to check-in time on a same-day turn. "
                 "Under the warning hours on Settings and it goes red.",
            group="The pressure"),
        Col("ready", "Ready for the next guest", 15, "status",
            formula=blank(
                "booking",
                'IF({completion}{r}>=1,"Ready",'
                'IF({next_arrival}{r}="","Not cleaned",'
                'IF(TodayDate>{next_arrival}{r},"Ready",'
                'IF({next_arrival}{r}-TodayDate<=1,'
                '"Not cleaned - guest due","Not cleaned"))))'),
            cf=alert_cf(), group="The pressure"),

        # ----------------------------------------------------- the cost --
        Col("cost_in", "Cleaning cost", 12, "money",
            validate=number_check(0, 100000, "Cleaning cost"),
            note="Blank uses the property's standard cleaning cost.",
            group="The cost"),
        Col("clean_cost", "Cost used", 11, "money",
            formula=blank(
                "booking",
                'IF({cost_in}{r}<>"",{cost_in}{r},'
                'IF({property}{r}="",0,'
                'IFERROR(INDEX([[Property Database.cleaning_cost]],'
                'MATCH({property}{r},[[Property Database.id]],0))*1,0)))'),
            group="The cost"),
        Col("laundry_cost", "Laundry", 10, "money",
            validate=number_check(0, 100000, "Laundry"),
            note="Blank uses the standard laundry cost on Settings.",
            group="The cost"),
        Col("supplies_cost", "Consumables", 11, "money",
            validate=number_check(0, 100000, "Consumables"),
            group="The cost"),
        Col("total_cost", "Turnover cost", 13, "money",
            formula=blank(
                "booking",
                'ROUND({clean_cost}{r}+'
                'IF({laundry_cost}{r}="",LaundryPerTurn,{laundry_cost}{r})+'
                'IF({supplies_cost}{r}="",SuppliesPerTurn,'
                '{supplies_cost}{r}),2)'),
            note="Cleaning plus laundry plus consumables. This is the "
                 "real cost of a changeover, and it is what the profit "
                 "per booking is built on.",
            group="The cost"),
        Col("fee_charged", "Fee charged to guest", 14, "money",
            formula=lookup_booking_num("cleaning_fee"),
            hidden=True, group="The cost"),
        Col("cost_gap", "Over or under recovered", 15, "money",
            formula=blank("booking",
                          'ROUND({fee_charged}{r}-{total_cost}{r},2)'),
            cf=[{"type": "cell", "criteria": "<", "value": -0.004,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 0.004,
                 "fmt": "good"}],
            note="The cleaning fee you charged the guest less what the "
                 "turnover actually cost you. Red means you are "
                 "subsidising the clean out of the room rate.",
            group="The cost"),
        Col("paid", "Cleaner paid", 11, "flag",
            validate=yesno(), cf=flag_yes_no_cf(), group="The cost"),

        # ------------------------------------------------ what to do now --
        Col("alert", "Action", 28, "status",
            formula=blank(
                "booking",
                'IF({checkout_date}{r}="","Check - booking not found",'
                'IF({status}{r}="Cancelled","",'
                'IF({cleaner}{r}="","Add - nobody assigned",'
                'IF({clean_date}{r}>TodayDate,'
                'IF({same_day}{r}="Yes","Same day - tight turnaround '
                'coming","OK - scheduled"),'
                'IF({completion}{r}>=1,'
                'IF({paid}{r}<>"Yes","Unpaid - cleaner not paid yet",'
                '"OK"),'
                'IF({next_arrival}{r}="","Action - not finished",'
                'IF({next_arrival}{r}<=TodayDate+1,'
                '"Not cleaned - guest due","Action - not finished")))))))'),
            cf=alert_cf(), group="What to do"),
        Col("notes", "Notes", 30, "wrap",
            note="Damage found, items left behind, anything the next "
                 "cleaner should know.",
            group="What to do"),
    ]

    return SheetSpec(
        name=NAME,
        title="Cleaning Schedule",
        subtitle="Every turnover: who is doing it, how long they have, "
                 "and what it really costs",
        columns=columns,
        rows=cfg.row_count(NAME),
        freeze_cols=2,
        samples=sample("cleaning", cfg),
        help_text="One row per turnover. Pick the booking that is ending "
                  "and the property, guest and check-out date arrive "
                  "with it - leave the cleaning date blank and the job "
                  "lands on the check-out day by itself. Tick the nine "
                  "tasks as they are done; the printable checklist on "
                  "the next tab reads the same ticks.",
        extra_notes=[
            "The turnover cost is cleaning plus laundry plus "
            "consumables, not just the cleaner's fee. Compare it with "
            "the cleaning fee you charge the guest - the 'over or "
            "under recovered' column is one of the quietest money "
            "leaks in this business.",
            "Hours in the window only appears on a same-day turn. It is "
            "the gap between your check-out and check-in times, so if "
            "it is too tight the answer is usually to move one of them "
            "on the Property Database tab.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Turnover totals and the cleaners' league table."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index

    row = sheet.section(row, "Every turnover on this tab", last)
    row = sheet.table_head(row, 0, [
        "Turnovers", "Complete", "Outstanding", "Same-day turns",
        "Cleaning paid", "Laundry & consumables", "Total cost",
        "Fees charged", "Recovered", "Avg cost"],
        [11, 10, 12, 14, 13, 18, 12, 13, 11, 10])
    total = '=ROUND(SUM([[Cleaning Schedule.total_cost]]),2)'
    jobs = '=COUNTIF([[Cleaning Schedule.booking]],"?*")'
    sheet.calc_row(row, 0, [
        jobs,
        '=COUNTIFS([[Cleaning Schedule.completion]],1)',
        '=COUNTIFS([[Cleaning Schedule.completion]],"<1")',
        '=COUNTIFS([[Cleaning Schedule.same_day]],"Yes")',
        '=ROUND(SUM([[Cleaning Schedule.clean_cost]]),2)',
        '=ROUND(SUM([[Cleaning Schedule.total_cost]])-'
        'SUM([[Cleaning Schedule.clean_cost]]),2)',
        total,
        '=ROUND(SUM([[Cleaning Schedule.fee_charged]]),2)',
        '=IF(SUM([[Cleaning Schedule.total_cost]])<=0,"",'
        'SUM([[Cleaning Schedule.fee_charged]])/'
        'SUM([[Cleaning Schedule.total_cost]]))',
        '=IF(COUNTIF([[Cleaning Schedule.booking]],"?*")<=0,"",'
        'ROUND(SUM([[Cleaning Schedule.total_cost]])/'
        'COUNTIF([[Cleaning Schedule.booking]],"?*"),2))',
    ], ["int", "int", "int", "int", "money", "money", "money", "money",
        "percent", "money"])
    row += 2

    sheet.note(row, "'Recovered' is the cleaning fee you charge guests "
                    "as a percentage of what turnovers actually cost "
                    "you. Under 100% means the room rate is quietly "
                    "paying for the cleaning.", last)
````

### `airbnb_rental_spreadsheet/sheets/turnover.py`

*245 lines · 2 functions*

> `Turnover Checklist` - one changeover, on one sheet of paper.
>
> Type a cleaning job ID at the top and this page becomes that job:
> the property, the address, the wifi, the access code, when the next
> guest arrives and how long there is to do it.  Underneath, five
> groups of tick boxes - bedroom, bathroom, kitchen, living space and
> a final inspection - with somewhere to sign.
>
> Print it, hand it over, get it back.  The nine summary flags at the
> bottom are the ones you transfer to the Cleaning Schedule afterwards,
> which is what makes the completion percentage on that tab mean
> something.

| Defines | What it does |
|---------|--------------|
| `_job(field, cast)` | A field from the chosen cleaning job, blank when nothing chosen |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Turnover Checklist'), `LAST_COL` (7).

````python
"""`Turnover Checklist` - one changeover, on one sheet of paper.

Type a cleaning job ID at the top and this page becomes that job:
the property, the address, the wifi, the access code, when the next
guest arrives and how long there is to do it.  Underneath, five
groups of tick boxes - bedroom, bathroom, kitchen, living space and
a final inspection - with somewhere to sign.

Print it, hand it over, get it back.  The nine summary flags at the
bottom are the ones you transfer to the Cleaning Schedule afterwards,
which is what makes the completion percentage on that tab mean
something.
"""

from __future__ import annotations

from ..config import (CLEANING_TASKS, Config, Palette, TURNOVER_CHECKLIST)
from ..layout import Registry, col_letter
from ..writer import SheetWriter

NAME = "Turnover Checklist"
LAST_COL = 7


def _job(field: str, cast: str = '&""') -> str:
    """A field from the chosen cleaning job, blank when nothing chosen."""
    return (f'=IF(JobId="","",IFERROR(INDEX('
            f'[[Cleaning Schedule.{field}]],MATCH(JobId,'
            f'[[Cleaning Schedule.id]],0)){cast},""))')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_soft)
    ws = s.ws

    s.banner("Turnover Checklist",
             "One job, one page. Type the cleaning job ID and print "
             "it.", LAST_COL)
    s.set_col(0, 0, 4)
    s.set_col(1, 1, 30)
    s.set_col(2, 2, 10)
    s.set_col(3, 3, 26)
    s.set_col(4, 4, 10)
    s.set_col(5, 5, 24)
    s.set_col(6, LAST_COL, 16)

    row = 4
    s.merge(row, 1, 1, "Cleaning job ID", fmts.chrome("label_bold"))
    s.write(row, 2, None, fmts.chrome("form_box", locked=False))
    book.define_name("JobId", f"='{NAME}'!$C${row + 1}")
    s.merge(row, 3, LAST_COL,
            "From the Cleaning Schedule - for example CLN-0001. "
            "Everything below fills itself in.",
            fmts.chrome("note"))
    s.row_height(row, 24)
    row += 2

    # ----------------------------------------------------- the job card --
    row = s.section(row, "The job", LAST_COL)
    card = [
        ("Property", _job("property_name"), "Address",
         '=IF(JobId="","",IFERROR(INDEX('
         '[[Property Database.address]],MATCH(INDEX('
         '[[Cleaning Schedule.property]],MATCH(JobId,'
         '[[Cleaning Schedule.id]],0)),'
         '[[Property Database.id]],0))&"",""))'),
        ("Clean on", _job("clean_date", "*1"), "Cleaner",
         _job("cleaner_name")),
        ("Guest leaving", _job("guest_name"), "Checked out at",
         '=IF(JobId="","",IFERROR(INDEX('
         '[[Property Database.check_out_time]],MATCH(INDEX('
         '[[Cleaning Schedule.property]],MATCH(JobId,'
         '[[Cleaning Schedule.id]],0)),'
         '[[Property Database.id]],0))&"",""))'),
        ("Next guest arrives", _job("next_arrival", "*1"),
         "Hours to do it", _job("window_hours", "*1")),
        ("Same-day turnover", _job("same_day"), "Access code",
         '=IF(JobId="","",IFERROR(INDEX('
         '[[Property Database.access]],MATCH(INDEX('
         '[[Cleaning Schedule.property]],MATCH(JobId,'
         '[[Cleaning Schedule.id]],0)),'
         '[[Property Database.id]],0))&"",""))'),
        ("Wifi network",
         '=IF(JobId="","",IFERROR(INDEX('
         '[[Property Database.wifi_name]],MATCH(INDEX('
         '[[Cleaning Schedule.property]],MATCH(JobId,'
         '[[Cleaning Schedule.id]],0)),'
         '[[Property Database.id]],0))&"",""))',
         "Who to ring", "=HostName&\" - \"&OwnerName"),
    ]
    styles = {"Clean on": "date", "Next guest arrives": "date",
              "Hours to do it": "number1"}
    for left_label, left_formula, right_label, right_formula in card:
        s.write(row, 1, left_label, fmts.chrome("label_bold"))
        s.merge(row, 2, 2, None,
                fmts.calc(styles.get(left_label, "text")))
        s.formula(row, 2, s.reg.render(left_formula),
                         fmts.calc(styles.get(left_label, "text")))
        s.write(row, 3, right_label, fmts.chrome("label_bold"))
        s.merge(row, 4, LAST_COL, None,
                fmts.calc(styles.get(right_label, "text")))
        s.formula(row, 4, s.reg.render(right_formula),
                         fmts.calc(styles.get(right_label, "text")))
        s.row_height(row, 20)
        row += 1

    s.write(row, 1, "Warning", fmts.chrome("label_bold"))
    s.merge(row, 2, LAST_COL, None, fmts.chrome("note_box"))
    s.formula(
        row, 2,
        '=IF(JobId="","Type a cleaning job ID above.",'
        'IF(COUNTIF([[Cleaning Schedule.id]],JobId)=0,'
        '"No job with that ID - check the Cleaning Schedule.",'
        'IF(IFERROR(INDEX([[Cleaning Schedule.same_day]],MATCH(JobId,'
        '[[Cleaning Schedule.id]],0))&"","")="Yes",'
        '"SAME-DAY TURNOVER. A guest arrives today - finish and '
        'report before check-in time.",'
        '"No arrival today. Normal turnaround.")))',
        fmts.chrome("note_box"))
    s.row_height(row, 28)
    ws.conditional_format(row, 2, row, LAST_COL, {
        "type": "text", "criteria": "containing", "value": "SAME-DAY",
        "format": fmts.cf("bold_bad")})
    row += 2

    # ----------------------------------------------- the tick boxes -----
    row = s.section(row, "The checklist", LAST_COL)
    row = s.note(
        row,
        "Tick as you go. Anything you cannot finish goes in the notes "
        "at the bottom - a cleaner who writes 'kettle broken' saves "
        "the next guest's review.",
        LAST_COL, height=24)
    row += 1

    for key, group, items in TURNOVER_CHECKLIST:
        s.merge(row, 0, 3, group, fmts.chrome("subsection"))
        s.merge(row, 4, 5, "Tick", fmts.chrome("subsection"))
        s.merge(row, 6, LAST_COL, None, fmts.chrome("subsection"))
        s.formula(
            row, 6,
            s.reg.render(
                f'=IF(JobId="","",IFERROR(IF(INDEX('
                f'[[Cleaning Schedule.{key}]],MATCH(JobId,'
                f'[[Cleaning Schedule.id]],0))="Yes","Signed off",'
                f'"Outstanding"),""))'),
            fmts.chrome("subsection"))
        s.row_height(row, 20)
        row += 1
        for item in items:
            s.write(row, 0, None, fmts.chrome("form_line", locked=False))
            s.merge(row, 1, 3, item, fmts.chrome("label"))
            s.merge(row, 4, LAST_COL, None,
                    fmts.chrome("form_line", locked=False))
            s.row_height(row, 19)
            row += 1
        row += 1

    # --------------------------------------------- the nine flags -------
    row = s.section(row, "Report back", LAST_COL)
    row = s.note(
        row,
        "These nine are the ones that go onto the Cleaning Schedule. "
        "Everything above is detail for the person doing the work; "
        "these are what the workbook counts.",
        LAST_COL, height=24)
    row = s.table_head(row, 0, ["", "Task", "Done?", "Notes"],
                       [4, 30, 10, 66])
    for key, label in CLEANING_TASKS:
        s.write(row, 0, None, fmts.chrome("form_line", locked=False))
        s.write(row, 1, label, fmts.chrome("label_bold"))
        s.merge(row, 2, 2, None, fmts.calc("flag"))
        s.formula(
            row, 2,
            s.reg.render(f'=IF(JobId="","",IFERROR(INDEX('
                         f'[[Cleaning Schedule.{key}]],MATCH(JobId,'
                         f'[[Cleaning Schedule.id]],0))&"",""))'),
            fmts.calc("flag"))
        s.merge(row, 3, LAST_COL, None,
                fmts.chrome("form_line", locked=False))
        s.row_height(row, 20)
        row += 1

    s.write(row, 1, "Completed", fmts.chrome("total_label"))
    s.merge(row, 2, 2, None,
            fmts.chrome("total_value", num_format="0.0%"))
    s.formula(
        row, 2,
        s.reg.render('=IF(JobId="","",IFERROR(INDEX('
                     '[[Cleaning Schedule.completion]],MATCH(JobId,'
                     '[[Cleaning Schedule.id]],0))*1,""))'),
        fmts.chrome("total_value", num_format="0.0%"))
    s.merge(row, 3, LAST_COL,
            "Red until every one of the nine is ticked on the "
            "Cleaning Schedule.", fmts.chrome("note"))
    ws.conditional_format(row, 2, row, 2, {
        "type": "cell", "criteria": ">=", "value": 1,
        "format": fmts.cf("good")})
    ws.conditional_format(row, 2, row, 2, {
        "type": "cell", "criteria": "<", "value": 1,
        "format": fmts.cf("bad")})
    row += 2

    # ---------------------------------------------------- sign off ------
    row = s.section(row, "Sign off", LAST_COL)
    for label in ("Anything broken or missing",
                  "Anything left behind by the guest",
                  "Supplies that need restocking"):
        s.write(row, 1, label, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, None,
                fmts.chrome("form_box", locked=False))
        s.row_height(row, 26)
        row += 1
    row += 1
    for left, right in (("Started at", "Finished at"),
                        ("Cleaner's signature", "Checked by")):
        s.write(row, 1, left, fmts.chrome("label_bold"))
        s.merge(row, 2, 3, None, fmts.chrome("form_line", locked=False))
        s.write(row, 4, right, fmts.chrome("label_bold"))
        s.merge(row, 5, LAST_COL, None,
                fmts.chrome("form_line", locked=False))
        s.row_height(row, 26)
        row += 1
    row += 1

    s.merge(row, 0, LAST_COL, None, fmts.chrome("print_sub"))
    s.formula(
        row, 0,
        '=HostName&"  -  turnover checklist  -  printed "&'
        'TEXT(TodayDate,"d mmm yyyy")', fmts.chrome("print_sub"))
    row += 1

    row = s.note(
        row,
        "Blank is fine. With no job ID typed this prints as an empty "
        "checklist you can keep a stack of in the cupboard - the "
        "property, the date and the cleaner just get written in by "
        "hand.",
        LAST_COL, height=26)

    s.finish(freeze_row=None, gridlines=False, fit_width=True,
             print_area=f"A1:{col_letter(LAST_COL)}{row + 2}")
    s.protect()
    return s
````

---

## 6. The money

Revenue first, cash second. A booking earns net revenue; a payout is what actually arrived, and the variance between the two is the column that pays for the spreadsheet. Expenses carry a tax category, a recurring frequency and a flag that stops a cleaning or maintenance cost being counted twice.

### `airbnb_rental_spreadsheet/sheets/payouts.py`

*248 lines · 2 functions*

> `Platform Revenue & Fees` - what the platform said, and what arrived.
>
> Every host has had the moment where the money in the bank does not
> match the money on the booking.  A cancelled night, a currency
> conversion, a resolution centre adjustment, a fee that changed when the
> account moved plan - none of it is visible from the reservation screen.
>
> This is the reconciliation ledger.  One row per payout: what the
> booking said you should get, what actually landed, and the difference
> between the two, with the difference shown in money and as a
> percentage, because a consistent two per cent gap is a fee rate that
> needs correcting on Settings rather than a mystery.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Reconciliation summary, then a per-platform fee check |

**Module constants:** `NAME` ('Platform Revenue & Fees').

````python
"""`Platform Revenue & Fees` - what the platform said, and what arrived.

Every host has had the moment where the money in the bank does not
match the money on the booking.  A cancelled night, a currency
conversion, a resolution centre adjustment, a fee that changed when the
account moved plan - none of it is visible from the reservation screen.

This is the reconciliation ledger.  One row per payout: what the
booking said you should get, what actually landed, and the difference
between the two, with the difference shown in money and as a
percentage, because a consistent two per cent gap is a fee rate that
needs correcting on Settings rather than a mystery.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, flag_yes_no_cf, id_col, lookup_booking,
                     lookup_booking_num, money_cf, month_col, owing_cf,
                     status_cf, year_col)

NAME = "Platform Revenue & Fees"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("booking", "PayoutPrefix", digits=cfg.digits("payout"),
               width=10, header="Payout ID"),
        Col("booking", "Booking", 11, "text",
            validate=list_from("BookingList"),
            note="The reservation this payout belongs to. One row per "
                 "payout - split a part payment across two rows.",
            group="The payout"),
        Col("property", "Property", 11, "text",
            formula=lookup_booking("property"), hidden=True,
            group="The payout"),
        Col("property_name", "Property", 19, "text",
            formula=lookup_booking("property_name"), group="The payout"),
        Col("guest_name", "Guest", 18, "text",
            formula=lookup_booking("guest_name"), group="The payout"),
        Col("platform", "Platform", 15, "text",
            formula=lookup_booking("platform"), group="The payout"),
        Col("check_out", "Check-out", 12, "date",
            formula=lookup_booking("check_out"), hidden=True,
            group="The payout"),

        # -------------------------------------------- what was expected --
        Col("exp_gross", "Gross booked", 12, "money",
            formula=lookup_booking_num("gross"),
            note="Straight from the booking - everything the guest paid.",
            group="Expected"),
        Col("exp_fee", "Fees expected", 12, "money",
            formula=lookup_booking_num("platform_fee_total"),
            group="Expected"),
        Col("exp_net", "Payout expected", 13, "money",
            formula=lookup_booking_num("net_revenue"),
            note="What the booking tab says should reach your account.",
            group="Expected"),

        # --------------------------------------------- what turned up --
        Col("payout_date", "Paid on", 12, "date",
            validate=date_check(),
            note="The date the money actually reached your account.",
            group="Received"),
        month_col("payout_date"),
        year_col("payout_date"),
        Col("amount_in", "Amount received", 14, "money",
            validate=number_check(-1000000, 1000000, "Amount received"),
            note="What landed. Take it from the bank statement, not "
                 "from the platform dashboard.",
            group="Received"),
        Col("fee_actual", "Fees actually taken", 14, "money",
            validate=number_check(0, 1000000, "Fees taken"),
            note="Optional. If the payout statement itemises the fee, "
                 "put it here and the fee-rate column below will tell "
                 "you whether your Settings rate is right.",
            group="Received"),
        Col("refund", "Refund given", 12, "money",
            validate=number_check(0, 1000000, "Refund"),
            note="Money you gave back - a cancellation, a goodwill "
                 "gesture, a resolution centre award. Positive number.",
            group="Received"),
        Col("adjustment", "Adjustment", 12, "money",
            validate=number_check(-1000000, 1000000, "Adjustment"),
            note="Anything else: a damage claim paid to you, a currency "
                 "difference, a host fee refund. Can be negative.",
            group="Received"),
        Col("received", "Net received", 13, "money",
            formula=blank(
                "booking",
                'ROUND({amount_in}{r}-{refund}{r}+{adjustment}{r},2)'),
            cf=money_cf(),
            note="Amount in, less refunds, plus adjustments. This is the "
                 "figure the financial reports treat as cash.",
            group="Received"),
        Col("status", "Status", 14, "status",
            validate=list_from("PayoutStatusList"),
            cf=status_cf(), group="Received"),
        Col("reference", "Statement ref", 14, "text", group="Received"),

        # -------------------------------------------------- the gap ---
        Col("variance", "Difference", 12, "money",
            formula=blank(
                "booking",
                'IF({payout_date}{r}="","",'
                'ROUND({received}{r}-{exp_net}{r},2))'),
            cf=owing_cf(),
            note="Received less expected. Amber when the platform kept "
                 "more than you thought, blue when it kept less.",
            group="The gap"),
        Col("variance_pct", "Difference %", 11, "percent",
            formula=blank(
                "booking",
                'IF({payout_date}{r}="","",IF({exp_net}{r}<=0,"",'
                '{variance}{r}/{exp_net}{r}))'),
            hidden=True, group="The gap"),
        Col("fee_rate_actual", "Real fee rate", 12, "percent",
            formula=blank(
                "booking",
                'IF({fee_actual}{r}="","",IF({exp_gross}{r}<=0,"",'
                '{fee_actual}{r}/{exp_gross}{r}))'),
            note="What the platform really charged, as a percentage of "
                 "the gross. If this keeps disagreeing with your "
                 "Settings rate, change the Settings rate.",
            group="The gap"),
        Col("days_to_pay", "Days to pay", 11, "int",
            formula=blank(
                "booking",
                'IF({payout_date}{r}="","",IF({check_out}{r}="","",'
                '{payout_date}{r}-{check_out}{r}))'),
            group="The gap"),
        Col("is_paid", "Money in", 9, "flag",
            formula=blank("booking",
                          'IF({payout_date}{r}="","No","Yes")'),
            cf=flag_yes_no_cf(), hidden=True, group="The gap"),

        Col("alert", "Action", 28, "status",
            formula=blank(
                "booking",
                'IF({exp_net}{r}="","Check - booking not found",'
                'IF({payout_date}{r}="",'
                'IF({check_out}{r}="","Check - no check-out date",'
                'IF(TodayDate-{check_out}{r}>PayoutChaseDays,'
                '"Chase - payout is overdue","Due - payout expected '
                'shortly")),'
                'IF({status}{r}="Disputed","Action - disputed payout",'
                'IF(ABS({variance}{r})<=0.5,"OK",'
                'IF({variance}{r}<0,"Check - less arrived than expected",'
                '"Check - more arrived than expected")))))'),
            cf=alert_cf(), group="The gap"),
        Col("notes", "Notes", 30, "wrap", group="The gap"),
    ]

    return SheetSpec(
        name=NAME,
        title="Platform Revenue & Fees",
        subtitle="What each platform said it would pay you, and what "
                 "actually landed",
        columns=columns,
        rows=cfg.row_count(NAME),
        freeze_cols=2,
        samples=sample("payouts", cfg),
        help_text="One row per payout. Pick the booking and the expected "
                  "figures arrive from the booking tab - then type what "
                  "really hit your bank account. The Difference column "
                  "is the point of this tab: small and random is "
                  "normal, consistent and one-sided means a fee rate "
                  "on Settings needs correcting.",
        extra_notes=[
            "Take the amount received from your bank statement rather "
            "than the platform dashboard. The dashboard shows what the "
            "platform intends to send; the statement shows what "
            "survived the currency conversion.",
            "A payout with no date is chased automatically once it is "
            "more days past check-out than the chase window on "
            "Settings.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Reconciliation summary, then a per-platform fee check."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index

    row = sheet.section(row, "Reconciliation", last)
    row = sheet.table_head(row, 0, [
        "Payouts logged", "Received", "Still expected", "Expected total",
        "Actually received", "Refunds", "Adjustments", "Difference",
        "Avg days to pay"],
        [13, 11, 13, 13, 15, 11, 12, 12, 14])
    sheet.calc_row(row, 0, [
        '=COUNTIF([[Platform Revenue & Fees.booking]],"?*")',
        '=COUNTIFS([[Platform Revenue & Fees.is_paid]],"Yes")',
        '=COUNTIFS([[Platform Revenue & Fees.is_paid]],"No")',
        '=ROUND(SUM([[Platform Revenue & Fees.exp_net]]),2)',
        '=ROUND(SUM([[Platform Revenue & Fees.received]]),2)',
        '=ROUND(SUM([[Platform Revenue & Fees.refund]]),2)',
        '=ROUND(SUM([[Platform Revenue & Fees.adjustment]]),2)',
        '=ROUND(SUM([[Platform Revenue & Fees.received]])-'
        'SUMIFS([[Platform Revenue & Fees.exp_net]],'
        '[[Platform Revenue & Fees.is_paid]],"Yes"),2)',
        '=IF(COUNTIFS([[Platform Revenue & Fees.is_paid]],"Yes")<=0,"",'
        'ROUND(SUM([[Platform Revenue & Fees.days_to_pay]])/'
        'COUNTIFS([[Platform Revenue & Fees.is_paid]],"Yes"),1))',
    ], ["int", "int", "int", "money", "money", "money", "money", "money",
        "number1"])
    row += 2

    row = sheet.section(row, "Is your fee rate right?", last)
    row = sheet.table_head(row, 0, [
        "Platform", "Payouts", "Gross booked", "Fees you assumed",
        "Fees really taken", "Assumed rate", "Real rate", "Difference"],
        [18, 10, 13, 15, 16, 12, 11, 12])
    for name, _commission, _payment, _counts, _note in cfg.platform_table:
        key = f'"{name}"'
        plat = f'[[Platform Revenue & Fees.platform]],{key}'
        gross = f'SUMIFS([[Platform Revenue & Fees.exp_gross]],{plat})'
        assumed = f'SUMIFS([[Platform Revenue & Fees.exp_fee]],{plat})'
        real = (f'SUMIFS([[Platform Revenue & Fees.fee_actual]],{plat})')
        with_real = (f'SUMIFS([[Platform Revenue & Fees.exp_gross]],{plat},'
                     f'[[Platform Revenue & Fees.fee_actual]],">0")')
        sheet.calc_row(row, 0, [
            name,
            f"=COUNTIFS({plat})",
            f"=ROUND({gross},2)",
            f"=ROUND({assumed},2)",
            f"=ROUND({real},2)",
            f'=IF({gross}<=0,"",{assumed}/{gross})',
            f'=IF({with_real}<=0,"",{real}/{with_real})',
            f'=IF({with_real}<=0,"",ROUND({real}-'
            f'SUMIFS([[Platform Revenue & Fees.exp_fee]],{plat},'
            f'[[Platform Revenue & Fees.fee_actual]],">0"),2))',
        ], ["text", "int", "money", "money", "money", "percent",
            "percent", "money"])
        row += 1

    row += 1
    sheet.note(row, "The real rate only counts payouts where you typed "
                    "the fee the statement showed, so it is blank until "
                    "you have filled a few in. When it settles a long "
                    "way from your assumed rate, change the rate on "
                    "Settings and every future booking prices itself "
                    "correctly.", last)
````

### `airbnb_rental_spreadsheet/sheets/expenses.py`

*252 lines · 2 functions*

> `Expense Tracker` - everything the properties cost you.
>
> Two rules keep this tab honest.
>
> First, **nothing is counted twice**.  Cleaning lives on the Cleaning
> Schedule and maintenance lives on the Maintenance Log, because both
> need their own workflow.  If you also type a cleaning invoice in here
> it is flagged and kept out of the profit figure - unless you tell
> Settings you would rather do it the other way round, in which case the
> other tabs step back instead.
>
> Second, **a recurring cost is shown as what it really costs a month**.
> An annual insurance premium and a weekly laundry bill are not
> comparable until they both say "per month", and a host staring at an
> expense breakdown is always asking the per-month question.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Totals, then the full category breakdown |

**Module constants:** `NAME` ('Expense Tracker').

````python
"""`Expense Tracker` - everything the properties cost you.

Two rules keep this tab honest.

First, **nothing is counted twice**.  Cleaning lives on the Cleaning
Schedule and maintenance lives on the Maintenance Log, because both
need their own workflow.  If you also type a cleaning invoice in here
it is flagged and kept out of the profit figure - unless you tell
Settings you would rather do it the other way round, in which case the
other tabs step back instead.

Second, **a recurring cost is shown as what it really costs a month**.
An annual insurance premium and a weekly laundry bill are not
comparable until they both say "per month", and a host staring at an
expense breakdown is always asking the per-month question.
"""

from __future__ import annotations

from ..config import EXPENSE_CATEGORIES, Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, flag_yes_no_cf, id_col, 
                    is_capital, lookup_booking, money_cf, 
                    monthly_equivalent, month_col, weekday_col, 
                    year_col)

NAME = "Expense Tracker"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "ExpensePrefix", digits=cfg.digits("expense"),
               width=10, header="Expense ID"),
        Col("date", "Date", 12, "date",
            validate=date_check(),
            note="The date on the receipt. Every row on this tab starts "
                 "here.",
            group="The cost"),
        month_col("date"),
        year_col("date"),
        weekday_col("date", "day", "Day", 6, hidden=True),
        Col("property", "Property", 11, "text",
            validate=list_from("PropertyList"),
            note="Which property it belongs to. Leave blank for a cost "
                 "that covers the whole portfolio - accounting fees, "
                 "your software subscription, the camera.",
            group="The cost"),
        Col("property_name", "Property name", 18, "text",
            formula=('=IF({date}{r}="","",IF({property}{r}="",'
                     '"Whole portfolio",'
                     'IFERROR(INDEX([[Property Database.name]],'
                     'MATCH({property}{r},[[Property Database.id]],0)),'
                     '"Not in Property Database")))'),
            group="The cost"),
        Col("category", "Category", 18, "text",
            validate=list_from("ExpenseCategoryList"),
            note="Pick from the list so the breakdowns add up. Edit the "
                 "list on Settings if your accountant wants different "
                 "headings.",
            group="The cost"),
        Col("vendor", "Paid to", 20, "text",
            note="Who you paid. Free text - it does not have to be "
                 "someone on the Staff & Vendors tab.",
            group="The cost"),
        Col("description", "What it was for", 30, "text",
            group="The cost"),
        Col("amount", "Amount", 12, "money",
            validate=number_check(-1000000, 1000000, "Amount"),
            cf=money_cf(),
            note="What you paid, including tax. A refund goes in as a "
                 "negative number.",
            group="The cost"),
        Col("tax_amount", "of which tax", 11, "money",
            validate=number_check(0, 1000000, "Tax"),
            note="Optional. The VAT or sales tax inside the amount, if "
                 "you reclaim it.",
            hidden=True, group="The cost"),
        Col("method", "Paid by", 14, "text",
            validate=list_from("PaymentMethodList"), group="The cost"),
        Col("receipt", "Receipt ref", 13, "text",
            note="Where the paperwork is - a file name, a folder, an "
                 "email subject. The thing you will want in March.",
            group="The cost"),

        # ----------------------------------------------- how it behaves --
        Col("frequency", "How often", 13, "text",
            validate=list_from("FrequencyList"),
            note="One off, or how often it repeats. This is what turns "
                 "the amount into a monthly equivalent.",
            group="How it behaves"),
        Col("monthly_equiv", "Per month", 11, "money",
            formula=blank("date",
                          monthly_equivalent("{amount}{r}",
                                             "{frequency}{r}")),
            note="What this cost really is each month. A yearly bill "
                 "divided by twelve, a weekly one times 52 over 12. One "
                 "off costs show nothing here.",
            group="How it behaves"),
        Col("booking", "Booking", 11, "text",
            validate=list_from("BookingList"),
            note="Optional. For a cost caused by one specific stay - a "
                 "replacement, a late check-out cleaner, a damage "
                 "repair. The true-profit-per-booking figure uses it.",
            group="How it behaves"),
        Col("booking_prop", "From booking", 11, "text",
            formula=lookup_booking("property"), hidden=True,
            group="How it behaves"),

        # ----------------------------------------------- how it counts --
        Col("tax_category", "Tax treatment", 17, "text",
            validate=list_from("TaxCategoryList"),
            note="How your accountant will want it grouped. The Tax "
                 "Summary tab totals these - it does not decide them "
                 "for you, because deductibility depends on where you "
                 "are and how you hold the property.",
            group="How it counts"),
        Col("is_cleaning", "Is cleaning", 11, "flag",
            formula=blank(
                "date",
                'IF({category}{r}="Cleaning","Yes",'
                'IF({category}{r}="Laundry","Yes","No"))'),
            hidden=True, group="How it counts"),
        Col("is_maintenance", "Is maintenance", 12, "flag",
            formula=blank(
                "date",
                'IF({category}{r}="Maintenance","Yes",'
                'IF({category}{r}="Repairs","Yes","No"))'),
            hidden=True, group="How it counts"),
        Col("is_capital", "Is capital", 10, "flag",
            formula=blank("date",
                          f'IF({is_capital("{category}{r}")},"Yes","No")'),
            hidden=True,
            note="Furniture and appliances are an asset, not a running "
                 "cost. They are totalled for tax but kept out of "
                 "operating profit.",
            group="How it counts"),
        Col("in_pl", "Counts in profit", 12, "flag",
            formula=blank(
                "date",
                'IF({is_capital}{r}="Yes","No",'
                'IF({is_cleaning}{r}="Yes",'
                'IF(CleaningInExpenses="Yes","Yes","No"),'
                'IF({is_maintenance}{r}="Yes",'
                'IF(MaintenanceInExpenses="Yes","Yes","No"),"Yes")))'),
            cf=flag_yes_no_cf(),
            note="No means this row is already counted somewhere else, "
                 "so adding it to profit would double count it. Both "
                 "switches live on Settings.",
            group="How it counts"),
        Col("pl_amount", "Counted as", 12, "money",
            formula=blank("date",
                          'IF({in_pl}{r}="Yes",{amount}{r},0)'),
            note="The figure the dashboard, the monthly report and the "
                 "per-property profit actually add up.",
            group="How it counts"),

        Col("alert", "Action", 26, "status",
            formula=blank(
                "date",
                'IF({amount}{r}="","Add - no amount",'
                'IF({category}{r}="","Add - no category",'
                'IF({in_pl}{r}="No",'
                'IF({is_capital}{r}="Yes","OK - capital, not a running '
                'cost","OK - counted on its own tab"),'
                'IF({receipt}{r}="","Check - no receipt reference",'
                'IF({tax_category}{r}="","Check - no tax treatment",'
                '"OK")))))'),
            cf=alert_cf(), group="How it counts"),
        Col("notes", "Notes", 28, "wrap", group="How it counts"),
    ]

    return SheetSpec(
        name=NAME,
        title="Expense Tracker",
        subtitle="Everything the properties cost you, counted once and "
                 "shown per month",
        columns=columns,
        rows=cfg.row_count(NAME),
        freeze_cols=2,
        samples=sample("expenses", cfg),
        help_text="One row per cost. Attach it to a property, or leave "
                  "the property blank for something that covers the "
                  "whole portfolio. Cleaning and maintenance typed in "
                  "here are flagged and excluded from profit, because "
                  "they already have their own tabs - the two switches "
                  "on Settings reverse that if you would rather run it "
                  "the other way.",
        extra_notes=[
            "Furniture and appliances are treated as capital, not as a "
            "running cost. They are still totalled on the Tax Summary, "
            "but they do not drag down the month they were bought in.",
            "Attach a cost to a booking when one particular stay caused "
            "it. That is what makes the true-profit-per-booking figure "
            "on the Monthly Reports tab worth looking at.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Totals, then the full category breakdown."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index

    row = sheet.section(row, "Everything on this tab", last)
    row = sheet.table_head(row, 0, [
        "Rows", "Total spent", "Counted in profit", "Capital items",
        "Already counted elsewhere", "Recurring per month",
        "Tax inside it", "No receipt"],
        [9, 12, 15, 12, 20, 17, 12, 11])
    sheet.calc_row(row, 0, [
        '=COUNT([[Expense Tracker.date]])',
        '=ROUND(SUM([[Expense Tracker.amount]]),2)',
        '=ROUND(SUM([[Expense Tracker.pl_amount]]),2)',
        '=ROUND(SUMIFS([[Expense Tracker.amount]],'
        '[[Expense Tracker.is_capital]],"Yes"),2)',
        '=ROUND(SUMIFS([[Expense Tracker.amount]],'
        '[[Expense Tracker.in_pl]],"No")-'
        'SUMIFS([[Expense Tracker.amount]],'
        '[[Expense Tracker.is_capital]],"Yes"),2)',
        '=ROUND(SUM([[Expense Tracker.monthly_equiv]]),2)',
        '=ROUND(SUM([[Expense Tracker.tax_amount]]),2)',
        '=COUNTIFS([[Expense Tracker.receipt]],"",'
        '[[Expense Tracker.date]],">0")',
    ], ["int", "money", "money", "money", "money", "money", "money",
        "int"])
    row += 2

    row = sheet.section(row, "Where the money goes", last)
    row = sheet.table_head(row, 0, [
        "Category", "Rows", "Total", "Counted in profit", "Per month",
        "Share of spend", "Biggest single"],
        [20, 8, 12, 15, 11, 13, 14])
    total = 'SUM([[Expense Tracker.amount]])'
    for name in EXPENSE_CATEGORIES:
        key = f'"{name}"'
        cat = f'[[Expense Tracker.category]],{key}'
        amt = f'SUMIFS([[Expense Tracker.amount]],{cat})'
        sheet.calc_row(row, 0, [
            name,
            f"=COUNTIFS({cat})",
            f"=ROUND({amt},2)",
            f'=ROUND(SUMIFS([[Expense Tracker.pl_amount]],{cat}),2)',
            f'=ROUND(SUMIFS([[Expense Tracker.monthly_equiv]],{cat}),2)',
            f'=IF({total}<=0,"",{amt}/{total})',
            f'=IF(COUNTIFS({cat})<=0,"",'
            f'SUMPRODUCT(MAX(([[Expense Tracker.category]]={key})*'
            f'IFERROR([[Expense Tracker.amount]]*1,0))))',
        ], ["text", "int", "money", "money", "money", "percent", "money"])
        row += 1
````

### `airbnb_rental_spreadsheet/sheets/tax.py`

*340 lines · 2 functions*

> `Tax Summary` - your figures, grouped the way an accountant wants them.
>
> Pick a year. This page then adds up gross rental income, what the
> platforms kept, every category of cost under its tax heading, and what
> is left - with capital items kept separate, because a new boiler is
> not a repair and your accountant will say so.
>
> Read the disclaimer. This tab sorts your own numbers into sensible
> piles; it does not know your country, your entity, whether you are
> trading or letting, or what your local rules allow. It will save your
> accountant two hours. It will not replace them.

| Defines | What it does |
|---------|--------------|
| `_year_window(column)` | `range,">="&start,range,"<="&end` for the chosen tax year |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Tax Summary'), `LAST_COL` (8).

````python
"""`Tax Summary` - your figures, grouped the way an accountant wants them.

Pick a year. This page then adds up gross rental income, what the
platforms kept, every category of cost under its tax heading, and what
is left - with capital items kept separate, because a new boiler is
not a repair and your accountant will say so.

Read the disclaimer. This tab sorts your own numbers into sensible
piles; it does not know your country, your entity, whether you are
trading or letting, or what your local rules allow. It will save your
accountant two hours. It will not replace them.
"""

from __future__ import annotations

from ..config import (CAPITAL_CATEGORIES, Config, DISCLAIMER,
                      EXPENSE_CATEGORIES, OWNERSHIP_CATEGORIES,
                      Palette, TAX_CATEGORIES)
from ..layout import Registry, col_letter
from ..writer import SheetWriter, number_check

NAME = "Tax Summary"
LAST_COL = 8


def _year_window(column: str) -> str:
    """`range,">="&start,range,"<="&end` for the chosen tax year."""
    return (f'{column},">="&TaxYearStart,{column},"<="&TaxYearEnd')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.deep)
    ws = s.ws

    s.banner("Tax & Financial Summary",
             "A year of income and deductible costs, grouped for "
             "whoever does your return.", LAST_COL)
    s.set_col(0, 0, 4)
    s.set_col(1, 1, 34)
    s.set_col(2, 2, 15)
    s.set_col(3, 3, 15)
    s.set_col(4, 4, 12)
    s.set_col(5, LAST_COL, 16)

    # -------------------------------------------------------- the year --
    row = 4
    s.merge(row, 1, 1, "Tax year", fmts.chrome("label_bold"))
    s.write(row, 2, cfg.today.year, fmts.chrome("form_box", locked=False,
                                                num_format="0"))
    ws.data_validation(row, 2, row, 2, number_check(2000, 2100,
                                                    decimal=False))
    book.define_name("TaxYear", f"='{NAME}'!$C${row + 1}")
    s.merge(row, 3, LAST_COL,
            "The calendar year to report on. If your tax year does "
            "not start in January, change the two dates below.",
            fmts.chrome("note"))
    row += 1

    s.merge(row, 1, 1, "Year starts", fmts.chrome("label_bold"))
    s.write(row, 2, None, fmts.calc("date"))
    s.formula(row, 2, "=DATE(TaxYear,1,1)", fmts.calc("date"))
    book.define_name("TaxYearStart", f"='{NAME}'!$C${row + 1}")
    s.merge(row, 3, 3, "Year ends", fmts.chrome("label_right"))
    s.write(row, 4, None, fmts.calc("date"))
    s.formula(row, 4, "=DATE(TaxYear+1,1,1)-1", fmts.calc("date"))
    book.define_name("TaxYearEnd", f"='{NAME}'!$E${row + 1}")
    s.merge(row, 5, LAST_COL,
            "Edit these two formulas for a 6 April or 1 July year.",
            fmts.chrome("note"))
    row += 2

    bk_year = _year_window("[[Booking Tracker.check_in]]")
    earns = '[[Booking Tracker.earns]],"Yes"'
    ex_year = _year_window("[[Expense Tracker.date]]")
    cl_year = _year_window("[[Cleaning Schedule.clean_date]]")
    mt_year = _year_window("[[Maintenance Log.report_date]]")

    gross = f'SUMIFS([[Booking Tracker.gross]],{bk_year},{earns})'
    accom = f'SUMIFS([[Booking Tracker.accommodation]],{bk_year},{earns})'
    cleaning_income = (f'SUMIFS([[Booking Tracker.cleaning_fee]],'
                       f'{bk_year},{earns})')
    extra_income = (f'SUMIFS([[Booking Tracker.extra_fee]],{bk_year},'
                    f'{earns})')
    fees = (f'SUMIFS([[Booking Tracker.platform_fee_total]],{bk_year},'
            f'{earns})')
    lodging = f'SUMIFS([[Booking Tracker.tax]],{bk_year},{earns})'
    net_income = f'({gross}-{fees}-{lodging})'
    clean_cost = f'SUMIFS([[Cleaning Schedule.total_cost]],{cl_year})'
    maint_cost = f'SUMIFS([[Maintenance Log.cost_used]],{mt_year})'
    other_cost = f'SUMIFS([[Expense Tracker.pl_amount]],{ex_year})'
    capital = (f'SUMIFS([[Expense Tracker.amount]],{ex_year},'
               f'[[Expense Tracker.is_capital]],"Yes")')
    total_cost = f'({clean_cost}+{maint_cost}+{other_cost})'

    # ----------------------------------------------------- headline -----
    row = s.section(row, "The headline", LAST_COL)
    s.kpi(row, 1, "Gross rental income", f"=ROUND({gross},2)", "money0",
          note="Everything a guest paid you")
    s.kpi(row, 3, "Deductible costs", f"=ROUND({total_cost}+{fees},2)",
          "money0", note="Including what the platforms kept")
    s.kpi(row, 5, "Taxable profit",
          f"=ROUND({net_income}-{total_cost},2)", "money0")
    row = s.kpi(row, 7, "Capital items",
                f"=ROUND({capital},2)", "money0",
                note="Not deducted here - ask about allowances")
    row += 1

    # ------------------------------------------------------- income -----
    row = s.section(row, "Income", LAST_COL)
    row = s.table_head(row, 1, ["Line", "Amount", "Per property",
                                "What it is"],
                       [34, 15, 15, 12])
    income_lines = [
        ("Accommodation", f"={accom}",
         "Room rate only. This is the number most tax systems call "
         "rental income."),
        ("Cleaning fees charged", f"={cleaning_income}",
         "Income in most jurisdictions, even though it recovers a "
         "cost. The cost itself is deducted below."),
        ("Extra guest and pet fees", f"={extra_income}",
         "Charged on top of the rate."),
        ("Gross received from guests", f"={gross}",
         "The three above, less discounts given."),
        ("Less platform commission and card fees", f"=-{fees}",
         "What Airbnb, Vrbo and the rest kept. Deductible almost "
         "everywhere, and easy to forget because you never see it."),
        ("Less lodging tax collected", f"=-{lodging}",
         "Only non-zero if you collect it yourself. Tax you remit is "
         "not your income."),
        ("Net rental income", f"={net_income}",
         "What actually reached you."),
    ]
    for label, formula, what in income_lines:
        bold = label.startswith(("Gross", "Net"))
        s.write(row, 1, label,
                fmts.chrome("total_label" if bold else "label"))
        fmt = (fmts.chrome("total_value",
                           num_format=s.num_format("money"))
               if bold else fmts.calc("money"))
        s.formula(row, 2, s.reg.render(f"=ROUND({formula[1:]},2)"),
                         fmt)
        s.formula(
            row, 3,
            s.reg.render(f'=IF(PropertyCount<=0,"",ROUND(({formula[1:]})'
                         f'/PropertyCount,2))'),
            fmts.calc("money"))
        s.merge(row, 4, LAST_COL, what, fmts.chrome("note"))
        s.row_height(row, 26)
        row += 1
    row += 1

    # ----------------------------------------------------- expenses -----
    row = s.section(row, "Costs by tax heading", LAST_COL)
    row = s.note(
        row,
        "Every expense row carries a tax heading, which you set on "
        "the Expense Tracker. Cleaning and maintenance are added from "
        "their own tabs so they are counted once and only once.",
        LAST_COL, height=26)
    row = s.table_head(row, 1, ["Tax heading", "Amount", "Share",
                                "Entries"],
                       [34, 15, 15, 12])
    for heading in TAX_CATEGORIES:
        amount = (f'SUMIFS([[Expense Tracker.amount]],{ex_year},'
                  f'[[Expense Tracker.tax_category]],"{heading}")')
        s.write(row, 1, heading, fmts.chrome("label"))
        s.formula(row, 2, s.reg.render(f"=ROUND({amount},2)"),
                         fmts.calc("money"))
        s.formula(
            row, 3,
            s.reg.render(f'=IF({total_cost}<=0,"",{amount}/'
                         f'{total_cost})'),
            fmts.calc("percent"))
        s.formula(
            row, 4,
            s.reg.render(f'=COUNTIFS({ex_year},'
                         f'[[Expense Tracker.tax_category]],'
                         f'"{heading}")'), fmts.calc("int"))
        row += 1
    row += 1

    # -------------------------------------------- the full breakdown ----
    row = s.section(row, "Costs by category", LAST_COL)
    row = s.table_head(row, 1, ["Category", "Amount", "Share",
                                "Entries", "Treatment"],
                       [34, 15, 15, 12, 20])
    for category in EXPENSE_CATEGORIES:
        amount = (f'SUMIFS([[Expense Tracker.amount]],{ex_year},'
                  f'[[Expense Tracker.category]],"{category}")')
        if category in CAPITAL_CATEGORIES:
            treatment = "Capital - not deducted here"
        elif category in OWNERSHIP_CATEGORIES:
            treatment = "Ownership cost - rules vary"
        else:
            treatment = "Running cost"
        s.write(row, 1, category, fmts.chrome("label"))
        s.formula(row, 2, s.reg.render(f"=ROUND({amount},2)"),
                         fmts.calc("money"))
        s.formula(
            row, 3,
            s.reg.render(f'=IF({total_cost}<=0,"",{amount}/'
                         f'{total_cost})'), fmts.calc("percent"))
        s.formula(
            row, 4,
            s.reg.render(f'=COUNTIFS({ex_year},'
                         f'[[Expense Tracker.category]],'
                         f'"{category}")'), fmts.calc("int"))
        s.merge(row, 5, LAST_COL, treatment, fmts.chrome("note"))
        row += 1

    s.write(row, 1, "Cleaning and turnovers", fmts.chrome("label"))
    s.formula(row, 2, s.reg.render(f"=ROUND({clean_cost},2)"),
                     fmts.calc("money"))
    s.formula(
        row, 3,
        s.reg.render(f'=IF({total_cost}<=0,"",{clean_cost}/'
                     f'{total_cost})'), fmts.calc("percent"))
    s.formula(row, 4,
                     s.reg.render(f'=COUNTIFS({cl_year})'),
                     fmts.calc("int"))
    s.merge(row, 5, LAST_COL, "From the Cleaning Schedule",
            fmts.chrome("note"))
    row += 1
    s.write(row, 1, "Repairs and maintenance", fmts.chrome("label"))
    s.formula(row, 2, s.reg.render(f"=ROUND({maint_cost},2)"),
                     fmts.calc("money"))
    s.formula(
        row, 3,
        s.reg.render(f'=IF({total_cost}<=0,"",{maint_cost}/'
                     f'{total_cost})'), fmts.calc("percent"))
    s.formula(row, 4,
                     s.reg.render(f'=COUNTIFS({mt_year})'),
                     fmts.calc("int"))
    s.merge(row, 5, LAST_COL, "From the Maintenance Log",
            fmts.chrome("note"))
    row += 1
    s.write(row, 1, "Total costs", fmts.chrome("total_label"))
    s.formula(row, 2, s.reg.render(f"=ROUND({total_cost},2)"),
                     fmts.chrome("total_value",
                                 num_format=s.num_format("money")))
    row += 2

    # ------------------------------------------------- the bottom line --
    row = s.section(row, "The bottom line", LAST_COL)
    row = s.table_head(row, 1, ["Line", "Amount", "Per property",
                                "Note"], [34, 15, 15, 12])
    bottom = [
        ("Net rental income", f"={net_income}",
         "After platform fees and any lodging tax you remit."),
        ("Less running costs", f"=-{total_cost}",
         "Cleaning, repairs and everything on the expense tab, "
         "excluding capital."),
        ("Taxable profit", f"={net_income}-{total_cost}",
         "Before any allowance, relief or depreciation your "
         "jurisdiction may give you."),
        ("Capital spend in the year", f"={capital}",
         "Furniture, appliances, refurbishment. Usually relieved "
         "over several years, not in one. Not deducted above."),
        ("Mortgage interest paid", '=SUMIFS([[Expense Tracker.amount]],'
         + ex_year + ',[[Expense Tracker.category]],'
                     '"Mortgage interest")',
         "Shown separately because how much of it is deductible is "
         "one of the most country-specific rules there is."),
        ("Rent paid to a landlord (rent-to-rent)",
         '=SUMIFS([[Expense Tracker.amount]],' + ex_year +
         ',[[Expense Tracker.category]],"Rent")',
         "For arbitrage operators, usually the largest single cost."),
    ]
    for label, formula, note in bottom:
        bold = label in ("Taxable profit", "Net rental income")
        s.write(row, 1, label,
                fmts.chrome("total_label" if bold else "label"))
        fmt = (fmts.chrome("total_value",
                           num_format=s.num_format("money"))
               if bold else fmts.calc("money"))
        s.formula(row, 2,
                         s.reg.render(f"=ROUND({formula[1:]},2)"), fmt)
        s.formula(
            row, 3,
            s.reg.render(f'=IF(PropertyCount<=0,"",ROUND(({formula[1:]})'
                         f'/PropertyCount,2))'), fmts.calc("money"))
        s.merge(row, 4, LAST_COL, note, fmts.chrome("note"))
        s.row_height(row, 26)
        row += 1
    row += 1

    # ------------------------------------------------- per property -----
    row = s.section(row, "By property, for the year", LAST_COL)
    row = s.table_head(row, 1, [
        "Property", "Gross", "Platform fees", "Costs", "Profit",
        "Nights sold"], [26, 15, 15, 14, 14, 12])
    for n in range(1, 9):
        pid = f'INDEX([[Property Database.id]],{n})'
        guard = f'INDEX([[Property Database.name]],{n})=""'
        p_gross = (f'SUMIFS([[Booking Tracker.gross]],{bk_year},{earns},'
                   f'[[Booking Tracker.property]],{pid})')
        p_fees = (f'SUMIFS([[Booking Tracker.platform_fee_total]],'
                  f'{bk_year},{earns},[[Booking Tracker.property]],'
                  f'{pid})')
        p_nights = (f'SUMIFS([[Booking Tracker.nights]],{bk_year},'
                    f'{earns},[[Booking Tracker.property]],{pid})')
        p_costs = (f'(SUMIFS([[Cleaning Schedule.total_cost]],{cl_year},'
                   f'[[Cleaning Schedule.property]],{pid})+'
                   f'SUMIFS([[Maintenance Log.cost_used]],{mt_year},'
                   f'[[Maintenance Log.property]],{pid})+'
                   f'SUMIFS([[Expense Tracker.pl_amount]],{ex_year},'
                   f'[[Expense Tracker.property]],{pid}))')
        s.formula(
            row, 1,
            s.reg.render(f'=IFERROR(IF({guard},"",INDEX('
                         f'[[Property Database.name]],{n})),"")'),
            fmts.calc("text"))
        for offset, (formula, style) in enumerate((
                (f"ROUND({p_gross},2)", "money"),
                (f"ROUND({p_fees},2)", "money"),
                (f"ROUND({p_costs},2)", "money"),
                (f"ROUND({p_gross}-{p_fees}-{p_costs},2)", "money"),
                (p_nights, "int"))):
            s.formula(
                row, 2 + offset,
                s.reg.render(f'=IFERROR(IF({guard},"",{formula}),"")'),
                fmts.calc(style))
        row += 1
    row += 1

    row = s.note(
        row,
        "Everything on this page is attributed to the year a stay "
        "checked in and the date a cost was dated. If your accountant "
        "wants revenue spread across the nights of a stay that "
        "straddles new year, tell them which convention this file "
        "uses - it is consistent, which is what matters.",
        LAST_COL, height=32)
    row += 1
    row = s.note(row, DISCLAIMER, LAST_COL, height=112)

    s.finish(freeze_row=4, gridlines=False, fit_width=True,
             print_area=f"A1:{col_letter(LAST_COL)}{row + 2}")
    s.protect()
    return s
````

---

## 7. The property itself

Stock counted the honest way with weeks of cover and a reorder flag, repairs with an estimate against an actual and a warranty countdown, the reviews with seven sub-ratings and a weakest category, and the cleaners and contractors with their rates, their completion scores and their insurance expiry dates.

### `airbnb_rental_spreadsheet/sheets/supplies.py`

*234 lines · 2 functions*

> `Supplies Inventory` - linen, loo roll and the things that run out.
>
> Running out of toilet paper on a Friday afternoon is a one-star review
> and a forty mile round trip.  This tab works out how fast each line is
> actually being consumed - from the number of turnovers that property
> has had, not from a guess - and tells you how many weeks of cover is
> left before it becomes a problem.
>
> Stock can belong to one property or to a shared store cupboard that
> serves the lot.  Leave the property blank and it is treated as shared.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Supplies Inventory').

````python
"""`Supplies Inventory` - linen, loo roll and the things that run out.

Running out of toilet paper on a Friday afternoon is a one-star review
and a forty mile round trip.  This tab works out how fast each line is
actually being consumed - from the number of turnovers that property
has had, not from a guess - and tells you how many weeks of cover is
left before it becomes a problem.

Stock can belong to one property or to a shared store cupboard that
serves the lot.  Leave the property blank and it is treated as shared.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, id_col, money_cf, status_cf)

NAME = "Supplies Inventory"


def spec(cfg: Config) -> SheetSpec:
    cl_prop = "[[Cleaning Schedule.property]]"
    cl_date = "[[Cleaning Schedule.clean_date]]"

    columns = [
        id_col("name", "SupplyPrefix", digits=cfg.digits("supply"),
               width=10, header="Item ID"),
        Col("name", "Item", 24, "text",
            note="Bath towels, king sheets, loo roll, dishwasher "
                 "tablets. One row per thing you re-buy.",
            group="The item"),
        Col("category", "Category", 18, "text",
            validate=list_from("SupplyCategoryList"), group="The item"),
        Col("property", "Property", 11, "text",
            validate=list_from("PropertyList"),
            note="Blank means a shared store cupboard that serves every "
                 "property.",
            group="The item"),
        Col("property_name", "Kept for", 18, "text",
            formula=('=IF({name}{r}="","",IF({property}{r}="",'
                     '"Shared store",'
                     'IFERROR(INDEX([[Property Database.name]],'
                     'MATCH({property}{r},[[Property Database.id]],0)),'
                     '"Not in Property Database")))'),
            group="The item"),
        Col("unit", "Unit", 10, "text",
            note="Each, pack, roll, pair, litre. Whatever you count it "
                 "in.",
            group="The item"),
        Col("supplier", "Supplier", 18, "text", group="The item"),
        Col("sku", "Code / SKU", 13, "text", hidden=True, group="The item"),

        # ----------------------------------------------------- the stock --
        Col("on_hand", "In stock", 10, "number1",
            validate=number_check(0, 100000, "Quantity in stock"),
            note="What is there now. Update it when you count, not "
                 "every time you take one.",
            cf=bar_cf(), group="Stock"),
        Col("min_level", "Minimum", 10, "number1",
            validate=number_check(0, 100000, "Minimum level"),
            note="The level at which you want to be told to order. For "
                 "linen, two full changes per bed is a sensible floor.",
            group="Stock"),
        Col("last_count", "Last counted", 12, "date",
            validate=date_check(), group="Stock"),
        Col("last_restock", "Last restocked", 12, "date",
            validate=date_check(), group="Stock"),
        Col("count_age", "Days since count", 13, "int",
            formula=blank("name",
                          'IF({last_count}{r}="","",'
                          'MAX(0,TodayDate-{last_count}{r}))'),
            hidden=True, group="Stock"),

        # ----------------------------------------------------- the usage --
        Col("per_turn", "Used per turnover", 14, "number1",
            validate=number_check(0, 1000, "Used per turnover"),
            note="How many of these a single changeover gets through. "
                 "Four towels, two sheet sets, one roll of bin bags. "
                 "This is what makes the forecast real.",
            group="Usage"),
        Col("turns_month", "Turnovers a month", 14, "number1",
            formula=blank(
                "name",
                f'IF({{property}}{{r}}="",'
                f'ROUND(COUNTIFS({cl_date},">="&(TodayDate-90),'
                f'{cl_date},"<="&TodayDate)/3,1),'
                f'ROUND(COUNTIFS({cl_prop},{{property}}{{r}},'
                f'{cl_date},">="&(TodayDate-90),'
                f'{cl_date},"<="&TodayDate)/3,1))'),
            note="Counted from the Cleaning Schedule over the last "
                 "ninety days, so it follows your real season rather "
                 "than an assumption.",
            group="Usage"),
        Col("usage_month", "Used a month", 12, "number1",
            formula=blank("name",
                          'ROUND({per_turn}{r}*{turns_month}{r},1)'),
            group="Usage"),
        Col("weeks_cover", "Weeks of cover", 13, "number1",
            formula=blank(
                "name",
                'IF({usage_month}{r}<=0,"",'
                'ROUND({on_hand}{r}/({usage_month}{r}/4.33),1))'),
            cf=[{"type": "cell", "criteria": "<", "value": 1,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<",
                 "value": "=RestockWeeks", "fmt": "warn"},
                {"type": "cell", "criteria": ">=",
                 "value": "=RestockWeeks", "fmt": "good"}],
            note="How long the stock on the shelf will last at the rate "
                 "you are actually turning the property over.",
            group="Usage"),

        # ----------------------------------------------------- the money --
        Col("unit_cost", "Unit cost", 11, "money",
            validate=number_check(0, 100000, "Unit cost"), group="Money"),
        Col("value", "Stock value", 12, "money",
            formula=blank("name",
                          'ROUND({on_hand}{r}*{unit_cost}{r},2)'),
            cf=money_cf(), group="Money"),
        Col("reorder_qty", "Order quantity", 13, "number1",
            validate=number_check(0, 100000, "Order quantity"),
            note="How many you buy at a time. Blank falls back to "
                 "enough to get back above the minimum.",
            group="Money"),
        Col("order_qty_used", "To order", 10, "number1",
            formula=blank(
                "name",
                'IF({needs_order}{r}<>"Yes",0,'
                'IF({reorder_qty}{r}<>"",{reorder_qty}{r},'
                'MAX(0,{min_level}{r}*2-{on_hand}{r})))'),
            group="Money"),
        Col("order_cost", "Order cost", 11, "money",
            formula=blank("name",
                          'ROUND({order_qty_used}{r}*{unit_cost}{r},2)'),
            group="Money"),
        Col("cost_per_turn", "Cost a turnover", 13, "money",
            formula=blank("name",
                          'ROUND({per_turn}{r}*{unit_cost}{r},2)'),
            hidden=True,
            note="What this one line adds to every changeover.",
            group="Money"),

        # --------------------------------------------------- the verdict --
        Col("needs_order", "Order now", 10, "flag",
            formula=blank(
                "name",
                'IF({min_level}{r}<=0,"No",'
                'IF({on_hand}{r}<={min_level}{r},"Yes",'
                'IF({weeks_cover}{r}="","No",'
                'IF({weeks_cover}{r}<RestockWeeks,"Yes","No"))))'),
            hidden=True, group="Verdict"),
        Col("status", "Stock status", 13, "status",
            formula=blank(
                "name",
                'IF({on_hand}{r}<=0,"Out of stock",'
                'IF({min_level}{r}<=0,"Not tracked",'
                'IF({on_hand}{r}<={min_level}{r},"Order now",'
                'IF({needs_order}{r}="Yes","Low",'
                'IF({on_hand}{r}>{min_level}{r}*4,"Overstocked",'
                '"In stock")))))'),
            cf=status_cf(), group="Verdict"),
        Col("alert", "Action", 28, "status",
            formula=blank(
                "name",
                'IF({on_hand}{r}<=0,"Out of stock - order today",'
                'IF({min_level}{r}<=0,"Add - set a minimum level",'
                'IF({needs_order}{r}="Yes",'
                'IF({order_cost}{r}>0,"Order - '
                '"&TEXT({order_qty_used}{r},"0")&" "&{unit}{r},'
                '"Order - stock is low"),'
                'IF({count_age}{r}="","Check - never counted",'
                'IF({count_age}{r}>90,"Check - not counted in three '
                'months","OK")))))'),
            cf=alert_cf(), group="Verdict"),
        Col("notes", "Notes", 26, "wrap", group="Verdict"),
    ]

    return SheetSpec(
        name=NAME,
        title="Supplies Inventory",
        subtitle="Linen, consumables and the things that run out on a "
                 "Friday afternoon",
        columns=columns,
        rows=cfg.row_count(NAME),
        freeze_cols=2,
        samples=sample("supplies", cfg),
        help_text="One row per thing you re-buy. Type what is on the "
                  "shelf, the minimum you want to keep and how many a "
                  "single turnover gets through - the weeks of cover, "
                  "the order quantity and the order cost work "
                  "themselves out from the number of changeovers that "
                  "property has actually had.",
        extra_notes=[
            "Turnovers a month is counted from the Cleaning Schedule "
            "over the last ninety days, so a property that is busy in "
            "August and quiet in February gets a different answer in "
            "each.",
            "Leave the property blank for a shared store cupboard. Its "
            "usage is then based on every turnover across the "
            "portfolio.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index

    row = sheet.section(row, "The store cupboard", last)
    row = sheet.table_head(row, 0, [
        "Lines tracked", "In stock", "Low", "Order now", "Out of stock",
        "Stock value", "Cost to restock", "Cost a turnover"],
        [13, 10, 8, 11, 13, 12, 15, 14])
    sheet.calc_row(row, 0, [
        '=COUNTIF([[Supplies Inventory.name]],"?*")',
        '=COUNTIFS([[Supplies Inventory.status]],"In stock")',
        '=COUNTIFS([[Supplies Inventory.status]],"Low")',
        '=COUNTIFS([[Supplies Inventory.status]],"Order now")',
        '=COUNTIFS([[Supplies Inventory.status]],"Out of stock")',
        '=ROUND(SUM([[Supplies Inventory.value]]),2)',
        '=ROUND(SUM([[Supplies Inventory.order_cost]]),2)',
        '=ROUND(SUM([[Supplies Inventory.cost_per_turn]]),2)',
    ], ["int", "int", "int", "int", "int", "money", "money", "money"])
    row += 2

    sheet.note(row, "'Cost a turnover' adds up what every consumable "
                    "line contributes to a single changeover. Compare "
                    "it with the consumables figure on Settings - if "
                    "they disagree, the Settings number is the one "
                    "making your turnover costs wrong.", last)
````

### `airbnb_rental_spreadsheet/sheets/maintenance.py`

*243 lines · 2 functions*

> `Maintenance Log` - what is broken, who is fixing it, what it cost.
>
> Short-term rentals wear out roughly five times faster than a home,
> because a hundred different people a year use the shower, the sofa and
> the front door.  The jobs that sink a property are rarely the big ones;
> they are the twelve small ones nobody wrote down.
>
> Every row carries a priority, a contractor, an estimate and an actual.
> The gap between the estimate and the actual is shown, because that gap
> is the most useful thing you will ever learn about a tradesperson.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Maintenance Log').

````python
"""`Maintenance Log` - what is broken, who is fixing it, what it cost.

Short-term rentals wear out roughly five times faster than a home,
because a hundred different people a year use the shower, the sofa and
the front door.  The jobs that sink a property are rarely the big ones;
they are the twelve small ones nobody wrote down.

Every row carries a priority, a contractor, an estimate and an actual.
The gap between the estimate and the actual is shown, because that gap
is the most useful thing you will ever learn about a tradesperson.
"""

from __future__ import annotations

from ..config import MAINTENANCE_CATEGORIES, Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, flag_yes_no_cf, id_col,
                     lookup_property, lookup_vendor, money_cf, month_col,
                     owing_cf, priority_cf, status_cf, year_col, yesno)

NAME = "Maintenance Log"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("report_date", "MaintenancePrefix",
               digits=cfg.digits("maintenance"), width=10, header="Job ID"),
        Col("report_date", "Reported", 12, "date",
            validate=date_check(),
            note="When it was noticed. Every row starts here.",
            group="The job"),
        month_col("report_date"),
        year_col("report_date"),
        Col("property", "Property", 11, "text",
            validate=list_from("PropertyList"), group="The job"),
        Col("property_name", "Property name", 19, "text",
            formula=lookup_property("name"), group="The job"),
        Col("category", "Category", 17, "text",
            validate=list_from("MaintenanceCategoryList"),
            group="The job"),
        Col("description", "What is wrong", 34, "wrap", group="The job"),
        Col("priority", "Priority", 12, "status",
            validate=list_from("PriorityList"),
            cf=priority_cf(),
            note="Urgent means a guest cannot use the property. Be "
                 "strict about it or the word stops meaning anything.",
            group="The job"),
        Col("reported_by", "Reported by", 16, "text",
            note="The guest, the cleaner, you, the neighbour.",
            group="The job"),
        Col("booking", "During booking", 12, "text",
            validate=list_from("BookingList"),
            note="Optional. If a particular stay caused it, say so - "
                 "it is the first thing you will want if you ever claim "
                 "on a deposit.",
            hidden=True, group="The job"),

        # ---------------------------------------------------- who fixes --
        Col("vendor", "Contractor", 11, "text",
            validate=list_from("VendorList"), group="Who fixes it"),
        Col("vendor_name", "Contractor name", 18, "text",
            formula=lookup_vendor("name"), group="Who fixes it"),
        Col("vendor_phone", "Phone", 14, "text",
            formula=lookup_vendor("phone"), hidden=True,
            group="Who fixes it"),
        Col("scheduled", "Booked in for", 12, "date",
            validate=date_check(), group="Who fixes it"),
        Col("completed", "Finished", 12, "date",
            validate=date_check(), group="Who fixes it"),
        Col("status", "Status", 16, "status",
            validate=list_from("MaintenanceStatusList"),
            cf=status_cf(), group="Who fixes it"),
        Col("blocks_calendar", "Blocks the calendar", 15, "flag",
            validate=yesno("Does this stop you letting the property?"),
            cf=[{"type": "text", "criteria": "containing", "value": "Yes",
                 "fmt": "bad"}],
            note="Yes means you cannot take a booking until it is done. "
                 "Put a Maintenance block on the Booking Tracker to "
                 "hold the dates.",
            group="Who fixes it"),

        # ----------------------------------------------------- the cost --
        Col("est_cost", "Estimate", 11, "money",
            validate=number_check(0, 1000000, "Estimated cost"),
            group="Cost"),
        Col("act_cost", "Actual", 11, "money",
            validate=number_check(0, 1000000, "Actual cost"),
            group="Cost"),
        Col("cost_used", "Cost counted", 12, "money",
            formula=blank(
                "report_date",
                'IF({act_cost}{r}<>"",{act_cost}{r},'
                'IF({status}{r}="Complete",0,0))'),
            cf=money_cf(),
            note="Only the actual figure counts towards profit. An "
                 "estimate is a plan, not a cost.",
            group="Cost"),
        Col("overrun", "Over the estimate", 13, "money",
            formula=blank(
                "report_date",
                'IF({act_cost}{r}="","",IF({est_cost}{r}="","",'
                'ROUND({act_cost}{r}-{est_cost}{r},2)))'),
            cf=owing_cf(),
            note="Actual less estimate. Worth watching per contractor.",
            hidden=True, group="Cost"),
        Col("invoice_ref", "Invoice ref", 13, "text", group="Cost"),
        Col("paid", "Paid", 8, "flag",
            validate=yesno(), cf=flag_yes_no_cf(), group="Cost"),
        Col("warranty_until", "Under warranty to", 14, "date",
            validate=date_check(),
            note="For a new appliance or a job with a guarantee. Saves "
                 "you paying twice.",
            hidden=True, group="Cost"),

        # ----------------------------------------------- where it stands --
        Col("is_open", "Still open", 10, "flag",
            formula=blank(
                "report_date",
                'IF({completed}{r}<>"","No",'
                'IF({status}{r}="Complete","No",'
                'IF({status}{r}="Cancelled","No","Yes")))'),
            cf=flag_yes_no_cf(), hidden=True, group="Where it stands"),
        Col("days_open", "Days open", 10, "int",
            formula=blank(
                "report_date",
                'IF({is_open}{r}="Yes",MAX(0,TodayDate-{report_date}{r}),'
                'IF({completed}{r}="","",'
                'MAX(0,{completed}{r}-{report_date}{r})))'),
            cf=[{"type": "cell", "criteria": ">", "value": 30,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 14,
                 "fmt": "warn"}],
            group="Where it stands"),
        Col("due_in", "Due in", 9, "int",
            formula=blank(
                "report_date",
                'IF({scheduled}{r}="","",IF({is_open}{r}<>"Yes","",'
                '{scheduled}{r}-TodayDate))'),
            hidden=True, group="Where it stands"),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "report_date",
                'IF({status}{r}="Cancelled","",'
                'IF({is_open}{r}="No",'
                'IF({paid}{r}<>"Yes",'
                'IF({act_cost}{r}="","Add - no final cost",'
                '"Unpaid - invoice not settled"),"OK"),'
                'IF({priority}{r}="Urgent","Urgent - guest cannot use '
                'the property",'
                'IF({vendor}{r}="","Action - nobody assigned",'
                'IF({scheduled}{r}="","Schedule - no date booked in",'
                'IF({due_in}{r}<0,"Overdue - booked in for '
                '"&TEXT({scheduled}{r},"d mmm"),'
                'IF({due_in}{r}<=MaintenanceWarnDays,"Due - '
                '"&TEXT({scheduled}{r},"d mmm"),'
                'IF({days_open}{r}>30,"Chase - open over a month",'
                '"OK - scheduled"))))))))'),
            cf=alert_cf(), group="Where it stands"),
        Col("notes", "Notes", 30, "wrap", group="Where it stands"),
    ]

    return SheetSpec(
        name=NAME,
        title="Maintenance Log",
        subtitle="What is broken, who is fixing it, when, and what it "
                 "actually cost",
        columns=columns,
        rows=cfg.row_count(NAME),
        freeze_cols=2,
        samples=sample("maintenance", cfg),
        help_text="One row per job, from the moment somebody notices it. "
                  "Only the actual cost counts towards profit - an "
                  "estimate is a plan. If a job stops you letting the "
                  "property, mark it as blocking and put a Maintenance "
                  "block on the Booking Tracker so the dates come off "
                  "the calendar.",
        extra_notes=[
            "Days open counts from the report date, and keeps counting "
            "until the job is finished. Anything over a month goes red, "
            "because a job open that long is usually a job nobody owns.",
            "Keep the Urgent priority for things that stop a guest "
            "using the property. If everything is urgent, nothing is.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index

    row = sheet.section(row, "The job list", last)
    row = sheet.table_head(row, 0, [
        "Jobs logged", "Still open", "Urgent open", "Blocking lettings",
        "Avg days to fix", "Estimated", "Actually spent", "Over estimate",
        "Unpaid invoices"],
        [12, 11, 12, 16, 15, 12, 14, 13, 14])
    sheet.calc_row(row, 0, [
        '=COUNT([[Maintenance Log.report_date]])',
        '=COUNTIFS([[Maintenance Log.is_open]],"Yes")',
        '=COUNTIFS([[Maintenance Log.is_open]],"Yes",'
        '[[Maintenance Log.priority]],"Urgent")',
        '=COUNTIFS([[Maintenance Log.is_open]],"Yes",'
        '[[Maintenance Log.blocks_calendar]],"Yes")',
        '=IF(COUNTIFS([[Maintenance Log.is_open]],"No")<=0,"",'
        'ROUND(SUMIFS([[Maintenance Log.days_open]],'
        '[[Maintenance Log.is_open]],"No")/'
        'COUNTIFS([[Maintenance Log.is_open]],"No"),1))',
        '=ROUND(SUM([[Maintenance Log.est_cost]]),2)',
        '=ROUND(SUM([[Maintenance Log.cost_used]]),2)',
        '=ROUND(SUM([[Maintenance Log.overrun]]),2)',
        '=COUNTIFS([[Maintenance Log.is_open]],"No",'
        '[[Maintenance Log.paid]],"<>Yes")',
    ], ["int", "int", "int", "int", "number1", "money", "money", "money",
        "int"])
    row += 2

    row = sheet.section(row, "What keeps breaking", last)
    row = sheet.table_head(row, 0, [
        "Category", "Jobs", "Open", "Spent", "Avg cost", "Avg days",
        "Share of spend"],
        [20, 8, 8, 12, 11, 10, 13])
    total = 'SUM([[Maintenance Log.cost_used]])'
    for name in MAINTENANCE_CATEGORIES:
        key = f'"{name}"'
        cat = f'[[Maintenance Log.category]],{key}'
        spend = f'SUMIFS([[Maintenance Log.cost_used]],{cat})'
        jobs = f'COUNTIFS({cat})'
        closed = f'COUNTIFS({cat},[[Maintenance Log.is_open]],"No")'
        sheet.calc_row(row, 0, [
            name,
            f"={jobs}",
            f'=COUNTIFS({cat},[[Maintenance Log.is_open]],"Yes")',
            f"=ROUND({spend},2)",
            f'=IF({jobs}<=0,"",ROUND({spend}/{jobs},2))',
            f'=IF({closed}<=0,"",ROUND(SUMIFS('
            f'[[Maintenance Log.days_open]],{cat},'
            f'[[Maintenance Log.is_open]],"No")/{closed},1))',
            f'=IF({total}<=0,"",{spend}/{total})',
        ], ["text", "int", "int", "money", "money", "number1", "percent"])
        row += 1
````

### `airbnb_rental_spreadsheet/sheets/reviews.py`

*250 lines · 2 functions*

> `Reviews Tracker` - the seven scores, and which one is costing you.
>
> An overall rating is nearly useless on its own.  A 4.6 made of straight
> 4.6s is a property that needs re-photographing; a 4.6 made of six 4.9s
> and a 3.4 for cleanliness is one cleaner having a bad month.  This tab
> keeps all seven category scores side by side so the weakest one is
> obvious, per property and per platform.
>
> It also tracks the thing hosts forget: whether you asked for the review
> at all, and whether you replied to it.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Reviews Tracker').

````python
"""`Reviews Tracker` - the seven scores, and which one is costing you.

An overall rating is nearly useless on its own.  A 4.6 made of straight
4.6s is a property that needs re-photographing; a 4.6 made of six 4.9s
and a 3.4 for cleanliness is one cleaner having a bad month.  This tab
keeps all seven category scores side by side so the weakest one is
obvious, per property and per platform.

It also tracks the thing hosts forget: whether you asked for the review
at all, and whether you replied to it.
"""

from __future__ import annotations

from ..config import RATING_CATEGORIES, Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from
from .common import (alert_cf, blank, flag_yes_no_cf, id_col, lookup_booking,
                     month_col, rating_cf, score_validation, status_cf,
                     year_col, yesno)

NAME = "Reviews Tracker"


def spec(cfg: Config) -> SheetSpec:
    score_keys = [k for k, _h in RATING_CATEGORIES if k != "rate_overall"]
    detail_sum = "+".join(f'IFERROR({{{k}}}{{r}}*1,0)' for k in score_keys)
    detail_count = "+".join(f'({{{k}}}{{r}}<>"")' for k in score_keys)

    columns = [
        id_col("booking", "ReviewPrefix", digits=cfg.digits("review"),
               width=10, header="Review ID"),
        Col("booking", "Booking", 11, "text",
            validate=list_from("BookingList"),
            note="The stay being reviewed. One row per review.",
            group="The stay"),
        Col("property", "Property", 11, "text",
            formula=lookup_booking("property"), hidden=True,
            group="The stay"),
        Col("property_name", "Property", 19, "text",
            formula=lookup_booking("property_name"), group="The stay"),
        Col("guest", "Guest", 11, "text",
            formula=lookup_booking("guest"), hidden=True,
            group="The stay"),
        Col("guest_name", "Guest", 18, "text",
            formula=lookup_booking("guest_name"), group="The stay"),
        Col("platform", "Platform", 14, "text",
            formula=lookup_booking("platform"), group="The stay"),
        Col("checkout", "Checked out", 12, "date",
            formula=lookup_booking("check_out"), group="The stay"),

        # ------------------------------------------------- the process --
        Col("requested", "Review asked for", 13, "date",
            validate=date_check(),
            note="The date you asked. Most platforms prompt the guest "
                 "automatically, but a short personal message still "
                 "moves the needle.",
            group="The process"),
        Col("status", "Status", 15, "status",
            validate=list_from("ReviewStatusList"),
            cf=status_cf(), group="The process"),
        Col("review_date", "Review received", 13, "date",
            validate=date_check(), group="The process"),
        month_col("review_date"),
        year_col("review_date"),
        Col("days_to_review", "Days to review", 13, "int",
            formula=blank(
                "booking",
                'IF({review_date}{r}="","",IF({checkout}{r}="","",'
                'MAX(0,{review_date}{r}-{checkout}{r})))'),
            hidden=True, group="The process"),
    ]

    for key, header in RATING_CATEGORIES:
        columns.append(
            Col(key, header, 11, "number1",
                validate=score_validation(5),
                cf=rating_cf() if key == "rate_overall" else [],
                note=f"{header} out of 5. Leave blank if the platform "
                     f"does not score it.",
                group="The scores"))

    columns += [
        Col("has_score", "Scored", 8, "flag",
            formula=blank("booking",
                          'IF({rate_overall}{r}="","No","Yes")'),
            cf=flag_yes_no_cf(), hidden=True, group="The scores"),
        Col("detail_avg", "Category average", 14, "number",
            formula=blank(
                "booking",
                f'IF(({detail_count})<=0,"",'
                f'ROUND(({detail_sum})/({detail_count}),2))'),
            cf=rating_cf(),
            note="The six category scores averaged. Compare it with the "
                 "overall - a gap means the guest liked the place more "
                 "than the sum of its parts, or a lot less.",
            group="The scores"),
        Col("weakest", "Weakest category", 16, "text",
            formula=blank(
                "booking",
                'IF({rate_clean}{r}="","",'
                'IF({rate_clean}{r}<=MIN(IFERROR({rate_accuracy}{r}*1,5),'
                'IFERROR({rate_checkin}{r}*1,5),'
                'IFERROR({rate_comm}{r}*1,5),'
                'IFERROR({rate_location}{r}*1,5),'
                'IFERROR({rate_value}{r}*1,5)),"Cleanliness",'
                'IF(IFERROR({rate_value}{r}*1,5)<=MIN('
                'IFERROR({rate_accuracy}{r}*1,5),'
                'IFERROR({rate_checkin}{r}*1,5),'
                'IFERROR({rate_comm}{r}*1,5),'
                'IFERROR({rate_location}{r}*1,5)),"Value",'
                'IF(IFERROR({rate_accuracy}{r}*1,5)<=MIN('
                'IFERROR({rate_checkin}{r}*1,5),'
                'IFERROR({rate_comm}{r}*1,5),'
                'IFERROR({rate_location}{r}*1,5)),"Accuracy",'
                'IF(IFERROR({rate_checkin}{r}*1,5)<=MIN('
                'IFERROR({rate_comm}{r}*1,5),'
                'IFERROR({rate_location}{r}*1,5)),"Check-in",'
                'IF(IFERROR({rate_comm}{r}*1,5)<='
                'IFERROR({rate_location}{r}*1,5),"Communication",'
                '"Location"))))))'),
            note="Which of the six scored categories came out lowest on "
                 "this review. The column to filter when a rating "
                 "starts sliding.",
            group="The scores"),
        Col("below_target", "Under target", 11, "flag",
            formula=blank(
                "booking",
                'IF({rate_overall}{r}="","No",'
                'IF({rate_overall}{r}<TargetRating,"Yes","No"))'),
            cf=[{"type": "text", "criteria": "containing", "value": "Yes",
                 "fmt": "bad"}],
            hidden=True, group="The scores"),

        # ------------------------------------------------- what was said --
        Col("public_review", "What they wrote", 36, "wrap",
            group="What was said"),
        Col("private_note", "Private feedback", 30, "wrap",
            note="The bit only you see. Usually more useful than the "
                 "public review.",
            group="What was said"),
        Col("responded", "You replied", 11, "flag",
            validate=yesno("Have you replied publicly?"),
            cf=flag_yes_no_cf(),
            note="Always reply to a poor review. Future guests read "
                 "the reply, not the complaint.",
            group="What was said"),
        Col("response_date", "Replied on", 12, "date",
            validate=date_check(), hidden=True,
            group="What was said"),
        Col("response", "Your reply", 32, "wrap",
            group="What was said"),
        Col("action_taken", "What you changed", 28, "wrap",
            note="The point of a bad review. New mattress, earlier "
                 "cleaner, clearer directions.",
            group="What was said"),

        Col("alert", "Action", 28, "status",
            formula=blank(
                "booking",
                'IF({checkout}{r}="","Check - booking not found",'
                'IF({has_score}{r}="No",'
                'IF({status}{r}="No review left","",'
                'IF(TodayDate-{checkout}{r}>ReviewChaseDays,'
                '"Chase - no review yet","Due - review window open")),'
                'IF({below_target}{r}="Yes",'
                'IF({responded}{r}<>"Yes","Action - reply to this '
                'review","OK - replied"),'
                'IF({responded}{r}<>"Yes","Follow up - thank them",'
                '"OK"))))'),
            cf=alert_cf(), group="What was said"),
    ]

    return SheetSpec(
        name=NAME,
        title="Reviews Tracker",
        subtitle="All seven scores, per property and per platform - so "
                 "you can see which one is actually costing you",
        columns=columns,
        rows=cfg.row_count(NAME),
        freeze_cols=2,
        samples=sample("reviews", cfg),
        help_text="One row per review. Pick the booking and the "
                  "property, guest and platform come with it. Fill in "
                  "whichever of the seven scores the platform gives "
                  "you - the weakest category is worked out for you, "
                  "and that is the column worth sorting on when a "
                  "rating starts to slide.",
        extra_notes=[
            "Always reply, and always reply to the bad ones. Future "
            "guests read the reply far more carefully than the "
            "complaint.",
            "A review with no score is still worth a row. The 'No "
            "review left' status stops the chase without pretending "
            "they gave you five stars.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    scored = 'COUNTIFS([[Reviews Tracker.has_score]],"Yes")'

    row = sheet.section(row, "How you are doing", last)
    row = sheet.table_head(row, 0, [
        "Reviews logged", "Scored", "Review rate", "Average overall",
        "Under target", "5 stars", "Replied to", "Avg days to review"],
        [13, 9, 12, 14, 12, 9, 11, 16])
    sheet.calc_row(row, 0, [
        '=COUNTIF([[Reviews Tracker.booking]],"?*")',
        f"={scored}",
        '=IF(COUNTIF([[Reviews Tracker.booking]],"?*")<=0,"",'
        f'{scored}/COUNTIF([[Reviews Tracker.booking]],"?*"))',
        f'=IF({scored}<=0,"",ROUND(SUMIFS('
        '[[Reviews Tracker.rate_overall]],'
        f'[[Reviews Tracker.has_score]],"Yes")/{scored},2))',
        '=COUNTIFS([[Reviews Tracker.below_target]],"Yes")',
        '=COUNTIFS([[Reviews Tracker.rate_overall]],5)',
        '=COUNTIFS([[Reviews Tracker.responded]],"Yes")',
        f'=IF({scored}<=0,"",ROUND(SUM('
        f'[[Reviews Tracker.days_to_review]])/{scored},1))',
    ], ["int", "int", "percent", "number", "int", "int", "int", "number1"])
    row += 2

    row = sheet.section(row, "Where the stars are going", last)
    headers = ["Category", "Reviews scored", "Average", "Under 4",
               "5 stars", "Weakest on"]
    row = sheet.table_head(row, 0, headers, [20, 14, 11, 10, 9, 12])
    for key, label in RATING_CATEGORIES:
        col = f"[[Reviews Tracker.{key}]]"
        count = f'COUNT({col})'
        sheet.calc_row(row, 0, [
            label,
            f"={count}",
            f'=IF({count}<=0,"",ROUND(SUM({col})/{count},2))',
            f'=COUNTIFS({col},"<4",{col},">0")',
            f'=COUNTIFS({col},5)',
            f'=COUNTIFS([[Reviews Tracker.weakest]],"{label}")',
        ], ["text", "int", "number", "int", "int", "int"])
        row += 1

    row += 1
    sheet.note(row, "'Weakest on' counts how many reviews put this "
                    "category bottom. One category holding the bottom "
                    "spot again and again is the thing to fix - and it "
                    "is almost always cleanliness or value, both of "
                    "which you control.", last)
````

### `airbnb_rental_spreadsheet/sheets/vendors.py`

*240 lines · 2 functions*

> `Staff & Vendors` - the people who keep the properties running.
>
> Cleaners, the laundry, a handyman, the plumber you only ring in
> February, the co-host who covers you in August.  One row each, with
> what they charge, where they work and how much work they have actually
> done for you.
>
> The job counts and the money are pulled from the Cleaning Schedule and
> the Maintenance Log, so this tab is never out of date and never has to
> be filled in twice.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Staff & Vendors').

````python
"""`Staff & Vendors` - the people who keep the properties running.

Cleaners, the laundry, a handyman, the plumber you only ring in
February, the co-host who covers you in August.  One row each, with
what they charge, where they work and how much work they have actually
done for you.

The job counts and the money are pulled from the Cleaning Schedule and
the Maintenance Log, so this tab is never out of date and never has to
be filled in twice.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec, max_if
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, flag_yes_no_cf, id_col,
                     money_cf, status_cf, yesno)

NAME = "Staff & Vendors"


def spec(cfg: Config) -> SheetSpec:
    cl_cleaner = "[[Cleaning Schedule.cleaner]]"
    cl_cost = "[[Cleaning Schedule.clean_cost]]"
    cl_hours = "[[Cleaning Schedule.hours]]"
    cl_paid = "[[Cleaning Schedule.paid]]"
    cl_comp = "[[Cleaning Schedule.completion]]"
    mt_vendor = "[[Maintenance Log.vendor]]"
    mt_cost = "[[Maintenance Log.cost_used]]"
    mt_open = "[[Maintenance Log.is_open]]"
    mt_paid = "[[Maintenance Log.paid]]"

    _is_me = cl_cleaner + "={id}{r}"
    _last_clean = max_if("[[Cleaning Schedule.clean_num]]", _is_me)

    columns = [
        id_col("name", "VendorPrefix", digits=cfg.digits("vendor"),
               width=10, header="Vendor ID"),
        Col("name", "Name", 22, "text",
            note="A person or a company. Cleaning jobs and maintenance "
                 "jobs both point at this ID.",
            group="Who they are"),
        Col("role", "Role", 18, "text",
            validate=list_from("StaffRoleList"), group="Who they are"),
        Col("company", "Company", 20, "text", group="Who they are"),
        Col("phone", "Phone", 15, "text",
            note="Printed on the turnover checklist, so a cleaner "
                 "standing in a flat with a broken boiler has the "
                 "plumber's number.",
            group="Who they are"),
        Col("email", "Email", 24, "text", group="Who they are"),
        Col("properties", "Covers", 22, "text",
            note="Which properties they work on. Free text - 'all', or "
                 "the names.",
            group="Who they are"),
        Col("availability", "Available", 15, "text",
            validate=list_from("AvailabilityList"), group="Who they are"),

        # ----------------------------------------------------- the deal --
        Col("pay_basis", "Paid", 18, "text",
            validate=list_from("PayBasisList"), group="The deal"),
        Col("rate", "Rate", 11, "money",
            validate=number_check(0, 100000, "Rate"),
            note="Whatever the basis says - per turnover, per hour, per "
                 "month. New cleaning jobs do not read this "
                 "automatically; the property's cleaning cost does "
                 "that, so a cleaner who charges differently per "
                 "property still works.",
            group="The deal"),
        Col("contract", "Contract", 14, "status",
            validate=list_from("ContractStatusList"),
            cf=status_cf(), group="The deal"),
        Col("started", "Working with you since", 14, "date",
            validate=date_check(), hidden=True, group="The deal"),
        Col("insured", "Insured", 9, "flag",
            validate=yesno("Do they carry their own liability cover?"),
            cf=flag_yes_no_cf(),
            note="Public liability cover. Worth asking for, and worth "
                 "writing down when you do.",
            group="The deal"),
        Col("insurance_expiry", "Cover expires", 13, "date",
            validate=date_check(), group="The deal"),
        Col("cover_days", "Days left", 10, "int",
            formula=blank("name",
                          'IF({insurance_expiry}{r}="","",'
                          '{insurance_expiry}{r}-TodayDate)'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": 30,
                 "fmt": "warn"}],
            hidden=True, group="The deal"),
        Col("keyholder", "Holds keys", 11, "flag",
            validate=yesno(), cf=flag_yes_no_cf(), group="The deal"),

        # ------------------------------------------------- the workload --
        Col("cleans", "Turnovers done", 13, "int",
            formula=blank("name", f'COUNTIFS({cl_cleaner},{{id}}{{r}})'),
            cf=bar_cf(), group="Workload"),
        Col("cleans_done", "Finished properly", 14, "int",
            formula=blank("name", f'COUNTIFS({cl_cleaner},{{id}}{{r}},'
                                  f'{cl_comp},1)'),
            hidden=True, group="Workload"),
        Col("clean_quality", "Completed in full", 14, "percent",
            formula=blank("name",
                          'IF({cleans}{r}<=0,"",'
                          '{cleans_done}{r}/{cleans}{r})'),
            cf=[{"type": "cell", "criteria": ">=", "value": 0.98,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0.9,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": 0.98,
                 "fmt": "warn"}],
            note="How often every one of the nine turnover tasks got "
                 "ticked. The most useful number on this tab.",
            group="Workload"),
        Col("clean_hours", "Hours on site", 12, "number1",
            formula=blank("name", f'ROUND(SUMIFS({cl_hours},'
                                  f'{cl_cleaner},{{id}}{{r}}),1)'),
            hidden=True, group="Workload"),
        Col("avg_hours", "Avg a turnover", 13, "number1",
            formula=blank("name",
                          'IF({clean_hours}{r}<=0,"",'
                          'ROUND({clean_hours}{r}/MAX(1,{cleans}{r}),1))'),
            hidden=True, group="Workload"),
        Col("jobs", "Repairs done", 12, "int",
            formula=blank("name", f'COUNTIFS({mt_vendor},{{id}}{{r}})'),
            group="Workload"),
        Col("jobs_open", "Repairs open", 12, "int",
            formula=blank("name", f'COUNTIFS({mt_vendor},{{id}}{{r}},'
                                  f'{mt_open},"Yes")'),
            group="Workload"),
        Col("last_job", "Last turnover", 13, "date",
            formula=blank("name",
                          f'IF({{cleans}}{{r}}<=0,"",{_last_clean})'),
            group="Workload"),

        # ---------------------------------------------------- the money --
        Col("paid_cleaning", "Paid for cleaning", 14, "money",
            formula=blank("name", f'ROUND(SUMIFS({cl_cost},'
                                  f'{cl_cleaner},{{id}}{{r}}),2)'),
            group="Money"),
        Col("paid_repairs", "Paid for repairs", 14, "money",
            formula=blank("name", f'ROUND(SUMIFS({mt_cost},'
                                  f'{mt_vendor},{{id}}{{r}}),2)'),
            group="Money"),
        Col("paid_total", "Total paid", 12, "money",
            formula=blank("name",
                          'ROUND({paid_cleaning}{r}+{paid_repairs}{r},2)'),
            cf=money_cf(), group="Money"),
        Col("owed", "Still owed", 11, "money",
            formula=blank(
                "name",
                f'ROUND(SUMIFS({cl_cost},{cl_cleaner},{{id}}{{r}},'
                f'{cl_paid},"<>Yes")+SUMIFS({mt_cost},'
                f'{mt_vendor},{{id}}{{r}},{mt_paid},"<>Yes"),2)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0.004,
                 "fmt": "warn"}],
            note="Cleaning jobs and repairs that are not marked paid.",
            group="Money"),
        Col("spend_rank", "Rank", 7, "int",
            formula=('=IF({name}{r}="","",COUNTIFS({paid_total_first}:'
                     '{paid_total_last},">"&{paid_total}{r})+'
                     'COUNTIFS({paid_total_first}:{paid_total}{r},'
                     '{paid_total}{r}))'),
            hidden=True, group="Money"),

        Col("alert", "Action", 28, "status",
            formula=blank(
                "name",
                'IF({contract}{r}="Ended","",'
                'IF({contract}{r}="Do not use","Check - marked do not '
                'use",'
                'IF({owed}{r}>0.004,"Unpaid - "&TEXT({owed}{r},"0")&'
                '" outstanding",'
                'IF({cover_days}{r}="",'
                'IF({insured}{r}="Yes","Add - no insurance expiry date",'
                '"OK"),'
                'IF({cover_days}{r}<0,"Expired - insurance cover ran '
                'out",'
                'IF({cover_days}{r}<30,"Due - insurance expires soon",'
                'IF({jobs_open}{r}>0,"Follow up - "&{jobs_open}{r}&'
                '" open repair(s)",'
                'IF({clean_quality}{r}="","OK",'
                'IF({clean_quality}{r}<0.9,"Check - turnovers often '
                'left incomplete","OK")))))))))'),
            cf=alert_cf(), group="Money"),
        Col("notes", "Notes", 28, "wrap", group="Money"),
    ]

    return SheetSpec(
        name=NAME,
        title="Staff & Vendors",
        subtitle="Cleaners, contractors and co-hosts - what they charge "
                 "and what they have actually done",
        columns=columns,
        rows=cfg.row_count(NAME),
        freeze_cols=2,
        samples=sample("vendors", cfg),
        help_text="One row per person or company. The turnover count, "
                  "the repair count and everything you have paid them "
                  "are read straight from the Cleaning Schedule and "
                  "the Maintenance Log, so this tab never needs "
                  "updating by hand.",
        extra_notes=[
            "'Completed in full' is how often a cleaner ticked all "
            "nine turnover tasks. It is the closest thing to a quality "
            "score you will get without standing in the flat yourself.",
            "The rate here is for reference. New cleaning jobs price "
            "themselves from the property's cleaning cost instead, so "
            "a cleaner who charges more for the big house still comes "
            "out right.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index

    row = sheet.section(row, "Your team", last)
    row = sheet.table_head(row, 0, [
        "On the books", "Active", "Cleaners", "Contractors", "Key holders",
        "Insurance expiring", "Total paid", "Still owed"],
        [12, 9, 10, 12, 12, 17, 12, 11])
    sheet.calc_row(row, 0, [
        '=COUNTIF([[Staff & Vendors.name]],"?*")',
        '=COUNTIFS([[Staff & Vendors.contract]],"Active")',
        '=COUNTIFS([[Staff & Vendors.role]],"Cleaner")+'
        'COUNTIFS([[Staff & Vendors.role]],"Head cleaner")',
        '=COUNTIFS([[Staff & Vendors.jobs]],">0")',
        '=COUNTIFS([[Staff & Vendors.keyholder]],"Yes")',
        '=COUNTIFS([[Staff & Vendors.cover_days]],"<30",'
        '[[Staff & Vendors.cover_days]],">=-3650")',
        '=ROUND(SUM([[Staff & Vendors.paid_total]]),2)',
        '=ROUND(SUM([[Staff & Vendors.owed]]),2)',
    ], ["int", "int", "int", "int", "int", "int", "money", "money"])
````

---

## 8. The pages

The hand-built tabs: the portfolio dashboard, the welcome, the settings tab every formula reads, the property-by-month performance register, and the monthly and annual reports.

### `airbnb_rental_spreadsheet/sheets/performance.py`

*348 lines · 2 functions*

> `Occupancy & Performance` - one row per property per month.
>
> This is the tab a short-term rental investor opens first.  Type a
> property and the first of a month; everything else - available nights,
> booked nights, occupancy, ADR, RevPAR, average stay, the cleaning cost
> per booking, the operating cost per night and the profit - is worked
> out from the bookings and the costs you have already entered.
>
> Three definitions are worth stating plainly, because half the
> spreadsheets sold online get them wrong:
>
> * **Occupancy** is booked nights over *available* nights, and a night
>   you blocked for yourself was never available.
> * **ADR** is accommodation revenue over booked nights.  Cleaning fees
>   are not accommodation revenue.
> * **RevPAR** is accommodation revenue over available nights, which is
>   the only one of the three that cannot be flattered by taking fewer
>   bookings at a higher price.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Occupancy & Performance').

````python
"""`Occupancy & Performance` - one row per property per month.

This is the tab a short-term rental investor opens first.  Type a
property and the first of a month; everything else - available nights,
booked nights, occupancy, ADR, RevPAR, average stay, the cleaning cost
per booking, the operating cost per night and the profit - is worked
out from the bookings and the costs you have already entered.

Three definitions are worth stating plainly, because half the
spreadsheets sold online get them wrong:

* **Occupancy** is booked nights over *available* nights, and a night
  you blocked for yourself was never available.
* **ADR** is accommodation revenue over booked nights.  Cleaning fees
  are not accommodation revenue.
* **RevPAR** is accommodation revenue over available nights, which is
  the only one of the three that cannot be flattered by taking fewer
  bookings at a higher price.
"""

from __future__ import annotations

from ..config import Config
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from
from .common import (alert_cf, bar_cf, blank, id_col, loss_cf, 
                    lookup_property, money_cf, target_cf)

NAME = "Occupancy & Performance"


def spec(cfg: Config) -> SheetSpec:
    bk_prop = "[[Booking Tracker.property]]"
    bk_month = "[[Booking Tracker.month]]"
    bk_earns = "[[Booking Tracker.earns]]"
    bk_block = "[[Booking Tracker.is_block]]"
    bk_cancel = "[[Booking Tracker.is_cancelled]]"
    bk_nights = "[[Booking Tracker.nights]]"
    bk_accom = "[[Booking Tracker.accommodation]]"
    bk_gross = "[[Booking Tracker.gross]]"
    bk_net = "[[Booking Tracker.net_revenue]]"
    bk_guests = "[[Booking Tracker.guests_total]]"
    bk_fees = "[[Booking Tracker.platform_fee_total]]"
    cl_prop = "[[Cleaning Schedule.property]]"
    cl_month = "[[Cleaning Schedule.month]]"
    cl_cost = "[[Cleaning Schedule.total_cost]]"
    ex_prop = "[[Expense Tracker.property]]"
    ex_month = "[[Expense Tracker.month]]"
    ex_pl = "[[Expense Tracker.pl_amount]]"
    mt_prop = "[[Maintenance Log.property]]"
    mt_month = "[[Maintenance Log.month]]"
    mt_cost = "[[Maintenance Log.cost_used]]"
    rv_prop = "[[Reviews Tracker.property]]"
    rv_month = "[[Reviews Tracker.month]]"
    rv_score = "[[Reviews Tracker.rate_overall]]"
    rv_has = "[[Reviews Tracker.has_score]]"

    def by_month(value: str, prop_range: str, month_range: str,
                 *extra: str) -> str:
        tail = "".join(f",{c}" for c in extra)
        return (f'SUMIFS({value},{prop_range},{{property}}{{r}},'
                f'{month_range},{{month}}{{r}}{tail})')

    def count_month(prop_range: str, month_range: str, *extra: str) -> str:
        tail = "".join(f",{c}" for c in extra)
        return (f'COUNTIFS({prop_range},{{property}}{{r}},'
                f'{month_range},{{month}}{{r}}{tail})')

    # a COUNTIFS/SUMIFS criterion is TWO arguments - the range and the
    # thing it is compared against.  Collapsing them into one argument
    # as `range="Yes"` makes Excel answer #VALUE! on every row.
    _scored = (rv_has, '"Yes"')
    _rv_count = count_month(rv_prop, rv_month, *_scored)
    _rv_total = by_month(rv_score, rv_prop, rv_month, *_scored)

    columns = [
        id_col("period", "PeriodPrefix", digits=cfg.digits("period"),
               width=10, header="Row ID"),
        Col("property", "Property", 11, "text",
            validate=list_from("PropertyList"),
            note="Which property this row measures.",
            group="The period"),
        Col("property_name", "Property name", 19, "text",
            formula=lookup_property("name"), group="The period"),
        Col("period", "Month", 12, "date",
            validate=date_check(),
            note="The first of the month, e.g. 01 Mar 2026. Every row "
                 "on this tab starts here. One row per property per "
                 "month.",
            group="The period"),
        Col("month", "Month key", 10, "text",
            formula=blank("period", 'TEXT({period}{r},"yyyy-mm")'),
            hidden=True, group="The period"),
        Col("month_label", "Period", 13, "text",
            formula=blank("period", 'TEXT({period}{r},"mmm yyyy")'),
            group="The period"),
        Col("days_in_month", "Days", 7, "int",
            formula=blank(
                "period",
                'DAY(DATE(YEAR({period}{r}),MONTH({period}{r})+1,1)-1)'),
            hidden=True, group="The period"),

        # ------------------------------------------------- the nights --
        Col("nights_blocked", "Blocked", 10, "int",
            formula=blank(
                "period",
                by_month(bk_nights, bk_prop, bk_month,
                         f'{bk_block},"Yes"')),
            note="Owner stays and maintenance closures that started in "
                 "this month.",
            group="Nights"),
        Col("nights_available", "Available", 11, "int",
            formula=blank(
                "period",
                'MAX(0,{days_in_month}{r}-{nights_blocked}{r})'),
            note="Days in the month less the nights you took off the "
                 "market yourself.",
            group="Nights"),
        Col("nights_booked", "Booked", 10, "int",
            formula=blank(
                "period",
                by_month(bk_nights, bk_prop, bk_month,
                         f'{bk_earns},"Yes"')),
            cf=bar_cf(), group="Nights"),
        Col("occupancy", "Occupancy", 11, "percent",
            formula=blank(
                "period",
                'IF({nights_available}{r}<=0,"",'
                '{nights_booked}{r}/{nights_available}{r})'),
            cf=target_cf("TargetOccupancy"), group="Nights"),
        Col("bookings", "Bookings", 10, "int",
            formula=blank("period",
                          count_month(bk_prop, bk_month,
                                      f'{bk_earns},"Yes"')),
            group="Nights"),
        Col("cancellations", "Cancelled", 10, "int",
            formula=blank("period",
                          count_month(bk_prop, bk_month,
                                      f'{bk_cancel},"Yes"')),
            group="Nights"),
        Col("cancel_rate", "Cancel rate", 11, "percent",
            formula=blank(
                "period",
                'IF({bookings}{r}+{cancellations}{r}<=0,"",'
                '{cancellations}{r}/({bookings}{r}+{cancellations}{r}))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0.1,
                 "fmt": "warn"}],
            hidden=True, group="Nights"),
        Col("guests", "Guests", 9, "int",
            formula=blank("period",
                          by_month(bk_guests, bk_prop, bk_month,
                                   f'{bk_earns},"Yes"')),
            hidden=True, group="Nights"),
        Col("alos", "Avg stay", 9, "number1",
            formula=blank(
                "period",
                'IF({bookings}{r}<=0,"",'
                'ROUND({nights_booked}{r}/{bookings}{r},1))'),
            group="Nights"),

        # ------------------------------------------------ the headline --
        Col("accommodation", "Room revenue", 13, "money",
            formula=blank("period",
                          'ROUND(' + by_month(bk_accom, bk_prop, bk_month,
                                              f'{bk_earns},"Yes"') + ',2)'),
            group="Rates"),
        Col("adr", "ADR", 11, "money",
            formula=blank(
                "period",
                'IF({nights_booked}{r}<=0,"",'
                'ROUND({accommodation}{r}/{nights_booked}{r},2))'),
            cf=target_cf("TargetADR"),
            note="Average daily rate. Accommodation only - the cleaning "
                 "fee is a cost recovery and stays out of it.",
            group="Rates"),
        Col("revpar", "RevPAR", 11, "money",
            formula=blank(
                "period",
                'IF({nights_available}{r}<=0,"",'
                'ROUND({accommodation}{r}/{nights_available}{r},2))'),
            cf=target_cf("TargetRevPAR"),
            note="Revenue per available night. The honest one: you "
                 "cannot improve it by taking fewer bookings at a "
                 "higher price.",
            group="Rates"),

        # ---------------------------------------------------- the money --
        Col("gross", "Gross", 12, "money",
            formula=blank("period",
                          'ROUND(' + by_month(bk_gross, bk_prop, bk_month,
                                              f'{bk_earns},"Yes"') + ',2)'),
            group="Money"),
        Col("platform_fees", "Platform fees", 12, "money",
            formula=blank("period",
                          'ROUND(' + by_month(bk_fees, bk_prop, bk_month,
                                              f'{bk_earns},"Yes"') + ',2)'),
            group="Money"),
        Col("net_revenue", "Net revenue", 12, "money",
            formula=blank("period",
                          'ROUND(' + by_month(bk_net, bk_prop, bk_month,
                                              f'{bk_earns},"Yes"') + ',2)'),
            cf=money_cf(), group="Money"),
        Col("cleaning_cost", "Cleaning", 11, "money",
            formula=blank("period",
                          'ROUND(' + by_month(cl_cost, cl_prop,
                                              cl_month) + ',2)'),
            group="Money"),
        Col("maintenance_cost", "Maintenance", 12, "money",
            formula=blank("period",
                          'ROUND(' + by_month(mt_cost, mt_prop,
                                              mt_month) + ',2)'),
            group="Money"),
        Col("other_cost", "Other costs", 12, "money",
            formula=blank("period",
                          'ROUND(' + by_month(ex_pl, ex_prop,
                                              ex_month) + ',2)'),
            group="Money"),
        Col("total_cost", "Total costs", 12, "money",
            formula=blank("period",
                          'ROUND({cleaning_cost}{r}+'
                          '{maintenance_cost}{r}+{other_cost}{r},2)'),
            group="Money"),
        Col("profit", "Operating profit", 14, "money",
            formula=blank("period",
                          'ROUND({net_revenue}{r}-{total_cost}{r},2)'),
            cf=loss_cf(), group="Money"),
        Col("margin", "Margin", 9, "percent",
            formula=blank("period",
                          'IF({net_revenue}{r}<=0,"",'
                          '{profit}{r}/{net_revenue}{r})'),
            cf=target_cf("TargetMargin"), group="Money"),

        # -------------------------------------------------- unit costs --
        Col("cost_per_night", "Cost a night", 12, "money",
            formula=blank(
                "period",
                'IF({nights_booked}{r}<=0,"",'
                'ROUND({total_cost}{r}/{nights_booked}{r},2))'),
            note="What a single booked night costs you to deliver. "
                 "Compare it with ADR - the gap is your real margin.",
            group="Unit costs"),
        Col("clean_per_booking", "Cleaning a booking", 14, "money",
            formula=blank(
                "period",
                'IF({bookings}{r}<=0,"",'
                'ROUND({cleaning_cost}{r}/{bookings}{r},2))'),
            note="The reason a run of one-night stays can lose money at "
                 "a perfectly good nightly rate.",
            group="Unit costs"),
        Col("profit_night", "Profit a night", 12, "money",
            formula=blank(
                "period",
                'IF({nights_booked}{r}<=0,"",'
                'ROUND({profit}{r}/{nights_booked}{r},2))'),
            cf=loss_cf(), group="Unit costs"),
        Col("rating", "Rating", 9, "number",
            formula=blank(
                "period",
                f'IF({_rv_count}<=0,"",'
                f'ROUND({_rv_total}/{_rv_count},2))'),
            note="Average review score for stays at this property that "
                 "were reviewed in this month.",
            hidden=True, group="Unit costs"),

        Col("alert", "Action", 28, "status",
            formula=blank(
                "period",
                'IF({property}{r}="","Add - no property on this row",'
                'IF({nights_available}{r}<=0,"Check - no available '
                'nights in this month",'
                'IF({nights_booked}{r}<=0,"Check - nothing booked at '
                'all",'
                'IF({profit}{r}<0,"Loss - costs beat net revenue",'
                'IF({occupancy}{r}<TargetOccupancy*0.7,"Action - '
                'occupancy well under target",'
                'IF({adr}{r}<TargetADR*0.8,"Check - rate well under '
                'target",'
                'IF({occupancy}{r}>0.95,"Check - almost full, your '
                'rate is probably too low","OK")))))))'),
            cf=alert_cf(), group="Unit costs"),
        Col("notes", "Notes", 28, "wrap",
            note="What happened that month - a heatwave, a festival, a "
                 "boiler, a price experiment. In twelve months this is "
                 "the column that explains the numbers.",
            group="Unit costs"),
    ]

    return SheetSpec(
        name=NAME,
        title="Occupancy & Performance",
        subtitle="One row per property per month: occupancy, ADR, "
                 "RevPAR, costs and profit",
        columns=columns,
        rows=cfg.row_count(NAME),
        freeze_cols=3,
        samples=sample("performance", cfg),
        help_text="Type a property and the first of a month. Everything "
                  "else is read from the bookings and costs you have "
                  "already entered, so a row takes two cells and gives "
                  "you the full picture of that property in that "
                  "month. Keep one row per property per month and the "
                  "year builds itself.",
        extra_notes=[
            "A stay is counted in the month it checks in, so a "
            "Hogmanay booking lands in December. That keeps one "
            "booking in one month and makes the totals add up; it is "
            "the same convention the platforms use on their earnings "
            "summaries.",
            "An occupancy figure above 95% is not a triumph, it is a "
            "pricing error, and the Action column says so.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index

    row = sheet.section(row, "Every period on this tab", last)
    row = sheet.table_head(row, 0, [
        "Rows", "Available nights", "Booked nights", "Occupancy",
        "Room revenue", "ADR", "RevPAR", "Net revenue", "Costs", "Profit"],
        [8, 16, 14, 11, 13, 11, 11, 12, 11, 12])
    avail = 'SUM([[Occupancy & Performance.nights_available]])'
    booked = 'SUM([[Occupancy & Performance.nights_booked]])'
    accom = 'SUM([[Occupancy & Performance.accommodation]])'
    sheet.calc_row(row, 0, [
        '=COUNT([[Occupancy & Performance.period]])',
        f"={avail}",
        f"={booked}",
        f'=IF({avail}<=0,"",{booked}/{avail})',
        f"=ROUND({accom},2)",
        f'=IF({booked}<=0,"",ROUND({accom}/{booked},2))',
        f'=IF({avail}<=0,"",ROUND({accom}/{avail},2))',
        '=ROUND(SUM([[Occupancy & Performance.net_revenue]]),2)',
        '=ROUND(SUM([[Occupancy & Performance.total_cost]]),2)',
        '=ROUND(SUM([[Occupancy & Performance.profit]]),2)',
    ], ["int", "int", "int", "percent", "money", "money", "money",
        "money", "money", "money"])
    row += 2

    sheet.note(row, "These totals cover whatever rows you have typed, "
                    "so they are only a portfolio view if you keep one "
                    "row per property per month. The Monthly and Annual "
                    "Reports tabs read the bookings directly and do not "
                    "depend on this tab being complete.", last)
````

### `airbnb_rental_spreadsheet/sheets/dashboard.py`

*459 lines · 5 functions*

> `Dashboard` - the page that answers "how is the portfolio doing?"
>
> Twenty numbers across the top, the jobs that need doing before you put
> the kettle on, every property side by side on revenue, occupancy,
> costs and profit, and twelve months of income against target.
> Everything reads from the registers, so nothing here is ever out of
> date and there is nothing to refresh.
>
> Five questions, in the order a host actually asks them: who is in
> tonight, who arrives tomorrow and is the flat clean, are we full, are
> we making money, and which property is quietly losing it.

| Defines | What it does |
|---------|--------------|
| `_nth_property(field, n)` | Row n of the Property Database, blank when the row is empty |
| `_by_rank(sheet, rank_key, field, n)` | The nth row of a register by one of its hidden rank columns |
| `_month_key(n)` | The `yyyy-mm` key for n months before this one |
| `_month_sum(value, months, n, *extra)` | One month's total out of a register, by its month key column |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Dashboard'), `LAST_COL` (12), `MONTHS` (12), `PROPERTY_ROWS` (8), `LIST_ROWS` (6).

````python
"""`Dashboard` - the page that answers "how is the portfolio doing?"

Twenty numbers across the top, the jobs that need doing before you put
the kettle on, every property side by side on revenue, occupancy,
costs and profit, and twelve months of income against target.
Everything reads from the registers, so nothing here is ever out of
date and there is nothing to refresh.

Five questions, in the order a host actually asks them: who is in
tonight, who arrives tomorrow and is the flat clean, are we full, are
we making money, and which property is quietly losing it.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, col_letter
from ..writer import SheetWriter

NAME = "Dashboard"
LAST_COL = 12
MONTHS = 12
PROPERTY_ROWS = 8
LIST_ROWS = 6


def _nth_property(field: str, n: int) -> str:
    """Row n of the Property Database, blank when the row is empty."""
    return (f'IFERROR(IF(INDEX([[Property Database.name]],{n})="","",'
            f'INDEX([[Property Database.{field}]],{n})),"")')


def _by_rank(sheet: str, rank_key: str, field: str, n: int) -> str:
    """The nth row of a register by one of its hidden rank columns."""
    return (f'IFERROR(INDEX([[{sheet}.{field}]],MATCH({n},'
            f'[[{sheet}.{rank_key}]],0)),"")')


def _month_key(n: int) -> str:
    """The `yyyy-mm` key for n months before this one."""
    if n == 0:
        return 'TEXT(TodayDate,"yyyy-mm")'
    return (f'TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{n},1),'
            f'"yyyy-mm")')


def _month_sum(value: str, months: str, n: int, *extra: str) -> str:
    """One month's total out of a register, by its month key column."""
    tail = "".join(f",{c}" for c in extra)
    return f'SUMIFS({value},{months},{_month_key(n)}{tail})'


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws

    s.banner("Short-Term Rental Dashboard",
             "Bookings, turnovers, occupancy and profit - all worked "
             "out from the registers you have already filled in.",
             LAST_COL)
    s.set_col(0, 0, 24)
    for column in range(1, LAST_COL + 1):
        s.set_col(column, column, 13)

    bk_earns = '[[Booking Tracker.earns]],"Yes"'
    bk_month = "[[Booking Tracker.month]]"

    # ---------------------------------------------------------- header --
    row = 4
    s.merge(row, 0, 0, "Business", fmts.chrome("label_bold"))
    s.merge(row, 1, 3, None, fmts.calc("text"))
    s.formula(row, 1, "=HostName", fmts.calc("text"))
    s.merge(row, 4, 4, "Month", fmts.chrome("label_right"))
    s.formula(row, 5, '=TEXT(TodayDate,"mmmm yyyy")',
                     fmts.calc("text"))
    s.merge(row, 6, 6, "Properties", fmts.chrome("label_right"))
    s.formula(row, 7, "=PropertyCount", fmts.calc("int"))
    s.merge(row, 8, 8, "Today", fmts.chrome("label_right"))
    s.formula(row, 9, "=TodayDate", fmts.calc("date"))
    s.merge(row, 10, LAST_COL, None, fmts.chrome("note"))
    s.formula(
        row, 10,
        '="Figures update the moment you type. Nothing to refresh."',
        fmts.chrome("note"))
    row += 2

    # ------------------------------------------------------- tonight ----
    row = s.section(row, "Tonight and the next few days", LAST_COL)
    in_house = ('COUNTIFS([[Booking Tracker.check_in]],"<="&TodayDate,'
                '[[Booking Tracker.check_out]],">"&TodayDate,'
                f'{bk_earns})')
    arriving = ('COUNTIFS([[Booking Tracker.check_in]],TodayDate,'
                f'{bk_earns})')
    leaving = ('COUNTIFS([[Booking Tracker.check_out]],TodayDate,'
               f'{bk_earns})')
    upcoming = ('COUNTIFS([[Booking Tracker.check_in]],">"&TodayDate,'
                '[[Booking Tracker.check_in]],"<="&TodayDate+'
                f'UpcomingDays,{bk_earns})')

    s.kpi(row, 0, "Guests in tonight", f"={in_house}", "int",
          note="Stays that started on or before today and have not "
               "checked out")
    s.kpi(row, 2, "Arriving today", f"={arriving}", "int")
    s.kpi(row, 4, "Departing today", f"={leaving}", "int")
    s.kpi(row, 6, "Same-day turnovers",
          '=COUNTIFS([[Cleaning Schedule.same_day]],"Yes",'
          '[[Cleaning Schedule.clean_date]],TodayDate)', "int",
          note="Out and in on the same day - the ones that go wrong")
    s.kpi(row, 8, "Cleans due today",
          '=COUNTIFS([[Cleaning Schedule.clean_date]],TodayDate,'
          '[[Cleaning Schedule.status]],"<>Completed")', "int")
    row = s.kpi(row, 10, f"Arriving in {cfg.upcoming_window_days} days",
                f"={upcoming}", "int")

    s.kpi(row, 0, "Properties guest-ready",
          '=COUNTIFS([[Property Database.alert]],"OK*")', "int",
          note="Nothing outstanding on the property at all")
    s.kpi(row, 2, "Turnovers unfinished",
          '=COUNTIFS([[Cleaning Schedule.clean_date]],"<="&TodayDate,'
          '[[Cleaning Schedule.status]],"<>Completed")', "int")
    s.kpi(row, 4, "Pre-arrival jobs open",
          '=COUNTIFS([[Check-in & Check-out.arrival_date]],'
          '">="&TodayDate,[[Check-in & Check-out.pre_arrival_done]],'
          '"No")', "int",
          note="Access codes, directions, wifi, house rules, welcome "
               "message")
    s.kpi(row, 6, "Double bookings",
          '=COUNTIFS([[Booking Tracker.clash]],"Yes")', "int",
          note="Two paying stays overlapping in one property")
    s.kpi(row, 8, "Urgent repairs",
          '=COUNTIFS([[Maintenance Log.is_open]],"Yes",'
          '[[Maintenance Log.priority]],"Urgent")', "int")
    row = s.kpi(row, 10, "Blocking a letting",
                '=COUNTIFS([[Maintenance Log.is_open]],"Yes",'
                '[[Maintenance Log.blocks_calendar]],"Yes")', "int")
    row += 1

    # ------------------------------------------------- this month ----
    row = s.section(row, "This month", LAST_COL)
    nights_m = _month_sum("[[Booking Tracker.nights]]", bk_month, 0,
                          bk_earns)
    accom_m = _month_sum("[[Booking Tracker.accommodation]]", bk_month,
                         0, bk_earns)
    gross_m = _month_sum("[[Booking Tracker.gross]]", bk_month, 0,
                         bk_earns)
    net_m = _month_sum("[[Booking Tracker.net_revenue]]", bk_month, 0,
                       bk_earns)
    fees_m = _month_sum("[[Booking Tracker.platform_fee_total]]",
                        bk_month, 0, bk_earns)
    clean_m = _month_sum("[[Cleaning Schedule.total_cost]]",
                         "[[Cleaning Schedule.month]]", 0)
    maint_m = _month_sum("[[Maintenance Log.cost_used]]",
                         "[[Maintenance Log.month]]", 0)
    other_m = _month_sum("[[Expense Tracker.pl_amount]]",
                         "[[Expense Tracker.month]]", 0)
    costs_m = f"({clean_m}+{maint_m}+{other_m})"
    profit_m = f"({net_m}-{costs_m})"
    avail_m = ("MAX(1,PropertyCount*DAY(DATE(YEAR(TodayDate),"
               "MONTH(TodayDate)+1,1)-1))")

    s.kpi(row, 0, "Bookings",
          f'=COUNTIFS({bk_month},{_month_key(0)},{bk_earns})', "int")
    s.kpi(row, 2, "Nights sold", f"={nights_m}", "int")
    s.kpi(row, 4, "Occupancy",
          f'=IF({avail_m}<=0,"",{nights_m}/{avail_m})', "percent",
          note="Nights sold over every night every property could have "
               "sold")
    s.kpi(row, 6, "ADR",
          f'=IF({nights_m}<=0,"",ROUND({accom_m}/{nights_m},2))',
          "money", note="Accommodation only - cleaning fees excluded")
    s.kpi(row, 8, "RevPAR",
          f'=IF({avail_m}<=0,"",ROUND({accom_m}/{avail_m},2))',
          "money")
    row = s.kpi(row, 10, "Average stay",
                f'=IF(COUNTIFS({bk_month},{_month_key(0)},{bk_earns})'
                f'<=0,"",ROUND({nights_m}/COUNTIFS({bk_month},'
                f'{_month_key(0)},{bk_earns}),1))', "number1")

    s.kpi(row, 0, "Gross booked", f"=ROUND({gross_m},2)", "money0")
    s.kpi(row, 2, "Platform fees", f"=ROUND({fees_m},2)", "money0",
          note="What the platforms kept")
    s.kpi(row, 4, "Net revenue", f"=ROUND({net_m},2)", "money0")
    s.kpi(row, 6, "Running costs", f"=ROUND({costs_m},2)", "money0",
          note="Cleaning, repairs and everything on the expense tab")
    s.kpi(row, 8, "Profit", f"=ROUND({profit_m},2)", "money0")
    row = s.kpi(row, 10, "Margin",
                f'=IF({net_m}<=0,"",{profit_m}/{net_m})', "percent")
    row += 1

    # ---------------------------------------------------- the money ----
    row = s.section(row, "Money still moving", LAST_COL)
    s.kpi(row, 0, "Unpaid bookings",
          '=ROUND(SUMIFS([[Booking Tracker.gross]],'
          '[[Booking Tracker.payment_status]],"Unpaid",'
          f'{bk_earns}),2)', "money0")
    s.kpi(row, 2, "Payouts owed",
          '=ROUND(SUMIFS([[Platform Revenue & Fees.exp_net]],'
          '[[Platform Revenue & Fees.is_paid]],"No"),2)', "money0",
          note="Stays that have ended but have not landed in the bank")
    s.kpi(row, 4, "Payouts overdue",
          '=COUNTIFS([[Platform Revenue & Fees.alert]],"Chase*")',
          "int")
    s.kpi(row, 6, "Owed to cleaners",
          '=ROUND(SUMIFS([[Cleaning Schedule.clean_cost]],'
          '[[Cleaning Schedule.paid]],"<>Yes",'
          '[[Cleaning Schedule.status]],"Completed"),2)', "money0")
    s.kpi(row, 8, "Fee variance",
          '=ROUND(SUM([[Platform Revenue & Fees.variance]]),2)',
          "money0",
          note="Real payouts against the fees you assumed. A big "
               "number means a rate on Settings is wrong")
    row = s.kpi(row, 10, "Security deposits held",
                '=ROUND(SUMIFS([[Property Database.deposit]],'
                '[[Property Database.guest_in]],"Yes"),2)', "money0")
    row += 1

    # --------------------------------------------- what needs doing ----
    row = s.section(row, "What needs doing", LAST_COL)
    row = s.note(
        row,
        "Every Action column in the workbook, counted. These are the "
        "only numbers on this page you should want to be zero.",
        LAST_COL, height=24)
    row = s.table_head(row, 0, [
        "Where", "Needs doing", "Of which urgent", "The usual cause"],
        [26, 13, 15, 56])
    jobs = [
        ("Property Database", '[[Property Database.alert]]',
         "Missing", "A property with no rate, no cleaning fee or an "
                    "open job on it"),
        ("Booking Tracker", '[[Booking Tracker.alert]]',
         "Duplicate", "Overlapping dates, a missing rate, or an "
                      "unconfirmed arrival"),
        ("Check-in & Check-out", '[[Check-in & Check-out.alert]]',
         "Send", "Access codes and directions not sent yet"),
        ("Cleaning Schedule", '[[Cleaning Schedule.alert]]',
         "Not cleaned", "A turnover due today or overdue"),
        ("Platform Revenue & Fees", '[[Platform Revenue & Fees.alert]]',
         "Chase", "A payout later than your platform's usual terms"),
        ("Supplies Inventory", '[[Supplies Inventory.alert]]',
         "Out of stock", "Below the minimum, or not enough to cover "
                         "the turnovers booked"),
        ("Maintenance Log", '[[Maintenance Log.alert]]',
         "Urgent", "An open job, or one blocking a letting"),
        ("Reviews Tracker", '[[Reviews Tracker.alert]]',
         "Action", "A poor review with no reply"),
        ("Staff & Vendors", '[[Staff & Vendors.alert]]',
         "Expired", "Lapsed insurance or an unpaid cleaner"),
    ]
    for label, column, urgent, cause in jobs:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.formula(
            row, 1,
            f'=COUNTIF({column},"?*")-COUNTIF({column},"OK*")',
            fmts.calc("int"))
        s.formula(row, 2, f'=COUNTIF({column},"{urgent}*")',
                         fmts.calc("int"))
        s.merge(row, 3, LAST_COL, cause, fmts.chrome("note"))
        row += 1
    row += 1

    # -------------------------------------------- property comparison --
    row = s.section(row, "Every property side by side", LAST_COL)
    row = s.note(
        row,
        "The whole portfolio on one line each, for the life of the "
        "file. Sort your attention by the last column: a property can "
        "be busy, well reviewed and still be the one losing money.",
        LAST_COL, height=26)
    row = s.table_head(row, 0, [
        "Property", "Nights", "Occupancy", "ADR", "RevPAR", "Net revenue",
        "Cleaning", "Repairs", "Other", "Profit", "Margin",
        "Profit a night", "Rating"],
        [24, 9, 10, 10, 10, 12, 10, 9, 9, 11, 9, 12, 8])
    for n in range(1, PROPERTY_ROWS + 1):
        s.write(row, 0, None, fmts.calc("text"))
        s.formula(row, 0, "=" + _nth_property("name", n),
                         fmts.calc("text"))
        for offset, (field, style) in enumerate((
                ("nights_booked", "int"), ("occupancy", "percent"),
                ("adr", "money"), ("revpar", "money"),
                ("net_revenue", "money0"), ("cleaning_spend", "money0"),
                ("maintenance_spend", "money0"),
                ("other_expenses", "money0"), ("profit", "money0"),
                ("margin", "percent"), ("profit_night", "money"),
                ("rating", "number"))):
            s.formula(row, offset + 1,
                             "=" + _nth_property(field, n),
                             fmts.calc(style))
        row += 1
    s.write(row, 0, "Portfolio", fmts.chrome("total_label"))
    total_nights = 'SUM([[Property Database.nights_booked]])'
    total_avail = 'SUM([[Property Database.nights_available]])'
    total_accom = 'SUM([[Property Database.accommodation]])'
    total_net = 'SUM([[Property Database.net_revenue]])'
    for offset, formula in enumerate((
            f"={total_nights}",
            f'=IF({total_avail}<=0,"",{total_nights}/{total_avail})',
            f'=IF({total_nights}<=0,"",ROUND({total_accom}/'
            f'{total_nights},2))',
            f'=IF({total_avail}<=0,"",ROUND({total_accom}/'
            f'{total_avail},2))',
            f"=ROUND({total_net},2)",
            '=ROUND(SUM([[Property Database.cleaning_spend]]),2)',
            '=ROUND(SUM([[Property Database.maintenance_spend]]),2)',
            '=ROUND(SUM([[Property Database.other_expenses]]),2)',
            '=ROUND(SUM([[Property Database.profit]]),2)',
            f'=IF({total_net}<=0,"",SUM([[Property Database.profit]])/'
            f'{total_net})',
            f'=IF({total_nights}<=0,"",ROUND(SUM('
            f'[[Property Database.profit]])/{total_nights},2))',
            '=IF(COUNTIFS([[Reviews Tracker.has_score]],"Yes")<=0,"",'
            'ROUND(SUMIFS([[Reviews Tracker.rate_overall]],'
            '[[Reviews Tracker.has_score]],"Yes")/'
            'COUNTIFS([[Reviews Tracker.has_score]],"Yes"),2))')):
        style = ("int", "percent", "money", "money", "money0", "money0",
                 "money0", "money0", "money0", "percent", "money",
                 "number")[offset]
        s.formula(row, offset + 1, formula,
                         fmts.chrome("total_value",
                                     num_format=s.num_format(style)))
    row += 2

    # ---------------------------------------------- twelve months ------
    row = s.section(row, "The last twelve months", LAST_COL)
    row = s.table_head(row, 0, [
        "Month", "Bookings", "Nights", "Occupancy", "ADR", "Net revenue",
        "Costs", "Profit", "Margin", "Against target"],
        [12, 10, 9, 10, 10, 12, 11, 11, 9, 13])
    first_month_row = row
    for back in range(MONTHS - 1, -1, -1):
        key = _month_key(back)
        nights = _month_sum("[[Booking Tracker.nights]]", bk_month,
                            back, bk_earns)
        accom = _month_sum("[[Booking Tracker.accommodation]]",
                           bk_month, back, bk_earns)
        net = _month_sum("[[Booking Tracker.net_revenue]]", bk_month,
                         back, bk_earns)
        costs = ("(" + _month_sum("[[Cleaning Schedule.total_cost]]",
                                  "[[Cleaning Schedule.month]]", back)
                 + "+" + _month_sum("[[Maintenance Log.cost_used]]",
                                    "[[Maintenance Log.month]]", back)
                 + "+" + _month_sum("[[Expense Tracker.pl_amount]]",
                                    "[[Expense Tracker.month]]", back)
                 + ")")
        avail = ("MAX(1,PropertyCount*DAY(DATE(YEAR(TodayDate),"
                 f"MONTH(TodayDate)-{back}+1,1)-1))")
        s.formula(
            row, 0,
            f'=TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{back},1),'
            f'"mmm yy")', fmts.calc("text"))
        for offset, (formula, style) in enumerate((
                (f'=COUNTIFS({bk_month},{key},{bk_earns})', "int"),
                (f"={nights}", "int"),
                (f'=IF({avail}<=0,"",{nights}/{avail})', "percent"),
                (f'=IF({nights}<=0,"",ROUND({accom}/{nights},2))',
                 "money"),
                (f"=ROUND({net},2)", "money0"),
                (f"=ROUND({costs},2)", "money0"),
                (f"=ROUND({net}-{costs},2)", "money0"),
                (f'=IF({net}<=0,"",({net}-{costs})/{net})', "percent"),
                (f'=IF(TargetRevenueMonth<=0,"",{net}/'
                 f'TargetRevenueMonth)', "percent"))):
            s.formula(row, offset + 1, formula, fmts.calc(style))
        row += 1
    last_month_row = row - 1

    ws.conditional_format(
        first_month_row, 3, last_month_row, 3,
        {"type": "cell", "criteria": ">=", "value": "TargetOccupancy",
         "format": fmts.cf("good")})
    ws.conditional_format(
        first_month_row, 7, last_month_row, 7,
        {"type": "cell", "criteria": "<", "value": 0,
         "format": fmts.cf("bad")})
    ws.conditional_format(
        first_month_row, 9, last_month_row, 9,
        {"type": "cell", "criteria": ">=", "value": 1,
         "format": fmts.cf("good")})
    ws.conditional_format(
        first_month_row, 9, last_month_row, 9,
        {"type": "cell", "criteria": "<", "value": 0.8,
         "format": fmts.cf("warn")})
    ws.add_sparkline(last_month_row + 1, 5, {
        "range": (f"'{NAME}'!F{first_month_row + 1}:"
                  f"F{last_month_row + 1}"),
        "type": "column", "style": 12})
    s.write(last_month_row + 1, 0, "Net revenue trend",
            fmts.chrome("label_bold"))
    ws.add_sparkline(last_month_row + 1, 7, {
        "range": (f"'{NAME}'!H{first_month_row + 1}:"
                  f"H{last_month_row + 1}"),
        "type": "column", "style": 12,
        "negative_points": True})
    row = last_month_row + 3

    # ----------------------------------------------- the short lists ---
    row = s.section(row, "Worth knowing", LAST_COL)

    s.write(row, 0, "Best performing properties",
            fmts.chrome("subsection"))
    s.write(row, 4, "Guests worth looking after",
            fmts.chrome("subsection"))
    s.write(row, 8, "Arriving next", fmts.chrome("subsection"))
    row += 1
    s.table_head(row, 0, ["Property", "Profit", "Occupancy"],
                 [22, 11, 10])
    s.table_head(row, 4, ["Guest", "Stays", "Net"], [20, 8, 11])
    s.table_head(row, 8, ["Guest", "Property", "Arrives", "Nights"],
                 [18, 18, 11, 8])
    row += 1
    for n in range(1, LIST_ROWS + 1):
        s.formula(
            row, 0,
            "=" + _by_rank("Property Database", "profit_rank", "name", n),
            fmts.calc("text"))
        s.formula(
            row, 1,
            "=" + _by_rank("Property Database", "profit_rank", "profit",
                           n), fmts.calc("money0"))
        s.formula(
            row, 2,
            "=" + _by_rank("Property Database", "profit_rank",
                           "occupancy", n), fmts.calc("percent"))
        s.formula(
            row, 4,
            "=" + _by_rank("Guest Database", "value_rank", "name", n),
            fmts.calc("text"))
        s.formula(
            row, 5,
            "=" + _by_rank("Guest Database", "value_rank", "stays", n),
            fmts.calc("int"))
        s.formula(
            row, 6,
            "=" + _by_rank("Guest Database", "value_rank", "net", n),
            fmts.calc("money0"))
        for offset, (field, style) in enumerate((
                ("guest_name", "text"), ("property_name", "text"),
                ("check_in", "date"), ("nights", "int"))):
            s.formula(
                row, 8 + offset,
                f'=IFERROR(INDEX([[Booking Tracker.{field}]],MATCH('
                f'{n},[[Booking Tracker.arrival_rank]],0)),"")',
                fmts.calc(style))
        row += 1
    row += 1

    row = s.note(
        row,
        "Nothing on this page is typed. Every figure traces back to a "
        "register through the property ID and the booking ID, which "
        "is why adding a fifth property or a third platform changes "
        "the dashboard without anyone editing a formula.",
        LAST_COL, height=28)

    s.finish(freeze_row=4, gridlines=False, fit_width=True,
             print_area=f"A1:{col_letter(LAST_COL)}{row + 2}")
    s.protect()
    return s
````

### `airbnb_rental_spreadsheet/sheets/getting_started.py`

*286 lines · 1 function*

> `Getting Started` - twenty minutes from download to running.
>
> Nobody reads a manual for a spreadsheet, so this one is short, in
> order, and tells you the two or three things that are genuinely not
> obvious: that every tab joins on the property ID and the booking ID,
> that the cleaning schedule is driven by your check-outs, and that the
> platform fee table on Settings is the one thing you must change before
> the numbers mean anything.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Getting Started'), `LAST_COL` (9), `STEPS` (11 items), `CONCEPTS` (6 items).

````python
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
````

### `airbnb_rental_spreadsheet/sheets/settings.py`

*513 lines · 2 functions*

> `Settings & Lists` - every number and every list the workbook runs on.
>
> Nothing here is hard coded to one country, one platform or one pricing
> model.  The commission each platform keeps, the card processing fee on
> top, the extra guest charge, the lodging tax and who remits it, the
> cost of a turnover, the targets, the alert windows and all forty-odd
> drop-downs are set here, once, and every other tab follows.
>
> The platform fee table matters most.  Airbnb runs a split fee in some
> markets and a host-only fee in others; Vrbo has two completely
> different models; Booking.com commission is negotiated per property.
> A spreadsheet that hard codes "3%" is wrong for most of the people who
> buy it, so this one ships with sensible defaults in cells you can type
> over.

| Defines | What it does |
|---------|--------------|
| `_month_keys(today, back, ahead)` | `yyyy-mm` keys around today, oldest first |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Settings & Lists'), `LAST_COL` (11), `LISTS` (30 items).

````python
"""`Settings & Lists` - every number and every list the workbook runs on.

Nothing here is hard coded to one country, one platform or one pricing
model.  The commission each platform keeps, the card processing fee on
top, the extra guest charge, the lodging tax and who remits it, the
cost of a turnover, the targets, the alert windows and all forty-odd
drop-downs are set here, once, and every other tab follows.

The platform fee table matters most.  Airbnb runs a split fee in some
markets and a host-only fee in others; Vrbo has two completely
different models; Booking.com commission is negotiated per property.
A spreadsheet that hard codes "3%" is wrong for most of the people who
buy it, so this one ships with sensible defaults in cells you can type
over.
"""

from __future__ import annotations

import datetime as _dt
from typing import List, Tuple

from ..config import (AVAILABILITY, BOOKING_STATUSES, CLEANING_STATUSES,
                      CONTRACT_STATUSES, COUNTRIES, DISCLAIMER,
                      EXPENSE_CATEGORIES, FREQUENCY_PER_YEAR,
                      GUEST_TYPES, ID_PREFIX_FIELDS,
                      MAINTENANCE_CATEGORIES, MAINTENANCE_STATUSES,
                      MONTH_NAMES, OWNERSHIP_TYPES, PAY_BASIS,
                      PAYMENT_METHODS, PAYMENT_STATUSES, PAYOUT_STATUSES,
                      PRIORITY_LEVELS, PRIVACY_NOTE, PROPERTY_STATUSES,
                      PROPERTY_TYPES, RECURRING_FREQUENCIES,
                      REVIEW_STATUSES, SEASONS, STAFF_ROLES, STAY_BANDS,
                      STOCK_STATUSES, SUPPLY_CATEGORIES, TAX_CATEGORIES,
                      TRIP_PURPOSES, YES_NO, YES_NO_NA, Config, Palette)
from ..layout import Registry, col_letter, dynamic_name, quote_sheet
from ..writer import SheetWriter, list_from, number_check

NAME = "Settings & Lists"
LAST_COL = 11

#: name -> values.  Everything the workbook offers you in a drop-down.
LISTS: List[Tuple[str, List[str]]] = [
    ("YesNoList", list(YES_NO)),
    ("YesNoNaList", list(YES_NO_NA)),
    ("BookingStatusList", list(BOOKING_STATUSES)),
    ("PaymentStatusList", list(PAYMENT_STATUSES)),
    ("PropertyTypeList", list(PROPERTY_TYPES)),
    ("PropertyStatusList", list(PROPERTY_STATUSES)),
    ("OwnershipList", list(OWNERSHIP_TYPES)),
    ("CleaningStatusList", list(CLEANING_STATUSES)),
    ("MaintenanceCategoryList", list(MAINTENANCE_CATEGORIES)),
    ("MaintenanceStatusList", list(MAINTENANCE_STATUSES)),
    ("PriorityList", list(PRIORITY_LEVELS)),
    ("ExpenseCategoryList", list(EXPENSE_CATEGORIES)),
    ("TaxCategoryList", list(TAX_CATEGORIES)),
    ("PaymentMethodList", list(PAYMENT_METHODS)),
    ("FrequencyList", list(RECURRING_FREQUENCIES)),
    ("SupplyCategoryList", list(SUPPLY_CATEGORIES)),
    ("StockStatusList", list(STOCK_STATUSES)),
    ("StaffRoleList", list(STAFF_ROLES)),
    ("PayBasisList", list(PAY_BASIS)),
    ("ContractStatusList", list(CONTRACT_STATUSES)),
    ("AvailabilityList", list(AVAILABILITY)),
    ("ReviewStatusList", list(REVIEW_STATUSES)),
    ("GuestTypeList", list(GUEST_TYPES)),
    ("TripPurposeList", list(TRIP_PURPOSES)),
    ("PayoutStatusList", list(PAYOUT_STATUSES)),
    ("CountryList", list(COUNTRIES)),
    ("MonthNameList", list(MONTH_NAMES)),
    ("SeasonList", [s[0] for s in SEASONS]),
    ("StayBandList", [b[0] for b in STAY_BANDS]),
    ("RatingList", ["5", "4.5", "4", "3.5", "3", "2", "1"]),
]


def _month_keys(today: _dt.date, back: int = 15, ahead: int = 9
                ) -> List[str]:
    """`yyyy-mm` keys around today, oldest first."""
    keys: List[str] = []
    year, month = today.year, today.month
    start = month - back
    while start <= 0:
        start += 12
        year -= 1
    for step in range(back + ahead + 1):
        m = start + step
        y = year + (m - 1) // 12
        m = (m - 1) % 12 + 1
        keys.append(f"{y:04d}-{m:02d}")
    return keys


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.ink)
    sheet_ref = quote_sheet(NAME)

    s.banner("Settings & Lists",
             "Your properties, your platforms, your fees, your lists. "
             "Set these once and every tab in the workbook follows.",
             LAST_COL)
    s.set_col(0, 0, 3)
    s.set_col(1, 2, 24)
    s.set_col(3, 3, 20)
    s.set_col(4, LAST_COL, 17)

    row = 4
    row = s.note(
        row,
        "White cells are yours to change; everything else is locked and "
        "calculated. Nothing here assumes a country, a platform or a "
        "pricing model - set your own commission rates, your own tax "
        "position, your own currency and your own targets, and the "
        "rest of the workbook follows.",
        LAST_COL, height=34)
    row += 1

    # ------------------------------------------------------ the business --
    row = s.section(row, "Your business", LAST_COL)
    row = s.field(row, "Business name", cfg.host.name, name="HostName",
                  value_end=5, help_col=7,
                  help_text="Printed at the top of every report and "
                            "every printable checklist.")
    row = s.field(row, "Owner / host", cfg.host.owner, name="OwnerName",
                  value_end=5, help_col=7)
    row = s.field(row, "Address", cfg.host.address, value_end=5,
                  help_col=7)
    row = s.field(row, "Phone", cfg.host.phone, value_end=5, help_col=7)
    row = s.field(row, "Email", cfg.host.email, value_end=5, help_col=7)
    row = s.field(row, "Website", cfg.host.website, value_end=5,
                  help_col=7)
    row = s.field(row, "Let registration / permit", cfg.host.licence,
                  value_end=5, help_col=7,
                  help_text="Your short-term let registration number, "
                            "if your city requires one. Per-property "
                            "licences live on the Property Database.")
    row = s.field(row, "Tax / business number", cfg.host.tax_id,
                  value_end=5, help_col=7)
    row = s.field(row, "Strapline", cfg.host.tagline, value_end=5,
                  help_col=7)
    row = s.field(row, "Currency symbol", cfg.currency_symbol,
                  name="CurrencySign",
                  help_text="Shown on screen. The cell formats follow "
                            "the symbol you build the file with.")
    row = s.field(row, "Currency code", cfg.currency_code)
    row = s.field(row, "Tracking started",
                  cfg.today - _dt.timedelta(days=120), style="date",
                  name="TrackingStart",
                  help_text="The day you began logging in this file.")
    row = s.field(row, "Today", None, style="date", name="TodayDate",
                  mode="calc", formula="=TODAY()",
                  help_text="Live. Everything that says 'overdue', "
                            "'arriving' or 'this month' reads this "
                            "cell.")
    row = s.field(row, "This month", None, style="text",
                  name="ThisMonthKey", mode="calc",
                  formula='=TEXT(TODAY(),"yyyy-mm")',
                  help_text="The month key every monthly total groups "
                            "on.")
    row = s.field(row, "Properties on the books", None, style="int",
                  name="PropertyCount", mode="calc",
                  formula='=COUNTIF([[Property Database.name]],"?*")')
    row = s.field(row, "Properties active", None, style="int",
                  name="ActiveProperties", mode="calc",
                  formula='=COUNTIFS([[Property Database.status]],'
                          '"Active")')
    row += 1

    # ------------------------------------------------- the platform fees --
    row = s.section(row, "Platform fees - the most important table here",
                    LAST_COL)
    row = s.note(
        row,
        "Every booking prices its commission and card fee from this "
        "table. The defaults below are the common cases, not the law: "
        "Airbnb runs a split fee in most markets and a host-only fee "
        "in others, Vrbo has two different models, and Booking.com "
        "commission is negotiated per property. Put your own numbers "
        "in - the Platform Revenue & Fees tab will tell you when they "
        "are wrong, because it compares what you assumed with what the "
        "payout actually showed.",
        LAST_COL, height=50)
    row += 1
    plat_head = row
    s.table_head(plat_head, 1, [
        "Platform", "Commission", "Card fee", "Is a platform?",
        "Notes"],
        widths=[22, 14, 12, 15, 46])
    first_plat = plat_head + 1
    table = cfg.platform_table
    plat_rows = len(table) + 2
    for index in range(plat_rows):
        line = first_plat + index
        if index < len(table):
            label, commission, payment, counts, note = table[index]
            s.write(line, 1, label, fmts.input("text"))
            s.write(line, 2, commission, fmts.input("percent"))
            s.write(line, 3, payment, fmts.input("percent"))
            s.write(line, 4, counts, fmts.input("flag"))
            s.write(line, 5, note, fmts.chrome("note"))
        else:
            s.write(line, 1, None, fmts.input("text"))
            s.write(line, 2, None, fmts.input("percent"))
            s.write(line, 3, None, fmts.input("percent"))
            s.write(line, 4, None, fmts.input("flag"))
            s.write(line, 5, None, fmts.chrome("note"))
    last_plat = first_plat + plat_rows
    for key, column in (("PlatformList", 1), ("PlatformCommission", 2),
                        ("PlatformPayment", 3), ("PlatformIsOTA", 4)):
        book.define_name(
            key,
            f"={sheet_ref}!${col_letter(column)}${first_plat + 1}:"
            f"${col_letter(column)}${last_plat}")
    row = last_plat + 1

    # --------------------------------------------------- pricing defaults --
    row = s.section(row, "Pricing defaults", LAST_COL)
    row = s.note(
        row,
        "Each property has its own nightly rate and cleaning fee on the "
        "Property Database. These are the portfolio-wide rules that "
        "sit behind them, and the charges a booking works out for "
        "itself when you leave the cell blank.",
        LAST_COL, height=32)
    row += 1
    row = s.field(row, "Standard nightly rate", cfg.default_nightly_rate,
                  style="money", name="DefaultNightly",
                  validate=number_check(0, 100000),
                  help_text="Only a starting point for a new property. "
                            "The booking reads the property's rate, "
                            "not this one.")
    row = s.field(row, "Standard cleaning fee", cfg.default_cleaning_fee,
                  style="money", name="DefaultCleaningFee",
                  validate=number_check(0, 100000))
    row = s.field(row, "Standard cleaning cost",
                  cfg.default_cleaning_cost, style="money",
                  name="DefaultCleaningCost",
                  validate=number_check(0, 100000),
                  help_text="What you pay a cleaner, as opposed to "
                            "what you charge the guest.")
    row = s.field(row, "Guests included in the rate",
                  cfg.extra_guest_after, style="int",
                  name="ExtraGuestAfter",
                  validate=number_check(0, 40, decimal=False),
                  help_text="Guests above this number are charged the "
                            "extra guest fee, per night.")
    row = s.field(row, "Extra guest fee a night", cfg.extra_guest_fee,
                  style="money", name="ExtraGuestFee",
                  validate=number_check(0, 10000))
    row = s.field(row, "Pet fee a stay", cfg.pet_fee, style="money",
                  name="PetFee", validate=number_check(0, 10000))
    row = s.field(row, "Security deposit",
                  cfg.default_security_deposit, style="money",
                  name="DefaultDeposit", validate=number_check(0, 100000))
    row = s.field(row, "Minimum nights", cfg.default_min_nights,
                  style="int", name="DefaultMinNights",
                  validate=number_check(1, 365, decimal=False))
    row = s.field(row, "Management fee", cfg.management_fee_percent,
                  style="percent", name="ManagementFee",
                  validate=number_check(0, 1),
                  help_text="For co-hosts: the share of revenue you "
                            "keep, or that an agency keeps from you. "
                            "Zero if it does not apply.")
    row += 1

    # ------------------------------------------------------------- tax ----
    row = s.section(row, "Lodging tax", LAST_COL)
    row = s.note(
        row,
        "Occupancy, lodging, tourist or sales tax on the accommodation "
        "total. Most platforms collect and remit it for you, which is "
        "why the default says so and the rate is zero. If you collect "
        "it yourself, put your rate in and switch the second cell to "
        "No - every booking will then show the tax it owes and keep it "
        "out of your net revenue.",
        LAST_COL, height=42)
    row += 1
    row = s.field(row, "Tax name", cfg.tax_label, name="TaxLabel")
    row = s.field(row, "Tax rate", cfg.tax_rate, style="percent",
                  name="TaxRate", validate=number_check(0, 1))
    row = s.field(row, "Platform collects and remits it",
                  cfg.tax_collected_by_platform, style="flag",
                  name="TaxCollectedByPlatform",
                  validate=list_from("YesNoList"))
    row += 1

    # ---------------------------------------------------- turnover costs --
    row = s.section(row, "What a turnover really costs", LAST_COL)
    row = s.note(
        row,
        "A changeover is never just the cleaner's fee. These two "
        "numbers are added to every cleaning job unless you type "
        "something different on the job itself, and they are what "
        "makes the true cost per booking honest.",
        LAST_COL, height=32)
    row += 1
    row = s.field(row, "Laundry a turnover", cfg.laundry_cost_turnover,
                  style="money", name="LaundryPerTurn",
                  validate=number_check(0, 10000))
    row = s.field(row, "Consumables a turnover",
                  cfg.supplies_cost_turnover, style="money",
                  name="SuppliesPerTurn", validate=number_check(0, 10000),
                  help_text="Loo roll, coffee, soap, bin bags. Compare "
                            "it with the figure the Supplies tab works "
                            "out from your actual stock usage.")
    row = s.field(row, "Welcome basket", cfg.welcome_cost_turnover,
                  style="money", name="WelcomePerTurn",
                  validate=number_check(0, 10000))
    row = s.field(row, "Default cleaner rate", cfg.cleaner_default_rate,
                  style="money", name="CleanerRate",
                  validate=number_check(0, 10000))
    row = s.field(row, "A tight turnaround is under",
                  cfg.turnover_warning_hours, style="number1",
                  name="TightTurnHours", validate=number_check(0, 48),
                  help_text="Hours between check-out and the next "
                            "check-in on a same-day turn. Below this "
                            "the cleaning job goes red.")
    row += 1

    # -------------------------------------------------- double counting --
    row = s.section(row, "Counting costs once", LAST_COL)
    row = s.note(
        row,
        "Cleaning has its own tab and so does maintenance, because "
        "both need a workflow rather than a line in a ledger. If you "
        "would rather type them as ordinary expenses instead, switch "
        "the relevant cell to Yes and the expense rows will start "
        "counting towards profit - but then stop using the other tab "
        "for money, or you will count the same cost twice.",
        LAST_COL, height=46)
    row += 1
    row = s.field(row, "Cleaning expenses count in profit", "No",
                  style="flag", name="CleaningInExpenses",
                  validate=list_from("YesNoList"),
                  help_text="No = the Cleaning Schedule provides the "
                            "cleaning cost, and cleaning rows on the "
                            "expense tab are ignored.")
    row = s.field(row, "Maintenance expenses count in profit", "No",
                  style="flag", name="MaintenanceInExpenses",
                  validate=list_from("YesNoList"),
                  help_text="No = the Maintenance Log provides the "
                            "repair costs.")
    row += 1

    # --------------------------------------------------------- targets --
    row = s.section(row, "Targets", LAST_COL)
    row = s.field(row, "Occupancy target", cfg.occupancy_target,
                  style="percent", name="TargetOccupancy",
                  validate=number_check(0, 1),
                  help_text="Every occupancy figure in the workbook is "
                            "green at or above this and amber below.")
    row = s.field(row, "ADR target", cfg.adr_target, style="money",
                  name="TargetADR", validate=number_check(0, 100000))
    row = s.field(row, "RevPAR target", cfg.revpar_target, style="money",
                  name="TargetRevPAR", validate=number_check(0, 100000))
    row = s.field(row, "Margin target", cfg.margin_target,
                  style="percent", name="TargetMargin",
                  validate=number_check(0, 1))
    row = s.field(row, "Rating target", cfg.rating_target,
                  style="number", name="TargetRating",
                  validate=number_check(0, 5))
    row = s.field(row, "Revenue target a month",
                  cfg.revenue_target_month, style="money0",
                  name="TargetRevenueMonth",
                  validate=number_check(0, 100000000))
    row = s.field(row, "Revenue target a year", cfg.revenue_target_year,
                  style="money0", name="TargetRevenueYear",
                  validate=number_check(0, 100000000))
    row = s.field(row, "Profit target a month", cfg.profit_target_month,
                  style="money0", name="TargetProfitMonth",
                  validate=number_check(0, 100000000))
    row += 1

    # ---------------------------------------------------------- alerts --
    row = s.section(row, "When to start nagging", LAST_COL)
    row = s.field(row, "'Arriving soon' window (days)",
                  cfg.upcoming_window_days, style="int",
                  name="UpcomingDays",
                  validate=number_check(0, 90, decimal=False),
                  help_text="A booking inside this window counts as "
                            "arriving soon, and its pre-arrival "
                            "messages start being chased.")
    row = s.field(row, "Chase a review after (days)",
                  cfg.review_chase_days, style="int",
                  name="ReviewChaseDays",
                  validate=number_check(0, 90, decimal=False))
    row = s.field(row, "Chase a payout after (days)",
                  cfg.payout_chase_days, style="int",
                  name="PayoutChaseDays",
                  validate=number_check(0, 180, decimal=False))
    row = s.field(row, "Warn about maintenance (days)",
                  cfg.maintenance_warning_days, style="int",
                  name="MaintenanceWarnDays",
                  validate=number_check(0, 180, decimal=False))
    row = s.field(row, "Reorder when cover drops below (weeks)",
                  cfg.restock_warning_weeks, style="number1",
                  name="RestockWeeks", validate=number_check(0, 52))
    row = s.field(row, "A long stay is (nights)", cfg.long_stay_nights,
                  style="int", name="LongStayNights",
                  validate=number_check(1, 365, decimal=False))
    row += 1

    # ---------------------------------------------- recurring frequency --
    row = s.section(row, "How often a recurring cost repeats", LAST_COL)
    row = s.note(
        row,
        "This is what turns any expense into a monthly equivalent, so "
        "a yearly insurance premium and a weekly laundry bill can sit "
        "in the same breakdown. 'One off' and 'Per booking' are zero "
        "because they are not a standing monthly cost.",
        LAST_COL, height=32)
    row += 1
    freq_head = row
    s.table_head(freq_head, 1, ["How often", "Times a year"],
                 widths=[22, 16])
    first_freq = freq_head + 1
    for index, label in enumerate(RECURRING_FREQUENCIES):
        line = first_freq + index
        s.write(line, 1, label, fmts.input("text"))
        s.write(line, 2, FREQUENCY_PER_YEAR.get(label, 0.0),
                fmts.input("number1"))
    last_freq = first_freq + len(RECURRING_FREQUENCIES) - 1
    book.define_name(
        "FrequencyPerYear",
        f"={sheet_ref}!$C${first_freq + 1}:$C${last_freq + 1}")
    row = last_freq + 2

    # ------------------------------------------------------ ID prefixes --
    row = s.section(row, "ID prefixes", LAST_COL)
    row = s.note(
        row,
        "The letters in front of every automatic ID. Change them here "
        "and the IDs on that tab change with them - but do it before "
        "you have typed three hundred bookings, not after.",
        LAST_COL, height=28)
    row += 1
    half = (len(ID_PREFIX_FIELDS) + 1) // 2
    top = row
    for index, (label, default, name, _digits) in enumerate(
            ID_PREFIX_FIELDS):
        column = 0 if index < half else 4
        line = top + (index if index < half else index - half)
        s.write(line, column + 1, label, fmts.chrome("label_bold"))
        s.write(line, column + 2, default, fmts.input("text"))
        book.define_name(
            name, f"={sheet_ref}!${col_letter(column + 2)}${line + 1}")
    row = top + half + 1

    # ----------------------------------------------------------- lists --
    row = s.section(row, "Drop-down lists", LAST_COL)
    row = s.note(
        row,
        "Add new entries directly under the last one in a column - "
        "every list has room to grow and blank rows are ignored. Do "
        "not delete a whole column: a drop-down pointing at a missing "
        "list stops working.",
        LAST_COL, height=30)
    row += 1

    per_row = 4
    spare = 6
    top = row
    column = 0
    tallest = 0
    for index, (name, values) in enumerate(LISTS):
        if index and index % per_row == 0:
            top += tallest + 2
            column = 0
            tallest = 0
        s.write(top, column, name.replace("List", ""),
                fmts.chrome("header"))
        s.set_col(column, column, 26)
        for offset, value in enumerate(values):
            s.write(top + 1 + offset, column, value, fmts.input("text"))
        for offset in range(len(values), len(values) + spare):
            s.write(top + 1 + offset, column, None, fmts.input("text"))
        first = top + 2
        last = top + 1 + len(values) + spare
        book.define_name(
            name,
            f"={sheet_ref}!${col_letter(column)}${first}:"
            f"${col_letter(column)}${last}")
        tallest = max(tallest, len(values) + spare)
        column += 3
    row = top + tallest + 3

    # ------------------------------------------------------ month keys --
    row = s.section(row, "Month keys", LAST_COL)
    row = s.note(
        row,
        "The months the reports pick from. They are plain text in "
        "yyyy-mm form so that every monthly total in the workbook "
        "groups on exactly the same thing.",
        LAST_COL, height=28)
    row += 1
    keys = _month_keys(cfg.today)
    s.write(row, 0, "Month", fmts.chrome("header"))
    s.set_col(0, 0, 14)
    for offset, key in enumerate(keys):
        s.write(row + 1 + offset, 0, key, fmts.input("text"))
    book.define_name(
        "MonthKeyList",
        f"={sheet_ref}!$A${row + 2}:$A${row + 1 + len(keys)}")
    book.define_name(
        "MonthKeys",
        dynamic_name(NAME, "A", row + 2, row + 1 + len(keys)))
    row = row + len(keys) + 2

    row = s.note(row, PRIVACY_NOTE, LAST_COL, height=42)
    row += 1
    row = s.note(row, DISCLAIMER, LAST_COL, height=96)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    s.protect()
    return s
````

### `airbnb_rental_spreadsheet/sheets/monthly.py`

*446 lines · 3 functions*

> `Monthly Reports` - any month you choose, in full.
>
> Type a month key at the top - `2026-03` - and the whole page becomes
> that month: occupancy and rate, revenue by platform, costs by
> category, every property compared, the guests who stayed and how the
> month sits against the one before it and the same month a year ago.
>
> It is the page to send an owner you co-host for, and the page to read
> yourself on the first of the month.

| Defines | What it does |
|---------|--------------|
| `_window(key)` | Date bounds for a `yyyy-mm` key cell, as first and last day |
| `_period(prefix, key)` | Every total for one month key, as a dict of formula fragments |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Monthly Reports'), `LAST_COL` (9), `PROPERTY_ROWS` (8).

````python
"""`Monthly Reports` - any month you choose, in full.

Type a month key at the top - `2026-03` - and the whole page becomes
that month: occupancy and rate, revenue by platform, costs by
category, every property compared, the guests who stayed and how the
month sits against the one before it and the same month a year ago.

It is the page to send an owner you co-host for, and the page to read
yourself on the first of the month.
"""

from __future__ import annotations

from ..config import Config, EXPENSE_CATEGORIES, Palette
from ..layout import Registry, col_letter
from ..writer import SheetWriter, list_from

NAME = "Monthly Reports"
LAST_COL = 9
PROPERTY_ROWS = 8


def _window(key: str) -> str:
    """Date bounds for a `yyyy-mm` key cell, as first and last day."""
    return (f'DATE(VALUE(LEFT({key},4)),VALUE(RIGHT({key},2)),1)',
            f'DATE(VALUE(LEFT({key},4)),VALUE(RIGHT({key},2))+1,1)-1')


def _period(prefix: str, key: str) -> dict:
    """Every total for one month key, as a dict of formula fragments."""
    earns = '[[Booking Tracker.earns]],"Yes"'
    bk = f'[[Booking Tracker.month]],{key}'
    nights = f'SUMIFS([[Booking Tracker.nights]],{bk},{earns})'
    accom = f'SUMIFS([[Booking Tracker.accommodation]],{bk},{earns})'
    gross = f'SUMIFS([[Booking Tracker.gross]],{bk},{earns})'
    fees = f'SUMIFS([[Booking Tracker.platform_fee_total]],{bk},{earns})'
    net = f'SUMIFS([[Booking Tracker.net_revenue]],{bk},{earns})'
    clean = f'SUMIFS([[Cleaning Schedule.total_cost]],' \
            f'[[Cleaning Schedule.month]],{key})'
    maint = f'SUMIFS([[Maintenance Log.cost_used]],' \
            f'[[Maintenance Log.month]],{key})'
    other = f'SUMIFS([[Expense Tracker.pl_amount]],' \
            f'[[Expense Tracker.month]],{key})'
    costs = f'({clean}+{maint}+{other})'
    blocked = f'SUMIFS([[Booking Tracker.nights]],{bk},' \
              f'[[Booking Tracker.is_block]],"Yes")'
    start, end = _window(key)
    days = f'DAY({end})'
    avail = f'MAX(1,PropertyCount*{days}-{blocked})'
    return {
        "bookings": f'COUNTIFS({bk},{earns})',
        "cancellations": f'COUNTIFS({bk},'
                         f'[[Booking Tracker.is_cancelled]],"Yes")',
        "nights": nights, "accom": accom, "gross": gross,
        "fees": fees, "net": net, "clean": clean, "maint": maint,
        "other": other, "costs": costs, "blocked": blocked,
        "avail": avail, "start": start, "end": end,
        "profit": f'({net}-{costs})',
        "occupancy": f'IF({avail}<=0,"",{nights}/{avail})',
        "adr": f'IF({nights}<=0,"",ROUND({accom}/{nights},2))',
        "revpar": f'IF({avail}<=0,"",ROUND({accom}/{avail},2))',
        "alos": f'IF(COUNTIFS({bk},{earns})<=0,"",ROUND({nights}/'
                f'COUNTIFS({bk},{earns}),1))',
        "guests": f'SUMIFS([[Booking Tracker.guests_total]],{bk},'
                  f'{earns})',
    }


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_mid)
    ws = s.ws

    s.banner("Monthly Report",
             "Pick a month. Occupancy, rate, revenue, costs and every "
             "property compared.", LAST_COL)
    s.set_col(0, 0, 26)
    for column in range(1, LAST_COL + 1):
        s.set_col(column, column, 14)

    row = 4
    s.write(row, 0, "Month", fmts.chrome("label_bold"))
    s.write(row, 1, cfg.today.strftime("%Y-%m"),
            fmts.chrome("form_box", locked=False))
    ws.data_validation(row, 1, row, 1, list_from("MonthKeyList"))
    book.define_name("ReportMonth", f"='{NAME}'!$B${row + 1}")
    s.merge(row, 2, 3, None, fmts.calc("text"))
    s.formula(
        row, 2,
        '=IF(ReportMonth="","Pick a month",TEXT(DATE(VALUE(LEFT('
        'ReportMonth,4)),VALUE(RIGHT(ReportMonth,2)),1),"mmmm yyyy"))',
        fmts.calc("text"))
    s.merge(row, 4, LAST_COL,
            "In yyyy-mm form. The list comes from Settings.",
            fmts.chrome("note"))
    row += 1

    # hidden previous-month and same-month-last-year keys
    s.write(row, 20, None, fmts.calc("text"))
    s.formula(
        row, 20,
        '=TEXT(DATE(VALUE(LEFT(ReportMonth,4)),VALUE(RIGHT('
        'ReportMonth,2))-1,1),"yyyy-mm")', fmts.calc("text"))
    book.define_name("PrevMonthKey", f"='{NAME}'!$U${row + 1}")
    s.write(row + 1, 20, None, fmts.calc("text"))
    s.formula(
        row + 1, 20,
        '=TEXT(DATE(VALUE(LEFT(ReportMonth,4))-1,VALUE(RIGHT('
        'ReportMonth,2)),1),"yyyy-mm")', fmts.calc("text"))
    book.define_name("LastYearMonthKey", f"='{NAME}'!$U${row + 2}")
    ws.set_column(20, 20, None, None, {"hidden": True})
    row += 1

    now = _period("now", "ReportMonth")
    prev = _period("prev", "PrevMonthKey")
    year = _period("year", "LastYearMonthKey")

    # ------------------------------------------------------ headline ----
    row = s.section(row, "The month", LAST_COL)
    s.kpi(row, 0, "Bookings", f'={now["bookings"]}', "int")
    s.kpi(row, 2, "Nights sold", f'={now["nights"]}', "int")
    s.kpi(row, 4, "Occupancy", f'={now["occupancy"]}', "percent")
    s.kpi(row, 6, "ADR", f'={now["adr"]}', "money",
          note="Accommodation only")
    row = s.kpi(row, 8, "RevPAR", f'={now["revpar"]}', "money")

    s.kpi(row, 0, "Gross", f'=ROUND({now["gross"]},2)', "money0")
    s.kpi(row, 2, "Platform fees", f'=ROUND({now["fees"]},2)', "money0")
    s.kpi(row, 4, "Net revenue", f'=ROUND({now["net"]},2)', "money0")
    s.kpi(row, 6, "Costs", f'=ROUND({now["costs"]},2)', "money0")
    row = s.kpi(row, 8, "Profit", f'=ROUND({now["profit"]},2)',
                "money0")
    row += 1

    # ------------------------------------------------- the comparison ---
    row = s.section(row, "Against last month and last year", LAST_COL)
    row = s.table_head(row, 0, [
        "Measure", "This month", "Last month", "Change",
        "Same month last year", "Change", "Target", "Against target"],
        [26, 14, 14, 12, 18, 12, 13, 14])
    measures = [
        ("Bookings", "bookings", "int", ""),
        ("Nights sold", "nights", "int", ""),
        ("Occupancy", "occupancy", "percent", "TargetOccupancy"),
        ("ADR", "adr", "money", "TargetADR"),
        ("RevPAR", "revpar", "money", "TargetRevPAR"),
        ("Average stay", "alos", "number1", ""),
        ("Guests hosted", "guests", "int", ""),
        ("Gross", "gross", "money0", ""),
        ("Platform fees", "fees", "money0", ""),
        ("Net revenue", "net", "money0", "TargetRevenueMonth"),
        ("Cleaning", "clean", "money0", ""),
        ("Repairs", "maint", "money0", ""),
        ("Other costs", "other", "money0", ""),
        ("Total costs", "costs", "money0", ""),
        ("Profit", "profit", "money0", "TargetProfitMonth"),
    ]
    compare_top = row
    for label, key, style, target in measures:
        a, b, c = now[key], prev[key], year[key]
        pct = (style in ("percent",))
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.formula(row, 1, (f"={a}"),
                         fmts.calc(style))
        s.formula(row, 2, s.reg.render(f"={b}"),
                         fmts.calc(style))
        if pct:
            s.formula(
                row, 3,
                s.reg.render(f'=IF({b}="","",IF({a}="","",{a}-{b}))'),
                fmts.calc("percent"))
        else:
            s.formula(
                row, 3,
                s.reg.render(f'=IF({b}="","",IF({b}<=0,"",IF({a}="",'
                             f'"",{a}/{b}-1)))'),
                fmts.calc("percent"))
        s.formula(row, 4, s.reg.render(f"={c}"),
                         fmts.calc(style))
        if pct:
            s.formula(
                row, 5,
                s.reg.render(f'=IF({c}="","",IF({a}="","",{a}-{c}))'),
                fmts.calc("percent"))
        else:
            s.formula(
                row, 5,
                s.reg.render(f'=IF({c}="","",IF({c}<=0,"",IF({a}="",'
                             f'"",{a}/{c}-1)))'),
                fmts.calc("percent"))
        if target:
            s.formula(row, 6, f"={target}", fmts.calc(style))
            s.formula(
                row, 7,
                s.reg.render(f'=IF({target}<=0,"",IF({a}="","",{a}/'
                             f'{target}))'), fmts.calc("percent"))
        else:
            s.write(row, 6, None, fmts.calc("text"))
            s.write(row, 7, None, fmts.calc("text"))
        row += 1
    for column in (3, 5):
        ws.conditional_format(compare_top, column, row - 1, column, {
            "type": "cell", "criteria": "<", "value": 0,
            "format": fmts.cf("bad")})
        ws.conditional_format(compare_top, column, row - 1, column, {
            "type": "cell", "criteria": ">", "value": 0,
            "format": fmts.cf("good")})
    ws.conditional_format(compare_top, 7, row - 1, 7, {
        "type": "cell", "criteria": ">=", "value": 1,
        "format": fmts.cf("good")})
    ws.conditional_format(compare_top, 7, row - 1, 7, {
        "type": "cell", "criteria": "<", "value": 0.8,
        "format": fmts.cf("warn")})
    row += 1

    # ---------------------------------------------------- by platform ---
    row = s.section(row, "Where the bookings came from", LAST_COL)
    row = s.table_head(row, 0, [
        "Platform", "Bookings", "Nights", "Gross", "Commission",
        "Card fee", "Net", "Effective fee", "Share of net"],
        [26, 11, 10, 13, 13, 11, 13, 13, 12])
    earns = '[[Booking Tracker.earns]],"Yes"'
    bk = '[[Booking Tracker.month]],ReportMonth'
    for name, _c, _p, _counts, _note in cfg.platform_table:
        plat = f'[[Booking Tracker.platform]],"{name}"'
        p_gross = f'SUMIFS([[Booking Tracker.gross]],{bk},{earns},{plat})'
        p_comm = (f'SUMIFS([[Booking Tracker.platform_fee]],{bk},'
                  f'{earns},{plat})')
        p_card = (f'SUMIFS([[Booking Tracker.payment_fee]],{bk},'
                  f'{earns},{plat})')
        p_net = (f'SUMIFS([[Booking Tracker.net_revenue]],{bk},{earns},'
                 f'{plat})')
        s.calc_row(row, 0, [
            name,
            f"=COUNTIFS({bk},{earns},{plat})",
            f"=SUMIFS([[Booking Tracker.nights]],{bk},{earns},{plat})",
            f"=ROUND({p_gross},2)",
            f"=ROUND({p_comm},2)",
            f"=ROUND({p_card},2)",
            f"=ROUND({p_net},2)",
            f'=IF({p_gross}<=0,"",({p_comm}+{p_card})/{p_gross})',
            f'=IF({now["net"]}<=0,"",{p_net}/{now["net"]})',
        ], ["text", "int", "int", "money", "money", "money", "money",
            "percent", "percent"])
        row += 1
    s.write(row, 0, "All platforms", fmts.chrome("total_label"))
    for offset, (formula, style) in enumerate((
            (f'={now["bookings"]}', "int"),
            (f'={now["nights"]}', "int"),
            (f'=ROUND({now["gross"]},2)', "money"),
            ('=ROUND(SUMIFS([[Booking Tracker.platform_fee]],'
             f'{bk},{earns}),2)', "money"),
            ('=ROUND(SUMIFS([[Booking Tracker.payment_fee]],'
             f'{bk},{earns}),2)', "money"),
            (f'=ROUND({now["net"]},2)', "money"),
            (f'=IF({now["gross"]}<=0,"",{now["fees"]}/{now["gross"]})',
             "percent"),
            ("=1", "percent"))):
        s.formula(row, offset + 1, s.reg.render(formula),
                         fmts.chrome("total_value",
                                     num_format=s.num_format(style)))
    row += 2

    # ---------------------------------------------------- by property ---
    row = s.section(row, "Every property this month", LAST_COL)
    row = s.table_head(row, 0, [
        "Property", "Nights", "Occupancy", "ADR", "Net revenue",
        "Cleaning", "Other costs", "Profit", "Margin", "Rating"],
        [26, 10, 11, 11, 13, 11, 12, 12, 10, 9])
    start, end = _window("ReportMonth")
    for n in range(1, PROPERTY_ROWS + 1):
        pid = f'INDEX([[Property Database.id]],{n})'
        guard = f'INDEX([[Property Database.name]],{n})=""'
        prop = f'[[Booking Tracker.property]],{pid}'
        p_nights = (f'SUMIFS([[Booking Tracker.nights]],{bk},{earns},'
                    f'{prop})')
        p_accom = (f'SUMIFS([[Booking Tracker.accommodation]],{bk},'
                   f'{earns},{prop})')
        p_net = (f'SUMIFS([[Booking Tracker.net_revenue]],{bk},{earns},'
                 f'{prop})')
        p_clean = (f'SUMIFS([[Cleaning Schedule.total_cost]],'
                   f'[[Cleaning Schedule.month]],ReportMonth,'
                   f'[[Cleaning Schedule.property]],{pid})')
        p_other = (f'(SUMIFS([[Maintenance Log.cost_used]],'
                   f'[[Maintenance Log.month]],ReportMonth,'
                   f'[[Maintenance Log.property]],{pid})+'
                   f'SUMIFS([[Expense Tracker.pl_amount]],'
                   f'[[Expense Tracker.month]],ReportMonth,'
                   f'[[Expense Tracker.property]],{pid}))')
        p_blocked = (f'SUMIFS([[Booking Tracker.nights]],{bk},{prop},'
                     f'[[Booking Tracker.is_block]],"Yes")')
        p_avail = f'MAX(1,DAY({end})-{p_blocked})'
        p_profit = f'({p_net}-{p_clean}-{p_other})'
        p_scored = ('COUNTIFS([[Reviews Tracker.property]],' + pid +
                    ',[[Reviews Tracker.month]],ReportMonth,'
                    '[[Reviews Tracker.has_score]],"Yes")')
        p_rating = ('SUMIFS([[Reviews Tracker.rate_overall]],'
                    '[[Reviews Tracker.property]],' + pid +
                    ',[[Reviews Tracker.month]],ReportMonth,'
                    '[[Reviews Tracker.has_score]],"Yes")')
        s.formula(
            row, 0,
            s.reg.render(f'=IFERROR(IF({guard},"",INDEX('
                         f'[[Property Database.name]],{n})),"")'),
            fmts.calc("text"))
        cells = [
            (p_nights, "int"),
            (f'IF({p_avail}<=0,"",{p_nights}/{p_avail})', "percent"),
            (f'IF({p_nights}<=0,"",ROUND({p_accom}/{p_nights},2))',
             "money"),
            (f"ROUND({p_net},2)", "money"),
            (f"ROUND({p_clean},2)", "money"),
            (f"ROUND({p_other},2)", "money"),
            (f"ROUND({p_profit},2)", "money"),
            (f'IF({p_net}<=0,"",{p_profit}/{p_net})', "percent"),
            (f'IF({p_scored}<=0,"",ROUND({p_rating}/{p_scored},2))',
             "number"),
        ]
        for offset, (formula, style) in enumerate(cells):
            s.formula(
                row, offset + 1,
                s.reg.render(f'=IFERROR(IF({guard},"",{formula}),"")'),
                fmts.calc(style))
        row += 1
    row += 1

    # ------------------------------------------------- costs breakdown --
    row = s.section(row, "What the month cost", LAST_COL)
    row = s.table_head(row, 0, [
        "Category", "Amount", "Share of costs", "Share of net revenue",
        "Entries"], [26, 14, 15, 19, 11])
    cost_total = now["costs"]
    s.calc_row(row, 0, [
        "Cleaning and turnovers",
        f'=ROUND({now["clean"]},2)',
        f'=IF({cost_total}<=0,"",{now["clean"]}/{cost_total})',
        f'=IF({now["net"]}<=0,"",{now["clean"]}/{now["net"]})',
        '=COUNTIFS([[Cleaning Schedule.month]],ReportMonth)',
    ], ["text", "money", "percent", "percent", "int"])
    row += 1
    s.calc_row(row, 0, [
        "Repairs and maintenance",
        f'=ROUND({now["maint"]},2)',
        f'=IF({cost_total}<=0,"",{now["maint"]}/{cost_total})',
        f'=IF({now["net"]}<=0,"",{now["maint"]}/{now["net"]})',
        '=COUNTIFS([[Maintenance Log.month]],ReportMonth)',
    ], ["text", "money", "percent", "percent", "int"])
    row += 1
    for category in EXPENSE_CATEGORIES:
        amount = (f'SUMIFS([[Expense Tracker.pl_amount]],'
                  f'[[Expense Tracker.month]],ReportMonth,'
                  f'[[Expense Tracker.category]],"{category}")')
        count = (f'COUNTIFS([[Expense Tracker.month]],ReportMonth,'
                 f'[[Expense Tracker.category]],"{category}")')
        s.calc_row(row, 0, [
            category,
            f"=ROUND({amount},2)",
            f'=IF({cost_total}<=0,"",{amount}/{cost_total})',
            f'=IF({now["net"]}<=0,"",{amount}/{now["net"]})',
            f"={count}",
        ], ["text", "money", "percent", "percent", "int"])
        row += 1
    s.write(row, 0, "Total", fmts.chrome("total_label"))
    s.formula(row, 1, (f'=ROUND({cost_total},2)'),
                     fmts.chrome("total_value",
                                 num_format=s.num_format("money")))
    s.formula(row, 2, "=1",
                     fmts.chrome("total_value", num_format="0.0%"))
    s.formula(
        row, 3,
        s.reg.render(f'=IF({now["net"]}<=0,"",{cost_total}/'
                     f'{now["net"]})'),
        fmts.chrome("total_value", num_format="0.0%"))
    row += 2

    # ------------------------------------------------- operations -------
    row = s.section(row, "How the month ran", LAST_COL)
    row = s.table_head(row, 0, ["Measure", "Value", "What it tells you"],
                       [26, 14, 70])
    ops = [
        ("Turnovers", '=COUNTIFS([[Cleaning Schedule.month]],'
                      'ReportMonth)',
         "int", "Should be close to the number of check-outs."),
        ("Same-day turnovers",
         '=COUNTIFS([[Cleaning Schedule.month]],ReportMonth,'
         '[[Cleaning Schedule.same_day]],"Yes")', "int",
         "The ones that go wrong. Worth pricing for."),
        ("Turnovers left incomplete",
         '=COUNTIFS([[Cleaning Schedule.month]],ReportMonth,'
         '[[Cleaning Schedule.completion]],"<1")', "int",
         "A task unticked is a task not done."),
        ("Cleaning cost a booking",
         f'=IF({now["bookings"]}<=0,"",ROUND({now["clean"]}/'
         f'{now["bookings"]},2))', "money",
         "The number that makes one-night stays look expensive."),
        ("Cost a booked night",
         f'=IF({now["nights"]}<=0,"",ROUND({now["costs"]}/'
         f'{now["nights"]},2))', "money",
         "Compare with ADR. The gap is the real margin."),
        ("Profit a booked night",
         f'=IF({now["nights"]}<=0,"",ROUND({now["profit"]}/'
         f'{now["nights"]},2))', "money",
         "What one more night sold is actually worth."),
        ("Cancellations", f'={now["cancellations"]}', "int",
         "Rows marked cancelled with a check-in in this month."),
        ("Nights you blocked", f'={now["blocked"]}', "int",
         "Owner stays and closures. Taken out of occupancy."),
        ("Reviews received",
         '=COUNTIFS([[Reviews Tracker.month]],ReportMonth)', "int", ""),
        ("Average review score",
         '=IF(COUNTIFS([[Reviews Tracker.month]],ReportMonth,'
         '[[Reviews Tracker.has_score]],"Yes")<=0,"",ROUND(SUMIFS('
         '[[Reviews Tracker.rate_overall]],[[Reviews Tracker.month]],'
         'ReportMonth,[[Reviews Tracker.has_score]],"Yes")/COUNTIFS('
         '[[Reviews Tracker.month]],ReportMonth,'
         '[[Reviews Tracker.has_score]],"Yes"),2))', "number", ""),
        ("Payouts received",
         '=ROUND(SUMIFS([[Platform Revenue & Fees.received]],'
         '[[Platform Revenue & Fees.month]],ReportMonth),2)', "money",
         "What actually landed in the bank this month."),
        ("Fee variance",
         '=ROUND(SUMIFS([[Platform Revenue & Fees.variance]],'
         '[[Platform Revenue & Fees.month]],ReportMonth),2)', "money",
         "Real payouts against the fee rates you assumed."),
    ]
    for label, formula, style, note in ops:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.formula(row, 1, (formula),
                         fmts.calc(style))
        s.merge(row, 2, LAST_COL, note, fmts.chrome("note"))
        row += 1
    row += 1

    row = s.note(
        row,
        "A booking belongs to the month it checks in, which keeps one "
        "stay in one month and makes the columns add up. Costs belong "
        "to the month they are dated. If you want a stay split across "
        "two months, that is an accounting policy rather than a "
        "spreadsheet setting - note it and stay consistent.",
        LAST_COL, height=34)

    s.finish(freeze_row=6, gridlines=False, fit_width=True,
             print_area=f"A1:{col_letter(LAST_COL)}{row + 2}")
    s.protect()
    return s
````

### `airbnb_rental_spreadsheet/sheets/annual.py`

*534 lines · 5 functions*

> `Annual Reports` - twelve months across, and the year that was.
>
> Pick a year. Every month becomes a column: nights, occupancy, ADR,
> RevPAR, revenue, costs and profit, with the season visible at a
> glance because it is laid out the way a year actually behaves.
>
> Below that, the year in totals, the year against the year before,
> every property ranked, and the twelve questions worth asking in
> January - which month carried the year, what the quiet months really
> cost, and whether your rates moved at all.

| Defines | What it does |
|---------|--------------|
| `_key(month, year)` | The `yyyy-mm` key for one month of the chosen year |
| `_year_sum(value, months, year)` | A value summed over all twelve months of a year |
| `_year_sum_c(value, months, extra, year)` |  |
| `_year_count(months, extra, year)` |  |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Annual Reports'), `LAST_COL` (14), `PROPERTY_ROWS` (8).

````python
"""`Annual Reports` - twelve months across, and the year that was.

Pick a year. Every month becomes a column: nights, occupancy, ADR,
RevPAR, revenue, costs and profit, with the season visible at a
glance because it is laid out the way a year actually behaves.

Below that, the year in totals, the year against the year before,
every property ranked, and the twelve questions worth asking in
January - which month carried the year, what the quiet months really
cost, and whether your rates moved at all.
"""

from __future__ import annotations

from ..config import Config, Palette, SEASONS
from ..layout import Registry, col_letter
from ..writer import SheetWriter, number_check

NAME = "Annual Reports"
LAST_COL = 14
PROPERTY_ROWS = 8


def _key(month: int, year: str = "ReportYear") -> str:
    """The `yyyy-mm` key for one month of the chosen year."""
    return f'TEXT(DATE({year},{month},1),"yyyy-mm")'


def _year_sum(value: str, months: str, year: str = "ReportYear") -> str:
    """A value summed over all twelve months of a year. Bracketed."""
    parts = "+".join(
        f'SUMIFS({value},{months},{_key(m, year)})' for m in range(1, 13))
    return f"({parts})"


def _year_sum_c(value: str, months: str, extra: str,
                year: str = "ReportYear") -> str:
    parts = "+".join(
        f'SUMIFS({value},{months},{_key(m, year)},{extra})'
        for m in range(1, 13))
    return f"({parts})"


def _year_count(months: str, extra: str,
                year: str = "ReportYear") -> str:
    parts = "+".join(
        f'COUNTIFS({months},{_key(m, year)},{extra})'
        for m in range(1, 13))
    return f"({parts})"


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.deep)
    ws = s.ws

    s.banner("Annual Report",
             "Twelve months across. The year in full, against last "
             "year, property by property.", LAST_COL)
    s.set_col(0, 0, 26)
    for column in range(1, 13):
        s.set_col(column, column, 11)
    s.set_col(13, LAST_COL, 13)

    row = 4
    s.write(row, 0, "Year", fmts.chrome("label_bold"))
    s.write(row, 1, cfg.today.year,
            fmts.chrome("form_box", locked=False, num_format="0"))
    ws.data_validation(row, 1, row, 1,
                       number_check(2000, 2100, decimal=False))
    book.define_name("ReportYear", f"='{NAME}'!$B${row + 1}")
    s.merge(row, 2, LAST_COL,
            "Every figure on this page follows this one cell. A "
            "booking belongs to the year it checks in.",
            fmts.chrome("note"))
    row += 2

    bk_month = "[[Booking Tracker.month]]"
    earns = '[[Booking Tracker.earns]],"Yes"'
    cl_month = "[[Cleaning Schedule.month]]"
    mt_month = "[[Maintenance Log.month]]"
    ex_month = "[[Expense Tracker.month]]"

    y_nights = _year_sum_c("[[Booking Tracker.nights]]", bk_month, earns)
    y_accom = _year_sum_c("[[Booking Tracker.accommodation]]", bk_month,
                          earns)
    y_gross = _year_sum_c("[[Booking Tracker.gross]]", bk_month, earns)
    y_fees = _year_sum_c("[[Booking Tracker.platform_fee_total]]",
                         bk_month, earns)
    y_net = _year_sum_c("[[Booking Tracker.net_revenue]]", bk_month,
                        earns)
    y_guests = _year_sum_c("[[Booking Tracker.guests_total]]", bk_month,
                           earns)
    y_blocked = _year_sum_c("[[Booking Tracker.nights]]", bk_month,
                            '[[Booking Tracker.is_block]],"Yes"')
    y_bookings = _year_count(bk_month, earns)
    y_cancel = _year_count(bk_month,
                           '[[Booking Tracker.is_cancelled]],"Yes"')
    y_clean = _year_sum("[[Cleaning Schedule.total_cost]]", cl_month)
    y_maint = _year_sum("[[Maintenance Log.cost_used]]", mt_month)
    y_other = _year_sum("[[Expense Tracker.pl_amount]]", ex_month)
    y_costs = f"({y_clean}+{y_maint}+{y_other})"
    y_profit = f"({y_net}-{y_costs})"
    y_days = "(DATE(ReportYear+1,1,1)-DATE(ReportYear,1,1))"
    y_avail = f"MAX(1,PropertyCount*{y_days}-{y_blocked})"

    # --------------------------------------------------- the headline --
    row = s.section(row, "The year", LAST_COL)
    s.kpi(row, 0, "Bookings", f"={y_bookings}", "int")
    s.kpi(row, 2, "Nights sold", f"={y_nights}", "int")
    s.kpi(row, 4, "Occupancy",
          f'=IF({y_avail}<=0,"",{y_nights}/{y_avail})', "percent")
    s.kpi(row, 6, "ADR",
          f'=IF({y_nights}<=0,"",ROUND({y_accom}/{y_nights},2))',
          "money")
    s.kpi(row, 8, "RevPAR",
          f'=IF({y_avail}<=0,"",ROUND({y_accom}/{y_avail},2))', "money")
    s.kpi(row, 10, "Guests hosted", f"={y_guests}", "int")
    row = s.kpi(row, 12, "Average stay",
                f'=IF({y_bookings}<=0,"",ROUND({y_nights}/'
                f'{y_bookings},1))', "number1")

    s.kpi(row, 0, "Gross", f"=ROUND({y_gross},2)", "money0")
    s.kpi(row, 2, "Platform fees", f"=ROUND({y_fees},2)", "money0")
    s.kpi(row, 4, "Net revenue", f"=ROUND({y_net},2)", "money0")
    s.kpi(row, 6, "Costs", f"=ROUND({y_costs},2)", "money0")
    s.kpi(row, 8, "Profit", f"=ROUND({y_profit},2)", "money0")
    s.kpi(row, 10, "Margin",
          f'=IF({y_net}<=0,"",{y_profit}/{y_net})', "percent")
    row = s.kpi(row, 12, "Against target",
                f'=IF(TargetRevenueYear<=0,"",{y_net}/'
                f'TargetRevenueYear)', "percent")
    row += 1

    # ---------------------------------------------- the twelve months --
    row = s.section(row, "Month by month", LAST_COL)
    header = ["Measure"] + [
        f"{m}" for m in ("Jan", "Feb", "Mar", "Apr", "May", "Jun",
                         "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    ] + ["Year"]
    row = s.table_head(row, 0, header,
                       [26] + [11] * 12 + [13])

    def month_set(builder):
        """A row of twelve month formulas plus the year total."""
        return [builder(m) for m in range(1, 13)]

    def m_nights(m):
        return (f'SUMIFS([[Booking Tracker.nights]],{bk_month},'
                f'{_key(m)},{earns})')

    def m_accom(m):
        return (f'SUMIFS([[Booking Tracker.accommodation]],{bk_month},'
                f'{_key(m)},{earns})')

    def m_net(m):
        return (f'SUMIFS([[Booking Tracker.net_revenue]],{bk_month},'
                f'{_key(m)},{earns})')

    def m_gross(m):
        return (f'SUMIFS([[Booking Tracker.gross]],{bk_month},'
                f'{_key(m)},{earns})')

    def m_fees(m):
        return (f'SUMIFS([[Booking Tracker.platform_fee_total]],'
                f'{bk_month},{_key(m)},{earns})')

    def m_bookings(m):
        return f'COUNTIFS({bk_month},{_key(m)},{earns})'

    def m_costs(m):
        return (f'(SUMIFS([[Cleaning Schedule.total_cost]],{cl_month},'
                f'{_key(m)})+SUMIFS([[Maintenance Log.cost_used]],'
                f'{mt_month},{_key(m)})+'
                f'SUMIFS([[Expense Tracker.pl_amount]],{ex_month},'
                f'{_key(m)}))')

    def m_days(m):
        return f'DAY(DATE(ReportYear,{m + 1},1)-1)'

    def m_avail(m):
        blocked = (f'SUMIFS([[Booking Tracker.nights]],{bk_month},'
                   f'{_key(m)},[[Booking Tracker.is_block]],"Yes")')
        return f'MAX(1,PropertyCount*{m_days(m)}-{blocked})'

    lines = [
        ("Bookings", month_set(m_bookings), f"={y_bookings}", "int"),
        ("Nights sold", month_set(m_nights), f"={y_nights}", "int"),
        ("Occupancy",
         [f'IF({m_avail(m)}<=0,"",{m_nights(m)}/{m_avail(m)})'
          for m in range(1, 13)],
         f'=IF({y_avail}<=0,"",{y_nights}/{y_avail})', "percent"),
        ("ADR",
         [f'IF({m_nights(m)}<=0,"",ROUND({m_accom(m)}/{m_nights(m)},2))'
          for m in range(1, 13)],
         f'=IF({y_nights}<=0,"",ROUND({y_accom}/{y_nights},2))',
         "money"),
        ("RevPAR",
         [f'IF({m_avail(m)}<=0,"",ROUND({m_accom(m)}/{m_avail(m)},2))'
          for m in range(1, 13)],
         f'=IF({y_avail}<=0,"",ROUND({y_accom}/{y_avail},2))', "money"),
        ("Gross", [f"ROUND({m_gross(m)},2)" for m in range(1, 13)],
         f"=ROUND({y_gross},2)", "money0"),
        ("Platform fees",
         [f"ROUND({m_fees(m)},2)" for m in range(1, 13)],
         f"=ROUND({y_fees},2)", "money0"),
        ("Net revenue", [f"ROUND({m_net(m)},2)" for m in range(1, 13)],
         f"=ROUND({y_net},2)", "money0"),
        ("Costs", [f"ROUND({m_costs(m)},2)" for m in range(1, 13)],
         f"=ROUND({y_costs},2)", "money0"),
        ("Profit",
         [f"ROUND({m_net(m)}-{m_costs(m)},2)" for m in range(1, 13)],
         f"=ROUND({y_profit},2)", "money0"),
        ("Margin",
         [f'IF({m_net(m)}<=0,"",({m_net(m)}-{m_costs(m)})/{m_net(m)})'
          for m in range(1, 13)],
         f'=IF({y_net}<=0,"",{y_profit}/{y_net})', "percent"),
        ("Share of the year",
         [f'IF({y_net}<=0,"",{m_net(m)}/{y_net})'
          for m in range(1, 13)],
         "=1", "percent"),
    ]
    grid_top = row
    sparkline_rows = {}
    for label, cells, total, style in lines:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        for index, formula in enumerate(cells):
            s.formula(row, index + 1,
                             s.reg.render("=" + formula.lstrip("=")),
                             fmts.calc(style))
        s.formula(row, 13, s.reg.render(total),
                         fmts.chrome("total_value",
                                     num_format=s.num_format(style)))
        sparkline_rows[label] = row
        row += 1
    grid_bottom = row - 1

    for label, kind in (("Occupancy", "line"), ("ADR", "line"),
                        ("Net revenue", "column"),
                        ("Profit", "column")):
        line = sparkline_rows[label]
        ws.add_sparkline(line, 14, {
            "range": f"'{NAME}'!B{line + 1}:M{line + 1}",
            "type": kind, "style": 12,
            "negative_points": kind == "column"})
    s.write(grid_top - 1, 14, "Shape", fmts.chrome("header"))

    ws.conditional_format(sparkline_rows["Occupancy"], 1,
                          sparkline_rows["Occupancy"], 12, {
                              "type": "cell", "criteria": ">=",
                              "value": "TargetOccupancy",
                              "format": fmts.cf("good")})
    ws.conditional_format(sparkline_rows["Profit"], 1,
                          sparkline_rows["Profit"], 12, {
                              "type": "cell", "criteria": "<",
                              "value": 0, "format": fmts.cf("bad")})
    row = grid_bottom + 2

    # ------------------------------------------------------ seasons ----
    row = s.section(row, "By season", LAST_COL)
    row = s.note(
        row,
        "Short-term lets are a seasonal business and the averages "
        "hide it. Three strong months paying for nine quiet ones is a "
        "completely different business from twelve steady ones, and "
        "they need different pricing.",
        LAST_COL, height=28)
    row = s.table_head(row, 0, [
        "Season", "Months", "Nights", "Occupancy", "ADR", "RevPAR",
        "Net revenue", "Costs", "Profit", "Share of profit"],
        [20, 18, 10, 11, 11, 11, 13, 12, 12, 14])
    for season, first_month, last_month in SEASONS:
        if first_month <= last_month:
            months = list(range(first_month, last_month + 1))
        else:                       # winter wraps the year end
            months = (list(range(first_month, 13))
                      + list(range(1, last_month + 1)))
        nights = "+".join(m_nights(m) for m in months)
        accom = "+".join(m_accom(m) for m in months)
        net = "+".join(m_net(m) for m in months)
        costs = "+".join(m_costs(m) for m in months)
        avail = "+".join(m_avail(m) for m in months)
        label = ", ".join(
            ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug",
             "Sep", "Oct", "Nov", "Dec")[m - 1] for m in months)
        s.calc_row(row, 0, [
            season, label,
            f"=({nights})",
            f'=IF(({avail})<=0,"",({nights})/({avail}))',
            f'=IF(({nights})<=0,"",ROUND(({accom})/({nights}),2))',
            f'=IF(({avail})<=0,"",ROUND(({accom})/({avail}),2))',
            f"=ROUND({net},2)",
            f"=ROUND({costs},2)",
            f"=ROUND(({net})-({costs}),2)",
            f'=IF({y_profit}<=0,"",(({net})-({costs}))/{y_profit})',
        ], ["text", "text", "int", "percent", "money", "money",
            "money0", "money0", "money0", "percent"])
        row += 1
    row += 1

    # --------------------------------------------- against last year ---
    row = s.section(row, "Against last year", LAST_COL)
    row = s.table_head(row, 0, [
        "Measure", "This year", "Last year", "Change", "Change %"],
        [26, 15, 15, 14, 12])
    ly = "ReportYear-1"
    ly_nights = _year_sum_c("[[Booking Tracker.nights]]", bk_month,
                            earns, ly)
    ly_accom = _year_sum_c("[[Booking Tracker.accommodation]]",
                           bk_month, earns, ly)
    ly_gross = _year_sum_c("[[Booking Tracker.gross]]", bk_month,
                           earns, ly)
    ly_fees = _year_sum_c("[[Booking Tracker.platform_fee_total]]",
                          bk_month, earns, ly)
    ly_net = _year_sum_c("[[Booking Tracker.net_revenue]]", bk_month,
                         earns, ly)
    ly_bookings = _year_count(bk_month, earns, ly)
    ly_costs = ("(" + _year_sum("[[Cleaning Schedule.total_cost]]",
                                cl_month, ly)
                + "+" + _year_sum("[[Maintenance Log.cost_used]]",
                                  mt_month, ly)
                + "+" + _year_sum("[[Expense Tracker.pl_amount]]",
                                  ex_month, ly) + ")")
    ly_profit = f"({ly_net}-{ly_costs})"
    ly_days = "(DATE(ReportYear,1,1)-DATE(ReportYear-1,1,1))"
    ly_avail = f"MAX(1,PropertyCount*{ly_days})"

    pairs = [
        ("Bookings", y_bookings, ly_bookings, "int", False),
        ("Nights sold", y_nights, ly_nights, "int", False),
        ("Occupancy", f'IF({y_avail}<=0,0,{y_nights}/{y_avail})',
         f'IF({ly_avail}<=0,0,{ly_nights}/{ly_avail})', "percent", True),
        ("ADR", f'IF({y_nights}<=0,0,{y_accom}/{y_nights})',
         f'IF({ly_nights}<=0,0,{ly_accom}/{ly_nights})', "money", False),
        ("RevPAR", f'IF({y_avail}<=0,0,{y_accom}/{y_avail})',
         f'IF({ly_avail}<=0,0,{ly_accom}/{ly_avail})', "money", False),
        ("Gross", y_gross, ly_gross, "money0", False),
        ("Platform fees", y_fees, ly_fees, "money0", False),
        ("Net revenue", y_net, ly_net, "money0", False),
        ("Costs", y_costs, ly_costs, "money0", False),
        ("Profit", y_profit, ly_profit, "money0", False),
        ("Margin", f'IF({y_net}<=0,0,{y_profit}/{y_net})',
         f'IF({ly_net}<=0,0,{ly_profit}/{ly_net})', "percent", True),
    ]
    compare_top = row
    for label, this, last, style, absolute in pairs:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.formula(row, 1, (f"=ROUND({this},4)"),
                         fmts.calc(style))
        s.formula(row, 2, s.reg.render(f"=ROUND({last},4)"),
                         fmts.calc(style))
        s.formula(row, 3,
                         s.reg.render(f"=ROUND({this}-{last},4)"),
                         fmts.calc("percent" if absolute else style))
        s.formula(
            row, 4,
            s.reg.render(f'=IF({last}<=0,"",{this}/{last}-1)'),
            fmts.calc("percent"))
        row += 1
    for column in (3, 4):
        ws.conditional_format(compare_top, column, row - 1, column, {
            "type": "cell", "criteria": "<", "value": 0,
            "format": fmts.cf("bad")})
        ws.conditional_format(compare_top, column, row - 1, column, {
            "type": "cell", "criteria": ">", "value": 0,
            "format": fmts.cf("good")})
    row += 1

    # ------------------------------------------------- by property -----
    row = s.section(row, "Every property, for the year", LAST_COL)
    row = s.table_head(row, 0, [
        "Property", "Bookings", "Nights", "Occupancy", "ADR", "RevPAR",
        "Net revenue", "Costs", "Profit", "Margin", "Profit a night",
        "Share of profit"],
        [24, 11, 10, 11, 11, 11, 13, 12, 12, 10, 13, 14])
    for n in range(1, PROPERTY_ROWS + 1):
        pid = f'INDEX([[Property Database.id]],{n})'
        guard = f'INDEX([[Property Database.name]],{n})=""'
        prop = f'[[Booking Tracker.property]],{pid}'
        p_nights = "+".join(
            f'SUMIFS([[Booking Tracker.nights]],{bk_month},{_key(m)},'
            f'{earns},{prop})' for m in range(1, 13))
        p_accom = "+".join(
            f'SUMIFS([[Booking Tracker.accommodation]],{bk_month},'
            f'{_key(m)},{earns},{prop})' for m in range(1, 13))
        p_net = "+".join(
            f'SUMIFS([[Booking Tracker.net_revenue]],{bk_month},'
            f'{_key(m)},{earns},{prop})' for m in range(1, 13))
        p_bookings = "+".join(
            f'COUNTIFS({bk_month},{_key(m)},{earns},{prop})'
            for m in range(1, 13))
        p_costs = "+".join(
            f'SUMIFS([[Cleaning Schedule.total_cost]],{cl_month},'
            f'{_key(m)},[[Cleaning Schedule.property]],{pid})+'
            f'SUMIFS([[Maintenance Log.cost_used]],{mt_month},'
            f'{_key(m)},[[Maintenance Log.property]],{pid})+'
            f'SUMIFS([[Expense Tracker.pl_amount]],{ex_month},'
            f'{_key(m)},[[Expense Tracker.property]],{pid})'
            for m in range(1, 13))
        p_avail = f"MAX(1,{y_days})"
        p_profit = f"(({p_net})-({p_costs}))"
        s.formula(
            row, 0,
            s.reg.render(f'=IFERROR(IF({guard},"",INDEX('
                         f'[[Property Database.name]],{n})),"")'),
            fmts.calc("text"))
        cells = [
            (f"({p_bookings})", "int"),
            (f"({p_nights})", "int"),
            (f'IF({p_avail}<=0,"",({p_nights})/{p_avail})', "percent"),
            (f'IF(({p_nights})<=0,"",ROUND(({p_accom})/({p_nights}),2))',
             "money"),
            (f'IF({p_avail}<=0,"",ROUND(({p_accom})/{p_avail},2))',
             "money"),
            (f"ROUND({p_net},2)", "money0"),
            (f"ROUND({p_costs},2)", "money0"),
            (f"ROUND({p_profit},2)", "money0"),
            (f'IF(({p_net})<=0,"",{p_profit}/({p_net}))', "percent"),
            (f'IF(({p_nights})<=0,"",ROUND({p_profit}/({p_nights}),2))',
             "money"),
            (f'IF({y_profit}<=0,"",{p_profit}/{y_profit})', "percent"),
        ]
        for offset, (formula, style) in enumerate(cells):
            s.formula(
                row, offset + 1,
                s.reg.render(f'=IFERROR(IF({guard},"",{formula}),"")'),
                fmts.calc(style))
        row += 1
    row += 1

    # ----------------------------------------------- the twelve asks ---
    row = s.section(row, "The questions worth asking in January",
                    LAST_COL)
    month_head = f"$B${grid_top}:$M${grid_top}"
    net_band = (f"$B${sparkline_rows['Net revenue'] + 1}:"
                f"$M${sparkline_rows['Net revenue'] + 1}")
    profit_band = (f"$B${sparkline_rows['Profit'] + 1}:"
                   f"$M${sparkline_rows['Profit'] + 1}")
    row = s.table_head(row, 0, ["Question", "Answer", "Why it matters"],
                       [34, 16, 80])
    asks = [
        ("Which month carried the year?",
         f'=IFERROR(INDEX({month_head},MATCH(MAX({net_band}),'
         f'{net_band},0)),"")',
         "text",
         "If one month is more than a fifth of the year, your "
         "pricing the rest of the time is doing very little work."),
        ("Which month was the thinnest?",
         f'=IFERROR(INDEX({month_head},MATCH(MIN({net_band}),'
         f'{net_band},0)),"")',
         "text",
         "The one to attack with a minimum-stay change, a price "
         "drop or a different audience."),
        ("How many months lost money?",
         f'=COUNTIF({profit_band},"<0")',
         "int",
         "A month below zero is not a disaster; three in a row is a "
         "business model."),
        ("Best month against worst",
         f'=IF(MIN({net_band})<=0,"",MAX({net_band})/MIN({net_band}))',
         "number1",
         "How seasonal you really are. Above three and you are "
         "running a summer business with twelve months of costs."),
        ("Nights you never sold",
         f'=MAX(0,{y_avail}-{y_nights})', "int",
         "Each one is an ADR you will never get back."),
        ("What the empty nights cost",
         f'=ROUND(MAX(0,{y_avail}-{y_nights})*IF({y_nights}<=0,0,'
         f'{y_accom}/{y_nights}),2)', "money0",
         "Empty nights valued at your own average rate. Not a real "
         "loss - nobody sells every night - but the size of the "
         "prize."),
        ("Cleaning as a share of net revenue",
         f'=IF({y_net}<=0,"",{y_clean}/{y_net})', "percent",
         "Over about a fifth and either your stays are too short or "
         "your cleaning fee is too low."),
        ("Platform fees as a share of gross",
         f'=IF({y_gross}<=0,"",{y_fees}/{y_gross})', "percent",
         "Worth comparing with the rates on Settings. A gap means "
         "your mix has shifted towards the expensive platforms."),
        ("Cost of a booked night",
         f'=IF({y_nights}<=0,"",ROUND({y_costs}/{y_nights},2))',
         "money",
         "Your break-even nightly rate, near enough."),
        ("Profit a booked night",
         f'=IF({y_nights}<=0,"",ROUND({y_profit}/{y_nights},2))',
         "money", "What one more night sold is worth."),
        ("Profit a property",
         f'=IF(PropertyCount<=0,"",ROUND({y_profit}/PropertyCount,2))',
         "money0",
         "Divide by the hours you put in and decide whether this is "
         "a business or a hobby."),
        ("Cancellation rate",
         f'=IF({y_bookings}+{y_cancel}<=0,"",{y_cancel}/'
         f'({y_bookings}+{y_cancel}))', "percent",
         "Above one in ten, look at your deposit and cancellation "
         "terms."),
        ("Nights you blocked yourself",
         f"={y_blocked}", "int",
         "Your own stays and closures. Valued at your average rate "
         "they are the real cost of keeping the place available to "
         "you."),
        ("Average review score",
         '=IF(' + _year_count('[[Reviews Tracker.month]]',
                              '[[Reviews Tracker.has_score]],"Yes"') +
         '<=0,"",ROUND(' +
         _year_sum_c('[[Reviews Tracker.rate_overall]]',
                     '[[Reviews Tracker.month]]',
                     '[[Reviews Tracker.has_score]],"Yes"') + '/' +
         _year_count('[[Reviews Tracker.month]]',
                     '[[Reviews Tracker.has_score]],"Yes"') + ',2))',
         "number",
         "The number that quietly sets your occupancy next year."),
    ]
    for question, formula, style, why in asks:
        s.write(row, 0, question, fmts.chrome("label_bold"))
        s.formula(row, 1, (formula),
                         fmts.calc(style))
        s.merge(row, 2, LAST_COL, why, fmts.chrome("note"))
        s.row_height(row, 28)
        row += 1
    row += 1

    row = s.note(
        row,
        "Everything on this page attributes a stay to the year it "
        "checks in, and a cost to the date it carries. Change the "
        "year cell at the top and the whole page follows - including "
        "the comparison, which simply reads the year before.",
        LAST_COL, height=28)

    s.finish(freeze_row=6, freeze_col=1, gridlines=False,
             fit_width=True, print_area=f"A1:{col_letter(LAST_COL)}{row + 2}")
    s.protect()
    return s
````

---

## 9. The quality tools

Not shipped with the workbook, but the reason it can be trusted: a static validator that also enforces three promises the listing makes, a bracket and placeholder checker, a demo data checker that runs before the file is written, and a recalculation harness that really evaluates every formula.

### `tools/validate_airbnb_workbook.py`

*1,015 lines · 17 functions*

> Static checks on a generated Airbnb & Short-Term Rental workbook.
>
> Run it with any Python that has openpyxl installed::
>
>     python tools/validate_airbnb_workbook.py output/airbnb.xlsx
>
> It never opens Excel - it parses the XML and reports anything that would
> break for a buyer: unresolved template tokens, formulas pointing at the
> wrong sheet, defined names that do not exist, missing sheet protection,
> formula cells that are not locked, and input cells that are.

| Defines | What it does |
|---------|--------------|
| `_colnum(ref)` |  |
| `check_raw_xml(path)` | Overlapping merges and stray hyperlinks - Excel calls these corrupt |
| `paren_depth(formula)` | Open brackets minus closing ones, ignoring anything inside text |
| `_top_level(text, wanted)` | Positions of `wanted` characters outside brackets and strings |
| `_arguments(formula)` | Every comma separated argument, at every bracket depth |
| `precedence_slip(formula)` | `SUMIFS(..)+SUMIFS(..)/COUNTIFS(..)` - the sum is not divided |
| `_span(text)` | (rows, columns) of an A1 range, or None if it is not one |
| `_calls(formula, names)` | Yield (function, [argument, ...]) for every call in *formula* |
| `ragged_criteria(formula)` | COUNTIFS/SUMIFS whose ranges are not all the same shape |
| `collapsed_criteria(formula)` | A COUNTIFS/SUMIFS criteria *range* that carries its own test |
| `_name_target(wb, name)` | The one cell a defined name points at, or None |
| `check_pages(wb, problems, notes)` | The summary pages, checked through their defined names |
| `_specs()` | Every register specification, keyed by sheet name |
| `check_chain(wb, problems, notes)` | The enquiry-to-enrolment-to-invoice chain only holds while these do |
| `check_platform_table(wb, problems, notes)` | The platform fee table a buyer must be able to type over |
| `check_promises(wb, problems, notes)` | Three promises the product makes in writing |
| `check(path)` |  |

**Module constants:** `FUNCTION_RE` (re.compile(...)), `NAME_RE` (re.compile(...)), `BARE_COL_RE` (re.compile(...)), `KNOWN_FUNCTIONS` ({'IF', 'IFERROR', 'SUM', 'SUMIF', 'SUMIFS', 'COUNT', 'COU...), `MODERN_ONLY` ({'XLOOKUP', 'IFS', 'MAXIFS', 'MINIFS', 'TEXTJOIN', 'CONCA...), `PRECEDENCE_OK` (0 entries), `RANGE_RE` (re.compile(...)), `SEEDED_INPUTS` ({'Monthly Reports!B5', 'Annual Reports!B5', 'Tax Summary!...), `CHAIN_CHECKS` (81 items), `PAGE_FORMULAS` (10 items), `PAGE_INPUTS` (45 items), `REQUIRED_LIST_NAMES` (36 items).

````python
"""Static checks on a generated Airbnb & Short-Term Rental workbook.

Run it with any Python that has openpyxl installed::

    python tools/validate_airbnb_workbook.py output/airbnb.xlsx

It never opens Excel - it parses the XML and reports anything that would
break for a buyer: unresolved template tokens, formulas pointing at the
wrong sheet, defined names that do not exist, missing sheet protection,
formula cells that are not locked, and input cells that are.
"""

from __future__ import annotations

import os
import re
import sys
import zipfile
from collections import Counter, defaultdict

try:
    import openpyxl
except ImportError:                                   # pragma: no cover
    print("openpyxl is required: pip install openpyxl")
    raise SystemExit(2)

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from airbnb_rental_spreadsheet.build import SHEET_ORDER as EXPECTED_SHEETS  # noqa: E402,E501

FUNCTION_RE = re.compile(r"\b([A-Z][A-Z0-9\.]{1,})\s*\(")
NAME_RE = re.compile(r"(?<![A-Za-z0-9_.!$\"])([A-Za-z_][A-Za-z0-9_.]{2,})"
                     r"(?![A-Za-z0-9_.(!])")
#: a column letter with no row number - almost always a missing {r}
BARE_COL_RE = re.compile(r"(?<![A-Za-z0-9_.!$\"])([A-Z]{1,3})"
                         r"(?![A-Za-z0-9_.(!$:])")

KNOWN_FUNCTIONS = {
    "IF", "IFERROR", "SUM", "SUMIF", "SUMIFS", "COUNT", "COUNTA", "COUNTIF",
    "COUNTIFS", "AVERAGE", "AVERAGEIF", "AVERAGEIFS", "MAX", "MIN", "ROUND",
    "ROUNDUP", "ROUNDDOWN", "INDEX", "MATCH", "OFFSET", "TEXT", "VALUE",
    "LEFT", "RIGHT", "MID", "LEN", "TRIM", "UPPER", "LOWER", "PROPER",
    "CONCATENATE", "TODAY", "NOW", "DATE", "DAY", "MONTH", "YEAR", "WEEKDAY",
    "WEEKNUM", "EOMONTH", "EDATE", "DATEDIF", "DAYS", "NETWORKDAYS", "HOUR",
    "MINUTE", "TIME", "SUMPRODUCT", "LARGE", "SMALL", "RANK", "ABS", "AND",
    "OR", "NOT", "ISBLANK", "ISNUMBER", "ISTEXT", "ISERROR", "ISNA", "NA",
    "CHOOSE", "VLOOKUP", "HLOOKUP", "LOOKUP", "SUBSTITUTE", "REPT",
    "MEDIAN", "STDEV", "FLOOR", "CEILING", "INT", "MOD", "SIGN", "RAND",
    "SEARCH", "FIND", "EXACT", "N", "T", "TRUE", "FALSE", "ROW", "ROWS",
    "COLUMN", "COLUMNS", "YEARFRAC", "INDIRECT", "ADDRESS", "CELL",
    "SUBTOTAL", "TRANSPOSE", "PRODUCT", "POWER", "SQRT", "COUNTBLANK",
    "ISERR", "ISLOGICAL", "ISREF", "TYPE", "NUMBERVALUE", "DATEVALUE",
    "TIMEVALUE", "WORKDAY", "FREQUENCY", "PERCENTILE",
}

#: functions that do not exist in Excel 2010 / Google Sheets
MODERN_ONLY = {
    "XLOOKUP", "IFS", "MAXIFS", "MINIFS", "TEXTJOIN", "CONCAT", "SWITCH",
    "LET", "LAMBDA", "FILTER", "SORT", "SORTBY", "UNIQUE", "SEQUENCE",
    "TEXTSPLIT", "AGGREGATE", "IFNA",
}


def _colnum(ref: str) -> tuple:
    match = re.match(r"([A-Z]+)(\d+)", ref)
    letters, row = match.group(1), int(match.group(2))
    num = 0
    for char in letters:
        num = num * 26 + (ord(char) - 64)
    return num, row


def check_raw_xml(path: str) -> list:
    """Overlapping merges and stray hyperlinks - Excel calls these corrupt."""
    found = []
    with zipfile.ZipFile(path) as zf:
        book = zf.read("xl/workbook.xml").decode()
        sheets = re.findall(r'<sheet name="([^"]+)"[^>]*r:id="rId(\d+)"',
                            book)
        sheet_names = {name for name, _ in sheets}
        for name, rid in sheets:
            try:
                xml = zf.read(f"xl/worksheets/sheet{rid}.xml").decode()
            except KeyError:
                continue
            merges = re.findall(r'<mergeCell ref="([^"]+)"/>', xml)
            boxes = []
            for ref in merges:
                start, _, end = ref.partition(":")
                end = end or start
                c1, r1 = _colnum(start)
                c2, r2 = _colnum(end)
                boxes.append((c1, r1, c2, r2, ref))
            for i, a in enumerate(boxes):
                for b in boxes[i + 1:]:
                    if (a[0] <= b[2] and b[0] <= a[2]
                            and a[1] <= b[3] and b[1] <= a[3]):
                        found.append(f"{name}: merges overlap "
                                     f"{a[4]} / {b[4]}")
            for location in re.findall(
                    r'<hyperlink ref="[^"]+" location="([^"]+)"', xml):
                target = location.split("!")[0].strip("'").replace("''", "'")
                if target not in sheet_names:
                    found.append(f"{name}: hyperlink points at missing "
                                 f"sheet '{target}'")
            for link in re.findall(r'<hyperlink ref="([^"]+)"', xml):
                lc, lr = _colnum(link)
                for c1, r1, c2, r2, ref in boxes:
                    if (c1 <= lc <= c2 and r1 <= lr <= r2
                            and (lc, lr) != (c1, r1)):
                        found.append(f"{name}: hyperlink {link} is inside "
                                     f"merged range {ref}")
    return found


def paren_depth(formula: str) -> int:
    """Open brackets minus closing ones, ignoring anything inside text.

    Excel escapes a quote inside a string by doubling it, so `""` is not
    the end of the string. Deeply nested IF() chains built by string
    concatenation are the one bug this whole file exists to catch: Excel
    and openpyxl both accept a short formula silently, and the buyer
    sees the damage.
    """
    depth = 0
    in_text = False
    i = 0
    while i < len(formula):
        char = formula[i]
        if in_text:
            if char == '"':
                if i + 1 < len(formula) and formula[i + 1] == '"':
                    i += 1
                else:
                    in_text = False
        elif char == '"':
            in_text = True
        elif char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
        i += 1
    return depth


#: divisions that really are meant to read `a + b/c`.  Each one is the
#: formula with its numbers blanked out, so one entry covers the whole
#: column.
PRECEDENCE_OK = {
    # nothing here yet: every `a + b/c` on this workbook is bracketed
    # in the source, which is the honest fix rather than a waiver
}


def _top_level(text, wanted):
    """Positions of `wanted` characters outside brackets and strings."""
    out = []
    depth = 0
    in_text = False
    i = 0
    while i < len(text):
        char = text[i]
        if in_text:
            if char == '"':
                if i + 1 < len(text) and text[i + 1] == '"':
                    i += 1
                else:
                    in_text = False
        elif char == '"':
            in_text = True
        elif char == "(":
            depth += 1
        elif char == ")":
            depth -= 1
        elif depth == 0 and char in wanted:
            before = text[i - 1] if i else ""
            if char != "/" and (before == "" or before in "(,*/+-^<>=&"):
                # a leading sign, not an operator between two values
                i += 1
                continue
            out.append(i)
        i += 1
    return out


def _arguments(formula):
    """Every comma separated argument, at every bracket depth."""
    starts = {0: 0}
    depth = 0
    in_text = False
    i = 0
    while i < len(formula):
        char = formula[i]
        if in_text:
            if char == '"':
                if i + 1 < len(formula) and formula[i + 1] == '"':
                    i += 1
                else:
                    in_text = False
        elif char == '"':
            in_text = True
        elif char == "(":
            depth += 1
            starts[depth] = i + 1
        elif char == ")":
            yield formula[starts.get(depth, 0):i]
            depth -= 1
        elif char == ",":
            yield formula[starts.get(depth, 0):i]
            starts[depth] = i + 1
        i += 1
    yield formula[starts.get(0, 0):]


def precedence_slip(formula):
    """`SUMIFS(..)+SUMIFS(..)/COUNTIFS(..)` - the sum is not divided.

    Excel divides before it adds, so a total built by adding two
    SUMIFS together has to be bracketed before it is divided by
    anything.  The file still opens, every bracket balances and the
    calculation engine returns a number - it is simply the wrong
    number, which is the worst kind of bug to ship.
    """
    for argument in _arguments(formula):
        if "/" not in argument:
            continue
        shape = re.sub(r"\$?[A-Z]{1,3}\$?\d+", "#", argument)
        shape = re.sub(r"\d+", "#", shape)
        if shape in PRECEDENCE_OK:
            continue
        sums = _top_level(argument, "+-")
        if not sums:
            continue
        if any(divide > sums[0] for divide in _top_level(argument, "/")):
            return argument
    return ""


RANGE_RE = re.compile(
    r"^(?:(?:'[^']+'|[A-Za-z0-9_.]+)!)?"
    r"\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)$")


def _span(text: str):
    """(rows, columns) of an A1 range, or None if it is not one."""
    match = RANGE_RE.match(text.strip())
    if not match:
        return None
    c1, r1, c2, r2 = match.groups()

    def num(letters):
        value = 0
        for char in letters:
            value = value * 26 + (ord(char) - 64)
        return value
    return (abs(int(r2) - int(r1)) + 1, abs(num(c2) - num(c1)) + 1)


def _calls(formula: str, names):
    """Yield (function, [argument, ...]) for every call in *formula*."""
    depth = 0
    in_text = False
    stack = []
    i = 0
    while i < len(formula):
        char = formula[i]
        if in_text:
            if char == '"':
                if i + 1 < len(formula) and formula[i + 1] == '"':
                    i += 1
                else:
                    in_text = False
        elif char == '"':
            in_text = True
        elif char == "(":
            head = re.search(r"([A-Z][A-Z0-9\.]*)$", formula[:i])
            stack.append([head.group(1) if head else "", i + 1, [], depth])
            depth += 1
        elif char == ")":
            depth -= 1
            if stack:
                name, start, args, _ = stack.pop()
                args.append(formula[start:i])
                if name in names:
                    yield name, args
                if stack:
                    stack[-1][1] = stack[-1][1]
        elif char == "," and stack and depth == stack[-1][3] + 1:
            frame = stack[-1]
            frame[2].append(formula[frame[1]:i])
            frame[1] = i + 1
        i += 1


def ragged_criteria(formula: str) -> str:
    """COUNTIFS/SUMIFS whose ranges are not all the same shape.

    Excel answers #VALUE! the moment two criteria ranges differ in
    size, and it is an easy mistake to make when one of them is a
    running range that grows row by row.
    """
    for name, args in _calls(formula, ("COUNTIFS", "SUMIFS",
                                       "AVERAGEIFS")):
        first = 1 if name in ("SUMIFS", "AVERAGEIFS") else 0
        ranges = [args[0]] if first else []
        ranges += args[first::2]
        spans = {}
        for text in ranges:
            span = _span(text)
            if span:
                spans.setdefault(span, text)
        if len(spans) > 1:
            return f"{name}: " + " vs ".join(sorted(spans.values()))
    return ""


def collapsed_criteria(formula: str) -> str:
    """A COUNTIFS/SUMIFS criteria *range* that carries its own test.

    A criterion is two arguments - the range, then the thing it is
    compared against.  Writing it as one argument::

        COUNTIFS(prop, "PRP-01", scores="Yes")

    leaves the function an odd number of arguments and makes Excel
    answer #VALUE! on every row.  It reads perfectly naturally, which
    is exactly why it is worth a check: `SUMPRODUCT((range="Yes")*..)`
    really is written that way, so the habit carries over.
    """
    for name, args in _calls(formula, ("COUNTIFS", "SUMIFS",
                                       "AVERAGEIFS")):
        first = 1 if name in ("SUMIFS", "AVERAGEIFS") else 0
        if (len(args) - first) % 2:
            return f"{name}: odd number of range/criteria arguments"
        for text in args[first::2]:
            stripped = re.sub(r'"[^"]*"', '""', text)
            if re.search(r"[<>=]", stripped):
                return f"{name}: criteria range carries a test -> {text}"
    return ""


#: input cells that ship with a formula in them as a helpful default -
#: the buyer is meant to be able to type a fixed value straight over it
SEEDED_INPUTS = {
    "Monthly Reports!B5",        # the month the report opens on
    "Annual Reports!B5",         # the year the report opens on
    "Tax Summary!C5",            # the tax year
    "Booking Calendar!B10",      # the property the calendar shows
    "Booking Calendar!B11",      # the month the calendar shows
    "Turnover Checklist!C5",     # the cleaning job being printed
}

#: the automation chain, link by link: (sheet, column key, what the
#: formula in that column has to still be pointing at)
CHAIN_CHECKS = [
    # ---- properties: the hub every other tab joins to -------------------
    ("Property Database", "id", ("text", "PropertyPrefix")),
    ("Property Database", "bookings", ("sheet", "Booking Tracker")),
    ("Property Database", "nights_booked", ("sheet", "Booking Tracker")),
    ("Property Database", "occupancy", ("column", "Property Database",
                                        "nights_available")),
    ("Property Database", "adr", ("column", "Property Database",
                                  "accommodation")),
    ("Property Database", "revpar", ("column", "Property Database",
                                     "accommodation")),
    ("Property Database", "cleaning_spend", ("sheet",
                                             "Cleaning Schedule")),
    ("Property Database", "maintenance_spend", ("sheet",
                                                "Maintenance Log")),
    ("Property Database", "other_expenses", ("sheet",
                                             "Expense Tracker")),
    ("Property Database", "profit", ("column", "Property Database",
                                     "total_costs")),
    ("Property Database", "margin", ("column", "Property Database",
                                     "net_revenue")),
    ("Property Database", "rating", ("sheet", "Reviews Tracker")),
    ("Property Database", "low_stock", ("sheet", "Supplies Inventory")),
    ("Property Database", "next_arrival", ("sheet", "Booking Tracker")),
    # ---- bookings: the other half of the join ---------------------------
    ("Booking Tracker", "id", ("text", "BookingPrefix")),
    ("Booking Tracker", "property_name", ("sheet", "Property Database")),
    ("Booking Tracker", "guest_name", ("sheet", "Guest Database")),
    ("Booking Tracker", "nights", ("column", "Booking Tracker",
                                   "check_in")),
    ("Booking Tracker", "rate_used", ("sheet", "Property Database")),
    ("Booking Tracker", "accommodation", ("column", "Booking Tracker",
                                          "rate_used")),
    ("Booking Tracker", "extra_fee", ("text", "ExtraGuestFee")),
    ("Booking Tracker", "gross", ("column", "Booking Tracker",
                                  "accommodation")),
    ("Booking Tracker", "fee_pct", ("text", "PlatformCommission")),
    ("Booking Tracker", "pay_pct", ("text", "PlatformPayment")),
    ("Booking Tracker", "platform_fee", ("column", "Booking Tracker",
                                         "fee_pct")),
    ("Booking Tracker", "payment_fee", ("column", "Booking Tracker",
                                        "pay_pct")),
    ("Booking Tracker", "net_revenue", ("column", "Booking Tracker",
                                        "platform_fee_total")),
    ("Booking Tracker", "tax", ("text", "TaxRate")),
    ("Booking Tracker", "clash", ("column", "Booking Tracker",
                                  "ci_num")),
    ("Booking Tracker", "same_day", ("column", "Booking Tracker",
                                     "co_num")),
    ("Booking Tracker", "has_cleaning", ("sheet", "Cleaning Schedule")),
    ("Booking Tracker", "arrival_rank", ("column", "Booking Tracker",
                                         "arr_key")),
    # ---- guests ----------------------------------------------------------
    ("Guest Database", "stays", ("sheet", "Booking Tracker")),
    ("Guest Database", "nights", ("sheet", "Booking Tracker")),
    ("Guest Database", "repeat", ("column", "Guest Database", "stays")),
    ("Guest Database", "net", ("sheet", "Booking Tracker")),
    ("Guest Database", "first_stay", ("sheet", "Booking Tracker")),
    # ---- arrivals --------------------------------------------------------
    ("Check-in & Check-out", "property_name", ("sheet",
                                               "Booking Tracker")),
    ("Check-in & Check-out", "arrival_date", ("sheet",
                                              "Booking Tracker")),
    ("Check-in & Check-out", "steps_done", ("column",
                                            "Check-in & Check-out",
                                            "sent_directions")),
    ("Check-in & Check-out", "readiness", ("column",
                                           "Check-in & Check-out",
                                           "sent_directions")),
    # ---- cleaning: generated from the check-outs -------------------------
    ("Cleaning Schedule", "checkout_date", ("sheet", "Booking Tracker")),
    ("Cleaning Schedule", "clean_date", ("column", "Cleaning Schedule",
                                         "checkout_date")),
    ("Cleaning Schedule", "cleaner_name", ("sheet", "Staff & Vendors")),
    ("Cleaning Schedule", "hours", ("column", "Cleaning Schedule",
                                    "start_time")),
    ("Cleaning Schedule", "completion", ("column", "Cleaning Schedule",
                                         "task_bedroom")),
    ("Cleaning Schedule", "next_arrival", ("sheet", "Booking Tracker")),
    ("Cleaning Schedule", "window_hours", ("sheet",
                                           "Property Database")),
    ("Cleaning Schedule", "total_cost", ("text", "LaundryPerTurn")),
    ("Cleaning Schedule", "clean_cost", ("sheet", "Property Database")),
    ("Cleaning Schedule", "cost_gap", ("column", "Cleaning Schedule",
                                       "fee_charged")),
    # ---- payouts: assumed fees against real ones -------------------------
    ("Platform Revenue & Fees", "exp_net", ("sheet", "Booking Tracker")),
    ("Platform Revenue & Fees", "received", ("column",
                                             "Platform Revenue & Fees",
                                             "amount_in")),
    ("Platform Revenue & Fees", "variance", ("column",
                                             "Platform Revenue & Fees",
                                             "exp_net")),
    ("Platform Revenue & Fees", "alert", ("text", "PayoutChaseDays")),
    # ---- costs -----------------------------------------------------------
    ("Expense Tracker", "is_cleaning", ("text", "Cleaning")),
    ("Expense Tracker", "in_pl", ("text", "CleaningInExpenses")),
    ("Expense Tracker", "pl_amount", ("column", "Expense Tracker",
                                      "in_pl")),
    ("Expense Tracker", "monthly_equiv", ("column", "Expense Tracker",
                                          "frequency")),
    ("Supplies Inventory", "turns_month", ("sheet", "Cleaning Schedule")),
    ("Supplies Inventory", "weeks_cover", ("column", "Supplies Inventory",
                                           "usage_month")),
    ("Supplies Inventory", "status", ("column", "Supplies Inventory",
                                      "min_level")),
    ("Supplies Inventory", "needs_order", ("text", "RestockWeeks")),
    ("Supplies Inventory", "alert", ("column", "Supplies Inventory",
                                     "needs_order")),
    ("Maintenance Log", "vendor_name", ("sheet", "Staff & Vendors")),
    ("Maintenance Log", "cost_used", ("column", "Maintenance Log",
                                      "act_cost")),
    ("Maintenance Log", "is_open", ("column", "Maintenance Log",
                                    "status")),
    ("Maintenance Log", "alert", ("text", "MaintenanceWarnDays")),
    # ---- reviews and vendors ---------------------------------------------
    ("Reviews Tracker", "property", ("sheet", "Booking Tracker")),
    ("Reviews Tracker", "detail_avg", ("column", "Reviews Tracker",
                                       "rate_clean")),
    ("Reviews Tracker", "below_target", ("text", "TargetRating")),
    ("Reviews Tracker", "alert", ("text", "ReviewChaseDays")),
    ("Staff & Vendors", "cleans", ("sheet", "Cleaning Schedule")),
    ("Staff & Vendors", "jobs", ("sheet", "Maintenance Log")),
    ("Staff & Vendors", "paid_total", ("column", "Staff & Vendors",
                                       "paid_cleaning")),
    # ---- the investor tab -------------------------------------------------
    ("Occupancy & Performance", "nights_booked", ("sheet",
                                                  "Booking Tracker")),
    ("Occupancy & Performance", "occupancy", ("column",
                                              "Occupancy & Performance",
                                              "nights_available")),
    ("Occupancy & Performance", "adr", ("column",
                                        "Occupancy & Performance",
                                        "accommodation")),
    ("Occupancy & Performance", "revpar", ("column",
                                           "Occupancy & Performance",
                                           "accommodation")),
    ("Occupancy & Performance", "profit", ("column",
                                           "Occupancy & Performance",
                                           "total_cost")),
    ("Occupancy & Performance", "clean_per_booking",
     ("column", "Occupancy & Performance", "cleaning_cost")),
]

#: cells reached by a defined name that have to hold a live formula -
#: the summary pages are where the value of the file lives, and a name
#: pointing at an empty cell means a page quietly stopped working.
PAGE_FORMULAS = [
    ("TodayDate", "settings: the live date every alert reads"),
    ("ThisMonthKey", "settings: the month key the totals group on"),
    ("PropertyCount", "settings: properties on the books"),
    ("ActiveProperties", "settings: properties currently letting"),
    ("CalStart", "calendar: the first of the month being shown"),
    ("CalGridStart", "calendar: the Monday the grid starts on"),
    ("TaxYearStart", "tax summary: the first day of the tax year"),
    ("TaxYearEnd", "tax summary: the last day of the tax year"),
    ("PrevMonthKey", "monthly reports: last month, for comparison"),
    ("LastYearMonthKey", "monthly reports: the same month last year"),
]

#: cells a buyer has to be able to type in
PAGE_INPUTS = [
    ("HostName", "settings: the business name"),
    ("OwnerName", "settings: the host or owner"),
    ("CurrencySign", "settings: the currency symbol"),
    ("TrackingStart", "settings: the day tracking began"),
    ("DefaultNightly", "settings: the standard nightly rate"),
    ("DefaultCleaningFee", "settings: the cleaning fee charged"),
    ("DefaultCleaningCost", "settings: what a turnover costs you"),
    ("ExtraGuestAfter", "settings: guests included in the rate"),
    ("ExtraGuestFee", "settings: the extra guest charge"),
    ("PetFee", "settings: the pet fee"),
    ("DefaultDeposit", "settings: the security deposit"),
    ("DefaultMinNights", "settings: the minimum stay"),
    ("ManagementFee", "settings: the co-host share"),
    ("TaxLabel", "settings: what the lodging tax is called"),
    ("TaxRate", "settings: the lodging tax rate"),
    ("TaxCollectedByPlatform", "settings: who remits the lodging tax"),
    ("LaundryPerTurn", "settings: laundry cost a turnover"),
    ("SuppliesPerTurn", "settings: consumables a turnover"),
    ("WelcomePerTurn", "settings: the welcome basket"),
    ("CleanerRate", "settings: the default cleaner rate"),
    ("TightTurnHours", "settings: what counts as a tight turnaround"),
    ("CleaningInExpenses", "settings: does cleaning count twice"),
    ("MaintenanceInExpenses", "settings: does maintenance count twice"),
    ("TargetOccupancy", "settings: the occupancy you are aiming at"),
    ("TargetADR", "settings: the rate you are aiming at"),
    ("TargetRevPAR", "settings: the RevPAR you are aiming at"),
    ("TargetMargin", "settings: the margin you are aiming at"),
    ("TargetRating", "settings: the review score you are aiming at"),
    ("TargetRevenueMonth", "settings: the monthly revenue target"),
    ("TargetRevenueYear", "settings: the annual revenue target"),
    ("TargetProfitMonth", "settings: the monthly profit target"),
    ("UpcomingDays", "settings: the arriving-soon window"),
    ("ReviewChaseDays", "settings: when to chase a review"),
    ("PayoutChaseDays", "settings: when to chase a payout"),
    ("MaintenanceWarnDays", "settings: when a repair starts warning"),
    ("RestockWeeks", "settings: when to reorder a supply"),
    ("LongStayNights", "settings: what counts as a long stay"),
    ("PropertyPrefix", "settings: the property ID prefix"),
    ("BookingPrefix", "settings: the booking ID prefix"),
    ("CalProperty", "calendar: the property being shown"),
    ("CalMonth", "calendar: the month being shown"),
    ("JobId", "turnover checklist: the job being printed"),
    ("TaxYear", "tax summary: the year being reported on"),
    ("ReportMonth", "monthly reports: the month being reported on"),
    ("ReportYear", "annual reports: the year being reported on"),
]

#: every drop down on the workbook reads one of these - a missing one
#: means a blank list and a column nobody can fill in
REQUIRED_LIST_NAMES = [
    "PropertyList", "PropertyNameList", "BookingList", "GuestList",
    "VendorList", "VendorNameList", "SupplyList", "PlatformList",
    "PlatformCommission", "PlatformPayment", "BookingStatusList",
    "PaymentStatusList", "PropertyTypeList", "PropertyStatusList",
    "OwnershipList", "CleaningStatusList", "MaintenanceCategoryList",
    "MaintenanceStatusList", "PriorityList", "ExpenseCategoryList",
    "TaxCategoryList", "PaymentMethodList", "FrequencyList",
    "SupplyCategoryList", "StockStatusList", "StaffRoleList",
    "PayBasisList", "ContractStatusList", "AvailabilityList",
    "ReviewStatusList", "GuestTypeList", "TripPurposeList",
    "PayoutStatusList", "YesNoList", "MonthKeyList", "CountryList",
]


def _name_target(wb, name):
    """The one cell a defined name points at, or None."""
    defn = wb.defined_names.get(name)
    if defn is None:
        for ws in wb.worksheets:
            if name in ws.defined_names:
                defn = ws.defined_names[name]
                break
    if defn is None:
        return None
    try:
        sheet, ref = next(defn.destinations)
    except (StopIteration, Exception):          # noqa: B902
        return None
    if sheet not in wb.sheetnames or ":" in ref:
        return None
    return wb[sheet][ref.replace("$", "")]


def check_pages(wb, problems: list, notes: list) -> None:
    """The summary pages, checked through their defined names."""
    live = 0
    for name, what in PAGE_FORMULAS:
        cell = _name_target(wb, name)
        if cell is None:
            problems.append(f"{name} is missing - {what}")
            continue
        value = cell.value
        if not (isinstance(value, str) and value.startswith("=")):
            problems.append(f"{name} no longer holds a formula - {what}")
            continue
        if cell.protection and not cell.protection.locked:
            problems.append(f"{name} is a formula but is unlocked - {what}")
            continue
        live += 1
    typed = 0
    for name, what in PAGE_INPUTS:
        cell = _name_target(wb, name)
        if cell is None:
            problems.append(f"{name} is missing - {what}")
            continue
        if cell.protection and cell.protection.locked:
            problems.append(f"{name} is locked, so nobody can change it "
                            f"- {what}")
            continue
        typed += 1
    notes.append(f"summary pages: {live} live formulas, "
                 f"{typed} typeable inputs verified")


def _specs():
    """Every register specification, keyed by sheet name."""
    from airbnb_rental_spreadsheet.build import REGISTER_MODULES
    from airbnb_rental_spreadsheet.config import Config
    cfg = Config()
    return {module.NAME: module.spec(cfg) for module in REGISTER_MODULES}


def check_chain(wb, problems: list, notes: list) -> None:
    """The enquiry-to-enrolment-to-invoice chain only holds while these do.

    Every one of them is a formula in row 5 of a register, so the check
    is cheap: find the column by its key in the specification, read the
    formula, and make sure it still mentions the thing it depends on.
    """
    specs = _specs()
    checked = 0
    for sheet_name, key, expect in CHAIN_CHECKS:
        spec = specs.get(sheet_name)
        if spec is None or sheet_name not in wb.sheetnames:
            problems.append(f"missing register {sheet_name}")
            continue
        try:
            column = spec.index(key) + 1
        except Exception:
            problems.append(f"{sheet_name}: column '{key}' has gone")
            continue
        formula = wb[sheet_name].cell(row=5, column=column).value
        if not isinstance(formula, str) or not formula.startswith("="):
            problems.append(f"{sheet_name}.{key} is no longer a formula")
            continue
        if expect[0] == "sheet":
            found = expect[1] in formula
            label = f"the {expect[1]} tab"
        elif expect[0] == "text":
            found = expect[1] in formula
            label = expect[1]
        else:
            letter = specs[expect[1]].letter(expect[2])
            label = f"{expect[1]}.{expect[2]} (column {letter})"
            if expect[1] == sheet_name:
                # a link inside the same register is a plain A1 reference
                found = bool(re.search(r"(?<![A-Z$])\$?" + letter +
                                       r"\$?5(?![0-9])", formula))
            else:
                found = f"${letter}$" in formula
        if not found:
            problems.append(
                f"{sheet_name}.{key} no longer points at {label} - "
                f"the automation chain is broken")
        checked += 1
    notes.append(f"automation chain: {checked} links verified")


def check_platform_table(wb, problems: list, notes: list) -> None:
    """The platform fee table a buyer must be able to type over.

    This is the single most important editable thing in the product:
    commission varies by platform, by market, by account type and by
    negotiation, so a workbook that locks these cells is wrong for
    almost everybody who buys it.
    """
    ranges = ("PlatformList", "PlatformCommission", "PlatformPayment",
              "PlatformIsOTA")
    found = 0
    for name in ranges:
        defn = wb.defined_names.get(name)
        if defn is None:
            problems.append(f"the editable platform fee table has lost "
                            f"{name}")
            continue
        try:
            _sheet, ref = next(defn.destinations)
        except Exception:                        # noqa: B902
            problems.append(f"{name} does not point anywhere")
            continue
        if ":" not in ref:
            problems.append(f"{name} should be a range, got {ref}")
            continue
        found += 1

    locked = 0
    cells = 0
    for name in ("PlatformCommission", "PlatformPayment"):
        defn = wb.defined_names.get(name)
        if defn is None:
            continue
        sheet, ref = next(defn.destinations)
        for row in wb[sheet][ref.replace("$", "")]:
            for cell in row:
                cells += 1
                if cell.protection and cell.protection.locked:
                    locked += 1
    if locked:
        problems.append(
            f"{locked} platform fee cells are locked - commission "
            f"rates vary by platform, market and account, so buyers "
            f"have to be able to type their own")
    notes.append(f"platform fee table: {found}/4 named ranges, "
                 f"{cells} editable rate cells"
                 if found else "platform fee table: NOT FOUND")


def check_promises(wb, problems: list, notes: list) -> None:
    """Three promises the product makes in writing.

    1. ADR and RevPAR use accommodation revenue only - counting the
       cleaning fee inflates both and is the single most common error
       in rental spreadsheets sold online.
    2. the workbook never claims to synchronise with a platform.
    3. the tax page carries a jurisdiction disclaimer.
    """
    from airbnb_rental_spreadsheet.build import REGISTER_MODULES
    from airbnb_rental_spreadsheet.config import Config
    cfg = Config()
    specs = {m.NAME: m.spec(cfg) for m in REGISTER_MODULES}

    for sheet_name, keys in (("Property Database", ("adr", "revpar")),
                             ("Occupancy & Performance",
                              ("adr", "revpar"))):
        spec = specs.get(sheet_name)
        if spec is None:
            continue
        for key in keys:
            try:
                column = spec.index(key) + 1
            except Exception:                    # noqa: B902
                problems.append(f"{sheet_name}.{key} has gone")
                continue
            formula = wb[sheet_name].cell(row=5, column=column).value
            if not isinstance(formula, str):
                continue
            if "cleaning" in formula.lower():
                problems.append(
                    f"{sheet_name}.{key} mentions cleaning - ADR and "
                    f"RevPAR must use accommodation revenue only")

    banned = ("sync", "synchronise", "synchronize", "live feed",
              "api connection", "automatically imports")
    hits = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if not isinstance(value, str) or value.startswith("="):
                    continue
                low = value.lower()
                for phrase in banned:
                    if phrase in low and "not" not in low \
                            and "cannot" not in low:
                        hits += 1
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: claims a "
                            f"platform connection -> {value[:70]}")
    disclaimer = False
    for ws in (wb["Tax Summary"], wb["Settings & Lists"],
               wb["Getting Started"]):
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if isinstance(value, str) and "jurisdiction" in \
                        value.lower():
                    disclaimer = True
    if not disclaimer:
        problems.append("no tax disclaimer mentioning jurisdiction")
    notes.append(f"product promises: ADR/RevPAR exclude cleaning, "
                 f"{hits} false sync claims, tax disclaimer "
                 f"{'present' if disclaimer else 'MISSING'}")


def check(path: str) -> int:
    problems: list = list(check_raw_xml(path))
    notes: list = []
    if problems:
        print(f"--- {path}")
        for problem in problems[:40]:
            print(f"    - {problem}")
        return 1
    wb = openpyxl.load_workbook(path, data_only=False)

    # ---------------------------------------------------------- structure
    names = wb.sheetnames
    if names != EXPECTED_SHEETS:
        problems.append(f"sheet order/={names}")
    hidden = [ws.title for ws in wb.worksheets if ws.sheet_state != "visible"]
    if hidden:
        problems.append(f"no tab should be hidden, got {hidden}")

    check_chain(wb, problems, notes)
    check_pages(wb, problems, notes)
    check_platform_table(wb, problems, notes)
    check_promises(wb, problems, notes)

    # --------------------------------------------------------- properties
    props = wb.properties
    if props.creator != "Novality Store":
        problems.append(f"creator is {props.creator!r}")
    if props.lastModifiedBy != "Novality Store":
        problems.append(f"lastModifiedBy is {props.lastModifiedBy!r}")

    # --------------------------------------------------------- protection
    for ws in wb.worksheets:
        prot = ws.protection
        if not prot.sheet:
            problems.append(f"{ws.title}: sheet protection off")
        elif not prot.password:
            problems.append(f"{ws.title}: protected without a password")

    # -------------------------------------------------------- definednames
    defined = set(wb.defined_names)
    for ws in wb.worksheets:
        defined |= set(ws.defined_names)
    for name in REQUIRED_LIST_NAMES:
        if name not in defined:
            problems.append(f"drop down list {name} is missing")
    notes.append(f"{len(defined)} defined names")

    # ------------------------------------------------------------ formulas
    used_names: Counter = Counter()
    unknown_functions: Counter = Counter()
    bare_refs: dict = {}
    bad_functions: Counter = Counter()
    formula_cells = 0
    locked_formula_problem = 0
    unlocked_input_cells = 0
    per_sheet = defaultdict(int)

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if isinstance(value, str) and value.startswith("="):
                    formula_cells += 1
                    per_sheet[ws.title] += 1
                    if "[[" in value or "{" in value or "}" in value:
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: unresolved "
                            f"template token -> {value[:90]}")
                    if value != value.strip():
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: formula has "
                            f"stray whitespace around it, which some "
                            f"calculation engines refuse to parse -> "
                            f"{value[:70]!r}")
                    if "#REF" in value:
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: #REF -> "
                            f"{value[:90]}")
                    where = f"{ws.title}!{cell.coordinate}"
                    if (cell.protection and not cell.protection.locked
                            and where not in SEEDED_INPUTS):
                        locked_formula_problem += 1
                        if locked_formula_problem < 25:
                            problems.append(
                                f"{where}: formula cell is NOT locked")
                    stripped = re.sub(r'"[^"]*"', '""', value)
                    stripped = re.sub(r"'[^']*'!", "SHEET!", stripped)
                    if value.count('"') % 2:
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: odd number of "
                            f"quotes -> {value[:80]}")
                    depth = paren_depth(value)
                    if depth:
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: unbalanced "
                            f"brackets ({depth:+d}) -> {value[:80]}")
                    slip = precedence_slip(value)
                    if slip:
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: a+b/c where "
                            f"(a+b)/c was meant -> {slip[:90]}")
                    collapsed = collapsed_criteria(value)
                    if collapsed:
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: a criteria "
                            f"range and its test were written as one "
                            f"argument, Excel answers #VALUE! -> "
                            f"{collapsed[:90]}")
                    ragged = ragged_criteria(value)
                    if ragged:
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: criteria "
                            f"ranges are different sizes, Excel answers "
                            f"#VALUE! -> {ragged[:90]}")
                    if ",)" in stripped or "(," in stripped or \
                            ",," in stripped:
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: empty argument "
                            f"-> {value[:80]}")
                    for func in FUNCTION_RE.findall(stripped):
                        if func in MODERN_ONLY:
                            bad_functions[func] += 1
                        elif func not in KNOWN_FUNCTIONS:
                            unknown_functions[func] += 1
                    for token in NAME_RE.findall(stripped):
                        if token in KNOWN_FUNCTIONS:
                            continue
                        used_names[token] += 1
                    for token in BARE_COL_RE.findall(stripped):
                        if token in KNOWN_FUNCTIONS:
                            continue
                        bare_refs[f"{ws.title}!{cell.coordinate} {token}"] = \
                            value[:90]
                elif value is not None and cell.protection and \
                        not cell.protection.locked:
                    unlocked_input_cells += 1

    # sheet names used inside formulas must exist
    sheet_refs = Counter()
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if isinstance(value, str) and value.startswith("="):
                    body = re.sub(r'"[^"]*"', '""', value)
                    for ref in re.findall(r"'([^']+)'!", body):
                        sheet_refs[ref] += 1
                    for ref in re.findall(r"(?<![A-Za-z0-9_'])"
                                          r"([A-Za-z][A-Za-z0-9_]*)!", body):
                        sheet_refs[ref] += 1
    for ref, count in sheet_refs.items():
        if ref not in names:
            problems.append(f"formula points at missing sheet "
                            f"'{ref}' ({count}x)")
    notes.append("cross sheet references: " + ", ".join(
        f"{k} {v}" for k, v in sorted(sheet_refs.items(),
                                      key=lambda kv: -kv[1])[:6]))

    for where, formula in list(bare_refs.items())[:15]:
        if where.rsplit(" ", 1)[-1] not in defined:
            problems.append(f"{where}: column letter with no row "
                            f"number -> {formula}")
    for func, count in unknown_functions.items():
        problems.append(f"unknown function {func}() used {count}x")
    for func, count in bad_functions.items():
        problems.append(f"{func}() needs a modern Excel - used {count}x")

    missing = {}
    for token, count in used_names.items():
        if token in defined:
            continue
        if token in KNOWN_FUNCTIONS or token.upper() in KNOWN_FUNCTIONS:
            continue
        if re.match(r"^[A-Z]{1,3}\d+$", token):
            continue
        missing[token] = count
    for token, count in sorted(missing.items(), key=lambda kv: -kv[1])[:25]:
        problems.append(f"undefined name used {count}x: {token}")

    # ----------------------------------------------------- data validation
    for ws in wb.worksheets:
        for dv in ws.data_validations.dataValidation:
            formula = (dv.formula1 or "").lstrip("=")
            if dv.type == "list" and formula and \
                    re.match(r"^[A-Za-z_][A-Za-z0-9_.]*$", formula):
                if formula not in defined:
                    problems.append(
                        f"{ws.title}: drop down uses missing name "
                        f"{formula}")

    # ------------------------------------------------------------- report
    with zipfile.ZipFile(path) as zf:
        size = sum(i.file_size for i in zf.infolist())
    notes.append(f"{formula_cells:,} formula cells, "
                 f"{unlocked_input_cells:,} unlocked value cells")
    notes.append(f"uncompressed xml {size / 1_000_000:.1f} MB")

    print(f"--- {path}")
    for note in notes:
        print(f"  . {note}")
    print("  . formulas per sheet: " + ", ".join(
        f"{k} {v}" for k, v in sorted(per_sheet.items(),
                                      key=lambda kv: -kv[1])[:10]))
    if problems:
        print(f"  ! {len(problems)} problem(s):")
        for problem in problems[:60]:
            print(f"    - {problem}")
        return 1
    print("  OK - no problems found")
    return 0


if __name__ == "__main__":
    target = (sys.argv[1] if len(sys.argv) > 1
              else "output/Airbnb & Short-Term Rental Management Spreadsheet.xlsx")
    raise SystemExit(check(target))
````

### `tools/check_parens.py`

*121 lines · 4 functions*

> Parenthesis and quote balance check for every formula in a workbook.
>
> Concatenated nested IFs are the one mistake a generator makes that no
> amount of reading catches, so this walks the saved .xlsx and counts
> brackets and quotes on every formula cell.
>
>     python3 tools/check_parens.py output/Workbook.xlsx

| Defines | What it does |
|---------|--------------|
| `sheet_names(book)` | target path -> sheet name, in workbook order |
| `balanced(formula)` | Is the formula balanced for brackets and double quotes? |
| `suspicious(formula)` | Cheap smells that are always bugs in a generated workbook |
| `check(path)` |  |

**Module constants:** `NS` ('{http://schemas.openxmlformats.org/spreadsheetml/2006/ma...).

````python
#!/usr/bin/env python3
"""Parenthesis and quote balance check for every formula in a workbook.

Concatenated nested IFs are the one mistake a generator makes that no
amount of reading catches, so this walks the saved .xlsx and counts
brackets and quotes on every formula cell.

    python3 tools/check_parens.py output/Workbook.xlsx
"""

from __future__ import annotations

import re
import sys
import zipfile
from typing import Dict, Iterable, List, Tuple
from xml.etree import ElementTree as ET

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def sheet_names(book: zipfile.ZipFile) -> Dict[str, str]:
    """target path -> sheet name, in workbook order."""
    rels = ET.fromstring(book.read("xl/_rels/workbook.xml.rels"))
    by_id = {
        rel.get("Id"): rel.get("Target").lstrip("/")
        for rel in rels
    }
    root = ET.fromstring(book.read("xl/workbook.xml"))
    names: Dict[str, str] = {}
    for sheet in root.iter(f"{NS}sheet"):
        rid = sheet.get(
            "{http://schemas.openxmlformats.org/officeDocument/"
            "2006/relationships}id")
        target = by_id.get(rid, "")
        if not target.startswith("xl/"):
            target = "xl/" + target
        names[target] = sheet.get("name", "?")
    return names


def balanced(formula: str) -> Tuple[bool, str]:
    """Is the formula balanced for brackets and double quotes?"""
    depth = 0
    in_string = False
    index = 0
    while index < len(formula):
        char = formula[index]
        if char == '"':
            if in_string and index + 1 < len(formula) \
                    and formula[index + 1] == '"':
                index += 2
                continue
            in_string = not in_string
        elif not in_string:
            if char == "(":
                depth += 1
            elif char == ")":
                depth -= 1
                if depth < 0:
                    return False, "closes a bracket that never opened"
        index += 1
    if in_string:
        return False, "odd number of double quotes"
    if depth:
        return False, f"{depth} bracket(s) left open"
    return True, ""


def suspicious(formula: str) -> str:
    """Cheap smells that are always bugs in a generated workbook."""
    if "{" in formula or "}" in formula:
        return "unrendered template placeholder"
    if "[[" in formula or "]]" in formula:
        return "unrendered sheet reference"
    if re.search(r"#(REF|NAME|VALUE|DIV/0|N/A)", formula):
        return "hard coded error value"
    if ",," in formula.replace('",,"', ""):
        return "empty argument"
    return ""


def check(path: str) -> int:
    problems: List[str] = []
    total = 0
    with zipfile.ZipFile(path) as book:
        names = sheet_names(book)
        targets: Iterable[str] = [n for n in book.namelist()
                                  if n.startswith("xl/worksheets/sheet")]
        for target in sorted(targets):
            sheet = names.get(target, target)
            root = ET.fromstring(book.read(target))
            for cell in root.iter(f"{NS}c"):
                node = cell.find(f"{NS}f")
                if node is None or not node.text:
                    continue
                total += 1
                formula = node.text
                ok, why = balanced(formula)
                smell = suspicious(formula)
                if not ok or smell:
                    ref = cell.get("r", "?")
                    problems.append(
                        f"{sheet}!{ref}: {why or smell}\n    ={formula[:160]}")
    print(f"{path}: {total:,} formulas checked")
    if problems:
        print(f"  {len(problems)} problem(s):")
        for line in problems[:40]:
            print("  - " + line)
        if len(problems) > 40:
            print(f"  ... and {len(problems) - 40} more")
        return 1
    print("  all balanced, no placeholders left behind")
    return 0


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        raise SystemExit(2)
    raise SystemExit(max(check(arg) for arg in sys.argv[1:]))
````

### `tools/check_airbnb_sample.py`

*96 lines · 1 function*

> Validate the Airbnb product's example data before it is built.
>
> Two checks, both of which have caught real bugs in this repo:
>
> 1. every key in a sample row is an *input* column on that register -
>    writing a sample value into a formula column silently destroys the
>    formula for that row;
> 2. every value that feeds a drop-down exists in that drop-down's list -
>    otherwise Excel shows the cell with a validation error the moment a
>    buyer clicks it.
>
>     python3 tools/check_airbnb_sample.py

| Defines | What it does |
|---------|--------------|
| `main()` |  |

**Module constants:** `TABLE_LISTS` (1 entries), `ID_LISTS` ({'PropertyList', 'BookingList', 'GuestList', 'VendorList'...), `ID_PATTERN` (re.compile(...)).

````python
#!/usr/bin/env python3
"""Validate the Airbnb product's example data before it is built.

Two checks, both of which have caught real bugs in this repo:

1. every key in a sample row is an *input* column on that register -
   writing a sample value into a formula column silently destroys the
   formula for that row;
2. every value that feeds a drop-down exists in that drop-down's list -
   otherwise Excel shows the cell with a validation error the moment a
   buyer clicks it.

    python3 tools/check_airbnb_sample.py
"""

from __future__ import annotations

import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from airbnb_rental_spreadsheet.config import Config          # noqa: E402
from airbnb_rental_spreadsheet.build import REGISTER_MODULES  # noqa: E402
from airbnb_rental_spreadsheet.sheets.settings import LISTS   # noqa: E402

#: lists that are built as a table on Settings rather than a plain column
TABLE_LISTS = {
    "PlatformList": lambda cfg: [name for name, *_ in cfg.platform_table],
}

#: names that resolve to a live column of IDs on another register
ID_LISTS = {"PropertyList", "BookingList", "GuestList", "VendorList",
            "SupplyList", "CleanTaskList", "MaintenanceList",
            "PropertyNameList", "GuestNameList", "VendorNameList",
            "SupplyNameList"}

ID_PATTERN = re.compile(r"^[A-Z]{3}-\d+$")


def main() -> int:
    cfg = Config()
    lists = {name: set(values) for name, values in LISTS}
    for name, builder in TABLE_LISTS.items():
        lists[name] = set(builder(cfg))

    problems = 0
    total_rows = 0
    for module in REGISTER_MODULES:
        spec = module.spec(cfg)
        inputs = {c.key for c in spec.columns if not c.formula}
        by_key = {c.key: c for c in spec.columns}
        rows = spec.samples
        total_rows += len(rows)

        if len(rows) > spec.rows:
            problems += 1
            print(f"{spec.name}: {len(rows)} sample rows will not fit "
                  f"in {spec.rows}")

        for index, row in enumerate(rows):
            extra = sorted(set(row) - inputs)
            if extra:
                problems += 1
                print(f"{spec.name} row {index}: not input columns "
                      f"-> {extra}")
            for key, value in row.items():
                column = by_key.get(key)
                if column is None or not column.validate:
                    continue
                source = column.validate.get("source")
                if not isinstance(source, str) or not source.startswith("="):
                    continue
                name = source.lstrip("=")
                if name in lists:
                    if value != "" and value not in lists[name]:
                        problems += 1
                        print(f"{spec.name}.{key} row {index}: "
                              f"{value!r} is not in {name}")
                elif name in ID_LISTS:
                    if value != "" and not ID_PATTERN.match(str(value)):
                        problems += 1
                        print(f"{spec.name}.{key} row {index}: "
                              f"{value!r} does not look like an ID "
                              f"({name})")

        print(f"  {spec.name:26s} {len(rows):4d} rows, "
              f"{len(inputs)} input columns")

    print(f"\n{total_rows} sample rows checked, {problems} problems")
    return 1 if problems else 0


if __name__ == "__main__":
    raise SystemExit(main())
````

### `tools/recalc_check.py`

*216 lines · 4 functions*

> Really calculate every formula in a generated workbook and report any
> cell that ends up as an Excel error (#REF!, #DIV/0!, #VALUE! ...).
>
> `validate_workbook.py` reads the file; this one *runs* it.  It loads the
> workbook into the `formulas` calculation engine, evaluates the whole
> dependency graph and then looks at the answer in every formula cell, so a
> formula that parses perfectly but divides by zero on an empty row is still
> caught.
>
>     pip install formulas openpyxl
>     python tools/recalc_check.py "output/Pet Sitting Business Spreadsheet.xlsx"
>
> Notes
> -----
> * The engine cannot resolve `OFFSET()` based defined names, so a working
>   copy of the file is made first with those names flattened to plain
>   ranges.  The copy is a test artefact - the real workbook keeps its
>   dynamic names and is never touched.
> * Only cells that actually hold a formula are judged, so help text that
>   merely mentions "#N/A" is not mistaken for a broken formula.
> * The engine is hungry: a full-size 400-row build is around 50,000
>   formulas and wants several GB of memory.  Check a smaller build
>   instead - the formulas are identical, there are just fewer rows of
>   them:
>
>       python build_pet_sitting_workbook.py --rows 120 -o /tmp/small.xlsx
>       python tools/recalc_check.py /tmp/small.xlsx
>
>   Pass --force to run it on a full-size file anyway.

| Defines | What it does |
|---------|--------------|
| `flatten_dynamic_names(src, dst)` | Save a copy of *src* at *dst* with OFFSET names turned into ranges |
| `formula_cells(path)` | `{(SHEET NAME upper, coordinate)}` for every cell holding a formula |
| `cell_text(value)` |  |
| `main(argv)` |  |

**Module constants:** `BIG` (25000), `ERROR_TEXT` (re.compile(...)), `DYNAMIC_NAME` (re.compile(...)).

````python
"""Really calculate every formula in a generated workbook and report any
cell that ends up as an Excel error (#REF!, #DIV/0!, #VALUE! ...).

`validate_workbook.py` reads the file; this one *runs* it.  It loads the
workbook into the `formulas` calculation engine, evaluates the whole
dependency graph and then looks at the answer in every formula cell, so a
formula that parses perfectly but divides by zero on an empty row is still
caught.

    pip install formulas openpyxl
    python tools/recalc_check.py "output/Pet Sitting Business Spreadsheet.xlsx"

Notes
-----
* The engine cannot resolve `OFFSET()` based defined names, so a working
  copy of the file is made first with those names flattened to plain
  ranges.  The copy is a test artefact - the real workbook keeps its
  dynamic names and is never touched.
* Only cells that actually hold a formula are judged, so help text that
  merely mentions "#N/A" is not mistaken for a broken formula.
* The engine is hungry: a full-size 400-row build is around 50,000
  formulas and wants several GB of memory.  Check a smaller build
  instead - the formulas are identical, there are just fewer rows of
  them:

      python build_pet_sitting_workbook.py --rows 120 -o /tmp/small.xlsx
      python tools/recalc_check.py /tmp/small.xlsx

  Pass --force to run it on a full-size file anyway.
"""

from __future__ import annotations

import argparse
import collections
import os
import re
import sys
import tempfile
import time

#: Formula count above which the engine needs more memory than a laptop
#: is likely to enjoy handing over.
BIG = 25000

ERROR_TEXT = re.compile(r"#(REF!|VALUE!|NAME\?|DIV/0!|NUM!|NULL!|N/A)")

# OFFSET(Sheet!$A$5,0,0,MAX(1,COUNTIF(Sheet!$A$5:$A$254,"?*")),1)
DYNAMIC_NAME = re.compile(
    r"^OFFSET\((?P<sheet>'[^']+'|[A-Za-z0-9_.]+)!"
    r"\$(?P<col>[A-Z]+)\$(?P<row>\d+),0,0,"
    r"MAX\(1,COUNTIF\([^,]+,\"\?\*\"\)\),1\)$")


def flatten_dynamic_names(src: str, dst: str) -> int:
    """Save a copy of *src* at *dst* with OFFSET names turned into ranges."""
    import warnings
    import openpyxl

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        book = openpyxl.load_workbook(src)
    changed = 0
    for _name, defined in list(book.defined_names.items()):
        match = DYNAMIC_NAME.match(defined.value or "")
        if not match:
            continue
        last = re.search(r":\$[A-Z]+\$(\d+)", defined.value)
        if not last:
            continue
        col = match.group("col")
        defined.value = (f"{match.group('sheet')}!"
                         f"${col}${match.group('row')}:${col}${last.group(1)}")
        changed += 1
    book.save(dst)
    return changed


def formula_cells(path: str) -> set:
    """`{(SHEET NAME upper, coordinate)}` for every cell holding a formula."""
    import warnings
    import openpyxl

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        book = openpyxl.load_workbook(path)
    cells = set()
    for sheet in book.worksheets:
        upper = sheet.title.upper()
        for row in sheet.iter_rows():
            for cell in row:
                value = cell.value
                if isinstance(value, str) and value.startswith("="):
                    cells.add((upper, cell.coordinate))
    return cells


def cell_text(value) -> str:
    try:
        return str(value.value[0, 0])
    except Exception:
        try:
            return str(value.value)
        except Exception:
            return ""


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("workbook")
    parser.add_argument("--list", type=int, default=12, metavar="N",
                        help="how many example cells to print per error type")
    parser.add_argument("--keep", action="store_true",
                        help="keep the flattened working copy")
    parser.add_argument("--show", default="", metavar="CELLS",
                        help="also print the calculated value of these "
                             "cells, e.g. \"Dashboard!A10,Reports!B12\"")
    parser.add_argument("--force", action="store_true",
                        help="check a big workbook even though the engine "
                             "will need several GB of memory")
    args = parser.parse_args(argv)

    path = args.workbook
    if not os.path.isfile(path):
        print(f"no such file: {path}")
        return 2
    try:
        import formulas
    except ImportError:
        print("this check needs the calculation engine: pip install formulas")
        return 2

    print(path, flush=True)
    started = time.time()
    folder = tempfile.mkdtemp(prefix="recalc_")
    working = os.path.join(folder, os.path.basename(path))
    count = flatten_dynamic_names(path, working)
    print(f"  . flattened {count} dynamic names into a working copy",
          flush=True)

    wanted = formula_cells(path)
    print(f"  . {len(wanted):,} formula cells to check", flush=True)
    if len(wanted) > BIG and not args.force:
        print("  ! that is a lot - the engine would need several GB of "
              "memory.\n"
              "    Check a smaller build instead (the formulas are the "
              "same, there\n"
              "    are just fewer rows of them):\n"
              "      python build_pet_sitting_workbook.py --rows 120 "
              "-o /tmp/small.xlsx\n"
              "      python tools/recalc_check.py /tmp/small.xlsx\n"
              "    or pass --force to go ahead anyway.")
        return 2

    model = formulas.ExcelModel().loads(working).finish()
    print(f"  . loaded in {time.time() - started:.0f}s", flush=True)
    clock = time.time()
    solution = model.calculate()
    print(f"  . calculated {len(solution):,} nodes in "
          f"{time.time() - clock:.0f}s", flush=True)

    book = os.path.basename(working).upper()
    counts = collections.Counter()
    examples = collections.defaultdict(list)
    seen = 0
    for key, value in solution.items():
        if "!" not in key or "]" not in key:
            continue
        sheet, _, coord = key.partition("!")
        sheet = sheet.split("]")[-1].rstrip("'").upper()
        if book and sheet.endswith(".XLSX"):
            continue
        coord = coord.replace("$", "")
        if (sheet, coord) not in wanted:
            continue
        seen += 1
        found = ERROR_TEXT.search(cell_text(value))
        if not found:
            continue
        token = f"{sheet} {found.group(0)}"
        counts[token] += 1
        if len(examples[token]) < args.list:
            examples[token].append(coord)

    if args.show:
        wanted_show = {}
        for item in args.show.split(","):
            item = item.strip()
            if not item:
                continue
            sheet, _, coord = item.rpartition("!")
            wanted_show[(sheet.strip("'").upper(),
                         coord.replace("$", "").upper())] = item
        print("  . values asked for:")
        for key, value in solution.items():
            if "!" not in key or "]" not in key:
                continue
            sheet, _, coord = key.partition("!")
            sheet = sheet.split("]")[-1].rstrip("'").upper()
            coord = coord.replace("$", "")
            label = wanted_show.get((sheet, coord))
            if label:
                print(f"    {label:28s} = {cell_text(value)[:70]}")

    print(f"  . {seen:,} of them evaluated by the engine", flush=True)
    if not counts:
        print("  OK - every formula worked out to a real answer")
        return 0
    print(f"  ! {sum(counts.values())} cell(s) ended up as an error:")
    for token, number in counts.most_common():
        print(f"    - {token} x{number}: {', '.join(examples[token])}")
    return 1


if __name__ == "__main__":
    sys.exit(main())
````

---

## Appendix A. Data dictionary

Every column of every register tab, straight from the specs that build them: the Excel column letter, the header a buyer sees, the key formulas use, the number format and whether the cell is typed in, chosen from a drop down or calculated. Pages (dashboards, reports, printables) are not listed - they have no fixed grid.

#### Property Database

Every place you let, what it costs to run and what it actually earns

63 columns, 30 rows of capacity, 31 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Property ID | `id` | id | formula |
| B | Property name | `name` | text | you type it |
| C | Short code | `nickname` | text | you type it |
| D | Address | `address` | text | you type it |
| E | City / area | `city` | text | you type it |
| F | Country | `country` | text | drop down `CountryList` |
| G | Property type | `type` | text | drop down `PropertyTypeList` |
| H | Bedrooms | `bedrooms` | int | you type it (checked) |
| I | Bathrooms | `bathrooms` | number1 | you type it (checked) |
| J | Beds | `beds` | int | you type it (checked) |
| K | Sleeps | `max_guests` | int | you type it (checked) |
| L | Status | `status` | status | drop down `PropertyStatusList` |
| M | Listed since | `listed_from` | date | you type it (checked) |
| N | Owner | `owner` | text | you type it |
| O | Managed by | `manager` | text | you type it |
| P | Ownership | `ownership` | text | drop down `OwnershipList` |
| Q | Main platform | `main_platform` | text | drop down `PlatformList` |
| R | Listing link | `listing_url` | link | you type it |
| S | Licence / permit | `licence` | text | you type it |
| T | Check-in | `check_in_time` | time | you type it (checked) |
| U | Check-out | `check_out_time` | time | you type it (checked) |
| V | Min nights | `min_nights` | int | you type it (checked) |
| W | Wi-Fi network | `wifi_name` | text | you type it |
| X | Access method | `access` | text | you type it |
| Y | Standard rate | `nightly_rate` | money | you type it (checked) |
| Z | Cleaning fee charged | `cleaning_fee` | money | you type it (checked) |
| AA | Cleaning cost paid | `cleaning_cost` | money | you type it (checked) |
| AB | Security deposit | `deposit` | money | you type it (checked) |
| AC | Monthly rent | `monthly_rent` | money | you type it (checked) |
| AD | Monthly mortgage | `monthly_mortgage` | money | you type it (checked) |
| AE | Purchase price | `purchase_price` | money0 | you type it (checked) |
| AF | Fixed cost a month | `fixed_monthly` | money | formula |
| AG | Bookings | `bookings` | int | formula |
| AH | Nights booked | `nights_booked` | int | formula |
| AI | Nights blocked | `nights_blocked` | int | formula |
| AJ | Guests hosted | `guests_hosted` | int | formula |
| AK | Cancellations | `cancellations` | int | formula |
| AL | Nights available | `nights_available` | int | formula |
| AM | Occupancy | `occupancy` | percent | formula |
| AN | ADR | `adr` | money | formula |
| AO | RevPAR | `revpar` | money | formula |
| AP | Avg stay | `alos` | number1 | formula |
| AQ | Rating | `rating` | number | formula |
| AR | Reviews | `reviews` | int | formula, hidden |
| AS | Room revenue | `accommodation` | money | formula |
| AT | Gross revenue | `gross` | money | formula |
| AU | Platform fees | `platform_fees` | money | formula |
| AV | Net revenue | `net_revenue` | money | formula |
| AW | Cleaning paid out | `cleaning_spend` | money | formula |
| AX | Maintenance | `maintenance_spend` | money | formula |
| AY | Other costs | `other_expenses` | money | formula |
| AZ | Total costs | `total_costs` | money | formula |
| BA | Operating profit | `profit` | money | formula |
| BB | Margin | `margin` | percent | formula |
| BC | Profit a night | `profit_night` | money | formula |
| BD | Gross yield | `yield_pct` | percent | formula, hidden |
| BE | Rank | `profit_rank` | int | formula |
| BF | Open jobs | `open_jobs` | int | formula |
| BG | Low stock | `low_stock` | int | formula |
| BH | Guest in tonight | `guest_in` | flag | formula |
| BI | Next arrival | `next_arrival` | date | formula |
| BJ | Action | `alert` | status | formula |
| BK | Notes | `notes` | wrap | you type it |

#### Booking Tracker

Every reservation, enquiry and block - priced, counted and checked for clashes

60 columns, 900 rows of capacity, 18 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Booking ID | `id` | id | formula |
| B | Property | `property` | text | drop down `PropertyList` |
| C | Property name | `property_name` | text | formula |
| D | Guest | `guest` | text | drop down `GuestList` |
| E | Guest name | `guest_name` | text | formula |
| F | Platform | `platform` | text | drop down `PlatformList` |
| G | Platform ref | `reference` | text | you type it |
| H | Status | `status` | status | drop down `BookingStatusList` |
| I | Booked on | `booked_date` | date | you type it (checked) |
| J | Check-in | `check_in` | date | you type it (checked) |
| K | Check-out | `check_out` | date | you type it (checked) |
| L | In | `ci_day` | text | formula, hidden |
| M | Out | `co_day` | text | formula, hidden |
| N | Nights | `nights` | int | formula |
| O | Check-in (number) | `ci_num` | number | formula, hidden |
| P | Check-out (number) | `co_num` | number | formula, hidden |
| Q | Row number | `seq_num` | int | formula, hidden |
| R | Month | `month` | text | formula, hidden |
| S | Year | `year` | int | formula, hidden |
| T | Season | `season` | text | formula, hidden |
| U | Lead time | `lead_time` | int | formula, hidden |
| V | Adults | `adults` | int | you type it (checked) |
| W | Children | `children` | int | you type it (checked) |
| X | Pets | `pets` | int | you type it (checked) |
| Y | Guests | `guests_total` | int | formula |
| Z | Over capacity | `over_capacity` | flag | formula, hidden |
| AA | Nightly rate | `rate_in` | money | you type it (checked) |
| AB | Rate used | `rate_used` | money | formula |
| AC | Room revenue | `accommodation` | money | formula |
| AD | Cleaning fee | `cleaning_in` | money | you type it (checked) |
| AE | Cleaning charged | `cleaning_fee` | money | formula |
| AF | Extra fees | `extra_in` | money | you type it (checked) |
| AG | Extra charged | `extra_fee` | money | formula |
| AH | Discount | `discount` | money | you type it (checked) |
| AI | Gross | `gross` | money | formula |
| AJ | Lodging tax | `tax` | money | formula, hidden |
| AK | Commission % | `fee_pct` | percent | formula, hidden |
| AL | Card fee % | `pay_pct` | percent | formula, hidden |
| AM | Platform fee | `platform_fee` | money | formula |
| AN | Card fee | `payment_fee` | money | formula |
| AO | Fees total | `platform_fee_total` | money | formula |
| AP | Net revenue | `net_revenue` | money | formula |
| AQ | Net a night | `net_night` | money | formula, hidden |
| AR | Payment | `payment_status` | status | drop down `PaymentStatusList` |
| AS | Holds dates | `holds` | flag | formula, hidden |
| AT | Earns | `earns` | flag | formula, hidden |
| AU | Is a block | `is_block` | flag | formula, hidden |
| AV | Cancelled | `is_cancelled` | flag | formula, hidden |
| AW | Stay length | `stay_band` | text | formula, hidden |
| AX | Clash | `clash` | int | formula, hidden |
| AY | Same-day turn | `same_day` | flag | formula |
| AZ | Cleaning booked | `has_cleaning` | flag | formula, hidden |
| BA | Arrival logged | `has_arrival` | flag | formula, hidden |
| BB | Review logged | `has_review` | flag | formula, hidden |
| BC | Where it is | `stay_state` | status | formula |
| BD | Arrival order key | `arr_key` | number | formula, hidden |
| BE | Arriving next | `arrival_rank` | int | formula, hidden |
| BF | Action | `alert` | status | formula |
| BG | Guest requests | `requests` | wrap | you type it |
| BH | Notes | `notes` | wrap | you type it |

#### Guest Database

Who has stayed with you, how often, and what they were worth after commission

29 columns, 400 rows of capacity, 13 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Guest ID | `id` | id | formula |
| B | Guest name | `name` | text | you type it |
| C | Email | `email` | text | you type it |
| D | Phone | `phone` | text | you type it |
| E | Country | `country` | text | drop down `CountryList` |
| F | Guest type | `guest_type` | text | drop down `GuestTypeList` |
| G | Usual trip | `purpose` | text | drop down `TripPurposeList` |
| H | Found you on | `first_platform` | text | drop down `PlatformList` |
| I | Happy to be contacted | `marketing_ok` | flag | drop down `YesNoList` |
| J | Stays | `stays` | int | formula |
| K | Cancelled | `cancellations` | int | formula, hidden |
| L | Nights | `nights` | int | formula |
| M | First stay | `first_stay` | date | formula |
| N | Last stay | `last_stay` | date | formula |
| O | Days since | `days_since` | int | formula |
| P | Repeat guest | `repeat` | flag | formula |
| Q | Lifetime gross | `gross` | money | formula |
| R | Lifetime net | `net` | money | formula |
| S | Avg per stay | `avg_stay_value` | money | formula |
| T | Avg nights | `nights_avg` | number1 | formula, hidden |
| U | Rank | `value_rank` | int | formula |
| V | Reviews left | `reviews` | int | formula, hidden |
| W | Rating given | `rating_given` | number | formula |
| X | Review status | `review_status` | status | drop down `ReviewStatusList` |
| Y | Preferences | `preferences` | wrap | you type it |
| Z | Standing requests | `requests` | wrap | you type it |
| AA | What they said | `feedback` | wrap | you type it |
| AB | Internal notes | `internal` | wrap | you type it |
| AC | Action | `alert` | status | formula |

#### Check-in & Check-out

The eight things every guest needs, and which of them you have actually done

32 columns, 900 rows of capacity, 17 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Arrival ID | `id` | id | formula |
| B | Booking | `booking` | text | drop down `BookingList` |
| C | Property | `property` | text | formula, hidden |
| D | Property | `property_name` | text | formula |
| E | Guest | `guest_name` | text | formula |
| F | Arrives | `arrival_date` | date | formula |
| G | Day | `arr_day` | text | formula, hidden |
| H | Leaves | `departure_date` | date | formula |
| I | Nights | `nights` | int | formula, hidden |
| J | Month | `month` | text | formula, hidden |
| K | Arrived at | `arrival_time` | time | you type it (checked) |
| L | Left at | `departure_time` | time | you type it (checked) |
| M | Late check-out agreed | `late_checkout` | flag | you type it (checked) |
| N | Early check-in agreed | `early_checkin` | flag | you type it (checked) |
| O | Access used | `key_method` | text | you type it |
| P | Directions sent | `sent_directions` | flag | you type it (checked) |
| Q | Access code sent | `sent_access` | flag | you type it (checked) |
| R | Wi-Fi details sent | `sent_wifi` | flag | you type it (checked) |
| S | House rules sent | `sent_rules` | flag | you type it (checked) |
| T | Welcome message sent | `sent_welcome` | flag | you type it (checked) |
| U | Pre-arrival score | `pre_arrival_done` | percent | formula |
| V | Check-out notes sent | `sent_checkout` | flag | you type it (checked) |
| W | Checked out | `checked_out` | flag | you type it (checked) |
| X | Property inspected | `inspected` | flag | you type it (checked) |
| Y | Condition photographed | `photographed` | flag | you type it (checked) |
| Z | Done | `steps_done` | int | formula, hidden |
| AA | Complete | `readiness` | percent | formula |
| AB | Still to do | `outstanding` | int | formula |
| AC | Where it is | `state` | status | formula |
| AD | Action | `alert` | status | formula |
| AE | Anything wrong | `issues` | wrap | you type it |
| AF | Notes | `notes` | wrap | you type it |

#### Cleaning Schedule

Every turnover: who is doing it, how long they have, and what it really costs

45 columns, 900 rows of capacity, 20 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Job ID | `id` | id | formula |
| B | Booking | `booking` | text | drop down `BookingList` |
| C | Property | `property` | text | formula, hidden |
| D | Property | `property_name` | text | formula |
| E | Guest leaving | `guest_name` | text | formula |
| F | Check-out | `checkout_date` | date | formula |
| G | Check-out (number) | `checkout_num` | number | formula, hidden |
| H | Cleaning date | `date_in` | date | you type it (checked) |
| I | Cleaned on | `clean_date` | date | formula |
| J | Day | `clean_day` | text | formula, hidden |
| K | Cleaned on (number) | `clean_num` | number | formula, hidden |
| L | Row number | `seq_num` | int | formula, hidden |
| M | Month | `month` | text | formula, hidden |
| N | Cleaner | `cleaner` | text | drop down `VendorList` |
| O | Cleaner name | `cleaner_name` | text | formula |
| P | Started | `start_time` | time | you type it (checked) |
| Q | Finished | `end_time` | time | you type it (checked) |
| R | Hours | `hours` | number1 | formula |
| S | Status | `status` | status | drop down `CleaningStatusList` |
| T | Bedrooms | `task_bedroom` | flag | you type it (checked) |
| U | Bathrooms | `task_bathroom` | flag | you type it (checked) |
| V | Kitchen | `task_kitchen` | flag | you type it (checked) |
| W | Living area | `task_living` | flag | you type it (checked) |
| X | Floors | `task_floors` | flag | you type it (checked) |
| Y | Laundry & linen | `task_laundry` | flag | you type it (checked) |
| Z | Rubbish out | `task_rubbish` | flag | you type it (checked) |
| AA | Restock supplies | `task_restock` | flag | you type it (checked) |
| AB | Final inspection | `task_inspection` | flag | you type it (checked) |
| AC | Tasks done | `tasks_done` | int | formula, hidden |
| AD | Complete | `completion` | percent | formula |
| AE | Next arrival | `next_arrival` | date | formula |
| AF | Days spare | `gap_days` | int | formula |
| AG | Same-day turn | `same_day` | flag | formula |
| AH | Hours in the window | `window_hours` | number1 | formula |
| AI | Ready for the next guest | `ready` | status | formula |
| AJ | Cleaning cost | `cost_in` | money | you type it (checked) |
| AK | Cost used | `clean_cost` | money | formula |
| AL | Laundry | `laundry_cost` | money | you type it (checked) |
| AM | Consumables | `supplies_cost` | money | you type it (checked) |
| AN | Turnover cost | `total_cost` | money | formula |
| AO | Fee charged to guest | `fee_charged` | money | formula, hidden |
| AP | Over or under recovered | `cost_gap` | money | formula |
| AQ | Cleaner paid | `paid` | flag | you type it (checked) |
| AR | Action | `alert` | status | formula |
| AS | Notes | `notes` | wrap | you type it |

#### Platform Revenue & Fees

What each platform said it would pay you, and what actually landed

27 columns, 900 rows of capacity, 9 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Payout ID | `id` | id | formula |
| B | Booking | `booking` | text | drop down `BookingList` |
| C | Property | `property` | text | formula, hidden |
| D | Property | `property_name` | text | formula |
| E | Guest | `guest_name` | text | formula |
| F | Platform | `platform` | text | formula |
| G | Check-out | `check_out` | date | formula, hidden |
| H | Gross booked | `exp_gross` | money | formula |
| I | Fees expected | `exp_fee` | money | formula |
| J | Payout expected | `exp_net` | money | formula |
| K | Paid on | `payout_date` | date | you type it (checked) |
| L | Month | `month` | text | formula, hidden |
| M | Year | `year` | int | formula, hidden |
| N | Amount received | `amount_in` | money | you type it (checked) |
| O | Fees actually taken | `fee_actual` | money | you type it (checked) |
| P | Refund given | `refund` | money | you type it (checked) |
| Q | Adjustment | `adjustment` | money | you type it (checked) |
| R | Net received | `received` | money | formula |
| S | Status | `status` | status | drop down `PayoutStatusList` |
| T | Statement ref | `reference` | text | you type it |
| U | Difference | `variance` | money | formula |
| V | Difference % | `variance_pct` | percent | formula, hidden |
| W | Real fee rate | `fee_rate_actual` | percent | formula |
| X | Days to pay | `days_to_pay` | int | formula |
| Y | Money in | `is_paid` | flag | formula, hidden |
| Z | Action | `alert` | status | formula |
| AA | Notes | `notes` | wrap | you type it |

#### Expense Tracker

Everything the properties cost you, counted once and shown per month

26 columns, 700 rows of capacity, 13 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Expense ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Month | `month` | text | formula, hidden |
| D | Year | `year` | int | formula, hidden |
| E | Day | `day` | text | formula, hidden |
| F | Property | `property` | text | drop down `PropertyList` |
| G | Property name | `property_name` | text | formula |
| H | Category | `category` | text | drop down `ExpenseCategoryList` |
| I | Paid to | `vendor` | text | you type it |
| J | What it was for | `description` | text | you type it |
| K | Amount | `amount` | money | you type it (checked) |
| L | of which tax | `tax_amount` | money | you type it (checked), hidden |
| M | Paid by | `method` | text | drop down `PaymentMethodList` |
| N | Receipt ref | `receipt` | text | you type it |
| O | How often | `frequency` | text | drop down `FrequencyList` |
| P | Per month | `monthly_equiv` | money | formula |
| Q | Booking | `booking` | text | drop down `BookingList` |
| R | From booking | `booking_prop` | text | formula, hidden |
| S | Tax treatment | `tax_category` | text | drop down `TaxCategoryList` |
| T | Is cleaning | `is_cleaning` | flag | formula, hidden |
| U | Is maintenance | `is_maintenance` | flag | formula, hidden |
| V | Is capital | `is_capital` | flag | formula, hidden |
| W | Counts in profit | `in_pl` | flag | formula |
| X | Counted as | `pl_amount` | money | formula |
| Y | Action | `alert` | status | formula |
| Z | Notes | `notes` | wrap | you type it |

#### Supplies Inventory

Linen, consumables and the things that run out on a Friday afternoon

27 columns, 300 rows of capacity, 14 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Item ID | `id` | id | formula |
| B | Item | `name` | text | you type it |
| C | Category | `category` | text | drop down `SupplyCategoryList` |
| D | Property | `property` | text | drop down `PropertyList` |
| E | Kept for | `property_name` | text | formula |
| F | Unit | `unit` | text | you type it |
| G | Supplier | `supplier` | text | you type it |
| H | Code / SKU | `sku` | text | you type it, hidden |
| I | In stock | `on_hand` | number1 | you type it (checked) |
| J | Minimum | `min_level` | number1 | you type it (checked) |
| K | Last counted | `last_count` | date | you type it (checked) |
| L | Last restocked | `last_restock` | date | you type it (checked) |
| M | Days since count | `count_age` | int | formula, hidden |
| N | Used per turnover | `per_turn` | number1 | you type it (checked) |
| O | Turnovers a month | `turns_month` | number1 | formula |
| P | Used a month | `usage_month` | number1 | formula |
| Q | Weeks of cover | `weeks_cover` | number1 | formula |
| R | Unit cost | `unit_cost` | money | you type it (checked) |
| S | Stock value | `value` | money | formula |
| T | Order quantity | `reorder_qty` | number1 | you type it (checked) |
| U | To order | `order_qty_used` | number1 | formula |
| V | Order cost | `order_cost` | money | formula |
| W | Cost a turnover | `cost_per_turn` | money | formula, hidden |
| X | Order now | `needs_order` | flag | formula, hidden |
| Y | Stock status | `status` | status | formula |
| Z | Action | `alert` | status | formula |
| AA | Notes | `notes` | wrap | you type it |

#### Maintenance Log

What is broken, who is fixing it, when, and what it actually cost

30 columns, 300 rows of capacity, 18 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Job ID | `id` | id | formula |
| B | Reported | `report_date` | date | you type it (checked) |
| C | Month | `month` | text | formula, hidden |
| D | Year | `year` | int | formula, hidden |
| E | Property | `property` | text | drop down `PropertyList` |
| F | Property name | `property_name` | text | formula |
| G | Category | `category` | text | drop down `MaintenanceCategoryList` |
| H | What is wrong | `description` | wrap | you type it |
| I | Priority | `priority` | status | drop down `PriorityList` |
| J | Reported by | `reported_by` | text | you type it |
| K | During booking | `booking` | text | drop down `BookingList`, hidden |
| L | Contractor | `vendor` | text | drop down `VendorList` |
| M | Contractor name | `vendor_name` | text | formula |
| N | Phone | `vendor_phone` | text | formula, hidden |
| O | Booked in for | `scheduled` | date | you type it (checked) |
| P | Finished | `completed` | date | you type it (checked) |
| Q | Status | `status` | status | drop down `MaintenanceStatusList` |
| R | Blocks the calendar | `blocks_calendar` | flag | you type it (checked) |
| S | Estimate | `est_cost` | money | you type it (checked) |
| T | Actual | `act_cost` | money | you type it (checked) |
| U | Cost counted | `cost_used` | money | formula |
| V | Over the estimate | `overrun` | money | formula, hidden |
| W | Invoice ref | `invoice_ref` | text | you type it |
| X | Paid | `paid` | flag | you type it (checked) |
| Y | Under warranty to | `warranty_until` | date | you type it (checked), hidden |
| Z | Still open | `is_open` | flag | formula, hidden |
| AA | Days open | `days_open` | int | formula |
| AB | Due in | `due_in` | int | formula, hidden |
| AC | Action | `alert` | status | formula |
| AD | Notes | `notes` | wrap | you type it |

#### Reviews Tracker

All seven scores, per property and per platform - so you can see which one is actually costing you

32 columns, 600 rows of capacity, 17 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Review ID | `id` | id | formula |
| B | Booking | `booking` | text | drop down `BookingList` |
| C | Property | `property` | text | formula, hidden |
| D | Property | `property_name` | text | formula |
| E | Guest | `guest` | text | formula, hidden |
| F | Guest | `guest_name` | text | formula |
| G | Platform | `platform` | text | formula |
| H | Checked out | `checkout` | date | formula |
| I | Review asked for | `requested` | date | you type it (checked) |
| J | Status | `status` | status | drop down `ReviewStatusList` |
| K | Review received | `review_date` | date | you type it (checked) |
| L | Month | `month` | text | formula, hidden |
| M | Year | `year` | int | formula, hidden |
| N | Days to review | `days_to_review` | int | formula, hidden |
| O | Overall | `rate_overall` | number1 | you type it (checked) |
| P | Cleanliness | `rate_clean` | number1 | you type it (checked) |
| Q | Accuracy | `rate_accuracy` | number1 | you type it (checked) |
| R | Check-in | `rate_checkin` | number1 | you type it (checked) |
| S | Communication | `rate_comm` | number1 | you type it (checked) |
| T | Location | `rate_location` | number1 | you type it (checked) |
| U | Value | `rate_value` | number1 | you type it (checked) |
| V | Scored | `has_score` | flag | formula, hidden |
| W | Category average | `detail_avg` | number | formula |
| X | Weakest category | `weakest` | text | formula |
| Y | Under target | `below_target` | flag | formula, hidden |
| Z | What they wrote | `public_review` | wrap | you type it |
| AA | Private feedback | `private_note` | wrap | you type it |
| AB | You replied | `responded` | flag | you type it (checked) |
| AC | Replied on | `response_date` | date | you type it (checked), hidden |
| AD | Your reply | `response` | wrap | you type it |
| AE | What you changed | `action_taken` | wrap | you type it |
| AF | Action | `alert` | status | formula |

#### Staff & Vendors

Cleaners, contractors and co-hosts - what they charge and what they have actually done

31 columns, 60 rows of capacity, 15 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Vendor ID | `id` | id | formula |
| B | Name | `name` | text | you type it |
| C | Role | `role` | text | drop down `StaffRoleList` |
| D | Company | `company` | text | you type it |
| E | Phone | `phone` | text | you type it |
| F | Email | `email` | text | you type it |
| G | Covers | `properties` | text | you type it |
| H | Available | `availability` | text | drop down `AvailabilityList` |
| I | Paid | `pay_basis` | text | drop down `PayBasisList` |
| J | Rate | `rate` | money | you type it (checked) |
| K | Contract | `contract` | status | drop down `ContractStatusList` |
| L | Working with you since | `started` | date | you type it (checked), hidden |
| M | Insured | `insured` | flag | you type it (checked) |
| N | Cover expires | `insurance_expiry` | date | you type it (checked) |
| O | Days left | `cover_days` | int | formula, hidden |
| P | Holds keys | `keyholder` | flag | you type it (checked) |
| Q | Turnovers done | `cleans` | int | formula |
| R | Finished properly | `cleans_done` | int | formula, hidden |
| S | Completed in full | `clean_quality` | percent | formula |
| T | Hours on site | `clean_hours` | number1 | formula, hidden |
| U | Avg a turnover | `avg_hours` | number1 | formula, hidden |
| V | Repairs done | `jobs` | int | formula |
| W | Repairs open | `jobs_open` | int | formula |
| X | Last turnover | `last_job` | date | formula |
| Y | Paid for cleaning | `paid_cleaning` | money | formula |
| Z | Paid for repairs | `paid_repairs` | money | formula |
| AA | Total paid | `paid_total` | money | formula |
| AB | Still owed | `owed` | money | formula |
| AC | Rank | `spend_rank` | int | formula, hidden |
| AD | Action | `alert` | status | formula |
| AE | Notes | `notes` | wrap | you type it |

#### Occupancy & Performance

One row per property per month: occupancy, ADR, RevPAR, costs and profit

34 columns, 400 rows of capacity, 3 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Row ID | `id` | id | formula |
| B | Property | `property` | text | drop down `PropertyList` |
| C | Property name | `property_name` | text | formula |
| D | Month | `period` | date | you type it (checked) |
| E | Month key | `month` | text | formula, hidden |
| F | Period | `month_label` | text | formula |
| G | Days | `days_in_month` | int | formula, hidden |
| H | Blocked | `nights_blocked` | int | formula |
| I | Available | `nights_available` | int | formula |
| J | Booked | `nights_booked` | int | formula |
| K | Occupancy | `occupancy` | percent | formula |
| L | Bookings | `bookings` | int | formula |
| M | Cancelled | `cancellations` | int | formula |
| N | Cancel rate | `cancel_rate` | percent | formula, hidden |
| O | Guests | `guests` | int | formula, hidden |
| P | Avg stay | `alos` | number1 | formula |
| Q | Room revenue | `accommodation` | money | formula |
| R | ADR | `adr` | money | formula |
| S | RevPAR | `revpar` | money | formula |
| T | Gross | `gross` | money | formula |
| U | Platform fees | `platform_fees` | money | formula |
| V | Net revenue | `net_revenue` | money | formula |
| W | Cleaning | `cleaning_cost` | money | formula |
| X | Maintenance | `maintenance_cost` | money | formula |
| Y | Other costs | `other_cost` | money | formula |
| Z | Total costs | `total_cost` | money | formula |
| AA | Operating profit | `profit` | money | formula |
| AB | Margin | `margin` | percent | formula |
| AC | Cost a night | `cost_per_night` | money | formula |
| AD | Cleaning a booking | `clean_per_booking` | money | formula |
| AE | Profit a night | `profit_night` | money | formula |
| AF | Rating | `rating` | number | formula, hidden |
| AG | Action | `alert` | status | formula |
| AH | Notes | `notes` | wrap | you type it |

---

*12,064 lines of Python in 37 files.*
