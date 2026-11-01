# Wedding Photographer Business Manager — source code

Complete source code of the generator that writes the `.xlsx` product.

Writes a 23 tab `.xlsx` workbook with nearly 115,000 formulas for a wedding photographer: a lead tracker that chases an enquiry on your own schedule and reports the win rate; a client CRM; a booking register of 87 columns where one wedding is one row; a twelve month calendar that turns red when two couples want the same Saturday; packages costed to a profit-per-hour figure; contracts and documents tracked through to counter-signed; deposits, instalments and balances with an overdue count in days; a wedding day timeline built from offsets against the ceremony time; a shot list with must-haves, family groups and who asked for them; second shooters with rates and insurance expiry; a ten step editing workflow against the turnaround you promised; gallery delivery with views, downloads, album orders and an expiry warning; equipment with shutter counts and replacement costs; travel and mileage against the travel fee you charged; a marketing calendar that reports enquiries per post; 27 expense categories; a profit analysis per wedding; a printable client questionnaire; reviews and referrals; and monthly and annual reports. Document author is **Novality Store**; every formula cell is locked and every sheet protected with the password **`premium`**, while the cells you are meant to type in stay open.

One technical decision holds the product together: **every wedding gets a unique Booking ID** (`WED-0001`), generated once on Wedding Bookings and chosen from a drop-down everywhere else. Eleven registers join back through it, so the payment, the timeline, the shot list, the editing job and the gallery are all demonstrably the same wedding. The validator checks that join in both directions on every build, because a Booking ID that can be mistyped is a spreadsheet full of quietly disconnected records.

Nothing is hard coded. The deposit percentage, the payment schedule, the extra hour rate, what a second shooter costs, the mileage rate and the free mileage allowance, the editing hours and rate, what your own hour is worth, the preview, gallery and album turnarounds, the tax rate and whether tax applies at all, every target and every alert window and all 48 drop-down lists are settings, not assumptions.

Two things it refuses to do, both deliberately. It never claims to connect to a booking platform or a calendar service - you type what happened and the formulas do the arithmetic - and it **tracks documents rather than pretending to generate legally binding contracts**. The Getting Started tab says both in plain words and the validator fails the build if any cell claims otherwise.

Pure Python, one dependency. Tested on XlsxWriter 3.2.0 and on the much older 0.3.0, so it does not rely on anything recent. The workbook itself is Excel 2010 safe — no dynamic arrays, no `XLOOKUP`, no `IFS`, nothing that needs array entry — and it opens in Google Sheets, LibreOffice and Numbers.

```bash
python build_wedding_workbook.py --help

  -o, --output PATH     where to write the file
  --blank               no example data, just the empty workbook
  --no-protect          leave the sheets unlocked
  --password TEXT       protection password (default: premium)
  --author TEXT         document author  (default: Novality Store)
  --studio TEXT         your studio name on the cover
  --owner TEXT          the photographer's name
  --currency SYMBOL     currency symbol  (default: $)
  --currency-code CODE  three letter code (default: USD)
  --deposit PERCENT     the deposit that holds a date
  --final-days N        days before the wedding the balance is due
  --extra-hour AMOUNT   the charge for an extra hour on the day
  --second-shooter AMT  what a second shooter costs you
  --mileage RATE        the mileage rate per mile
  --delivery-days N     the gallery turnaround you promise
  --editing-hours N     honest editing hours per wedding
  --tax PERCENT         sales tax or VAT on packages
  --weddings-target N   how many weddings a year you want
  --revenue-target AMT  revenue target for the year
  --margin-target P     the profit margin you are aiming at
  --rows N              rows of capacity in the big registers
  --today YYYY-MM-DD    pretend it is this date (for the example)
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
  * [`build_wedding_workbook.py`](#build_wedding_workbookpy)
* [2. The package](#2-the-package)
  * [`wedding_photography_spreadsheet/__init__.py`](#wedding_photography_spreadsheet__init__py)
  * [`wedding_photography_spreadsheet/__main__.py`](#wedding_photography_spreadsheet__main__py)
  * [`wedding_photography_spreadsheet/cli.py`](#wedding_photography_spreadsheetclipy)
  * [`wedding_photography_spreadsheet/config.py`](#wedding_photography_spreadsheetconfigpy)
  * [`wedding_photography_spreadsheet/layout.py`](#wedding_photography_spreadsheetlayoutpy)
  * [`wedding_photography_spreadsheet/formats.py`](#wedding_photography_spreadsheetformatspy)
  * [`wedding_photography_spreadsheet/writer.py`](#wedding_photography_spreadsheetwriterpy)
  * [`wedding_photography_spreadsheet/names.py`](#wedding_photography_spreadsheetnamespy)
  * [`wedding_photography_spreadsheet/sample_data.py`](#wedding_photography_spreadsheetsample_datapy)
  * [`wedding_photography_spreadsheet/build.py`](#wedding_photography_spreadsheetbuildpy)
* [3. Shared column factories](#3-shared-column-factories)
  * [`wedding_photography_spreadsheet/sheets/__init__.py`](#wedding_photography_spreadsheetsheets__init__py)
  * [`wedding_photography_spreadsheet/sheets/common.py`](#wedding_photography_spreadsheetsheetscommonpy)
* [4. Settings and the two front pages](#4-settings-and-the-two-front-pages)
  * [`wedding_photography_spreadsheet/sheets/settings.py`](#wedding_photography_spreadsheetsheetssettingspy)
  * [`wedding_photography_spreadsheet/sheets/dashboard.py`](#wedding_photography_spreadsheetsheetsdashboardpy)
  * [`wedding_photography_spreadsheet/sheets/getting_started.py`](#wedding_photography_spreadsheetsheetsgetting_startedpy)
* [5. Enquiry to confirmed booking](#5-enquiry-to-confirmed-booking)
  * [`wedding_photography_spreadsheet/sheets/leads.py`](#wedding_photography_spreadsheetsheetsleadspy)
  * [`wedding_photography_spreadsheet/sheets/clients.py`](#wedding_photography_spreadsheetsheetsclientspy)
  * [`wedding_photography_spreadsheet/sheets/bookings.py`](#wedding_photography_spreadsheetsheetsbookingspy)
* [6. Package to payment collection](#6-package-to-payment-collection)
  * [`wedding_photography_spreadsheet/sheets/packages.py`](#wedding_photography_spreadsheetsheetspackagespy)
  * [`wedding_photography_spreadsheet/sheets/contracts.py`](#wedding_photography_spreadsheetsheetscontractspy)
  * [`wedding_photography_spreadsheet/sheets/payments.py`](#wedding_photography_spreadsheetsheetspaymentspy)
* [7. The wedding day itself](#7-the-wedding-day-itself)
  * [`wedding_photography_spreadsheet/sheets/timeline.py`](#wedding_photography_spreadsheetsheetstimelinepy)
  * [`wedding_photography_spreadsheet/sheets/shots.py`](#wedding_photography_spreadsheetsheetsshotspy)
  * [`wedding_photography_spreadsheet/sheets/staff.py`](#wedding_photography_spreadsheetsheetsstaffpy)
  * [`wedding_photography_spreadsheet/sheets/calendar.py`](#wedding_photography_spreadsheetsheetscalendarpy)
  * [`wedding_photography_spreadsheet/sheets/questionnaire.py`](#wedding_photography_spreadsheetsheetsquestionnairepy)
* [8. Editing to final gallery delivery](#8-editing-to-final-gallery-delivery)
  * [`wedding_photography_spreadsheet/sheets/editing.py`](#wedding_photography_spreadsheetsheetseditingpy)
  * [`wedding_photography_spreadsheet/sheets/gallery.py`](#wedding_photography_spreadsheetsheetsgallerypy)
* [9. What each wedding actually made](#9-what-each-wedding-actually-made)
  * [`wedding_photography_spreadsheet/sheets/equipment.py`](#wedding_photography_spreadsheetsheetsequipmentpy)
  * [`wedding_photography_spreadsheet/sheets/travel.py`](#wedding_photography_spreadsheetsheetstravelpy)
  * [`wedding_photography_spreadsheet/sheets/marketing.py`](#wedding_photography_spreadsheetsheetsmarketingpy)
  * [`wedding_photography_spreadsheet/sheets/expenses.py`](#wedding_photography_spreadsheetsheetsexpensespy)
  * [`wedding_photography_spreadsheet/sheets/profit.py`](#wedding_photography_spreadsheetsheetsprofitpy)
  * [`wedding_photography_spreadsheet/sheets/reviews.py`](#wedding_photography_spreadsheetsheetsreviewspy)
  * [`wedding_photography_spreadsheet/sheets/reports.py`](#wedding_photography_spreadsheetsheetsreportspy)
* [10. The quality tools](#10-the-quality-tools)
  * [`tools/validate_wedding_workbook.py`](#toolsvalidate_wedding_workbookpy)
  * [`tools/check_parens.py`](#toolscheck_parenspy)
  * [`tools/check_wedding_specs.py`](#toolscheck_wedding_specspy)
  * [`tools/check_wedding_sample.py`](#toolscheck_wedding_samplepy)
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

`build_wedding_workbook.py` parses the command line into a `Config`, then
`wedding_photography_spreadsheet.build.build_workbook(path, cfg)` does six things in order:

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

23 tabs in the order they appear in the finished workbook. "Rows" is the capacity of a register at the default `--rows` setting.

| # | Tab | Kind | Written by | Columns | Rows |
|---|-----|------|------------|--------:|-----:|
| 1 | **Dashboard** | page | `wedding_photography_spreadsheet/sheets/dashboard.py` | - | - |
| 2 | **Getting Started** | page | `wedding_photography_spreadsheet/sheets/getting_started.py` | - | - |
| 3 | **Settings & Lists** | page | `wedding_photography_spreadsheet/sheets/settings.py` | - | - |
| 4 | **Client CRM** | register | `wedding_photography_spreadsheet/sheets/clients.py` | 33 | 300 |
| 5 | **Lead Tracker** | register | `wedding_photography_spreadsheet/sheets/leads.py` | 40 | 400 |
| 6 | **Booking Calendar** | page | `wedding_photography_spreadsheet/sheets/calendar.py` | - | - |
| 7 | **Wedding Bookings** | register | `wedding_photography_spreadsheet/sheets/bookings.py` | 87 | 200 |
| 8 | **Packages & Pricing** | register | `wedding_photography_spreadsheet/sheets/packages.py` | 41 | 30 |
| 9 | **Contracts & Documents** | register | `wedding_photography_spreadsheet/sheets/contracts.py` | 21 | 200 |
| 10 | **Payments & Invoices** | register | `wedding_photography_spreadsheet/sheets/payments.py` | 29 | 600 |
| 11 | **Wedding Timeline** | register | `wedding_photography_spreadsheet/sheets/timeline.py` | 33 | 900 |
| 12 | **Shot List Manager** | register | `wedding_photography_spreadsheet/sheets/shots.py` | 29 | 1400 |
| 13 | **Photographer & Staff** | register | `wedding_photography_spreadsheet/sheets/staff.py` | 38 | 40 |
| 14 | **Editing Workflow** | register | `wedding_photography_spreadsheet/sheets/editing.py` | 46 | 200 |
| 15 | **Gallery & Delivery** | register | `wedding_photography_spreadsheet/sheets/gallery.py` | 36 | 200 |
| 16 | **Equipment Inventory** | register | `wedding_photography_spreadsheet/sheets/equipment.py` | 33 | 150 |
| 17 | **Travel & Mileage** | register | `wedding_photography_spreadsheet/sheets/travel.py` | 33 | 300 |
| 18 | **Marketing Calendar** | register | `wedding_photography_spreadsheet/sheets/marketing.py` | 32 | 400 |
| 19 | **Expense Tracker** | register | `wedding_photography_spreadsheet/sheets/expenses.py` | 26 | 600 |
| 20 | **Profit Analysis** | register | `wedding_photography_spreadsheet/sheets/profit.py` | 44 | 200 |
| 21 | **Client Questionnaire** | page | `wedding_photography_spreadsheet/sheets/questionnaire.py` | - | - |
| 22 | **Reviews & Referrals** | register | `wedding_photography_spreadsheet/sheets/reviews.py` | 30 | 200 |
| 23 | **Monthly & Annual Reports** | page | `wedding_photography_spreadsheet/sheets/reports.py` | - | - |

### Every file at a glance

| File | Lines | What it is |
|------|------:|------------|
| [`build_wedding_workbook.py`](#build_wedding_workbookpy) | 22 | Build the Wedding Photographer Business Manager workbook |
| [`wedding_photography_spreadsheet/__init__.py`](#wedding_photography_spreadsheet__init__py) | 19 | Wedding Photographer Business Manager |
| [`wedding_photography_spreadsheet/__main__.py`](#wedding_photography_spreadsheet__main__py) | 5 | ``python -m wedding_photography_spreadsheet`` |
| [`wedding_photography_spreadsheet/cli.py`](#wedding_photography_spreadsheetclipy) | 200 | Command line front end |
| [`wedding_photography_spreadsheet/config.py`](#wedding_photography_spreadsheetconfigpy) | 746 | Every default, list, package and palette the workbook is built from |
| [`wedding_photography_spreadsheet/layout.py`](#wedding_photography_spreadsheetlayoutpy) | 253 | Layout engine |
| [`wedding_photography_spreadsheet/formats.py`](#wedding_photography_spreadsheetformatspy) | 202 | Cell format factory |
| [`wedding_photography_spreadsheet/writer.py`](#wedding_photography_spreadsheetwriterpy) | 496 | Sheet writing helpers shared by every tab of the workbook |
| [`wedding_photography_spreadsheet/names.py`](#wedding_photography_spreadsheetnamespy) | 192 | Workbook level defined names that point at the register sheets |
| [`wedding_photography_spreadsheet/sample_data.py`](#wedding_photography_spreadsheetsample_datapy) | 2,038 | A worked example: one studio, fifteen weddings, every tab populated |
| [`wedding_photography_spreadsheet/build.py`](#wedding_photography_spreadsheetbuildpy) | 165 | Workbook builder: puts the twenty-three tabs together in order |
| [`wedding_photography_spreadsheet/sheets/__init__.py`](#wedding_photography_spreadsheetsheets__init__py) | 0 |  |
| [`wedding_photography_spreadsheet/sheets/common.py`](#wedding_photography_spreadsheetsheetscommonpy) | 560 | Small building blocks shared by the register sheets |
| [`wedding_photography_spreadsheet/sheets/settings.py`](#wedding_photography_spreadsheetsheetssettingspy) | 655 | `Settings & Lists` - every number and every list the workbook runs on |
| [`wedding_photography_spreadsheet/sheets/dashboard.py`](#wedding_photography_spreadsheetsheetsdashboardpy) | 494 | `Dashboard` - the page that answers "how is the business doing?" Everything here is worked out from the... |
| [`wedding_photography_spreadsheet/sheets/getting_started.py`](#wedding_photography_spreadsheetsheetsgetting_startedpy) | 286 | `Getting Started` - what to do first, and what this file is not |
| [`wedding_photography_spreadsheet/sheets/leads.py`](#wedding_photography_spreadsheetsheetsleadspy) | 228 | Lead Tracker - every enquiry, and what happened to it |
| [`wedding_photography_spreadsheet/sheets/clients.py`](#wedding_photography_spreadsheetsheetsclientspy) | 198 | Client CRM - one row per couple, from first enquiry to final gallery |
| [`wedding_photography_spreadsheet/sheets/bookings.py`](#wedding_photography_spreadsheetsheetsbookingspy) | 479 | Wedding Bookings - one row per wedding, and the spine of the file |
| [`wedding_photography_spreadsheet/sheets/packages.py`](#wedding_photography_spreadsheetsheetspackagespy) | 270 | Packages & Pricing - what you sell, what it costs you, what it earns |
| [`wedding_photography_spreadsheet/sheets/contracts.py`](#wedding_photography_spreadsheetsheetscontractspy) | 151 | Contracts & Documents - a register of paperwork, not a contract tool |
| [`wedding_photography_spreadsheet/sheets/payments.py`](#wedding_photography_spreadsheetsheetspaymentspy) | 176 | Payments & Invoices - every instalment, due and received |
| [`wedding_photography_spreadsheet/sheets/timeline.py`](#wedding_photography_spreadsheetsheetstimelinepy) | 215 | Wedding Timeline - the running order of a wedding day, hour by hour |
| [`wedding_photography_spreadsheet/sheets/shots.py`](#wedding_photography_spreadsheetsheetsshotspy) | 202 | Shot List Manager - the list you work from, grouped and ticked off |
| [`wedding_photography_spreadsheet/sheets/staff.py`](#wedding_photography_spreadsheetsheetsstaffpy) | 222 | Photographer & Staff - who shoots, who edits, and what they cost |
| [`wedding_photography_spreadsheet/sheets/calendar.py`](#wedding_photography_spreadsheetsheetscalendarpy) | 186 | `Booking Calendar` - twelve months of dates, at a glance |
| [`wedding_photography_spreadsheet/sheets/questionnaire.py`](#wedding_photography_spreadsheetsheetsquestionnairepy) | 228 | `Client Questionnaire` - the printable form, and the answers |
| [`wedding_photography_spreadsheet/sheets/editing.py`](#wedding_photography_spreadsheetsheetseditingpy) | 258 | Editing Workflow - one row per wedding, ten steps from card to gallery |
| [`wedding_photography_spreadsheet/sheets/gallery.py`](#wedding_photography_spreadsheetsheetsgallerypy) | 213 | Gallery & Delivery - handing the photographs over, and what happens next |
| [`wedding_photography_spreadsheet/sheets/equipment.py`](#wedding_photography_spreadsheetsheetsequipmentpy) | 206 | Equipment Inventory - what you own, what it is worth and what it owes you |
| [`wedding_photography_spreadsheet/sheets/travel.py`](#wedding_photography_spreadsheetsheetstravelpy) | 221 | Travel & Mileage - getting to the wedding, and what it costs |
| [`wedding_photography_spreadsheet/sheets/marketing.py`](#wedding_photography_spreadsheetsheetsmarketingpy) | 186 | Marketing Calendar - what you posted, and whether anybody booked |
| [`wedding_photography_spreadsheet/sheets/expenses.py`](#wedding_photography_spreadsheetsheetsexpensespy) | 180 | Expense Tracker - everything the business pays out |
| [`wedding_photography_spreadsheet/sheets/profit.py`](#wedding_photography_spreadsheetsheetsprofitpy) | 277 | Profit Analysis - what each wedding actually made, line by line |
| [`wedding_photography_spreadsheet/sheets/reviews.py`](#wedding_photography_spreadsheetsheetsreviewspy) | 179 | Reviews & Referrals - the cheapest marketing there is, tracked |
| [`wedding_photography_spreadsheet/sheets/reports.py`](#wedding_photography_spreadsheetsheetsreportspy) | 386 | `Monthly & Annual Reports` - the year in numbers, and the year before |
| [`tools/validate_wedding_workbook.py`](#toolsvalidate_wedding_workbookpy) | 1,019 | Static checks on a generated Wedding Photographer Business Manager |
| [`tools/check_parens.py`](#toolscheck_parenspy) | 121 | Parenthesis and quote balance check for every formula in a workbook |
| [`tools/check_wedding_specs.py`](#toolscheck_wedding_specspy) | 170 | Fast structural sweep over the wedding product's register specs |
| [`tools/check_wedding_sample.py`](#toolscheck_wedding_samplepy) | 146 | Validate the wedding product's example data before it is built |
| [`tools/recalc_check.py`](#toolsrecalc_checkpy) | 216 | Really calculate every formula in a generated workbook and report any cell that ends up as an Excel error... |

---

## 1. Run it

Save the files below keeping these exact paths, then run the one liner. Nothing else is needed - no project file, no build step.

```bash
pip install XlsxWriter
python build_wedding_workbook.py
```

The workbook lands in `output/Wedding Photographer Business Manager.xlsx`.

### `build_wedding_workbook.py`

*22 lines*

> Build the Wedding Photographer Business Manager workbook.
>
>     python3 build_wedding_workbook.py
>     python3 build_wedding_workbook.py -o "out/Wedding.xlsx" --blank
>     python3 build_wedding_workbook.py --rows 60 --today 2026-11-01
>
> A thin wrapper around ``python -m wedding_photography_spreadsheet`` so
> the product can be built without worrying about the import path.

````python
#!/usr/bin/env python3
"""Build the Wedding Photographer Business Manager workbook.

    python3 build_wedding_workbook.py
    python3 build_wedding_workbook.py -o "out/Wedding.xlsx" --blank
    python3 build_wedding_workbook.py --rows 60 --today 2026-11-01

A thin wrapper around ``python -m wedding_photography_spreadsheet`` so
the product can be built without worrying about the import path.
"""

from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from wedding_photography_spreadsheet.cli import main   # noqa: E402

if __name__ == "__main__":
    raise SystemExit(main())
````

---

## 2. The package

Configuration, the formula template engine, cell formats, the sheet writer, the cross sheet names and the worked example. These are the parts every tab leans on. `config.py` holds every price, every turnaround promise and every disclaimer, so a photographer in another country, on another pricing model, can be matched without touching a formula.

### `wedding_photography_spreadsheet/__init__.py`

*19 lines*

> Wedding Photographer Business Manager.
>
> Build the workbook with::
>
>     python -m wedding_photography_spreadsheet -o "output/Wedding.xlsx"
>
> or from Python::
>
>     from wedding_photography_spreadsheet import build_workbook
>     build_workbook("Wedding Photographer Business Manager.xlsx")

````python
"""Wedding Photographer Business Manager.

Build the workbook with::

    python -m wedding_photography_spreadsheet -o "output/Wedding.xlsx"

or from Python::

    from wedding_photography_spreadsheet import build_workbook
    build_workbook("Wedding Photographer Business Manager.xlsx")
"""

from .build import build_workbook
from .config import (AUTHOR, LISTING_TITLE, PRODUCT_NAME,
                     PRODUCT_SUBTITLE, VERSION, Config)

__all__ = ["build_workbook", "Config", "AUTHOR", "PRODUCT_NAME",
           "PRODUCT_SUBTITLE", "LISTING_TITLE", "VERSION"]
__version__ = VERSION
````

### `wedding_photography_spreadsheet/__main__.py`

*5 lines*

> ``python -m wedding_photography_spreadsheet``.

````python
"""``python -m wedding_photography_spreadsheet``."""

from .cli import main

raise SystemExit(main())
````

### `wedding_photography_spreadsheet/cli.py`

*200 lines · 3 functions*

> Command line front end.
>
> python -m wedding_photography_spreadsheet
> python -m wedding_photography_spreadsheet -o "out/Wedding.xlsx" --blank
> python -m wedding_photography_spreadsheet --currency "£" --rows 60

| Defines | What it does |
|---------|--------------|
| `build_parser()` |  |
| `config_from_args(args)` |  |
| `main(argv)` |  |

**Module constants:** `DEFAULT_OUTPUT` (os.path.join(...)).

````python
"""Command line front end.

    python -m wedding_photography_spreadsheet
    python -m wedding_photography_spreadsheet -o "out/Wedding.xlsx" --blank
    python -m wedding_photography_spreadsheet --currency "£" --rows 60
"""

from __future__ import annotations

import argparse
import datetime as _dt
import os
import sys
from typing import List, Optional

from .build import SHEET_ORDER, build_workbook
from .config import (AUTHOR, DEFAULT_ROWS, LISTING_TITLE, PACKAGES,
                     PRODUCT_NAME, SCALE_ANCHOR, Config)

DEFAULT_OUTPUT = os.path.join(
    "output", "Wedding Photographer Business Manager.xlsx")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="wedding_photography_spreadsheet",
        description=f"Build the {PRODUCT_NAME} workbook.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument("-o", "--output", default=DEFAULT_OUTPUT,
                        help="where to write the .xlsx file")
    parser.add_argument("--blank", action="store_true",
                        help="leave the registers empty (no example "
                             "studio)")
    parser.add_argument("--no-protect", action="store_true",
                        help="do not protect the sheets (for "
                             "development)")
    parser.add_argument("--password", default=None,
                        help="sheet protection password")
    parser.add_argument("--author", default=AUTHOR,
                        help="document author / creator")
    parser.add_argument("--studio", default=None,
                        help="pre-fill the studio name on Settings")
    parser.add_argument("--owner", default=None,
                        help="pre-fill the photographer's name")
    parser.add_argument("--currency", default=None,
                        help="currency symbol, e.g. $ £ € A$")
    parser.add_argument("--currency-code", default=None,
                        help="three letter currency code, e.g. USD")
    parser.add_argument("--year", type=int, default=None,
                        help="the year the reports open on")
    parser.add_argument("--deposit", type=float, default=None,
                        help="deposit to hold a date, as a "
                             "percentage, e.g. 30 for 30%%")
    parser.add_argument("--final-days", type=int, default=None,
                        help="days before the wedding the balance "
                             "falls due")
    parser.add_argument("--extra-hour", type=float, default=None,
                        help="charge for an extra hour on the day")
    parser.add_argument("--second-shooter", type=float, default=None,
                        help="what a second shooter costs you")
    parser.add_argument("--mileage", type=float, default=None,
                        help="mileage rate per mile")
    parser.add_argument("--delivery-days", type=int, default=None,
                        help="your promised gallery turnaround")
    parser.add_argument("--editing-hours", type=float, default=None,
                        help="honest editing hours per wedding")
    parser.add_argument("--tax", type=float, default=None,
                        help="sales tax or VAT on packages, as a "
                             "percentage, e.g. 20 for 20%% (zero by "
                             "default)")
    parser.add_argument("--weddings-target", type=int, default=None,
                        help="how many weddings a year you want")
    parser.add_argument("--revenue-target", type=float, default=None,
                        help="revenue target for the year")
    parser.add_argument("--margin-target", type=float, default=None,
                        help="profit margin target as a percentage, "
                             "e.g. 55 for 55%%")
    parser.add_argument("--rows", type=int, default=None,
                        help="scale every register (the defaults are "
                             "scaled to fit this many wedding rows)")
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
        kwargs["currency"] = args.currency
    if args.currency_code:
        kwargs["currency_code"] = args.currency_code
    if args.year:
        kwargs["year"] = args.year
    if args.deposit is not None:
        kwargs["deposit_percent"] = round(args.deposit / 100.0, 5)
    if args.final_days is not None:
        kwargs["final_payment_days_before"] = args.final_days
    if args.extra_hour is not None:
        kwargs["extra_hour_rate"] = args.extra_hour
    if args.second_shooter is not None:
        kwargs["second_shooter_rate"] = args.second_shooter
    if args.mileage is not None:
        kwargs["mileage_rate"] = args.mileage
    if args.delivery_days is not None:
        kwargs["delivery_days"] = args.delivery_days
    if args.editing_hours is not None:
        kwargs["editing_hours_per_wedding"] = args.editing_hours
    if args.tax is not None:
        kwargs["tax_rate"] = round(args.tax / 100.0, 5)
        kwargs["tax_on_packages"] = args.tax > 0
    if args.weddings_target is not None:
        kwargs["target_weddings_year"] = args.weddings_target
    if args.revenue_target is not None:
        kwargs["target_revenue_year"] = args.revenue_target
        kwargs["target_revenue_month"] = round(
            args.revenue_target / 12.0, 2)
    if args.margin_target is not None:
        kwargs["target_profit_margin"] = round(
            args.margin_target / 100.0, 4)
    if args.today:
        kwargs["today"] = _dt.date.fromisoformat(args.today)
    cfg = Config(**kwargs)
    if args.studio:
        cfg.studio.name = args.studio
    if args.owner:
        cfg.studio.owner = args.owner
    if args.rows:
        # every register scales against Wedding Bookings, which is
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

    rows = sum(cfg.row_count(sheet) for sheet in DEFAULT_ROWS)
    cheapest = min(PACKAGES, key=lambda p: p[-1])
    dearest = max(PACKAGES, key=lambda p: p[-1])
    sign = cfg.currency

    print(f"{PRODUCT_NAME}")
    print(f"  listing title : {LISTING_TITLE}")
    print(f"  author        : {cfg.author}")
    print(f"  sheets        : {len(SHEET_ORDER)}")
    print(f"  studio        : {cfg.studio.name} "
          f"({cfg.studio.owner})")
    print(f"  today         : {cfg.today.isoformat()}")
    print(f"  register rows : {rows:,} across "
          f"{len(DEFAULT_ROWS)} registers "
          f"(scaled on {SCALE_ANCHOR})")
    print(f"  packages      : {len(PACKAGES)}, "
          f"{sign}{cheapest[-1]:,.0f} to {sign}{dearest[-1]:,.0f}")
    print(f"  deposit       : {cfg.deposit_percent:.0%} to hold, "
          f"balance {cfg.final_payment_days_before} days before")
    print(f"  turnaround    : preview {cfg.preview_days} days, "
          f"gallery {cfg.delivery_days} days")
    print(f"  extras        : extra hour {sign}"
          f"{cfg.extra_hour_rate:,.0f}, second shooter {sign}"
          f"{cfg.second_shooter_rate:,.0f}, mileage {sign}"
          f"{cfg.mileage_rate:.2f}")
    print(f"  tax           : {cfg.tax_rate:.2%} "
          f"({'on' if cfg.tax_on_packages else 'off'})")
    print(f"  targets       : {cfg.target_weddings_year} weddings, "
          f"{sign}{cfg.target_revenue_year:,.0f}, "
          f"{cfg.target_profit_margin:.0%} margin")
    print(f"  example data  : "
          f"{'no' if not cfg.include_sample_data else 'yes'}")
    print(f"  protection    : "
          f"{'off' if not cfg.protect_sheets else 'on, password ' + cfg.password}")
    print(f"  written to    : {os.path.abspath(path)}  "
          f"({size:,.0f} KB)")
    return 0


if __name__ == "__main__":        # pragma: no cover
    raise SystemExit(main())
````

### `wedding_photography_spreadsheet/config.py`

*746 lines · 3 classes*

> Every default, list, package and palette the workbook is built from.
>
> Nothing in the generated file is hard coded to one photographer's way
> of working: the packages, their inclusions and their prices, the
> payment schedule, the mileage rate, the editing turnaround, the shot
> list and the wedding-day timeline are all settings a buyer types over.

| Defines | What it does |
|---------|--------------|
| `class Palette` | fields: ink, deep, primary, primary_mid, primary_soft, accent, good, warn |
| `class Studio` | fields: name, owner, email, phone, website, instagram, city, region |
| `class Config` | Everything the build reads |

**Module constants:** `AUTHOR` ('Novality Store'), `PASSWORD` ('premium'), `VERSION` ('1.0.0'), `PRODUCT_NAME` ('Wedding Photographer Business Manager'), `PRODUCT_SHORT` ('Wedding Photography Business'), `PRODUCT_SUBTITLE` ('From first enquiry to final gallery - one workbook for t...), `LISTING_TITLE` ('Wedding Photographer Business Spreadsheet | Photography ...), `HONESTY_NOTE` ('This workbook is a business organisation tool. It is not...), `CONTRACT_NOTE` ('This tab tracks the status of your paperwork. It does no...), `PRIVACY_NOTE` ("Record only what you need to photograph the wedding well...), `DISCLAIMER` ('Nothing here is accounting or tax advice. How a cost is ...), `LEAD_SOURCES` (13 items), `LEAD_STATUSES` (9 items), `OPEN_LEAD_STATUSES` (5 items) and 61 more.

````python
"""Every default, list, package and palette the workbook is built from.

Nothing in the generated file is hard coded to one photographer's way
of working: the packages, their inclusions and their prices, the
payment schedule, the mileage rate, the editing turnaround, the shot
list and the wedding-day timeline are all settings a buyer types over.
"""

from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

# ---------------------------------------------------------------------------
# the product
# ---------------------------------------------------------------------------

AUTHOR = "Novality Store"
PASSWORD = "premium"
VERSION = "1.0.0"

PRODUCT_NAME = "Wedding Photographer Business Manager"
PRODUCT_SHORT = "Wedding Photography Business"
PRODUCT_SUBTITLE = ("From first enquiry to final gallery - one workbook "
                    "for the whole wedding photography business")
LISTING_TITLE = ("Wedding Photographer Business Spreadsheet | Photography "
                 "CRM, Booking Calendar, Pricing, Invoices & Client "
                 "Management")

#: said plainly on Getting Started, on Settings and in the listing
HONESTY_NOTE = (
    "This workbook is a business organisation tool. It is not a live "
    "booking integration, it is not a contract-signing platform and it "
    "is not a gallery host. It does not connect to your calendar, your "
    "email, your gallery provider or your bank, and it cannot send a "
    "contract for signature or take a payment. It records what you have "
    "sent, signed, delivered and been paid, so that nothing is lost "
    "between the enquiry and the final gallery.")

#: the Contracts tab says this at the top, every time
CONTRACT_NOTE = (
    "This tab tracks the status of your paperwork. It does not draft, "
    "generate or execute a contract and nothing in it is legal advice. "
    "Have your contract, model release and cancellation policy written "
    "or reviewed by a lawyer in your own country, then record here when "
    "each one was sent, signed and filed.")

#: the privacy line on the client tabs
PRIVACY_NOTE = (
    "Record only what you need to photograph the wedding well. Do not "
    "store card numbers, bank details, passport or ID scans in a "
    "spreadsheet. Guest lists and family notes are personal data - keep "
    "the file somewhere private, share it with nobody who does not need "
    "it, and delete a wedding's personal notes once the gallery has "
    "been delivered and the review is in.")

#: the tax line
DISCLAIMER = (
    "Nothing here is accounting or tax advice. How a cost is "
    "classified, what is deductible, whether a camera body is an "
    "expense or a capital asset, and how self-employed photography "
    "income is taxed all depend on your jurisdiction and your "
    "circumstances. The figures are organised so that the conversation "
    "with an accountant is a short one; they do not replace it.")


# ---------------------------------------------------------------------------
# lists - every drop down in the workbook
# ---------------------------------------------------------------------------

LEAD_SOURCES = [
    "Instagram", "Pinterest", "Website", "Google search", "Referral",
    "Facebook", "Wedding fair", "Venue recommendation", "Past client",
    "Wedding directory", "TikTok", "Word of mouth", "Other",
]

LEAD_STATUSES = [
    "New", "Contacted", "Consultation booked", "Consultation done",
    "Proposal sent", "Booked", "Lost", "Not a fit", "Date unavailable",
]

#: lead statuses that are still live, so still worth chasing
OPEN_LEAD_STATUSES = ["New", "Contacted", "Consultation booked",
                      "Consultation done", "Proposal sent"]
WON_LEAD_STATUSES = ["Booked"]
LOST_LEAD_STATUSES = ["Lost", "Not a fit", "Date unavailable"]

CLIENT_STATUSES = [
    "Enquiry", "Booked", "Wedding done", "Delivered", "Complete",
    "Cancelled", "Postponed",
]

PRIORITY_LEVELS = ["High", "Normal", "Low"]

BOOKING_STATUSES = [
    "Pencilled in", "Confirmed", "Wedding done", "In editing",
    "Delivered", "Complete", "Cancelled", "Postponed",
]
#: a booking that still occupies the date
HOLDING_STATUSES = ["Pencilled in", "Confirmed", "Wedding done",
                    "In editing", "Delivered", "Complete"]
#: a booking that counts towards revenue
EARNING_STATUSES = ["Confirmed", "Wedding done", "In editing",
                    "Delivered", "Complete"]
#: a booking whose wedding day has happened
SHOT_STATUSES = ["Wedding done", "In editing", "Delivered", "Complete"]

CONTRACT_STATUSES = [
    "Not sent", "Sent", "Viewed", "Signed", "Counter-signed",
    "Amended", "Cancelled",
]

DOC_STATUSES = ["Not needed", "Outstanding", "Sent", "Received",
                "Signed", "Filed"]

PAYMENT_STATUSES = [
    "Not due yet", "Due", "Part paid", "Paid in full", "Overdue",
    "Refunded", "Written off",
]

PAYMENT_STAGES = [
    "Booking deposit", "Second instalment", "Third instalment",
    "Final payment", "Add-on", "Album", "Prints", "Travel",
    "Overtime", "Refund", "Other",
]

PAYMENT_METHODS = [
    "Bank transfer", "Debit card", "Credit card", "Cash", "PayPal",
    "Stripe", "Cheque", "Direct debit", "Other",
]

PACKAGE_TIERS = ["Essential", "Classic", "Premium", "Luxury", "Elopement",
                 "Micro wedding", "Custom"]

#: the four packages the workbook ships with, all editable on Settings
#: (name, tier, hours, photographers, engagement session, edited images,
#:  album, print credit, travel allowance miles, price)
PACKAGES: List[Tuple[str, str, float, int, str, int, str, float,
                     float, float]] = [
    ("Essential Wedding", "Essential", 6, 1, "No", 350,
     "No album", 0, 30, 1200),
    ("Classic Wedding", "Classic", 8, 1, "Yes", 500,
     "20 page album", 100, 60, 1800),
    ("Premium Wedding", "Premium", 10, 2, "Yes", 700,
     "30 page album", 250, 100, 2500),
    ("Luxury Wedding", "Luxury", 12, 2, "Yes", 900,
     "40 page album + parent copies", 500, 150, 3500),
]

ALBUM_OPTIONS = ["No album", "10 page album", "20 page album",
                 "30 page album", "40 page album",
                 "40 page album + parent copies", "Album to be chosen"]

ADD_ON_TYPES = [
    "Extra hour", "Second shooter", "Engagement session",
    "Bridal portraits", "Rehearsal dinner", "Album upgrade",
    "Parent albums", "Extra prints", "Wall art", "Drone coverage",
    "Photo booth", "Video add-on", "Rush delivery", "Raw files",
    "Travel", "Accommodation", "Other",
]

STAFF_ROLES = [
    "Lead photographer", "Second shooter", "Third shooter", "Assistant",
    "Videographer", "Editor", "Retoucher", "Album designer",
    "Studio manager", "Intern",
]

PAY_BASIS = ["Per wedding", "Per hour", "Per day", "Per image",
             "Monthly retainer", "% of booking", "Not paid (owner)"]

AVAILABILITY = ["Any day", "Weekends only", "Weekdays only",
                "By arrangement", "Unavailable"]

#: the wedding-day timeline the workbook ships with.  Every row is
#: (key, label, default minutes before/after the ceremony, default
#:  duration in minutes, who usually covers it)
TIMELINE_EVENTS: List[Tuple[str, str, int, int, str]] = [
    ("venue_detail", "Venue and detail shots", -300, 30,
     "Lead photographer"),
    ("bride_prep", "Bridal preparation", -270, 75, "Lead photographer"),
    ("groom_prep", "Groom preparation", -240, 45, "Second shooter"),
    ("dress_details", "Dress, rings and detail shots", -210, 20,
     "Lead photographer"),
    ("first_look", "First look", -120, 30, "Lead photographer"),
    ("couple_portraits_1", "Couple portraits before the ceremony",
     -90, 30, "Lead photographer"),
    ("guests_arrive", "Guests arriving", -30, 20, "Second shooter"),
    ("ceremony", "Wedding ceremony", 0, 45, "Lead photographer"),
    ("confetti", "Confetti and congratulations", 50, 15,
     "Lead photographer"),
    ("family_groups", "Family group photographs", 70, 40,
     "Lead photographer"),
    ("wedding_party", "Wedding party photographs", 115, 25,
     "Lead photographer"),
    ("couple_portraits_2", "Couple portraits", 145, 30,
     "Lead photographer"),
    ("cocktail_hour", "Cocktail hour and candids", 145, 45,
     "Second shooter"),
    ("reception_detail", "Reception room details", 180, 20,
     "Second shooter"),
    ("entrance", "Reception entrance", 210, 10, "Lead photographer"),
    ("speeches", "Speeches", 240, 45, "Lead photographer"),
    ("cake_cutting", "Cake cutting", 300, 10, "Lead photographer"),
    ("first_dance", "First dance", 315, 15, "Lead photographer"),
    ("dancing", "Dancing and party coverage", 330, 60,
     "Lead photographer"),
    ("sparkler_exit", "Sparkler or night exit", 400, 15,
     "Lead photographer"),
]

#: the shot list, grouped exactly as a photographer works through it.
#: (key, group label, [shot, ...])
SHOT_LIST: List[Tuple[str, str, List[str]]] = [
    ("details", "Details and venue", [
        "Venue exterior",
        "Ceremony room before guests",
        "Reception room before guests",
        "Table settings and stationery",
        "Flowers and decorations",
        "The dress hanging",
        "Shoes, rings and jewellery",
        "Invitation suite flat lay",
    ]),
    ("prep", "Getting ready", [
        "Bride in robe with the wedding party",
        "Hair and make-up",
        "Mother or friend helping with the dress",
        "Bride fully ready, window light portrait",
        "Groom and party getting ready",
        "Groom's details - watch, cufflinks, buttonhole",
        "Letters, gifts and first-look-at-the-note moments",
    ]),
    ("couple", "Bride and groom", [
        "Bride solo portraits",
        "Groom solo portraits",
        "First look",
        "Couple walking together",
        "Romantic close-ups",
        "Full length couple portrait with the venue",
        "Golden hour couple portraits",
        "Night portrait with off-camera light",
    ]),
    ("family", "Family and groups", [
        "Bride with her parents",
        "Groom with his parents",
        "Both families together",
        "Siblings",
        "Grandparents",
        "Extended family",
        "Wedding party together",
        "Bride with bridesmaids",
        "Groom with groomsmen",
    ]),
    ("ceremony", "Ceremony", [
        "Guests arriving and being seated",
        "Groom waiting at the front",
        "Walking down the aisle",
        "The couple's faces during the vows",
        "Exchange of rings",
        "First kiss",
        "Signing the register",
        "Ceremony exit and confetti",
    ]),
    ("reception", "Reception", [
        "Reception room and cake",
        "Entrance of the couple",
        "Speeches - speakers and reactions",
        "First dance",
        "Cake cutting",
        "Guests dancing",
        "Candid guest moments",
        "Bouquet toss",
        "Send-off",
    ]),
]

EDITING_STAGES = [
    "Not started", "Cards backed up", "Imported", "Culling",
    "Culled", "Editing", "Colour correction", "Retouching",
    "Album design", "Preview sent", "Revisions", "Ready to deliver",
    "Delivered",
]

#: the editing workflow checklist, in order.  (key, label)
EDITING_STEPS: List[Tuple[str, str]] = [
    ("backed_up", "Memory cards backed up"),
    ("imported", "Photos imported"),
    ("culled", "Culling complete"),
    ("selected", "Selection complete"),
    ("edited", "Editing complete"),
    ("colour", "Colour correction done"),
    ("retouched", "Retouching done"),
    ("album_designed", "Album designed"),
    ("preview_sent", "Client preview sent"),
    ("final_uploaded", "Final gallery uploaded"),
]

GALLERY_STATUSES = [
    "Not created", "Created", "Preview sent", "Client reviewing",
    "Revisions requested", "Delivered", "Archived", "Expired",
]

EQUIPMENT_CATEGORIES = [
    "Camera body", "Lens", "Flash / speedlight", "Studio lighting",
    "Light modifier", "Tripod / monopod", "Memory card", "Battery",
    "Camera bag", "Backup drive", "Editing computer", "Monitor",
    "Audio", "Drone", "Software licence", "Other",
]

EQUIPMENT_CONDITIONS = ["New", "Excellent", "Good", "Fair",
                        "Needs repair", "Retired", "Lost or stolen"]

INSURANCE_STATUSES = ["Insured", "Not insured", "Pending", "Expired",
                      "Not needed"]

TRAVEL_MODES = ["Car", "Train", "Flight", "Taxi", "Hire car", "Other"]

REIMBURSEMENT_STATUSES = ["Not charged", "To invoice", "Invoiced",
                          "Reimbursed", "Absorbed"]

MARKETING_PLATFORMS = [
    "Instagram", "Pinterest", "Facebook", "TikTok", "Blog", "Website",
    "Google Business", "Wedding directory", "Email newsletter",
    "YouTube", "Print ad", "Wedding fair", "Other",
]

CONTENT_TYPES = [
    "Full wedding gallery", "Sneak peek", "Carousel", "Reel / video",
    "Story", "Blog post", "Behind the scenes", "Client testimonial",
    "Educational post", "Promotion", "Portfolio update", "Other",
]

CONTENT_STATUSES = ["Idea", "Drafted", "Scheduled", "Published",
                    "Evergreen", "Archived"]

EXPENSE_CATEGORIES = [
    "Camera equipment", "Lens purchases", "Lighting & grip",
    "Memory & storage", "Software subscriptions", "Editing services",
    "Second shooters", "Assistants", "Albums & prints",
    "Packaging & postage", "Marketing & advertising",
    "Website & hosting", "Insurance", "Equipment repair",
    "Equipment rental", "Studio rent", "Travel & mileage",
    "Accommodation", "Meals on shoot", "Office expenses",
    "Education & training", "Accounting & legal", "Bank & card fees",
    "Licences & permits", "Client gifts", "Tax & national insurance",
    "Miscellaneous",
]

TAX_CATEGORIES = [
    "Operating expense", "Capital / asset", "Cost of sale",
    "Owner drawing", "Not deductible", "Review with accountant",
]

#: expenses that are usually an asset rather than an expense
CAPITAL_CATEGORIES = ["Camera equipment", "Lens purchases",
                      "Lighting & grip", "Editing computer"]

#: expenses that belong to one wedding rather than to the business
DIRECT_CATEGORIES = ["Second shooters", "Assistants", "Albums & prints",
                     "Packaging & postage", "Travel & mileage",
                     "Accommodation", "Meals on shoot",
                     "Editing services"]

RECURRING_FREQUENCIES = ["One off", "Weekly", "Monthly", "Quarterly",
                         "Twice a year", "Annual"]

FREQUENCY_PER_YEAR = {
    "One off": 0, "Weekly": 52, "Monthly": 12, "Quarterly": 4,
    "Twice a year": 2, "Annual": 1,
}

REVIEW_PLATFORMS = ["Google", "Facebook", "Instagram", "Website",
                    "Wedding directory", "Email", "Word of mouth",
                    "Other"]

REVIEW_STATUSES = ["Not requested", "Requested", "Received",
                   "No review left", "Declined", "Chased"]

WEDDING_STYLES = [
    "Classic", "Documentary", "Fine art", "Editorial", "Rustic",
    "Boho", "Modern", "Vintage", "Destination", "Elopement",
    "Micro wedding", "Traditional cultural", "Other",
]

PHOTO_STYLES = ["Light and airy", "Dark and moody", "True to colour",
                "Film emulation", "Black and white heavy",
                "Mixed - photographer's choice"]

DELIVERY_PREFERENCES = ["Online gallery only", "Gallery + USB",
                        "Gallery + album", "Gallery + prints",
                        "Everything"]

VENUE_TYPES = ["Hotel", "Barn", "Country house", "Church", "Registry office",
               "Marquee", "Vineyard", "Beach", "Garden", "Restaurant",
               "Castle", "Other"]

YES_NO = ["Yes", "No"]
YES_NO_NA = ["Yes", "No", "Not needed"]
YES_NO_PENDING = ["Yes", "No", "Pending"]

MONTH_NAMES = ["January", "February", "March", "April", "May", "June",
               "July", "August", "September", "October", "November",
               "December"]

WEEKDAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday",
                 "Saturday", "Sunday"]

#: (name, first month, last month) - winter wraps the year end
SEASONS: List[Tuple[str, int, int]] = [
    ("Winter", 12, 2), ("Spring", 3, 5),
    ("Summer", 6, 8), ("Autumn", 9, 11),
]

COUNTRIES = [
    "United States", "United Kingdom", "Canada", "Australia", "Ireland",
    "New Zealand", "France", "Spain", "Italy", "Germany", "Portugal",
    "Netherlands", "Greece", "Mexico", "India", "South Africa", "Other",
]

#: how far ahead a wedding counts as "coming up"
LEAD_TIME_BANDS = [
    ("This week", 0, 7), ("Next 30 days", 8, 30),
    ("1-3 months", 31, 90), ("3-6 months", 91, 180),
    ("6-12 months", 181, 365), ("Over a year", 366, 3650),
]

#: every named drop-down list the workbook defines
LIST_NAMES = [
    "LeadSourceList", "LeadStatusList", "ClientStatusList",
    "PriorityList", "BookingStatusList", "ContractStatusList",
    "DocStatusList", "PaymentStatusList", "PaymentStageList",
    "PaymentMethodList", "PackageList", "PackageTierList",
    "AlbumList", "AddOnList", "StaffRoleList", "PayBasisList",
    "AvailabilityList", "EditingStageList", "GalleryStatusList",
    "EquipmentCategoryList", "EquipmentConditionList",
    "InsuranceStatusList", "TravelModeList", "ReimbursementList",
    "MarketingPlatformList", "ContentTypeList", "ContentStatusList",
    "ExpenseCategoryList", "TaxCategoryList", "FrequencyList",
    "ReviewPlatformList", "ReviewStatusList", "WeddingStyleList",
    "PhotoStyleList", "DeliveryPreferenceList", "VenueTypeList",
    "YesNoList", "YesNoNAList", "YesNoPendingList", "CountryList",
    "MonthKeyList", "ClientList", "ClientNameList", "LeadList",
    "BookingList", "StaffList", "StaffNameList", "EquipmentList",
    "InvoiceList", "ShotGroupList", "TimelineEventList",
]

#: (label, default prefix, defined name, digits)
ID_PREFIX_FIELDS: List[Tuple[str, str, str, int]] = [
    ("Client", "CLT-", "ClientPrefix", 4),
    ("Lead", "LED-", "LeadPrefix", 4),
    ("Booking", "WED-", "BookingPrefix", 4),
    ("Package", "PKG-", "PackagePrefix", 2),
    ("Contract", "CON-", "ContractPrefix", 4),
    ("Invoice", "INV-", "InvoicePrefix", 4),
    ("Timeline row", "TLN-", "TimelinePrefix", 4),
    ("Shot", "SHT-", "ShotPrefix", 4),
    ("Photographer", "PHG-", "StaffPrefix", 3),
    ("Editing job", "EDT-", "EditPrefix", 4),
    ("Gallery", "GAL-", "GalleryPrefix", 4),
    ("Equipment", "EQP-", "EquipmentPrefix", 3),
    ("Travel", "TRV-", "TravelPrefix", 4),
    ("Marketing", "MKT-", "MarketingPrefix", 4),
    ("Expense", "EXP-", "ExpensePrefix", 4),
    ("Profit row", "PRF-", "ProfitPrefix", 4),
    ("Review", "REV-", "ReviewPrefix", 4),
]

ID_PREFIXES: Dict[str, Tuple[str, int]] = {
    "client": ("CLT-", 4), "lead": ("LED-", 4), "booking": ("WED-", 4),
    "package": ("PKG-", 2), "contract": ("CON-", 4),
    "invoice": ("INV-", 4), "timeline": ("TLN-", 4), "shot": ("SHT-", 4),
    "staff": ("PHG-", 3), "edit": ("EDT-", 4), "gallery": ("GAL-", 4),
    "equipment": ("EQP-", 3), "travel": ("TRV-", 4),
    "marketing": ("MKT-", 4), "expense": ("EXP-", 4),
    "profit": ("PRF-", 4), "review": ("REV-", 4),
}


# ---------------------------------------------------------------------------
# rows
# ---------------------------------------------------------------------------

#: the register the `--rows` switch scales everything else against
SCALE_ANCHOR = "Wedding Bookings"

#: weddings shown side by side on the dashboard and the reports
MAX_WEDDINGS_SHOWN = 10

#: days on the dashboard's "what is coming up" strip
UPCOMING_STRIP_DAYS = 14

DEFAULT_ROWS: Dict[str, int] = {
    "Client CRM": 300,
    "Lead Tracker": 400,
    "Wedding Bookings": 200,
    "Packages & Pricing": 30,
    "Contracts & Documents": 200,
    "Payments & Invoices": 600,
    "Wedding Timeline": 900,
    "Shot List Manager": 1400,
    "Photographer & Staff": 40,
    "Editing Workflow": 200,
    "Gallery & Delivery": 200,
    "Equipment Inventory": 150,
    "Travel & Mileage": 300,
    "Marketing Calendar": 400,
    "Expense Tracker": 600,
    "Profit Analysis": 200,
    "Reviews & Referrals": 200,
}


# ---------------------------------------------------------------------------
# palette
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class Palette:
    ink: str = "#2B2118"
    deep: str = "#4A3528"
    primary: str = "#8C6A4F"
    primary_mid: str = "#B08B68"
    primary_soft: str = "#F3EBE3"
    accent: str = "#A8577A"
    good: str = "#3F7D5C"
    warn: str = "#B8862A"
    bad: str = "#B23A48"
    muted: str = "#8A7D72"
    paper: str = "#FFFDFB"
    rule: str = "#D9CBBE"

    # ---- derived tones the formats and the writer ask for ----------
    # A warm neutral scheme: inputs are plain white so they read as
    # "type here", calculated cells sit a shade warmer, and the four
    # status colours are muted enough to survive a whole page of them.
    white: str = "#FFFFFF"
    input_bg: str = "#FFFFFF"
    input_border: str = "#E3D5C7"
    calc_bg: str = "#FBF7F3"
    calc_border: str = "#EDE2D7"
    line: str = "#D9CBBE"
    note_bg: str = "#FCF9F6"
    grey_text: str = "#6E6259"
    good_bg: str = "#DEEDE3"
    good_text: str = "#26563E"
    warn_bg: str = "#F8EBD3"
    warn_text: str = "#7F5B12"
    bad_bg: str = "#F7DCDE"
    bad_text: str = "#8E2531"
    info_bg: str = "#EDE4EC"
    info_text: str = "#6B3A56"
    accent_soft: str = "#F7E9EF"


PALETTE = Palette()


# ---------------------------------------------------------------------------
# the business
# ---------------------------------------------------------------------------

@dataclass
class Studio:
    name: str = "Rowan & Fern Photography"
    owner: str = "Marisa Calloway"
    email: str = "hello@rowanandfern.example"
    phone: str = "(503) 555-0148"
    website: str = "rowanandfern.example"
    instagram: str = "@rowanandfern"
    city: str = "Portland"
    region: str = "Oregon"
    country: str = "United States"


@dataclass
class Config:
    """Everything the build reads.  One object, passed everywhere."""

    studio: Studio = field(default_factory=Studio)

    author: str = AUTHOR
    password: str = PASSWORD
    protect_sheets: bool = True
    include_sample_data: bool = True

    currency: str = "$"
    currency_code: str = "USD"

    _today: Optional[_dt.date] = None
    year: Optional[int] = None

    # ---- pricing ----------------------------------------------------------
    default_package: str = "Classic Wedding"
    extra_hour_rate: float = 250.0
    second_shooter_rate: float = 350.0
    assistant_rate: float = 150.0
    engagement_session_price: float = 400.0
    album_base_price: float = 650.0
    mileage_rate: float = 0.67
    free_travel_miles: float = 60.0
    rush_delivery_fee: float = 350.0

    # ---- the payment schedule ---------------------------------------------
    deposit_percent: float = 0.30
    second_percent: float = 0.35
    deposit_days_after_booking: int = 7
    second_payment_days_before: int = 90
    final_payment_days_before: int = 14
    payment_terms_days: int = 14

    # ---- tax ---------------------------------------------------------------
    tax_rate: float = 0.0
    tax_name: str = "Sales tax"
    tax_on_packages: bool = False

    # ---- the workflow ------------------------------------------------------
    preview_days: int = 7
    delivery_days: int = 42
    album_days: int = 60
    gallery_months_live: int = 12
    editing_hours_per_wedding: float = 18.0
    editing_cost_per_hour: float = 35.0
    images_per_hour_culled: int = 400

    # ---- alert windows -----------------------------------------------------
    lead_chase_days: int = 3
    consultation_chase_days: int = 7
    upcoming_days: int = 30
    contract_warn_days: int = 30
    payment_chase_days: int = 7
    review_chase_days: int = 14
    maintenance_warn_days: int = 45
    warranty_warn_days: int = 60

    # ---- targets -----------------------------------------------------------
    target_weddings_year: int = 24
    target_booking_value: float = 2400.0
    target_revenue_year: float = 72000.0
    target_revenue_month: float = 6000.0
    target_profit_margin: float = 0.55
    target_conversion: float = 0.30
    target_rating: float = 4.9
    target_delivery_days: int = 42

    rows_override: Optional[int] = None

    # ---------------------------------------------------------------- dates --
    @property
    def today(self) -> _dt.date:
        return self._today or _dt.date.today()

    @today.setter
    def today(self, value: _dt.date) -> None:
        self._today = value

    @property
    def report_year(self) -> int:
        return self.year or self.today.year

    # ------------------------------------------------------------- formats --
    @property
    def money_format(self) -> str:
        return f'{self.currency}#,##0.00;[Red]-{self.currency}#,##0.00'

    @property
    def money_format_0(self) -> str:
        return f'{self.currency}#,##0;[Red]-{self.currency}#,##0'

    @property
    def money_format_4(self) -> str:
        return f'{self.currency}#,##0.0000'

    @property
    def date_format(self) -> str:
        return 'dd mmm yyyy'

    @property
    def time_format(self) -> str:
        """24 hour, because a wedding timeline that says 7:30 twice
        is how a photographer ends up at the church in the morning."""
        return 'hh:mm'

    # -------------------------------------------------------------- naming --
    @property
    def version(self) -> str:
        return VERSION

    @property
    def product_name(self) -> str:
        return PRODUCT_NAME

    @property
    def listing_title(self) -> str:
        return LISTING_TITLE

    @property
    def studio_name(self) -> str:
        return self.studio.name

    # ------------------------------------------------------------ packages --
    @property
    def package_table(self) -> List[Tuple]:
        return list(PACKAGES)

    @property
    def package_names(self) -> List[str]:
        return [row[0] for row in PACKAGES]

    def package_price(self, name: str) -> float:
        for row in PACKAGES:
            if row[0] == name:
                return float(row[9])
        return 0.0

    # ---------------------------------------------------------------- rows --
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

    # ----------------------------------------------------------------- ids --
    def rid(self, kind: str, row: int) -> str:
        """The ID a given record row will show, as plain text."""
        prefix, digits = ID_PREFIXES[kind]
        return f"{prefix}{row:0{digits}d}"

    def prefix(self, kind: str) -> str:
        return ID_PREFIXES[kind][0]

    def digits(self, kind: str) -> int:
        return ID_PREFIXES[kind][1]
````

### `wedding_photography_spreadsheet/layout.py`

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

### `wedding_photography_spreadsheet/formats.py`

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

### `wedding_photography_spreadsheet/writer.py`

*496 lines · 1 class · 8 functions*

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

**Module constants:** `P` (Palette), `NAV_LINKS` (12 items), `PROTECT_OPTIONS` (15 entries).

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
    ("Clients", "Client CRM"),
    ("Leads", "Lead Tracker"),
    ("Calendar", "Booking Calendar"),
    ("Weddings", "Wedding Bookings"),
    ("Payments", "Payments & Invoices"),
    ("Timeline", "Wedding Timeline"),
    ("Editing", "Editing Workflow"),
    ("Delivery", "Gallery & Delivery"),
    ("Profit", "Profit Analysis"),
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

    def cf(self, cell_range: str, rule: Dict[str, Any]):
        """Apply one conditional format to a raw A1 range.

        Registers get their conditional formats through ``grid`` via
        ``Col.cf``.  Pages that build their own grids - the booking
        calendar, mainly - need this instead.  It resolves the ``fmt``
        style name exactly as ``grid`` does, because passing a Format
        object under the key ``fmt`` is silently ignored by XlsxWriter
        and the colour just never appears.
        """
        options = dict(rule)
        fmt_key = options.pop("fmt", None)
        if fmt_key is not None:
            options["format"] = (self.f.cf(fmt_key)
                                 if isinstance(fmt_key, str)
                                 else fmt_key)
        if "criteria" in options and isinstance(options["criteria"], str):
            options["criteria"] = self.reg.render(
                options["criteria"]).strip()
        self.ws.conditional_format(cell_range, options)

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

### `wedding_photography_spreadsheet/names.py`

*192 lines · 1 function*

> Workbook level defined names that point at the register sheets.
>
> The fixed drop-downs - venue types, payment methods, expense categories
> - define their own names on `Settings & Lists`.  The names here are the
> ones that have to follow live data: the Booking IDs, the couples, the
> photographers, the packages you actually sell and the marketing posts
> you actually made.  Each is an ``OFFSET`` range that grows and shrinks
> with the filled rows, which is why a wedding you book this morning is
> in the payments drop-down this afternoon.
>
> `BookingList` is the important one.  It is the Booking ID column from
> `Wedding Bookings`, and it feeds the drop-down on twelve other tabs.
> Every join in this workbook runs through it.

| Defines | What it does |
|---------|--------------|
| `define_cross_sheet_names(book, reg)` | Create the live lookup names |

**Module constants:** `CROSS_SHEET_NAMES` (135 entries).

````python
"""Workbook level defined names that point at the register sheets.

The fixed drop-downs - venue types, payment methods, expense categories
- define their own names on `Settings & Lists`.  The names here are the
ones that have to follow live data: the Booking IDs, the couples, the
photographers, the packages you actually sell and the marketing posts
you actually made.  Each is an ``OFFSET`` range that grows and shrinks
with the filled rows, which is why a wedding you book this morning is
in the payments drop-down this afternoon.

`BookingList` is the important one.  It is the Booking ID column from
`Wedding Bookings`, and it feeds the drop-down on twelve other tabs.
Every join in this workbook runs through it.
"""

from __future__ import annotations

from typing import Dict, Tuple

from .layout import Registry, dynamic_name

#: defined name -> (sheet, column key)
CROSS_SHEET_NAMES: Dict[str, Tuple[str, str]] = {
    # ---- the join keys, in order of how much work they do -------------
    "BookingList": ("Wedding Bookings", "id"),
    "ClientList": ("Client CRM", "id"),
    "ClientNameList": ("Client CRM", "name"),
    "LeadList": ("Lead Tracker", "id"),
    "StaffList": ("Photographer & Staff", "id"),
    "StaffNameList": ("Photographer & Staff", "name"),
    "PackageList": ("Packages & Pricing", "name"),
    "PackageNameList": ("Packages & Pricing", "name"),
    "CampaignList": ("Marketing Calendar", "id"),
    "EquipmentList": ("Equipment Inventory", "id"),

    # ---- package columns the pricing helpers read ---------------------
    "PackagePriceCol": ("Packages & Pricing", "total_price"),
    "PackageHoursCol": ("Packages & Pricing", "hours"),
    "PackageImagesCol": ("Packages & Pricing", "images"),
    "PackageMarginCol": ("Packages & Pricing", "margin"),
    "PackageSoldCol": ("Packages & Pricing", "sold"),
    "PackageRevenueCol": ("Packages & Pricing", "revenue"),
    "PackageActiveCol": ("Packages & Pricing", "active"),
    "PackageTierCol": ("Packages & Pricing", "tier"),

    # ---- booking columns - the busiest ranges in the workbook ---------
    "BookingClient": ("Wedding Bookings", "client"),
    "BookingCouple": ("Wedding Bookings", "client_name"),
    "BookingDate": ("Wedding Bookings", "wedding_date"),
    "BookingDateNum": ("Wedding Bookings", "wed_num"),
    "BookingStatusCol": ("Wedding Bookings", "status"),
    "BookingPackage": ("Wedding Bookings", "package"),
    "BookingVenue": ("Wedding Bookings", "venue"),
    "BookingVenueType": ("Wedding Bookings", "venue_type"),
    "BookingPhotographer": ("Wedding Bookings", "photographer"),
    "BookingSecond": ("Wedding Bookings", "second"),
    "BookingMonths": ("Wedding Bookings", "month"),
    "BookingYears": ("Wedding Bookings", "year"),
    "BookingSeason": ("Wedding Bookings", "season"),
    "BookingValue": ("Wedding Bookings", "total_value"),
    "BookingCollected": ("Wedding Bookings", "collected"),
    "BookingOutstanding": ("Wedding Bookings", "outstanding"),
    "BookingProfit": ("Wedding Bookings", "profit"),
    "BookingMargin": ("Wedding Bookings", "margin"),
    "BookingHours": ("Wedding Bookings", "coverage_hours"),
    "BookingHolds": ("Wedding Bookings", "holds"),
    "BookingEarns": ("Wedding Bookings", "earns"),
    "BookingShot": ("Wedding Bookings", "shot"),
    "BookingDaysAway": ("Wedding Bookings", "days_away"),
    "BookingAlert": ("Wedding Bookings", "alert"),
    "BookingRowNum": ("Wedding Bookings", "seq_num"),
    "BookingProfitRank": ("Wedding Bookings", "profit_rank"),
    "BookingCeremony": ("Wedding Bookings", "ceremony_time"),
    "BookingContract": ("Wedding Bookings", "contract_state"),
    "BookingGallery": ("Wedding Bookings", "gallery_state"),
    "BookingEditing": ("Wedding Bookings", "editing_state"),

    # ---- the enquiry funnel -------------------------------------------
    "LeadDate": ("Lead Tracker", "enquiry_date"),
    "LeadDateNum": ("Lead Tracker", "enq_num"),
    "LeadSourceCol": ("Lead Tracker", "source"),
    "LeadStatusCol": ("Lead Tracker", "status"),
    "LeadMonths": ("Lead Tracker", "month"),
    "LeadOpen": ("Lead Tracker", "is_open"),
    "LeadWon": ("Lead Tracker", "is_won"),
    "LeadValue": ("Lead Tracker", "booked_value"),
    "LeadCampaign": ("Lead Tracker", "campaign"),
    "LeadWeddingDate": ("Lead Tracker", "wedding_date"),

    # ---- money in and money out ---------------------------------------
    "PaymentBooking": ("Payments & Invoices", "booking"),
    "PaymentMonths": ("Payments & Invoices", "month"),
    "PaymentDueNum": ("Payments & Invoices", "due_num"),
    "PaymentPaidNum": ("Payments & Invoices", "paid_num"),
    "PaymentAmount": ("Payments & Invoices", "amount"),
    "PaymentReceived": ("Payments & Invoices", "received"),
    "PaymentNet": ("Payments & Invoices", "net"),
    "PaymentBalance": ("Payments & Invoices", "balance"),
    "PaymentFee": ("Payments & Invoices", "fee"),
    "PaymentStatusCol": ("Payments & Invoices", "status"),
    "PaymentLate": ("Payments & Invoices", "is_late"),
    "ExpenseMonths": ("Expense Tracker", "month"),
    "ExpenseYears": ("Expense Tracker", "year"),
    "ExpenseDateNum": ("Expense Tracker", "date_num"),
    "ExpenseCategoryCol": ("Expense Tracker", "category"),
    "ExpenseAmounts": ("Expense Tracker", "amount"),
    "ExpenseClaimable": ("Expense Tracker", "claimable"),
    "ExpenseBooking": ("Expense Tracker", "booking"),
    "ExpenseOverhead": ("Expense Tracker", "is_overhead"),
    "ExpenseCapital": ("Expense Tracker", "is_capital"),
    "ExpenseMonthly": ("Expense Tracker", "monthly"),

    # ---- the work -------------------------------------------------------
    "TimelineBooking": ("Wedding Timeline", "booking"),
    "TimelineKey": ("Wedding Timeline", "day_key"),
    "TimelineStart": ("Wedding Timeline", "start"),
    "TimelineFinish": ("Wedding Timeline", "finish"),
    "TimelineShown": ("Wedding Timeline", "shown"),
    "TimelineWho": ("Wedding Timeline", "staff_name"),
    "TimelineWhere": ("Wedding Timeline", "location"),
    "TimelineCritical": ("Wedding Timeline", "critical"),
    "TimelineDone": ("Wedding Timeline", "done"),
    "ShotBooking": ("Shot List Manager", "booking"),
    "ShotGroupCol": ("Shot List Manager", "group_name"),
    "ShotText": ("Shot List Manager", "shot"),
    "ShotPeople": ("Shot List Manager", "people"),
    "ShotMustHave": ("Shot List Manager", "must_have"),
    "ShotDone": ("Shot List Manager", "done"),
    "ShotOpen": ("Shot List Manager", "open"),
    "EditBooking": ("Editing Workflow", "booking"),
    "EditStage": ("Editing Workflow", "stage"),
    "EditProgress": ("Editing Workflow", "progress"),
    "EditDue": ("Editing Workflow", "due"),
    "EditDueNum": ("Editing Workflow", "due_num"),
    "EditLate": ("Editing Workflow", "late"),
    "EditHours": ("Editing Workflow", "hours"),
    "EditCost": ("Editing Workflow", "cost"),
    "GalleryBooking": ("Gallery & Delivery", "booking"),
    "GalleryStatusCol": ("Gallery & Delivery", "status"),
    "GalleryDelivered": ("Gallery & Delivery", "is_delivered"),
    "GalleryDeliveredNum": ("Gallery & Delivery", "delivered_num"),
    "GalleryAfterSales": ("Gallery & Delivery", "after_total"),
    "GalleryExpires": ("Gallery & Delivery", "expires"),

    # ---- everything else the dashboard and the reports scan -------------
    "TravelBooking": ("Travel & Mileage", "booking"),
    "TravelMonths": ("Travel & Mileage", "month"),
    "TravelMiles": ("Travel & Mileage", "all_miles"),
    "TravelCost": ("Travel & Mileage", "total_cost"),
    "TravelCharged": ("Travel & Mileage", "chargeable"),
    "MarketingMonths": ("Marketing Calendar", "month"),
    "MarketingPlatformCol": ("Marketing Calendar", "platform"),
    "MarketingSpend": ("Marketing Calendar", "true_cost"),
    "MarketingRevenue": ("Marketing Calendar", "revenue"),
    "MarketingEnquiries": ("Marketing Calendar", "all_enquiries"),
    "ReviewBooking": ("Reviews & Referrals", "booking"),
    "ReviewMonths": ("Reviews & Referrals", "month"),
    "ReviewRating": ("Reviews & Referrals", "rating"),
    "ReviewPhotographer": ("Reviews & Referrals", "photographer"),
    "ReviewQuote": ("Reviews & Referrals", "quote"),
    "ReviewHas": ("Reviews & Referrals", "has_review"),
    "ContractBooking": ("Contracts & Documents", "booking"),
    "ContractStatusCol": ("Contracts & Documents", "status"),
    "ContractDone": ("Contracts & Documents", "is_done"),
    "EquipmentStatusCol": ("Equipment Inventory", "status"),
    "EquipmentValue": ("Equipment Inventory", "price"),
    "EquipmentInsured": ("Equipment Inventory", "insured"),
    "ProfitBooking": ("Profit Analysis", "booking"),
    "ProfitNet": ("Profit Analysis", "net_profit"),
    "ProfitPerHour": ("Profit Analysis", "per_hour"),
    "ProfitRankCol": ("Profit Analysis", "net_profit_rank"),
}


def define_cross_sheet_names(book, reg: Registry) -> None:
    """Create the live lookup names. Call after every register is built."""
    missing = []
    for name, (sheet, key) in CROSS_SHEET_NAMES.items():
        if sheet not in reg:
            missing.append(f"{name}: no sheet {sheet!r}")
            continue
        spec = reg[sheet]
        if not any(c.key == key for c in spec.columns):
            missing.append(f"{name}: {sheet!r} has no column {key!r}")
            continue
        book.define_name(
            name,
            dynamic_name(sheet, spec.letter(key), spec.first_data_row,
                         spec.last_data_row))
    if missing:
        raise KeyError("defined names point at nothing: "
                       + "; ".join(missing))
````

### `wedding_photography_spreadsheet/sample_data.py`

*2,038 lines · 27 functions*

> A worked example: one studio, fifteen weddings, every tab populated.
>
> The studio is Rowan & Fern Photography and the year it is having is
> not a flattering one.  There is a wedding from three months ago still
> sitting in editing past the promised turnaround, a couple who never
> paid their balance, a date that has been promised to two couples at
> once, and an elopement that was priced on instinct and works out at
> less per hour than a supermarket shift.  A sample studio where
> everything runs smoothly teaches nobody anything; this one teaches you
> what each alert column is actually for.
>
> Every row here is checked twice before it ships.  The keys must all be
> *input* columns on their register - a value written into a formula
> column is silently dropped and takes the formula with it - and every
> value that feeds a drop-down must exist in that drop-down's list.
> ``tools/check_wedding_sample.py`` does both checks.
>
> Dates are all relative to ``cfg.today``, so the example stays current
> no matter when the workbook is built.  Weddings land on Saturdays
> because real ones mostly do.

| Defines | What it does |
|---------|--------------|
| `_day(cfg, offset)` | A date `offset` days from the build date |
| `_time(hour, minute)` |  |
| `_clt(n)` |  |
| `_wed(n)` |  |
| `_phg(n)` |  |
| `_mkt(n)` |  |
| `_booking_status(index)` |  |
| `clients(cfg)` |  |
| `leads(cfg)` |  |
| `bookings(cfg)` |  |
| `packages(cfg)` |  |
| `staff(cfg)` |  |
| `contracts(cfg)` |  |
| `payments(cfg)` |  |
| `_timeline_for(cfg, index, keys, *, done, drift)` | One wedding's running order, drawn from the standard plan |
| `timeline(cfg)` |  |
| `_shot_group(label)` | Map a config shot group onto a Shot group drop-down value |
| `shots(cfg)` |  |
| `editing(cfg)` |  |
| `gallery(cfg)` |  |
| `equipment(cfg)` |  |
| `travel(cfg)` |  |
| `marketing(cfg)` |  |
| `expenses(cfg)` |  |
| `profit(cfg)` | One line per wedding |
| `reviews(cfg)` |  |
| `get(name, cfg)` | Sample rows for a register, or [] when samples are switched off |

**Module constants:** `WEDDINGS` (15 items).

````python
"""A worked example: one studio, fifteen weddings, every tab populated.

The studio is Rowan & Fern Photography and the year it is having is
not a flattering one.  There is a wedding from three months ago still
sitting in editing past the promised turnaround, a couple who never
paid their balance, a date that has been promised to two couples at
once, and an elopement that was priced on instinct and works out at
less per hour than a supermarket shift.  A sample studio where
everything runs smoothly teaches nobody anything; this one teaches you
what each alert column is actually for.

Every row here is checked twice before it ships.  The keys must all be
*input* columns on their register - a value written into a formula
column is silently dropped and takes the formula with it - and every
value that feeds a drop-down must exist in that drop-down's list.
``tools/check_wedding_sample.py`` does both checks.

Dates are all relative to ``cfg.today``, so the example stays current
no matter when the workbook is built.  Weddings land on Saturdays
because real ones mostly do.
"""

from __future__ import annotations

import datetime as _dt
from typing import Any, Callable, Dict, List

from .config import SHOT_LIST, TIMELINE_EVENTS, Config

Rows = List[Dict[str, Any]]


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------

def _day(cfg: Config, offset: int) -> _dt.date:
    """A date `offset` days from the build date."""
    return cfg.today + _dt.timedelta(days=offset)


def _time(hour: int, minute: int = 0) -> _dt.time:
    return _dt.time(hour % 24, minute)


def _clt(n: int) -> str:
    return f"CLT-{n:04d}"


def _wed(n: int) -> str:
    return f"WED-{n:04d}"


def _phg(n: int) -> str:
    return f"PHG-{n:03d}"


def _mkt(n: int) -> str:
    return f"MKT-{n:04d}"


# The fifteen weddings, in the row order they appear in the register.
# (offset from today, couple, venue, venue type, city, guests, package)
# Offsets are all congruent to 6 mod 7, which is a Saturday when the
# build date is a Sunday, and close enough to one otherwise.
WEDDINGS = [
    (-400, "Nadia & Priya Raman-Okafor", "The Old Mill Barn", "Barn",
     "Hood River", 95, "Classic Wedding"),
    (-302, "Hollis & Jun Tanaka-Reed", "Cascade Vineyard Estate",
     "Vineyard", "Dundee", 130, "Premium Wedding"),
    (-190, "Imogen & Bram Castellanos", "St Agnes Church",
     "Church", "Portland", 150, "Premium Wedding"),
    (-92, "Saoirse & Eliot Vance", "Willow Creek Country House",
     "Country house", "Newberg", 110, "Classic Wedding"),
    (-64, "Margit & Tobias Lindqvist", "The Beaumont Hotel",
     "Hotel", "Portland", 80, "Classic Wedding"),
    (-29, "Delphine & Oscar Mbeki", "Rosegate Gardens",
     "Garden", "Lake Oswego", 70, "Essential Wedding"),
    (-1, "Yusra & Kaleem Haddad", "The Granary at Fir Point",
     "Barn", "Sherwood", 165, "Luxury Wedding"),
    (6, "Tessa & Rowan Whitfield", "Marchmont House",
     "Country house", "Salem", 100, "Classic Wedding"),
    (20, "Camille & Esme Fontaine", "Harbour Light Hotel",
     "Hotel", "Astoria", 120, "Premium Wedding"),
    (55, "Juno & Felix Arbuthnot", "Crater Rim Overlook",
     "Other", "Bend", 8, "Essential Wedding"),
    (118, "Beatrix & Idris Oyelaran", "The Winter Orangery",
     "Garden", "Portland", 90, "Classic Wedding"),
    (181, "Anneliese & Caspar von Hoff", "Thornbury Castle",
     "Castle", "Victoria", 180, "Luxury Wedding"),
    (181, "Marnie & Gus Pendleton", "Fernbrook Marquee",
     "Marquee", "Hillsboro", 140, "Premium Wedding"),
    (251, "Océane & Théo Belanger", "Villa Dorado",
     "Other", "Oaxaca", 60, "Luxury Wedding"),
    (342, "Sunniva & Abel Mackintosh", "The Assembly Rooms",
     "Other", "Portland", 115, "Classic Wedding"),
]

# Which booking is in which state.  Index is 1-based, matching WED-nnnn.
_DONE = (1, 2, 3)              # delivered, paid, reviewed
_EDITING_LATE = (4,)           # past the promised turnaround
_DELIVERED_WAITING = (5,)      # delivered, no review yet
_EDITING_OK = (6,)             # inside the turnaround, balance unpaid
_JUST_SHOT = (7,)              # cards backed up, nothing else
_PAST = _DONE + _EDITING_LATE + _DELIVERED_WAITING + _EDITING_OK \
    + _JUST_SHOT


def _booking_status(index: int) -> str:
    if index in _DONE:
        return "Complete"
    if index in _EDITING_LATE or index in _EDITING_OK:
        return "In editing"
    if index in _DELIVERED_WAITING:
        return "Delivered"
    if index in _JUST_SHOT:
        return "Wedding done"
    if index in (13, 15):
        return "Pencilled in"
    return "Confirmed"


# ---------------------------------------------------------------------------
# 1. Client CRM
# ---------------------------------------------------------------------------

_CLIENT_EXTRA = [
    # style, photo_style, delivery_pref, source, referred_by, priority
    ("Rustic", "Light and airy", "Gallery + album", "Instagram", "",
     "Normal"),
    ("Fine art", "Film emulation", "Everything", "Venue recommendation",
     "Cascade Vineyard Estate", "Normal"),
    ("Traditional cultural", "True to colour", "Gallery + prints",
     "Referral", "Nadia Raman-Okafor", "Normal"),
    ("Documentary", "True to colour", "Gallery + USB", "Website", "",
     "Normal"),
    ("Modern", "Dark and moody", "Online gallery only", "Google search",
     "", "Normal"),
    ("Boho", "Light and airy", "Gallery + USB", "Pinterest", "", "Low"),
    ("Editorial", "Dark and moody", "Everything", "Wedding fair", "",
     "High"),
    ("Classic", "True to colour", "Gallery + album", "Referral",
     "Hollis Tanaka-Reed", "High"),
    ("Documentary", "Light and airy", "Gallery + album", "Instagram",
     "", "Normal"),
    ("Elopement", "Film emulation", "Online gallery only", "TikTok", "",
     "Low"),
    ("Vintage", "Light and airy", "Gallery + prints",
     "Venue recommendation", "The Winter Orangery", "Normal"),
    ("Fine art", "Light and airy", "Everything", "Wedding directory",
     "", "High"),
    ("Rustic", "True to colour", "Gallery + USB", "Instagram", "",
     "Normal"),
    ("Destination", "Film emulation", "Everything", "Past client",
     "Imogen Castellanos", "High"),
    ("Modern", "Mixed - photographer's choice", "Gallery + album",
     "Word of mouth", "", "Normal"),
]

_CLIENT_STATUS = {
    1: "Complete", 2: "Complete", 3: "Complete", 4: "Wedding done",
    5: "Delivered", 6: "Wedding done", 7: "Wedding done",
}

_CLIENT_NOTES = {
    4: "Chasing me for the gallery. I am three weeks past my own "
       "promise and I have no good answer.",
    6: "Lovely couple, still owe the balance. Third reminder went "
       "out and I have not phoned yet.",
    8: "Questionnaire still not back and the wedding is next "
       "Saturday. Ring them, do not email again.",
    10: "Priced this on the phone in two minutes. See Profit "
        "Analysis before quoting the next elopement.",
    13: "Pencilled only. Same date as Anneliese & Caspar - whoever "
        "signs first gets it and I need to tell the other today.",
}


def clients(cfg: Config) -> Rows:
    rows: Rows = []
    for index, (offset, couple, venue, _vtype, city, _guests, _pkg) \
            in enumerate(WEDDINGS, start=1):
        first, second = couple.split(" & ")
        style, photo, delivery, source, referred, priority = \
            _CLIENT_EXTRA[index - 1]
        surname = second.split(" ")[-1]
        handle = first.split(" ")[0].lower()
        rows.append({
            "name": f"{first} {surname}" if " " not in first
                    else first,
            "partner": second,
            "email": f"{handle}.{surname.lower().replace(' ', '')}"
                     f"@example.com",
            "phone": f"+1 503 555 {1000 + index * 37:04d}",
            "country": "Mexico" if city == "Oaxaca"
                       else ("Canada" if city == "Victoria"
                             else "United States"),
            "status": _CLIENT_STATUS.get(index, "Booked"),
            "priority": priority,
            "wedding_date": _day(cfg, offset),
            "venue": venue,
            "venue_city": city,
            "source": source,
            "referred_by": referred,
            "first_contact": _day(cfg, offset - 330),
            "style": style,
            "photo_style": photo,
            "delivery_pref": delivery,
            "notes": _CLIENT_NOTES.get(index, ""),
            "marketing_ok": "No" if index in (3, 12) else "Yes",
        })

    # three live enquiries with no booking behind them yet
    rows += [
        {"name": "Wren Alcott", "partner": "Mateo Alcott",
         "email": "wren.alcott@example.com",
         "phone": "+1 503 555 6612", "country": "United States",
         "status": "Enquiry", "priority": "High",
         "wedding_date": _day(cfg, 236), "venue": "Still looking",
         "venue_city": "Portland", "source": "Instagram",
         "referred_by": "", "first_contact": _day(cfg, -11),
         "style": "Documentary", "photo_style": "Light and airy",
         "delivery_pref": "Gallery + album",
         "notes": "Consultation booked. Budget is real, date is a "
                  "Friday, I have it free.",
         "marketing_ok": "Yes"},
        {"name": "Fenella Brightwater", "partner": "Alice Nkemdirim",
         "email": "fenella.b@example.com",
         "phone": "+1 503 555 7741", "country": "United States",
         "status": "Enquiry", "priority": "Normal",
         "wedding_date": _day(cfg, 181), "venue": "Fernbrook Marquee",
         "venue_city": "Hillsboro", "source": "Venue recommendation",
         "referred_by": "Fernbrook Marquee", "first_contact":
         _day(cfg, -4),
         "style": "Boho", "photo_style": "Light and airy",
         "delivery_pref": "Online gallery only",
         "notes": "Third enquiry for that same Saturday. The date is "
                  "worth more than the package - price it that way.",
         "marketing_ok": "Yes"},
        {"name": "Isolde Marchetti", "partner": "Rafe Marchetti",
         "email": "isolde.m@example.com",
         "phone": "+1 503 555 8890", "country": "United States",
         "status": "Enquiry", "priority": "Low",
         "wedding_date": _day(cfg, 475), "venue": "Undecided",
         "venue_city": "Eugene", "source": "Google search",
         "referred_by": "", "first_contact": _day(cfg, -2),
         "style": "Classic", "photo_style": "True to colour",
         "delivery_pref": "Gallery + USB",
         "notes": "Eighteen months out and price-shopping. Worth one "
                  "good reply, not three.",
         "marketing_ok": "Yes"},
    ]
    return rows


# ---------------------------------------------------------------------------
# 2. Lead Tracker
# ---------------------------------------------------------------------------

def leads(cfg: Config) -> Rows:
    rows: Rows = []

    # the enquiries that turned into the fifteen bookings
    won = [
        (1, "Instagram", _mkt(2), "", 1400, 42),
        (2, "Venue recommendation", "", "", 2600, 31),
        (3, "Referral", "", _wed(1), 2700, 18),
        (4, "Website", _mkt(5), "", 1900, 26),
        (5, "Google search", "", "", 1800, 35),
        (6, "Pinterest", _mkt(8), "", 1300, 21),
        (7, "Wedding fair", _mkt(11), "", 3800, 47),
        (8, "Referral", "", _wed(2), 2000, 14),
        (9, "Instagram", _mkt(14), "", 2600, 29),
        (10, "TikTok", _mkt(17), "", 1250, 9),
        (11, "Venue recommendation", "", "", 1950, 23),
        (12, "Wedding directory", _mkt(20), "", 3900, 52),
        (13, "Instagram", _mkt(22), "", 2550, 16),
        (14, "Past client", "", _wed(3), 4100, 38),
        (15, "Word of mouth", "", "", 1900, 27),
    ]
    for index, source, campaign, referral, budget, lag in won:
        offset, couple, venue, _vt, _city, guests, package = \
            WEDDINGS[index - 1]
        first, second = couple.split(" & ")
        enquiry = offset - 330
        rows.append({
            "name": first, "partner": second,
            "email": f"{first.split(' ')[0].lower()}"
                     f".{second.split(' ')[-1].lower()}@example.com",
            "phone": f"+1 503 555 {1000 + index * 37:04d}",
            "client": _clt(index),
            "enquiry_date": _day(cfg, enquiry),
            "source": source, "campaign": campaign,
            "referred_by": referral,
            "wedding_date": _day(cfg, offset),
            "venue": venue, "guests": guests, "package": package,
            "budget": budget, "status": "Booked", "priority": "Normal",
            "replied": _day(cfg, enquiry + 1),
            "consultation": _day(cfg, enquiry + 6),
            "proposal_sent": _day(cfg, enquiry + 7),
            "decision_date": _day(cfg, enquiry + lag),
            "lost_reason": "", "follow_up": "",
            "notes": "",
        })

    # the ones that got away, and the ones still in play
    rows += [
        {"name": "Wren Alcott", "partner": "Mateo Alcott",
         "email": "wren.alcott@example.com",
         "phone": "+1 503 555 6612", "client": _clt(16),
         "enquiry_date": _day(cfg, -11), "source": "Instagram",
         "campaign": _mkt(24), "referred_by": "",
         "wedding_date": _day(cfg, 236), "venue": "Still looking",
         "guests": 85, "package": "Premium Wedding", "budget": 2800,
         "status": "Consultation booked", "priority": "High",
         "replied": _day(cfg, -11), "consultation": _day(cfg, 3),
         "proposal_sent": "", "decision_date": "", "lost_reason": "",
         "follow_up": _day(cfg, 4),
         "notes": "Best enquiry this month. Send the Premium "
                  "comparison before the call, not after."},
        {"name": "Fenella Brightwater", "partner": "Alice Nkemdirim",
         "email": "fenella.b@example.com",
         "phone": "+1 503 555 7741", "client": _clt(17),
         "enquiry_date": _day(cfg, -4),
         "source": "Venue recommendation", "campaign": "",
         "referred_by": "", "wedding_date": _day(cfg, 181),
         "venue": "Fernbrook Marquee", "guests": 110,
         "package": "Classic Wedding", "budget": 2100,
         "status": "Proposal sent", "priority": "Normal",
         "replied": _day(cfg, -4), "consultation": _day(cfg, -2),
         "proposal_sent": _day(cfg, -1), "decision_date": "",
         "lost_reason": "", "follow_up": _day(cfg, 6),
         "notes": "Third couple wanting that Saturday."},
        {"name": "Isolde Marchetti", "partner": "Rafe Marchetti",
         "email": "isolde.m@example.com",
         "phone": "+1 503 555 8890", "client": _clt(18),
         "enquiry_date": _day(cfg, -2), "source": "Google search",
         "campaign": "", "referred_by": "",
         "wedding_date": _day(cfg, 475), "venue": "Undecided",
         "guests": 60, "package": "", "budget": 0,
         "status": "New", "priority": "Low",
         "replied": "", "consultation": "", "proposal_sent": "",
         "decision_date": "", "lost_reason": "", "follow_up": "",
         "notes": "Not replied yet. This is the row the dashboard "
                  "is shouting about."},
        {"name": "Corrine Vasquez", "partner": "Dov Perlman",
         "email": "corrine.v@example.com",
         "phone": "+1 503 555 2214", "client": "",
         "enquiry_date": _day(cfg, -23), "source": "Instagram",
         "campaign": _mkt(22), "referred_by": "",
         "wedding_date": _day(cfg, 194), "venue": "The Beaumont Hotel",
         "guests": 140, "package": "Premium Wedding", "budget": 2500,
         "status": "Lost", "priority": "Normal",
         "replied": _day(cfg, -23), "consultation": _day(cfg, -18),
         "proposal_sent": _day(cfg, -17),
         "decision_date": _day(cfg, -9),
         "lost_reason": "Went with a cheaper photographer",
         "follow_up": "",
         "notes": "Lost on price by about $400. They asked twice "
                  "whether I could 'do it for less' - I should have "
                  "offered fewer hours, not a discount."},
        {"name": "Perpetua Odili", "partner": "Sam Odili",
         "email": "p.odili@example.com",
         "phone": "+1 503 555 3390", "client": "",
         "enquiry_date": _day(cfg, -37), "source": "Wedding fair",
         "campaign": _mkt(11), "referred_by": "",
         "wedding_date": _day(cfg, 118), "venue": "The Winter Orangery",
         "guests": 90, "package": "Classic Wedding", "budget": 1800,
         "status": "Date unavailable", "priority": "Normal",
         "replied": _day(cfg, -37), "consultation": "",
         "proposal_sent": "", "decision_date": _day(cfg, -36),
         "lost_reason": "Date already booked",
         "follow_up": "",
         "notes": "Offered them an associate shooter and they said "
                  "no. Worth asking the question anyway."},
        {"name": "Halle Bergstrom", "partner": "Nils Bergstrom",
         "email": "halle.b@example.com",
         "phone": "+1 503 555 4418", "client": "",
         "enquiry_date": _day(cfg, -19), "source": "Pinterest",
         "campaign": _mkt(8), "referred_by": "",
         "wedding_date": _day(cfg, 320), "venue": "Cascade Vineyard "
                                                 "Estate",
         "guests": 70, "package": "Classic Wedding", "budget": 1600,
         "status": "Contacted", "priority": "Normal",
         "replied": _day(cfg, -19), "consultation": "",
         "proposal_sent": "", "decision_date": "", "lost_reason": "",
         "follow_up": _day(cfg, -5),
         "notes": "Follow-up date has already passed. Chase today."},
        {"name": "Ottoline Reyes", "partner": "Gabriel Reyes",
         "email": "ottoline.r@example.com",
         "phone": "+1 503 555 5502", "client": "",
         "enquiry_date": _day(cfg, -6), "source": "Referral",
         "campaign": "", "referred_by": _wed(5),
         "wedding_date": _day(cfg, 285), "venue": "Marchmont House",
         "guests": 100, "package": "Premium Wedding", "budget": 2600,
         "status": "Consultation done", "priority": "High",
         "replied": _day(cfg, -6), "consultation": _day(cfg, -1),
         "proposal_sent": "", "decision_date": "", "lost_reason": "",
         "follow_up": _day(cfg, 1),
         "notes": "Proposal promised within 48 hours. Margit sent "
                  "them - thank her either way."},
        {"name": "Davina Achebe", "partner": "Kit Achebe",
         "email": "davina.a@example.com",
         "phone": "+1 503 555 9931", "client": "",
         "enquiry_date": _day(cfg, -58), "source": "Website",
         "campaign": _mkt(5), "referred_by": "",
         "wedding_date": _day(cfg, 410), "venue": "Thornbury Castle",
         "guests": 200, "package": "Luxury Wedding", "budget": 4500,
         "status": "Proposal sent", "priority": "High",
         "replied": _day(cfg, -58), "consultation": _day(cfg, -51),
         "proposal_sent": _day(cfg, -49), "decision_date": "",
         "lost_reason": "", "follow_up": _day(cfg, -21),
         "notes": "Biggest open proposal on the board and I have let "
                  "it go cold for seven weeks."},
        {"name": "Minna Soderberg", "partner": "Ray Soderberg",
         "email": "minna.s@example.com",
         "phone": "+1 503 555 1177", "client": "",
         "enquiry_date": _day(cfg, -71), "source": "Facebook",
         "campaign": "", "referred_by": "",
         "wedding_date": _day(cfg, 160), "venue": "Rosegate Gardens",
         "guests": 45, "package": "Essential Wedding", "budget": 900,
         "status": "Not a fit", "priority": "Low",
         "replied": _day(cfg, -70), "consultation": "",
         "proposal_sent": "", "decision_date": _day(cfg, -69),
         "lost_reason": "Budget below the Essential package",
         "follow_up": "",
         "notes": "Sent them two photographers who shoot at that "
                  "level. Costs nothing, comes back eventually."},
    ]
    return rows


# ---------------------------------------------------------------------------
# 3. Wedding Bookings
# ---------------------------------------------------------------------------

_BOOKING_TEAM = {
    # index: (lead, second)
    1: (_phg(1), _phg(2)), 2: (_phg(1), _phg(2)),
    3: (_phg(1), _phg(3)), 4: (_phg(1), _phg(2)),
    5: (_phg(1), ""), 6: (_phg(1), ""),
    7: (_phg(1), _phg(2)), 8: (_phg(1), _phg(3)),
    9: (_phg(1), _phg(2)), 10: (_phg(1), ""),
    11: (_phg(1), _phg(3)), 12: (_phg(1), _phg(2)),
    13: (_phg(3), _phg(5)), 14: (_phg(1), _phg(2)),
    15: (_phg(1), ""),
}

_BOOKING_MONEY = {
    # index: (addons, travel_fee, discount)
    1: (250, 0, 0), 2: (650, 0, 100), 3: (400, 45, 0),
    4: (250, 0, 150), 5: (0, 0, 0), 6: (0, 0, 0),
    7: (900, 120, 0), 8: (250, 0, 0), 9: (650, 95, 0),
    10: (0, 210, 0), 11: (250, 0, 0), 12: (1150, 380, 0),
    13: (400, 0, 0), 14: (650, 1250, 0), 15: (0, 0, 100),
}

_MUST_HAVES = {
    7: "Grandmother Haddad with both grandsons - she is 94 and "
       "will not last the evening.\nThe nikah signing, wide and "
       "close.\nKaleem's mother's gold bangles on the dress.",
    8: "Tessa's late father's watch on Rowan's wrist.\n"
       "The dog, Biscuit, in at least one family group.",
    9: "Both brides getting ready in the same room - they wanted "
       "no first look.\nThe lighthouse at sunset, with or without "
       "us.",
    12: "The castle from the drive, before anyone arrives.\n"
        "Four generations of von Hoff women together.",
}

_FAMILY = {
    7: "Couple + both sets of parents\nCouple + Haddad grandparents\n"
       "Couple + all siblings\nBride + her four aunts\n"
       "Groom + university friends",
    8: "Couple + Whitfield parents\nCouple + Tessa's mother alone\n"
       "Couple + Biscuit\nCouple + all four grandparents",
    9: "Couple + Fontaine parents\nCouple + Camille's brothers\n"
       "Couple + both mothers",
    12: "Couple + von Hoff parents\nCouple + all grandparents\n"
        "Four generations of women\nCouple + the whole family",
}

_SENSITIVE = {
    7: "Bride's father died in March. Do not ask the family to "
       "assemble 'both parents' - the list above is exact.",
    8: "Tessa's father died two years ago. Her mother will cry at "
       "the watch shot; give her a minute first.",
    12: "Groom's parents are divorced and will not stand together. "
        "Two separate parent groups, never one.",
}

_RESTRICTIONS = {
    3: "No flash during the service. Photographer stays behind the "
       "third pew throughout.",
    7: "No photographs of women who have not consented during the "
       "bridal preparation. Ask the aunt in the green first.",
    12: "Castle permit forbids drone flight. Confirmed in writing.",
}

_TRADITIONS = {
    3: "Full Catholic mass, about 70 minutes.",
    7: "Nikah signing, then a separate baraat arrival with drums "
       "at about 15:40. Allow 25 minutes, not 10.",
    14: "Mexican lazo and arras during the ceremony.",
}

_BOOKING_NOTES = {
    4: "Gallery is three weeks past the promise. Say so plainly "
       "and give a date you can actually keep.",
    6: "Balance unpaid 29 days after the wedding. Gallery stays "
       "unlisted until it clears - that was in the contract.",
    8: "Questionnaire not back and the wedding is in six days. "
       "Phone, do not email.",
    10: "Eight guests on a clifftop. Priced like a portrait shoot, "
        "costed like a wedding. Read Profit Analysis first.",
    13: "Pencilled in on a date already promised to Anneliese & "
        "Caspar. Whoever signs first takes it.",
    15: "Pencilled. No contract sent yet - that is the only thing "
        "standing between this and a real booking.",
}


def bookings(cfg: Config) -> Rows:
    rows: Rows = []
    for index, (offset, _couple, venue, vtype, city, guests, package) \
            in enumerate(WEDDINGS, start=1):
        lead, second = _BOOKING_TEAM[index]
        addons, travel, discount = _BOOKING_MONEY[index]
        status = _booking_status(index)
        ceremony = 14 if vtype != "Church" else 13
        if index == 10:
            ceremony = 17            # clifftop elopement at sunset
        hours = {"Essential Wedding": 6, "Classic Wedding": 8,
                 "Premium Wedding": 10, "Luxury Wedding": 12}[package]
        start_hour = ceremony - 4 if hours >= 10 else ceremony - 3
        if index == 10:
            start_hour = 16
            hours = 2
        row: Dict[str, Any] = {
            "client": _clt(index),
            "wedding_date": _day(cfg, offset),
            "status": status,
            "confirmed_date": "" if status == "Pencilled in"
                              else _day(cfg, offset - 300),
            "ceremony_time": _time(ceremony),
            "reception_time": _time(ceremony + 3),
            "photo_start": _time(start_hour),
            "photo_end": _time(start_hour + hours),
            "venue": venue,
            "venue_type": vtype,
            "prep_location": f"{venue} - bridal suite"
                             if vtype in ("Hotel", "Country house",
                                          "Castle")
                             else f"Family home, {city}",
            "ceremony_location": f"{venue} - "
                                 + ("chapel" if vtype == "Church"
                                    else "ceremony lawn"),
            "reception_location": f"{venue} - main room",
            "guests": guests,
            "photographer": lead,
            "second": second,
            "package": package,
            "addons": addons,
            "travel_fee": travel,
            "discount": discount,
            "must_haves": _MUST_HAVES.get(index, ""),
            "family_list": _FAMILY.get(index, ""),
            "key_guests": "Wedding planner: Odette Hale (on the day "
                          "from 11:00)" if index in (7, 12) else "",
            "traditions": _TRADITIONS.get(index, ""),
            "restrictions": _RESTRICTIONS.get(index, ""),
            "sensitivities": _SENSITIVE.get(index, ""),
            "special_requests": "Sparkler exit at 22:30 - venue has "
                                "agreed in writing."
                                if index in (7, 9, 12) else "",
            "questionnaire_sent": "" if index in (13, 15)
                                  else _day(cfg, offset - 60),
            "questionnaire_back": "",
            "notes": _BOOKING_NOTES.get(index, ""),
        }
        if index not in (8, 13, 15) and index <= 14:
            row["questionnaire_back"] = _day(cfg, offset - 44)
        rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# 4. Packages & Pricing
# ---------------------------------------------------------------------------

def packages(cfg: Config) -> Rows:
    rows: Rows = [
        {"name": "Essential Wedding", "tier": "Essential",
         "active": "Yes",
         "summary": "Six hours, one photographer, online gallery. "
                    "The smallest package I will still put my name "
                    "on.",
         "hours": 6, "photographers": 1, "engagement": "No",
         "images": 350, "album": "No album", "print_credit": 0,
         "travel_allowance": 40,
         "inclusions": "Six hours of coverage\nOne photographer\n"
                       "350+ edited images\nOnline gallery for "
                       "12 months\nPrint release",
         "base_price": 1200, "standard_discount": 0,
         "album_cost": 0, "editing_hours": 12, "other_cost": 40},
        {"name": "Classic Wedding", "tier": "Classic",
         "active": "Yes",
         "summary": "Eight hours, engagement shoot, 20 page album. "
                    "Two thirds of my bookings land here.",
         "hours": 8, "photographers": 1, "engagement": "Yes",
         "images": 500, "album": "20 page album", "print_credit": 100,
         "travel_allowance": 60,
         "inclusions": "Eight hours of coverage\nOne photographer\n"
                       "Engagement shoot\n500+ edited images\n"
                       "20 page album\n$100 print credit\n"
                       "Online gallery for 12 months",
         "base_price": 1800, "standard_discount": 0,
         "album_cost": 185, "editing_hours": 18, "other_cost": 55},
        {"name": "Premium Wedding", "tier": "Premium",
         "active": "Yes",
         "summary": "Ten hours, two photographers, 30 page album. "
                    "The one to push when the guest list passes 100.",
         "hours": 10, "photographers": 2, "engagement": "Yes",
         "images": 700, "album": "30 page album", "print_credit": 250,
         "travel_allowance": 80,
         "inclusions": "Ten hours of coverage\nLead + second "
                       "shooter\nEngagement shoot\n700+ edited "
                       "images\n30 page album\n$250 print credit\n"
                       "Sneak peek within 48 hours",
         "base_price": 2500, "standard_discount": 0,
         "album_cost": 265, "editing_hours": 24, "other_cost": 70},
        {"name": "Luxury Wedding", "tier": "Luxury",
         "active": "Yes",
         "summary": "Twelve hours, two photographers, 40 page album "
                    "with parent copies. Sell the album, not the "
                    "hours.",
         "hours": 12, "photographers": 2, "engagement": "Yes",
         "images": 900, "album": "40 page album + parent copies",
         "print_credit": 500, "travel_allowance": 120,
         "inclusions": "Twelve hours of coverage\nLead + second "
                       "shooter\nEngagement shoot\n900+ edited "
                       "images\n40 page album + two parent copies\n"
                       "$500 print credit\nSneak peek within 24 "
                       "hours\nUSB in a presentation box",
         "base_price": 3500, "standard_discount": 0,
         "album_cost": 540, "editing_hours": 30, "other_cost": 95},
        {"name": "Elopement Hour", "tier": "Elopement",
         "active": "Yes",
         "summary": "Two hours, one photographer, 80 images. Priced "
                    "on instinct in 2024 and never revisited - the "
                    "hourly rate on this is the worst on the sheet.",
         "hours": 2, "photographers": 1, "engagement": "No",
         "images": 80, "album": "No album", "print_credit": 0,
         "travel_allowance": 30,
         "inclusions": "Two hours of coverage\nOne photographer\n"
                       "80+ edited images\nOnline gallery for "
                       "12 months",
         "base_price": 650, "standard_discount": 0,
         "album_cost": 0, "editing_hours": 5, "other_cost": 25},
        {"name": "Classic 2024", "tier": "Custom", "active": "No",
         "summary": "Last year's Classic, kept only so old bookings "
                    "and old numbers still make sense. Do not quote "
                    "it.",
         "hours": 8, "photographers": 1, "engagement": "Yes",
         "images": 450, "album": "20 page album", "print_credit": 75,
         "travel_allowance": 50,
         "inclusions": "Eight hours of coverage\nOne photographer\n"
                       "Engagement shoot\n450+ edited images\n"
                       "20 page album",
         "base_price": 1550, "standard_discount": 0,
         "album_cost": 185, "editing_hours": 18, "other_cost": 55},
    ]
    return rows


# ---------------------------------------------------------------------------
# 5. Photographer & Staff
# ---------------------------------------------------------------------------

def staff(cfg: Config) -> Rows:
    return [
        {"name": "Marisa Calloway", "role": "Lead photographer",
         "status": "Active", "type": "Owner",
         "phone": "+1 503 555 0101",
         "email": "marisa@rowanandfern.example.com",
         "started": _day(cfg, -2920), "day_rate": 0, "hour_rate": 0,
         "edit_rate": 0, "is_owner": "Yes", "can_lead": "Yes",
         "can_second": "No", "can_edit": "Yes",
         "specialism": "Documentary coverage, family diplomacy",
         "kit": "R6 II + R6, 35mm, 50mm, 85mm, 70-200mm, two "
                "speedlights",
         "insured": "Yes", "insurance_expiry": _day(cfg, 97),
         "notes": "The whole business. Every hour she spends "
                  "editing is an hour she is not selling - that is "
                  "what Owner hourly value on Settings is for."},
        {"name": "Dev Ashworth", "role": "Second shooter",
         "status": "Active", "type": "Self-employed contractor",
         "phone": "+1 503 555 0188",
         "email": "dev.ashworth@example.com",
         "started": _day(cfg, -1460), "day_rate": 350,
         "hour_rate": 45, "edit_rate": 0, "is_owner": "No",
         "can_lead": "No", "can_second": "Yes", "can_edit": "No",
         "specialism": "Reportage, dance floor, never misses a "
                       "speech reaction",
         "kit": "Own R6 + 24-70mm, 135mm", "insured": "Yes",
         "insurance_expiry": _day(cfg, -12),
         "notes": "Insurance expired twelve days ago. He cannot "
                  "shoot the next wedding until the new "
                  "certificate is on file."},
        {"name": "Joelle Nakamura", "role": "Second shooter",
         "status": "Active", "type": "Self-employed contractor",
         "phone": "+1 503 555 0231",
         "email": "joelle.n@example.com",
         "started": _day(cfg, -980), "day_rate": 375,
         "hour_rate": 48, "edit_rate": 32, "is_owner": "No",
         "can_lead": "Yes", "can_second": "Yes", "can_edit": "Yes",
         "specialism": "Fine art, can lead if Marisa is double "
                       "booked",
         "kit": "Own Z6 III + 35mm, 85mm", "insured": "Yes",
         "insurance_expiry": _day(cfg, 211),
         "notes": "The answer to a date clash. Costs more than Dev "
                  "and is worth it."},
        {"name": "Tomas Bright", "role": "Editor",
         "status": "Active", "type": "Self-employed contractor",
         "phone": "+1 503 555 0344",
         "email": "tomas@brightedits.example.com",
         "started": _day(cfg, -610), "day_rate": 0, "hour_rate": 0,
         "edit_rate": 28, "is_owner": "No", "can_lead": "No",
         "can_second": "No", "can_edit": "Yes",
         "specialism": "Culling and colour, 48 hour turnaround",
         "kit": "Calibrated studio, remote", "insured": "No",
         "insurance_expiry": "",
         "notes": "At $28 an hour he is cheaper than Marisa's own "
                  "time by a wide margin. The backlog on "
                  "WED-0004 exists because she did not send it."},
        {"name": "Priya Raman", "role": "Assistant",
         "status": "Occasional", "type": "Unpaid helper",
         "phone": "+1 503 555 0415",
         "email": "priya.raman@example.com",
         "started": _day(cfg, -240), "day_rate": 80, "hour_rate": 12,
         "edit_rate": 0, "is_owner": "No", "can_lead": "No",
         "can_second": "Yes", "can_edit": "No",
         "specialism": "Lighting, bags, keeping Aunt Margaret busy",
         "kit": "None - uses studio spares", "insured": "No",
         "insurance_expiry": "",
         "notes": "Studying photography. Not insured, so she "
                  "assists and does not shoot."},
    ]


# ---------------------------------------------------------------------------
# 6. Contracts & Documents
# ---------------------------------------------------------------------------

def contracts(cfg: Config) -> Rows:
    rows: Rows = []
    for index, (offset, _couple, venue, vtype, _city, _g, _p) in \
            enumerate(WEDDINGS, start=1):
        sent = offset - 305
        store = f"Dropbox / Weddings / {_wed(index)}"
        if index == 15:
            # the pencilled booking with nothing sent
            rows.append({
                "booking": _wed(index), "doc_type": "Wedding contract",
                "title": "Wedding photography contract",
                "required": "Yes", "status": "Outstanding",
                "sent_date": "", "signed_date": "", "expires": "",
                "stored_at": store, "signers": "Both parties",
                "notes": "Not sent. Until this goes out the date is "
                         "not held, whatever the calendar says."})
            continue
        if index == 13:
            rows.append({
                "booking": _wed(index), "doc_type": "Wedding contract",
                "title": "Wedding photography contract",
                "required": "Yes", "status": "Sent",
                "sent_date": _day(cfg, -9), "signed_date": "",
                "expires": _day(cfg, 5), "stored_at": store,
                "signers": "Both parties",
                "notes": "Sent nine days ago. The same date is "
                         "pencilled for another couple."})
            continue
        rows.append({
            "booking": _wed(index), "doc_type": "Wedding contract",
            "title": "Wedding photography contract", "required": "Yes",
            "status": "Counter-signed", "sent_date": _day(cfg, sent),
            "signed_date": _day(cfg, sent + 4), "expires": "",
            "stored_at": store, "signers": "Both parties",
            "notes": ""})
        rows.append({
            "booking": _wed(index), "doc_type": "Booking form",
            "title": "Booking form and deposit confirmation",
            "required": "Yes", "status": "Filed",
            "sent_date": _day(cfg, sent), "signed_date":
            _day(cfg, sent + 4), "expires": "", "stored_at": store,
            "signers": "Couple", "notes": ""})
        if vtype in ("Church", "Castle", "Vineyard", "Country house"):
            late = index in (9, 12)
            rows.append({
                "booking": _wed(index), "doc_type": "Venue permit",
                "title": f"{venue} photography permit",
                "required": "Yes",
                "status": "Outstanding" if late else "Filed",
                "sent_date": "" if late else _day(cfg, offset - 70),
                "signed_date": "" if late else _day(cfg, offset - 62),
                "expires": _day(cfg, offset + 1),
                "stored_at": store, "signers": "Venue",
                "notes": "Venue will not let me past the gate "
                         "without this." if late else ""})
        if index in (7, 12, 14):
            rows.append({
                "booking": _wed(index),
                "doc_type": "Venue public liability",
                "title": f"{venue} - proof of my cover",
                "required": "Yes", "status": "Filed",
                "sent_date": _day(cfg, offset - 50),
                "signed_date": _day(cfg, offset - 50),
                "expires": _day(cfg, 97), "stored_at": store,
                "signers": "Insurer", "notes": ""})
        if index in _PAST:
            rows.append({
                "booking": _wed(index), "doc_type": "Model release",
                "title": "Permission to use images in marketing",
                "required": "No",
                "status": "Signed" if index not in (3, 6)
                          else "Not needed",
                "sent_date": _day(cfg, offset + 2)
                if index not in (3, 6) else "",
                "signed_date": _day(cfg, offset + 5)
                if index not in (3, 6) else "", "expires": "",
                "stored_at": store, "signers": "Couple",
                "notes": "Declined - nothing from this wedding goes "
                         "online." if index in (3, 6) else ""})
        if index in (8, 9, 12):
            rows.append({
                "booking": _wed(index), "doc_type": "Timeline sign-off",
                "title": "Wedding day timeline, agreed version",
                "required": "Yes",
                "status": "Sent" if index == 8 else "Signed",
                "sent_date": _day(cfg, offset - 30),
                "signed_date": "" if index == 8
                               else _day(cfg, offset - 26),
                "expires": _day(cfg, offset - 7), "stored_at": store,
                "signers": "Couple",
                "notes": "Unsigned with six days to go." if index == 8
                         else ""})
        if _BOOKING_TEAM[index][1]:
            rows.append({
                "booking": _wed(index),
                "doc_type": "Second shooter agreement",
                "title": "Second shooter terms and copyright "
                         "assignment",
                "required": "Yes",
                "status": "Signed" if index != 9 else "Outstanding",
                "sent_date": _day(cfg, offset - 40) if index != 9
                             else "",
                "signed_date": _day(cfg, offset - 38) if index != 9
                               else "", "expires": "",
                "stored_at": store, "signers": "Second shooter",
                "notes": "Dev's insurance certificate has expired "
                         "too - fix both together." if index == 9
                         else ""})
    return rows


# ---------------------------------------------------------------------------
# 7. Payments & Invoices
# ---------------------------------------------------------------------------

_PACKAGE_PRICE = {"Essential Wedding": 1200, "Classic Wedding": 1800,
                  "Premium Wedding": 2500, "Luxury Wedding": 3500}


def payments(cfg: Config) -> Rows:
    rows: Rows = []
    counter = 0
    for index, (offset, _couple, _venue, _vt, _city, _g, package) in \
            enumerate(WEDDINGS, start=1):
        if index == 15:
            continue                 # pencilled, nothing invoiced yet
        addons, travel, discount = _BOOKING_MONEY[index]
        total = _PACKAGE_PRICE[package] + addons + travel - discount
        deposit = round(total * 0.30, 2)
        second = round(total * 0.35, 2)
        final = round(total - deposit - second, 2)
        booked = offset - 300
        paid_all = index not in (6, 13)

        counter += 1
        rows.append({
            "booking": _wed(index), "stage": "Booking deposit",
            "reference": f"INV-{2000 + counter}",
            "description": "Deposit to hold the date",
            "amount": deposit, "due_date": _day(cfg, booked + 7),
            "received": deposit if index != 13 else 0,
            "received_date": _day(cfg, booked + 3) if index != 13
                             else "",
            "method": "Bank transfer" if index != 13 else "",
            "fee": 0, "reminders": 0, "last_reminder": "",
            "notes": "Date is not held until this clears."
                     if index == 13 else ""})

        counter += 1
        due_second = offset - 90
        got_second = due_second < 0 and index != 13
        rows.append({
            "booking": _wed(index), "stage": "Second instalment",
            "reference": f"INV-{2000 + counter}",
            "description": "Second instalment, 90 days before",
            "amount": second, "due_date": _day(cfg, due_second),
            "received": second if got_second else 0,
            "received_date": _day(cfg, due_second + 2)
                             if got_second else "",
            "method": "Bank transfer" if got_second else "",
            "fee": 0, "reminders": 0, "last_reminder": "",
            "notes": ""})

        counter += 1
        due_final = offset - 14
        got_final = due_final < 0 and paid_all
        rows.append({
            "booking": _wed(index), "stage": "Final payment",
            "reference": f"INV-{2000 + counter}",
            "description": "Balance, due two weeks before the day",
            "amount": final, "due_date": _day(cfg, due_final),
            "received": final if got_final else 0,
            "received_date": _day(cfg, due_final + 1)
                             if got_final else "",
            "method": "Bank transfer" if got_final else "",
            "fee": 0,
            "reminders": 3 if index == 6 else 0,
            "last_reminder": _day(cfg, -6) if index == 6 else "",
            "notes": "Forty-three days overdue. Three emails, no "
                     "phone call yet. Make the phone call."
                     if index == 6 else ""})

        if index in (3, 11):
            counter += 1
            rows.append({
                "booking": _wed(index), "stage": "Album",
                "reference": f"INV-{2000 + counter}",
                "description": "Album upgrade to 40 pages",
                "amount": 320, "due_date": _day(cfg, offset + 40),
                "received": 320 if index == 3 else 0,
                "received_date": _day(cfg, offset + 44)
                                 if index == 3 else "",
                "method": "Credit card" if index == 3 else "",
                "fee": 9.60 if index == 3 else 0,
                "reminders": 0, "last_reminder": "", "notes": ""})
        if index == 2:
            counter += 1
            rows.append({
                "booking": _wed(index), "stage": "Prints",
                "reference": f"INV-{2000 + counter}",
                "description": "Parent print sets, two of",
                "amount": 180, "due_date": _day(cfg, offset + 60),
                "received": 180, "received_date":
                _day(cfg, offset + 58), "method": "Stripe",
                "fee": 5.52, "reminders": 0, "last_reminder": "",
                "notes": "Print sales are the quietest money in "
                         "this business."})
        if index == 7:
            counter += 1
            rows.append({
                "booking": _wed(index), "stage": "Overtime",
                "reference": f"INV-{2000 + counter}",
                "description": "One extra hour on the night, agreed "
                               "at 22:10 by text",
                "amount": 250, "due_date": _day(cfg, 13),
                "received": 0, "received_date": "", "method": "",
                "fee": 0, "reminders": 0, "last_reminder": "",
                "notes": "Agreed in writing on the night. Invoice "
                         "it while they still remember saying yes."})
        if index == 4:
            counter += 1
            rows.append({
                "booking": _wed(index), "stage": "Refund",
                "reference": f"CRN-{2000 + counter}",
                "description": "Goodwill credit for the late gallery",
                "amount": -150, "due_date": _day(cfg, -5),
                "received": -150, "received_date": _day(cfg, -5),
                "method": "Bank transfer", "fee": 0, "reminders": 0,
                "last_reminder": "",
                "notes": "Cheaper than the review they would have "
                         "left. It still comes out of the margin."})
    return rows


# ---------------------------------------------------------------------------
# 8. Wedding Timeline
# ---------------------------------------------------------------------------

_TIMELINE_LOCATIONS = {
    "venue_detail": "Ceremony room", "bride_prep": "Bridal suite",
    "groom_prep": "Groom's room", "dress_details": "Bridal suite",
    "first_look": "Walled garden", "couple_portraits_1":
    "Walled garden", "guests_arrive": "Entrance",
    "ceremony": "Ceremony room", "confetti": "Front steps",
    "family_groups": "South lawn", "wedding_party": "South lawn",
    "couple_portraits_2": "Orchard", "cocktail_hour": "Terrace",
    "reception_detail": "Dining room", "entrance": "Dining room",
    "speeches": "Dining room", "cake_cutting": "Dining room",
    "first_dance": "Dance floor", "dancing": "Dance floor",
    "sparkler_exit": "Front drive",
}

_CRITICAL = {"ceremony", "family_groups", "first_dance", "speeches",
             "cake_cutting", "first_look"}


def _timeline_for(cfg: Config, index: int, keys, *, done: bool,
                  drift: Dict[str, int] = None) -> Rows:
    """One wedding's running order, drawn from the standard plan."""
    drift = drift or {}
    offset, _couple, _venue, vtype, _city, _g, _pkg = \
        WEDDINGS[index - 1]
    ceremony_hour = 13 if vtype == "Church" else 14
    lead, second = _BOOKING_TEAM[index]
    rows: Rows = []
    for key, label, minutes, duration, role in TIMELINE_EVENTS:
        if key not in keys:
            continue
        staff_id = second if role == "Second shooter" and second \
            else lead
        row: Dict[str, Any] = {
            "booking": _wed(index), "event": label, "label": "",
            "offset": minutes, "time_override": "",
            "duration": duration,
            "location": _TIMELINE_LOCATIONS.get(key, ""),
            "staff": staff_id,
            "critical": "Yes" if key in _CRITICAL else "No",
            "notes": "", "done": "Yes" if done else "No",
            "actual": "",
        }
        if done:
            total = ceremony_hour * 60 + minutes + drift.get(key, 0)
            row["actual"] = _time(total // 60, total % 60)
        rows.append(row)
    return rows


def timeline(cfg: Config) -> Rows:
    all_keys = [k for k, _l, _m, _d, _r in TIMELINE_EVENTS]
    rows: Rows = []

    # the wedding that happened yesterday, with what actually
    # happened written next to what was meant to happen
    rows += _timeline_for(
        cfg, 7, all_keys, done=True,
        drift={"bride_prep": 20, "groom_prep": 15, "first_look": 25,
               "couple_portraits_1": 30, "ceremony": 18,
               "confetti": 18, "family_groups": 22,
               "wedding_party": 35, "couple_portraits_2": 40,
               "speeches": 12, "first_dance": 5, "dancing": 5,
               "sparkler_exit": 0})
    rows[0]["notes"] = ("Baraat arrived at 15:40 with drums. Twenty "
                        "five minutes, not ten - put that in the "
                        "template for next time.")
    for row in rows:
        if row["event"] == "Family group photographs":
            row["notes"] = ("Ran 22 minutes late because nobody had "
                            "told the uncles to stay. Print the "
                            "family list next time.")

    # next Saturday: planned, nothing done
    rows += _timeline_for(cfg, 8, all_keys[:18], done=False)

    # three weeks out: two brides, no first look
    keys_9 = [k for k in all_keys
              if k not in ("first_look", "groom_prep")]
    rows += _timeline_for(cfg, 9, keys_9, done=False)

    # the elopement: four lines is the whole day
    rows += [
        {"booking": _wed(10), "event": "Venue and detail shots",
         "label": "Trailhead and rim", "offset": -45, "duration": 20,
         "time_override": "", "location": "Crater Rim Overlook",
         "staff": _phg(1), "critical": "No",
         "notes": "Forty minute walk in from the car park. That walk "
                  "is why this package loses money.",
         "done": "No", "actual": ""},
        {"booking": _wed(10), "event": "Wedding ceremony",
         "label": "Vows at the overlook", "offset": 0, "duration": 20,
         "time_override": _time(17), "location": "Crater Rim Overlook",
         "staff": _phg(1), "critical": "Yes", "notes": "",
         "done": "No", "actual": ""},
        {"booking": _wed(10), "event": "Couple portraits",
         "label": "Golden hour portraits", "offset": 25,
         "duration": 45, "time_override": "",
         "location": "Rim trail", "staff": _phg(1), "critical": "Yes",
         "notes": "Sunset is at 18:12. There is no second chance.",
         "done": "No", "actual": ""},
        {"booking": _wed(10), "event": "Confetti and congratulations",
         "label": "Champagne on the rock", "offset": 75,
         "duration": 15, "time_override": "",
         "location": "Rim trail", "staff": _phg(1), "critical": "No",
         "notes": "", "done": "No", "actual": ""},
    ]

    # the castle, six months out, skeleton only
    rows += _timeline_for(
        cfg, 12,
        ["venue_detail", "bride_prep", "ceremony", "family_groups",
         "speeches", "first_dance"], done=False)
    for row in rows[-6:]:
        row["notes"] = ("Skeleton only. Firm this up at the venue "
                        "visit in month four.")
    return rows


# ---------------------------------------------------------------------------
# 9. Shot List Manager
# ---------------------------------------------------------------------------

_SHOT_WHEN = {
    "Details": "Before the ceremony",
    "Preparation": "Getting ready",
    "Couple": "Couple portraits",
    "Family": "Family groups",
    "Ceremony": "Ceremony",
    "Reception": "Reception",
}

_GROUP_NAME = {
    "Details and venue": "Details", "Getting ready": "Preparation",
    "The couple": "Couple", "Family and groups": "Family",
    "The ceremony": "Ceremony", "The reception": "Reception",
}


def _shot_group(label: str) -> str:
    """Map a config shot group onto a Shot group drop-down value."""
    if label in _GROUP_NAME:
        return _GROUP_NAME[label]
    lowered = label.lower()
    for key, value in (("detail", "Details"), ("ready", "Preparation"),
                       ("prep", "Preparation"), ("couple", "Couple"),
                       ("portrait", "Couple"), ("family", "Family"),
                       ("group", "Family"), ("ceremony", "Ceremony"),
                       ("reception", "Reception"),
                       ("party", "Reception"),
                       ("evening", "Reception")):
        if key in lowered:
            return value
    return "Details"


def shots(cfg: Config) -> Rows:
    rows: Rows = []

    # WED-0009, three weeks out: the full standard list, nothing shot
    for _key, label, shot_names in SHOT_LIST:
        group = _shot_group(label)
        for shot in shot_names:
            rows.append({
                "booking": _wed(9), "group_name": group, "shot": shot,
                "people": "", "priority": "Normal", "must_have": "No",
                "requested_by": "Your standard list",
                "when": _SHOT_WHEN[group], "location": "",
                "notes": "", "done": "No", "taken_at": "",
                "skipped": "No", "skip_reason": "", "delivered": "No"})

    # the two things that wedding is actually about
    rows += [
        {"booking": _wed(9), "group_name": "Couple",
         "shot": "Both brides getting ready in the same room",
         "people": "Camille, Esme", "priority": "High",
         "must_have": "Yes", "requested_by": "Couple asked for it",
         "when": "Getting ready", "location": "Harbour suite",
         "notes": "They did not want a first look. This replaces it.",
         "done": "No", "taken_at": "", "skipped": "No",
         "skip_reason": "", "delivered": "No"},
        {"booking": _wed(9), "group_name": "Couple",
         "shot": "The lighthouse at sunset",
         "people": "Camille, Esme", "priority": "High",
         "must_have": "Yes", "requested_by": "Couple asked for it",
         "when": "Golden hour", "location": "Point headland",
         "notes": "Eleven minute walk. Leave the reception at 19:40.",
         "done": "No", "taken_at": "", "skipped": "No",
         "skip_reason": "", "delivered": "No"},
        {"booking": _wed(9), "group_name": "Family",
         "shot": "Couple + Fontaine parents", "people":
         "Camille, Esme, Ariane, Luc", "priority": "High",
         "must_have": "Yes", "requested_by": "Couple asked for it",
         "when": "Family groups", "location": "Terrace",
         "notes": "", "done": "No", "taken_at": "", "skipped": "No",
         "skip_reason": "", "delivered": "No"},
        {"booking": _wed(9), "group_name": "Family",
         "shot": "Couple + Camille's brothers",
         "people": "Camille, Esme, Luca, Theo", "priority": "Normal",
         "must_have": "Yes", "requested_by": "Couple asked for it",
         "when": "Family groups", "location": "Terrace", "notes": "",
         "done": "No", "taken_at": "", "skipped": "No",
         "skip_reason": "", "delivered": "No"},
        {"booking": _wed(9), "group_name": "Family",
         "shot": "Couple + both mothers",
         "people": "Camille, Esme, Ariane, Sylvie",
         "priority": "Normal", "must_have": "Yes",
         "requested_by": "Parent asked for it", "when":
         "Family groups", "location": "Terrace", "notes": "",
         "done": "No", "taken_at": "", "skipped": "No",
         "skip_reason": "", "delivered": "No"},
    ]

    # WED-0008, next Saturday: the short list that matters
    eight = [
        ("Details", "The dress hanging", "", "Normal", "No",
         "Your standard list"),
        ("Details", "Rings on the invitation suite", "", "Normal",
         "No", "Your standard list"),
        ("Preparation", "Tessa's father's watch going on Rowan's "
                        "wrist", "Rowan", "High", "Yes",
         "Couple asked for it"),
        ("Preparation", "Mother fastening the dress", "Tessa, Hilary",
         "High", "Yes", "Parent asked for it"),
        ("Ceremony", "The walk down the aisle", "", "High", "No",
         "Your standard list"),
        ("Ceremony", "The vows, both faces", "", "High", "No",
         "Your standard list"),
        ("Ceremony", "Ring exchange, close", "", "Normal", "No",
         "Your standard list"),
        ("Family", "Couple + Whitfield parents",
         "Tessa, Rowan, Gavin, Lorna", "High", "Yes",
         "Couple asked for it"),
        ("Family", "Couple + Tessa's mother alone", "Tessa, Hilary",
         "High", "Yes", "Couple asked for it"),
        ("Family", "Couple + Biscuit the dog", "Tessa, Rowan",
         "Normal", "Yes", "Couple asked for it"),
        ("Family", "Couple + all four grandparents", "", "Normal",
         "Yes", "Parent asked for it"),
        ("Couple", "Golden hour in the long meadow", "", "High", "No",
         "Your own idea"),
        ("Reception", "Room before the guests come in", "", "Normal",
         "No", "Your standard list"),
        ("Reception", "Speeches - reaction shots, not the speaker",
         "", "High", "No", "Your own idea"),
        ("Reception", "First dance, wide and close", "", "High", "No",
         "Your standard list"),
    ]
    for group, shot, people, priority, must, source in eight:
        rows.append({
            "booking": _wed(8), "group_name": group, "shot": shot,
            "people": people, "priority": priority, "must_have": must,
            "requested_by": source, "when": _SHOT_WHEN[group],
            "location": "", "notes": "", "done": "No", "taken_at": "",
            "skipped": "No", "skip_reason": "", "delivered": "No"})

    # WED-0007, shot yesterday: what got taken and what did not
    seven = [
        ("Family", "Grandmother Haddad with both grandsons",
         "Yusra, Kaleem, Nadra", "High", "Yes",
         "Couple asked for it", "Yes", _time(16, 42), "No", ""),
        ("Ceremony", "The nikah signing, wide", "", "High", "Yes",
         "Couple asked for it", "Yes", _time(14, 18), "No", ""),
        ("Ceremony", "The nikah signing, close on the pen", "",
         "High", "Yes", "Couple asked for it", "Yes",
         _time(14, 21), "No", ""),
        ("Details", "Kaleem's mother's gold bangles on the dress",
         "", "High", "Yes", "Parent asked for it", "Yes",
         _time(11, 30), "No", ""),
        ("Details", "Baraat drums arriving", "", "Normal", "No",
         "Your own idea", "Yes", _time(15, 44), "No", ""),
        ("Family", "Couple + both sets of parents", "", "High", "Yes",
         "Your standard list", "Yes", _time(16, 30), "No", ""),
        ("Family", "Couple + Haddad grandparents", "", "High", "Yes",
         "Couple asked for it", "Yes", _time(16, 36), "No", ""),
        ("Family", "Couple + all siblings", "", "Normal", "Yes",
         "Couple asked for it", "Yes", _time(16, 48), "No", ""),
        ("Family", "Bride + her four aunts", "", "Normal", "Yes",
         "Parent asked for it", "Yes", _time(16, 55), "No", ""),
        ("Family", "Groom + university friends", "", "Low", "No",
         "Couple asked for it", "No", "", "Yes",
         "Half of them were already on the dance floor and the "
         "light had gone. Offered a reshoot at the reception, "
         "they declined."),
        ("Couple", "Sparkler exit", "", "High", "No",
         "Couple asked for it", "Yes", _time(22, 31), "No", ""),
        ("Reception", "Speeches - the father of the bride's chair, "
                      "empty, with his photograph on it", "",
         "High", "Yes", "Your own idea", "Yes", _time(19, 52), "No",
         ""),
    ]
    for (group, shot, people, priority, must, source, done, taken,
         skipped, reason) in seven:
        rows.append({
            "booking": _wed(7), "group_name": group, "shot": shot,
            "people": people, "priority": priority, "must_have": must,
            "requested_by": source, "when": _SHOT_WHEN[group],
            "location": "", "notes": "", "done": done,
            "taken_at": taken, "skipped": skipped,
            "skip_reason": reason, "delivered": "No"})

    # WED-0012, the castle: the four shots booked six months early
    twelve = [
        ("Details", "The castle from the drive, before anyone "
                    "arrives", "High", "Yes", "Couple asked for it"),
        ("Family", "Four generations of von Hoff women",
         "High", "Yes", "Couple asked for it"),
        ("Family", "Couple + von Hoff parents - father's side only",
         "High", "Yes", "Couple asked for it"),
        ("Family", "Couple + von Hoff parents - mother's side only",
         "High", "Yes", "Couple asked for it"),
    ]
    for group, shot, priority, must, source in twelve:
        rows.append({
            "booking": _wed(12), "group_name": group, "shot": shot,
            "people": "", "priority": priority, "must_have": must,
            "requested_by": source, "when": _SHOT_WHEN[group],
            "location": "", "notes": "Parents are divorced and will "
                                     "not stand together."
            if "side only" in shot else "",
            "done": "No", "taken_at": "", "skipped": "No",
            "skip_reason": "", "delivered": "No"})
    return rows


# ---------------------------------------------------------------------------
# 10. Editing Workflow
# ---------------------------------------------------------------------------

def editing(cfg: Config) -> Rows:
    rows: Rows = []
    recipe = {
        # index: (editor, outsourced, steps done, shot, kept,
        #         delivered, hours, rush)
        1: (_phg(1), "No", 10, 3800, 620, 540, 19.5, "No"),
        2: (_phg(4), "Yes", 10, 5200, 880, 760, 24.0, "No"),
        3: (_phg(1), "No", 10, 4900, 820, 710, 27.5, "No"),
        4: (_phg(1), "No", 5, 4100, 690, 0, 9.0, "No"),
        5: (_phg(4), "Yes", 10, 3600, 610, 520, 16.5, "No"),
        6: (_phg(1), "No", 7, 2900, 470, 0, 8.0, "No"),
        7: (_phg(1), "No", 2, 6400, 0, 0, 1.5, "Yes"),
    }
    steps = ["step_backed_up", "step_imported", "step_culled",
             "step_selected", "step_edited", "step_colour",
             "step_retouched", "step_album_designed",
             "step_preview_sent", "step_final_uploaded"]
    notes = {
        4: "Ninety-two days since the wedding against a forty-two "
           "day promise. It stalled at 'selection complete' because "
           "I kept meaning to do it myself. Tomas charges $28 an "
           "hour. Send it.",
        6: "Deliberately parked at retouching until the balance "
           "clears. That is in the contract and the couple know.",
        7: "Cards backed up to two drives before I went to bed. "
           "That is the only step that cannot wait.",
    }
    for index, (editor, outsourced, done_steps, shot, kept,
                delivered, hours, rush) in recipe.items():
        offset = WEDDINGS[index - 1][0]
        row: Dict[str, Any] = {
            "booking": _wed(index), "editor": editor,
            "outsourced": outsourced,
            "priority": "High" if index in (4, 7) else "Normal",
            "rush": rush,
            "started": _day(cfg, offset + 2),
            "preview_sent": _day(cfg, offset + 7)
                            if done_steps >= 9 else "",
            "finished": _day(cfg, offset + 34)
                        if done_steps == 10 else "",
            "shot_count": shot, "kept": kept,
            "delivered_count": delivered, "hours": hours,
            "backup": "Yes",
            "notes": notes.get(index, ""),
        }
        for position, key in enumerate(steps, start=1):
            row[key] = "Yes" if position <= done_steps else "No"
        if index == 7:
            row["started"] = _day(cfg, 0)
        rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# 11. Gallery & Delivery
# ---------------------------------------------------------------------------

def gallery(cfg: Config) -> Rows:
    recipe = {
        1: ("Pic-Time", "Delivered", 540, 1240, 310, 96, "Yes",
            "Yes", 310, 0),
        2: ("Pic-Time", "Delivered", 760, 2180, 540, 171, "Yes",
            "Yes", 180, 140),
        3: ("Pic-Time", "Delivered", 710, 1890, 480, 142, "No",
            "Yes", 0, 320),
        4: ("Pic-Time", "Created", 0, 0, 0, 0, "No", "No", 0, 0),
        5: ("Pic-Time", "Delivered", 520, 640, 210, 58, "Yes", "No",
            0, 0),
        6: ("Pic-Time", "Preview sent", 40, 96, 0, 11, "No", "No",
            0, 0),
        7: ("Pic-Time", "Not created", 0, 0, 0, 0, "No", "No", 0, 0),
    }
    notes = {
        3: "Expires in a few weeks. They have not downloaded the "
           "full set - warn them before it closes, then sell the "
           "extension.",
        4: "Gallery exists and is empty. The couple have the link.",
        5: "Delivered and nobody has asked for a review yet. That "
           "is money sitting on the table.",
        6: "Preview only. The full gallery goes live when the "
           "balance clears.",
        7: "Shot yesterday. Sneak peek promised within 24 hours for "
           "Luxury - that clock is running.",
    }
    rows: Rows = []
    for index, (platform, status, images, views, downloads, favs,
                usb, album, print_sales, extra) in recipe.items():
        offset = WEDDINGS[index - 1][0]
        delivered_on = _day(cfg, offset + 34) \
            if status in ("Delivered", "Archived") else ""
        rows.append({
            "booking": _wed(index), "platform": platform,
            "gallery_name": f"{WEDDINGS[index - 1][1]} - "
                            f"{WEDDINGS[index - 1][2]}",
            "link": f"https://gallery.example.com/{_wed(index).lower()}",
            "has_password": "Yes", "status": status,
            "delivered_on": delivered_on, "images": images,
            "expiry_warned": "No", "archived": "No",
            "first_viewed": _day(cfg, offset + 35)
                            if delivered_on else "",
            "views": views, "downloads": downloads,
            "favourites": favs, "usb_sent": usb,
            "album_ordered": album,
            "album_sent": _day(cfg, offset + 95)
                          if album == "Yes" and index in (1, 2)
                          else "",
            "print_sales": print_sales, "extra_sales": extra,
            "review_asked": _day(cfg, offset + 40)
                            if index in (1, 2, 3) else "",
            "notes": notes.get(index, ""),
        })
    return rows


# ---------------------------------------------------------------------------
# 12. Equipment Inventory
# ---------------------------------------------------------------------------

def equipment(cfg: Config) -> Rows:
    kit = [
        ("Canon R6 Mark II (body 1)", "Camera body", "Canon",
         "R6M2-884217", "In use", "Excellent", _phg(1), -540, 2499,
         5, "Yes", 2100, 185000, 400000, "", 2499,
         "The one that goes to every wedding."),
        ("Canon R6 (body 2)", "Camera body", "Canon", "R6-771940",
         "In use", "Good", _phg(1), -1180, 2099, 5, "Yes", 1200,
         372000, 400000, "Canon R6 Mark II (body 1)", 2499,
         "Shutter count is at 93% of rated life. Budget for the "
         "replacement now, not in the week it dies."),
        ("Canon R8 (backup body)", "Camera body", "Canon",
         "R8-110562", "Spare", "Excellent", "", -300, 1499, 5,
         "Yes", 1300, 41000, 400000, "Canon R6 (body 2)", 1499,
         "Lives in the car. Has never been needed and that is the "
         "entire point."),
        ("Sigma 35mm f1.4 Art", "Lens", "Sigma", "SG35-22914",
         "In use", "Good", _phg(1), -1460, 849, 8, "Yes", 600, 0, 0,
         "", 899, ""),
        ("Canon RF 50mm f1.2L", "Lens", "Canon", "RF50-449021",
         "In use", "Excellent", _phg(1), -720, 2299, 8, "Yes", 2000,
         0, 0, "", 2299, "The look the whole brand is built on."),
        ("Canon RF 85mm f1.2L", "Lens", "Canon", "RF85-558130",
         "In use", "Excellent", _phg(1), -400, 2699, 8, "Yes", 2400,
         0, 0, "", 2699, ""),
        ("Canon RF 70-200mm f2.8L", "Lens", "Canon",
         "RF70200-663094", "Being repaired", "Needs repair", "",
         -980, 2599, 8, "Yes", 2200, 0, 0, "", 2699,
         "Focus motor is grinding. Away for three weeks and the "
         "next wedding is in six days - hire one if it is not back "
         "by Thursday."),
        ("Canon RF 24-70mm f2.8L", "Lens", "Canon", "RF2470-229417",
         "In use", "Good", _phg(2), -1300, 2299, 8, "Yes", 1700, 0,
         0, "", 2399, "On permanent loan to Dev."),
        ("Godox V1 speedlight (1)", "Flash / speedlight", "Godox",
         "V1-330881", "In use", "Good", _phg(1), -900, 259, 4, "No",
         0, 0, 0, "", 269, ""),
        ("Godox V1 speedlight (2)", "Flash / speedlight", "Godox",
         "V1-330882", "In use", "Fair", _phg(1), -900, 259, 4, "No",
         0, 0, 0, "Godox V1 speedlight (1)", 269,
         "Hot shoe is loose. Works, but I would not trust it for a "
         "first dance."),
        ("Godox AD200 Pro", "Studio lighting", "Godox",
         "AD200-771203", "In use", "Excellent", _phg(1), -610, 349,
         5, "No", 0, 0, 0, "", 369, ""),
        ("MagMod modifier kit", "Light modifier", "MagMod",
         "MM-KIT-9921", "In use", "Good", _phg(1), -1100, 229, 6,
         "No", 0, 0, 0, "", 249, ""),
        ("Manfrotto 055 tripod", "Tripod / monopod", "Manfrotto",
         "MN055-412098", "In use", "Good", _phg(1), -1800, 219, 10,
         "No", 0, 0, 0, "", 239, ""),
        ("SanDisk 128GB CFexpress x6", "Memory card", "SanDisk",
         "SD-CFE-BATCH-7", "In use", "Good", _phg(1), -420, 1080, 3,
         "No", 0, 0, 0, "", 1140,
         "Six cards. Never reformat one until the wedding is on two "
         "drives - that rule is the business."),
        ("SanDisk 64GB SD x8", "Memory card", "SanDisk",
         "SD-SDXC-BATCH-4", "In use", "Fair", _phg(2), -1500, 440,
         3, "No", 0, 0, 0, "", 480,
         "Past their rated life. Replace before next season."),
        ("Canon LP-E6NH batteries x8", "Battery", "Canon",
         "LPE6-BATCH-11", "In use", "Fair", _phg(1), -800, 640, 3,
         "No", 0, 0, 0, "", 680,
         "Two of the eight no longer hold a full day."),
        ("ThinkTank Airport Advantage", "Camera bag", "ThinkTank",
         "TT-AA-55302", "In use", "Good", _phg(1), -1600, 329, 10,
         "No", 0, 0, 0, "", 349, ""),
        ("Samsung T7 4TB (working)", "Backup drive", "Samsung",
         "T7-990214", "In use", "Excellent", _phg(1), -300, 299, 4,
         "No", 0, 0, 0, "", 299, ""),
        ("Samsung T7 4TB (mirror)", "Backup drive", "Samsung",
         "T7-990215", "In use", "Excellent", _phg(1), -300, 299, 4,
         "No", 0, 0, 0, "Samsung T7 4TB (working)", 299,
         "The second copy. A wedding is not backed up until it is "
         "on this one too."),
        ("Synology DS223 NAS 16TB", "Backup drive", "Synology",
         "DS223-440178", "In use", "Excellent", _phg(1), -200, 749,
         6, "Yes", 700, 0, 0, "", 799,
         "Third copy, off the desk. Cheaper than losing one "
         "wedding."),
        ("Mac Studio M2 Max", "Editing computer", "Apple",
         "MS-M2-227741", "In use", "Excellent", _phg(1), -480, 2399,
         5, "Yes", 2000, 0, 0, "", 2499, ""),
        ("BenQ SW271C monitor", "Monitor", "BenQ", "BQ-SW271-88120",
         "In use", "Good", _phg(1), -1020, 1099, 7, "No", 0, 0, 0,
         "", 1149, "Calibrate it monthly or the albums come back "
                   "magenta."),
        ("Adobe Creative Cloud Photography", "Software licence",
         "Adobe", "ADBE-SUB-2026", "In use", "New", _phg(1), -365,
         239, 1, "No", 0, 0, 0, "", 239,
         "Renews annually. Sits in Expense Tracker as a recurring "
         "cost, not here."),
        ("Nikon D750 (old body)", "Camera body", "Nikon",
         "D750-118827", "Retired", "Retired", "", -3100, 1899, 5,
         "No", 0, 486000, 400000, "", 0,
         "Kept for parts and nostalgia. Written off years ago."),
    ]
    rows: Rows = []
    for (item, category, make, serial, status, condition, assigned,
         bought, price, life, insured, value, shutter, limit,
         backup_for, replace, note) in kit:
        rows.append({
            "item": item, "category": category, "make": make,
            "serial": serial, "status": status, "condition": condition,
            "assigned": assigned, "location": "Studio"
            if not assigned else "With the photographer",
            "bought": _day(cfg, bought), "price": price,
            "life_years": life, "insured": insured,
            "insured_value": value,
            "warranty_to": _day(cfg, bought + 730),
            "serviced": _day(cfg, bought + 365)
            if category == "Camera body" else "",
            "shutter": shutter, "shutter_limit": limit,
            "backup_for": backup_for, "replace_cost": replace,
            "notes": note,
        })
    return rows


# ---------------------------------------------------------------------------
# 13. Travel & Mileage
# ---------------------------------------------------------------------------

def travel(cfg: Config) -> Rows:
    base = "Studio, Portland OR"
    plan = [
        # index, purpose, miles, mode, fuel, tolls, parking, accom,
        # other, hours, receipt
        (1, "Wedding day", 62, "Car", 0, 0, 12, 0, 0, 2.5, "Yes"),
        (1, "Venue visit", 62, "Car", 0, 0, 0, 0, 0, 2.5, "No"),
        (2, "Wedding day", 31, "Car", 0, 0, 0, 0, 0, 1.2, "No"),
        (2, "Engagement shoot", 31, "Car", 0, 0, 8, 0, 0, 1.2, "Yes"),
        (3, "Wedding day", 9, "Car", 0, 0, 22, 0, 0, 0.5, "Yes"),
        (4, "Wedding day", 26, "Car", 0, 0, 0, 0, 0, 1.0, "No"),
        (4, "Album delivery", 26, "Car", 0, 0, 6, 0, 0, 1.0, "Yes"),
        (5, "Wedding day", 7, "Car", 0, 0, 30, 0, 0, 0.4, "Yes"),
        (6, "Wedding day", 14, "Car", 0, 0, 0, 0, 0, 0.6, "No"),
        (7, "Wedding day", 19, "Car", 0, 0, 15, 0, 0, 0.8, "Yes"),
        (7, "Rehearsal", 19, "Car", 0, 0, 0, 0, 0, 0.8, "No"),
        (8, "Venue visit", 48, "Car", 0, 0, 0, 0, 0, 1.9, "No"),
        (9, "Venue visit", 97, "Car", 0, 4.50, 0, 0, 0, 3.6, "Yes"),
        (10, "Venue visit", 164, "Car", 0, 0, 10, 0, 0, 5.8, "Yes"),
        (12, "Venue visit", 288, "Car", 0, 18.00, 0, 145, 42,
         9.5, "Yes"),
        (14, "Venue visit", 0, "Flight", 0, 0, 0, 0, 0, 0, "No"),
    ]
    rows: Rows = []
    for (index, purpose, miles, mode, fuel, tolls, parking, accom,
         other, hours, receipt) in plan:
        offset = WEDDINGS[index - 1][0]
        when = offset if purpose == "Wedding day" else offset - 120
        rows.append({
            "date": _day(cfg, when), "booking": _wed(index),
            "purpose": purpose,
            "from_place": base,
            "to_place": f"{WEDDINGS[index - 1][2]}, "
                        f"{WEDDINGS[index - 1][4]}",
            "mode": mode, "one_way": miles,
            "return_trip": "Yes", "scout_trips": 0,
            "fuel": fuel, "tolls": tolls, "parking": parking,
            "accommodation": accom, "other_cost": other,
            "charge_override": 0, "hours": hours, "receipt": receipt,
            "notes": "",
        })

    # the two that teach something
    rows[-2]["notes"] = ("Victoria and back with an overnight. The "
                         "travel fee covers the fuel and nothing "
                         "else - ten hours of my life are free.")
    rows[-1].update({
        "mode": "Flight", "one_way": 0, "return_trip": "Yes",
        "fuel": 0, "other_cost": 680, "accommodation": 420,
        "hours": 14, "receipt": "Yes",
        "notes": "Oaxaca. Flights and three nights. The $1,250 "
                 "travel fee does not cover this and I agreed to "
                 "it anyway."})

    rows += [
        {"date": _day(cfg, -45), "booking": "",
         "purpose": "Training or workshop", "from_place": base,
         "to_place": "Seattle WA", "mode": "Train", "one_way": 0,
         "return_trip": "Yes", "scout_trips": 0, "fuel": 0,
         "tolls": 0, "parking": 0, "accommodation": 165,
         "other_cost": 128, "charge_override": 0, "hours": 7.5,
         "receipt": "Yes",
         "notes": "Flash workshop. Not chargeable to anybody, still "
                  "the best money spent this year."},
        {"date": _day(cfg, -12), "booking": "",
         "purpose": "Networking", "from_place": base,
         "to_place": "Portland city centre", "mode": "Car",
         "one_way": 6, "return_trip": "Yes", "scout_trips": 0,
         "fuel": 0, "tolls": 0, "parking": 9, "accommodation": 0,
         "other_cost": 0, "charge_override": 0, "hours": 2.5,
         "receipt": "Yes",
         "notes": "Venue coordinators' drinks. Two of my last five "
                  "bookings came from that room."},
        {"date": _day(cfg, -3), "booking": "",
         "purpose": "Equipment collection", "from_place": base,
         "to_place": "Pro Photo Supply", "mode": "Car", "one_way": 4,
         "return_trip": "Yes", "scout_trips": 0, "fuel": 0,
         "tolls": 0, "parking": 0, "accommodation": 0,
         "other_cost": 0, "charge_override": 0, "hours": 0.8,
         "receipt": "No", "notes": "Dropped the 70-200 in for "
                                   "repair."},
    ]
    return rows


# ---------------------------------------------------------------------------
# 14. Marketing Calendar
# ---------------------------------------------------------------------------

def marketing(cfg: Config) -> Rows:
    plan = [
        # offset, platform, type, title, booking index, status,
        # reach, engagement, saves, clicks, enquiries, spend, hours
        (-395, "Instagram", "Sneak peek", "Nadia & Priya - first "
         "look in the barn", 1, "Published", 4200, 388, 61, 44, 2,
         0, 1.0),
        (-360, "Blog", "Full wedding gallery", "A September barn "
         "wedding at The Old Mill", 1, "Evergreen", 1850, 42, 0,
         118, 3, 0, 3.5),
        (-298, "Instagram", "Carousel", "Cascade Vineyard in "
         "autumn light", 2, "Published", 6100, 540, 94, 72, 1, 0,
         1.2),
        (-280, "Pinterest", "Portfolio update", "Vineyard wedding "
         "inspiration board", 2, "Evergreen", 22400, 310, 1180,
         264, 4, 0, 2.0),
        (-250, "Instagram", "Reel / video", "Sixty seconds of a "
         "vineyard wedding", 2, "Published", 18700, 1420, 206, 88,
         2, 40, 2.5),
        (-186, "Instagram", "Sneak peek", "St Agnes, and the light "
         "through those windows", 3, "Published", 3900, 301, 44, 31,
         0, 0, 0.8),
        (-150, "Blog", "Blog post", "How to plan family groups "
         "without losing an hour", "", "Evergreen", 2400, 61, 0,
         190, 5, 0, 4.0),
        (-120, "Instagram", "Educational post", "Why your "
         "photographer keeps asking about the sunset", "",
         "Published", 5200, 410, 128, 52, 1, 0, 1.0),
        (-96, "Instagram", "Behind the scenes", "What is actually "
         "in the bag", "", "Published", 3100, 266, 38, 19, 0, 0,
         0.6),
        (-88, "Instagram", "Sneak peek", "Willow Creek in high "
         "summer", 4, "Published", 4600, 352, 57, 40, 1, 0, 0.9),
        (-70, "Wedding directory", "Promotion", "Autumn date "
         "availability listing", "", "Archived", 980, 0, 0, 37, 2,
         149, 0.5),
        (-62, "Instagram", "Client testimonial", "What Margit said "
         "about the gallery", 5, "Published", 2800, 240, 31, 26, 1,
         0, 0.5),
        (-48, "Facebook", "Full wedding gallery", "Rosegate Gardens "
         "- a small wedding that was not small at all", 6,
         "Published", 1600, 88, 0, 22, 0, 0, 1.5),
        (-40, "Email newsletter", "Promotion", "Two 2027 Saturdays "
         "left", "", "Published", 420, 96, 0, 58, 3, 0, 1.2),
        (-30, "Instagram", "Carousel", "Five things to do the week "
         "before your wedding", "", "Published", 7400, 610, 212,
         64, 2, 0, 1.5),
        (-21, "TikTok", "Reel / video", "A day in the life of a "
         "wedding photographer", "", "Published", 31200, 2840, 410,
         96, 1, 0, 2.0),
        (-14, "Google Business", "Portfolio update", "Refreshed "
         "the profile gallery", "", "Published", 1240, 18, 0, 71,
         2, 0, 0.5),
        (-7, "Instagram", "Story", "Behind the scenes at the "
         "Haddad wedding", 7, "Published", 2100, 180, 12, 14, 0, 0,
         0.3),
        (-2, "Instagram", "Sneak peek", "Yusra & Kaleem - the "
         "nikah", 7, "Drafted", 0, 0, 0, 0, 0, 0, 0.4),
        (1, "Instagram", "Sneak peek", "Yusra & Kaleem - the "
         "sparkler exit", 7, "Scheduled", 0, 0, 0, 0, 0, 0, 0.4),
        (4, "Pinterest", "Portfolio update", "Barn weddings in "
         "Oregon", "", "Scheduled", 0, 0, 0, 0, 0, 0, 1.0),
        (9, "Blog", "Blog post", "A winter wedding at The Granary",
         7, "Idea", 0, 0, 0, 0, 0, 0, 0),
        (16, "Instagram", "Educational post", "What a wedding "
         "photographer actually does in the six weeks after your "
         "wedding", "", "Idea", 0, 0, 0, 0, 0, 0, 0),
        (23, "Email newsletter", "Promotion", "Engagement shoot "
         "slots for spring", "", "Idea", 0, 0, 0, 0, 0, 0, 0),
        (30, "Wedding fair", "Promotion", "Portland Winter Wedding "
         "Fair stand", "", "Scheduled", 0, 0, 0, 0, 0, 495, 12.0),
        (45, "Instagram", "Client testimonial", "Ask Delphine for "
         "a line about the gallery", 6, "Idea", 0, 0, 0, 0, 0, 0,
         0),
    ]
    rows: Rows = []
    for (offset, platform, ctype, title, index, status, reach,
         engagement, saves, clicks, enquiries, spend, hours) in plan:
        booking = _wed(index) if index else ""
        rows.append({
            "date": _day(cfg, offset), "platform": platform,
            "type": ctype, "title": title,
            "content": "", "booking": booking,
            "permission": "Yes" if booking else "No",
            "status": status,
            "posted_on": _day(cfg, offset)
            if status in ("Published", "Evergreen", "Archived")
            else "",
            "reach": reach, "engagement": engagement, "saves": saves,
            "clicks": clicks, "manual_enquiries": enquiries,
            "spend": spend, "hours": hours, "notes": "",
        })
    rows[3]["notes"] = ("Pinterest is slow and then it is not. This "
                        "one board has outperformed a year of "
                        "stories.")
    rows[10]["notes"] = ("$149 for two enquiries, neither of which "
                         "booked. Do not renew.")
    rows[24]["notes"] = ("$495 and twelve hours. Last year's fair "
                         "paid for itself twice over - check "
                         "Marketing Calendar before committing "
                         "again.")
    return rows


# ---------------------------------------------------------------------------
# 15. Expense Tracker
# ---------------------------------------------------------------------------

def expenses(cfg: Config) -> Rows:
    rows: Rows = []

    # recurring overheads, thirteen months of them
    recurring = [
        ("Adobe", "Software subscriptions", "Creative Cloud "
         "Photography plan", 19.99, "Monthly", "Operating expense"),
        ("Pic-Time", "Software subscriptions", "Gallery hosting",
         29.00, "Monthly", "Operating expense"),
        ("Squarespace", "Website & hosting", "Website and domain",
         23.00, "Monthly", "Operating expense"),
        ("Backblaze", "Memory & storage", "Off-site backup",
         9.00, "Monthly", "Operating expense"),
    ]
    for month in range(13, 0, -1):
        when = -month * 30 + 4
        for supplier, category, description, amount, freq, tax_cat \
                in recurring:
            rows.append({
                "date": _day(cfg, when), "supplier": supplier,
                "category": category, "description": description,
                "booking": "", "amount": amount, "tax_amount": 0,
                "method": "Credit card", "receipt": "Yes",
                "reference": "", "recurring": freq,
                "business_pct": 1.0, "tax_category": tax_cat,
                "notes": ""})

    one_offs = [
        (-420, "Hiscox", "Insurance", "Annual public liability and "
         "kit cover", 612.00, "One off", "Operating expense", "",
         "Yes", "Renews in about three months. Dev's own cover has "
         "already lapsed - that is a separate problem."),
        (-400, "Dev Ashworth", "Second shooters", "Second shooter, "
         "Raman-Okafor wedding", 350.00, "One off", "Cost of sale",
         1, "Yes", ""),
        (-396, "Loxley Colour", "Albums & prints", "20 page album, "
         "Raman-Okafor", 185.00, "One off", "Cost of sale", 1,
         "Yes", ""),
        (-392, "Shell", "Travel & mileage", "Fuel", 58.20, "One off",
         "Operating expense", "", "Yes", ""),
        (-302, "Dev Ashworth", "Second shooters", "Second shooter, "
         "Tanaka-Reed wedding", 350.00, "One off", "Cost of sale",
         2, "Yes", ""),
        (-298, "Tomas Bright", "Editing services", "24 hours "
         "editing, Tanaka-Reed", 672.00, "One off", "Cost of sale",
         2, "Yes", "Twenty-eight dollars an hour against my own "
                   "sixty-five. This is the cheapest line on the "
                   "sheet."),
        (-290, "Loxley Colour", "Albums & prints", "30 page album, "
         "Tanaka-Reed", 265.00, "One off", "Cost of sale", 2, "Yes",
         ""),
        (-250, "Meta", "Marketing & advertising", "Boosted reel",
         40.00, "One off", "Operating expense", "", "Yes", ""),
        (-190, "Joelle Nakamura", "Second shooters", "Second "
         "shooter, Castellanos wedding", 375.00, "One off",
         "Cost of sale", 3, "Yes", ""),
        (-186, "Loxley Colour", "Albums & prints", "30 page album, "
         "Castellanos", 265.00, "One off", "Cost of sale", 3, "Yes",
         ""),
        (-180, "Pro Photo Supply", "Camera equipment", "Canon RF "
         "85mm f1.2L", 2699.00, "One off", "Capital / asset", "",
         "Yes", "Capital, not an expense. It belongs on Equipment "
                "Inventory too - put it in both places."),
        (-120, "Pro Photo Supply", "Memory & storage", "Six "
         "CFexpress cards", 1080.00, "One off", "Capital / asset",
         "", "Yes", ""),
        (-96, "Dev Ashworth", "Second shooters", "Second shooter, "
         "Vance wedding", 350.00, "One off", "Cost of sale", 4,
         "Yes", ""),
        (-92, "Pro Photo Supply", "Equipment rental", "70-200mm "
         "hire while mine was in for service", 95.00, "One off",
         "Cost of sale", 4, "Yes", ""),
        (-70, "Wedding Directory Co", "Marketing & advertising",
         "Autumn listing", 149.00, "One off", "Operating expense",
         "", "Yes", "Two enquiries, no bookings. Not renewing."),
        (-64, "Tomas Bright", "Editing services", "16.5 hours "
         "editing, Lindqvist", 462.00, "One off", "Cost of sale", 5,
         "Yes", ""),
        (-60, "Loxley Colour", "Albums & prints", "20 page album, "
         "Lindqvist", 185.00, "One off", "Cost of sale", 5, "Yes",
         ""),
        (-45, "Amtrak", "Education & training", "Seattle flash "
         "workshop - travel and ticket", 293.00, "One off",
         "Operating expense", "", "Yes", ""),
        (-40, "Canon Service", "Equipment repair", "70-200mm focus "
         "motor", 340.00, "One off", "Operating expense", "", "Yes",
         "Three week turnaround and a wedding in six days."),
        (-32, "The Flower Press", "Client gifts", "Thank-you "
         "bouquets, two", 96.00, "One off", "Operating expense", "",
         "Yes", ""),
        (-29, "Priya Raman", "Assistants", "Assistant, Mbeki "
         "wedding", 80.00, "One off", "Cost of sale", 6, "Yes", ""),
        (-21, "Peak Accounting", "Accounting & legal", "Quarterly "
         "bookkeeping", 240.00, "Quarterly", "Operating expense",
         "", "Yes", ""),
        (-14, "Stripe", "Bank & card fees", "Card processing fees, "
         "last month", 31.12, "Monthly", "Operating expense", "",
         "No", "No receipt - comes off the statement. Fine for the "
               "accountant, not fine for the receipt column."),
        (-8, "Noissue", "Packaging & postage", "USB boxes and "
         "tissue, 25 sets", 148.00, "One off", "Cost of sale", "",
         "Yes", ""),
        (-2, "Dev Ashworth", "Second shooters", "Second shooter, "
         "Haddad wedding", 350.00, "One off", "Cost of sale", 7,
         "Yes", ""),
        (-1, "The Granary at Fir Point", "Meals on shoot", "Supplier "
         "meals, two", 44.00, "One off", "Operating expense", 7,
         "No", "Lost the receipt. Forty-four dollars I cannot "
               "prove."),
        (-1, "Portland Parking", "Travel & mileage", "Venue car "
         "park", 15.00, "One off", "Operating expense", 7, "Yes",
         ""),
        (-310, "Marisa Calloway", "Tax & national insurance",
         "Quarterly estimated tax", 2400.00, "Quarterly",
         "Not deductible", "", "Yes", "Not an expense. It is here "
         "so the bank balance makes sense."),
        (-120, "Marisa Calloway", "Tax & national insurance",
         "Quarterly estimated tax", 2150.00, "Quarterly",
         "Not deductible", "", "Yes", ""),
        (-250, "Studio Collective", "Studio rent", "Desk and "
         "meeting room, six months up front", 1740.00,
         "Twice a year", "Operating expense", "", "Yes", ""),
        (-61, "Studio Collective", "Studio rent", "Desk and "
         "meeting room, six months up front", 1740.00,
         "Twice a year", "Operating expense", "", "Yes", ""),
        (-200, "Marisa Calloway", "Miscellaneous", "Monthly owner "
         "drawing", 2200.00, "Monthly", "Owner drawing", "", "No",
         "Owner drawings are not a business cost. Tagging them "
         "correctly is what keeps the profit figure honest."),
        (-5, "Pro Photo Supply", "Lens purchases", "Deposit on a "
         "replacement body", 500.00, "One off", "Capital / asset",
         "", "Yes", "Body 2 is at 93% of its rated shutter life. "
                    "Equipment Inventory has been saying so for "
                    "months."),
        (-18, "Apple", "Office expenses", "Printer ink and paper",
         62.40, "One off", "Operating expense", "", "No", ""),
        (-9, "City of Portland", "Licences & permits", "Business "
         "licence renewal", 110.00, "Annual", "Operating expense",
         "", "Yes", ""),
    ]
    for (offset, supplier, category, description, amount, freq,
         tax_cat, index, receipt, note) in one_offs:
        rows.append({
            "date": _day(cfg, offset), "supplier": supplier,
            "category": category, "description": description,
            "booking": _wed(index) if index else "",
            "amount": amount, "tax_amount": 0,
            "method": "Credit card" if category in
            ("Software subscriptions", "Website & hosting")
            else "Bank transfer",
            "receipt": receipt, "reference": "",
            "recurring": freq,
            "business_pct": 0.5 if category == "Office expenses"
            else 1.0,
            "tax_category": tax_cat, "notes": note})
    return rows


# ---------------------------------------------------------------------------
# 16. Profit Analysis
# ---------------------------------------------------------------------------

def profit(cfg: Config) -> Rows:
    """One line per wedding. Everything else on that tab is a
    formula - this register is deliberately one column wide."""
    return [{"booking": _wed(index)}
            for index in range(1, len(WEDDINGS) + 1)]


# ---------------------------------------------------------------------------
# 17. Reviews & Referrals
# ---------------------------------------------------------------------------

def reviews(cfg: Config) -> Rows:
    plan = [
        (1, "Email", 0, "Received", "Google", 5,
         "They caught my grandmother laughing",
         "We have one photograph of my grandmother laughing with "
         "both her grandsons and she died in January. Marisa did "
         "not ask us to pose for it. She just saw it.",
         "Yes", "Website homepage, Google profile", "Yes", 2,
         "Yes", ""),
        (2, "In the gallery", 1, "Received", "Google", 5,
         "Worth every penny and then some",
         "We booked the Premium package and spent the whole day "
         "forgetting there was a photographer there. The album "
         "arrived three weeks early.",
         "Yes", "Instagram, Google profile", "Yes", 1, "Yes", ""),
        (3, "Email", 2, "Received", "Wedding directory", 4,
         "Beautiful photographs, slow album",
         "The photographs are everything we hoped for. The album "
         "took longer than we were told, which is the only reason "
         "this is not five stars.",
         "No", "", "Yes", 1, "Yes",
         "Four stars for a reason I control. The album was 26 days "
         "late. Fix the process, not the review."),
        (4, "Email", 2, "Chased", "Google", "", "", "", "No", "",
         "No", 0, "No",
         "Asking for a review while the gallery is three weeks "
         "late is not a good idea. Deliver first, then ask."),
        (5, "", 0, "Not requested", "", "", "", "", "No", "", "No",
         0, "No",
         "Delivered 34 days ago, review never asked for. This is "
         "the easiest thing on the whole dashboard to fix."),
        (6, "", 0, "Not requested", "", "", "", "", "No", "", "No",
         0, "No",
         "Balance unpaid. Ask for the money before asking for the "
         "review."),
        (2, "Text message", 0, "Received", "Facebook", 5,
         "Second time we have used Rowan & Fern",
         "Used Marisa for our engagement shoot and then the "
         "wedding. Same person, same care, both times.",
         "Yes", "Facebook page", "Yes", 0, "Yes",
         "A second review from the same couple on a different "
         "platform. Perfectly normal and worth asking for."),
        (1, "Thank-you card", 0, "Declined", "", "", "", "", "No",
         "", "No", 1, "Yes",
         "Did not want to write a public review and sent two "
         "couples instead. That is a better outcome."),
        (3, "Automatic reminder", 3, "No review left", "Google", "",
         "", "", "No", "", "No", 0, "No",
         "Three automatic reminders and nothing. Stop at two."),
        (5, "In person", 0, "Requested", "Google", "", "", "", "No",
         "", "No", 0, "No",
         "Asked at the album handover. Give it a fortnight before "
         "chasing."),
    ]
    rows: Rows = []
    for (index, how, reminders, status, platform, rating, headline,
         quote, can_use, used_on, responded, referrals, thanked,
         note) in plan:
        offset = WEDDINGS[index - 1][0]
        asked = _day(cfg, offset + 40) if how else ""
        received = _day(cfg, offset + 46) if status == "Received" \
            else ""
        rows.append({
            "booking": _wed(index), "asked_on": asked,
            "how_asked": how, "reminders": reminders,
            "status": status, "received_on": received,
            "platform": platform, "rating": rating,
            "headline": headline, "quote": quote,
            "can_use": can_use, "used_on": used_on,
            "responded": responded, "referrals": referrals,
            "thanked": thanked, "notes": note})
    return rows


# ---------------------------------------------------------------------------
# registry
# ---------------------------------------------------------------------------

#: keyed by register *module* name, which is what each register's
#: ``spec()`` passes to :func:`get`.
_BUILDERS: Dict[str, Callable[[Any], List[Dict[str, Any]]]] = {
    "clients": clients,
    "leads": leads,
    "bookings": bookings,
    "packages": packages,
    "staff": staff,
    "contracts": contracts,
    "payments": payments,
    "timeline": timeline,
    "shots": shots,
    "editing": editing,
    "gallery": gallery,
    "equipment": equipment,
    "travel": travel,
    "marketing": marketing,
    "expenses": expenses,
    "profit": profit,
    "reviews": reviews,
}


def get(name: str, cfg) -> List[Dict[str, Any]]:
    """Sample rows for a register, or [] when samples are switched off."""
    if not getattr(cfg, "include_sample_data", True):
        return []
    builder = _BUILDERS.get(name)
    return builder(cfg) if builder else []
````

### `wedding_photography_spreadsheet/build.py`

*165 lines · 3 functions*

> Workbook builder: puts the twenty-three tabs together in order.
>
> Typical use::
>
>     from wedding_photography_spreadsheet import build_workbook
>     build_workbook("Wedding Photographer Business Manager.xlsx")

| Defines | What it does |
|---------|--------------|
| `build_register(book, fmts, cfg, reg, module)` | Create one register sheet from its :class:`SheetSpec` |
| `build_workbook(path, cfg)` | Write the whole product to *path* and return the path |
| `_set_properties(book, cfg)` |  |

**Module constants:** `REGISTER_MODULES` (17 items), `SHEET_ORDER` (23 items), `PAGE_BUILDERS` (6 entries), `REGISTER_BY_NAME` ({module.NAME: module for module in REGISTER_MODULES}).

````python
"""Workbook builder: puts the twenty-three tabs together in order.

Typical use::

    from wedding_photography_spreadsheet import build_workbook
    build_workbook("Wedding Photographer Business Manager.xlsx")
"""

from __future__ import annotations

import datetime as _dt
import inspect
import os
from typing import Callable, Dict, List, Optional

from xlsxwriter.workbook import Workbook

from .config import VERSION, Config
from .formats import FormatBook
from .layout import Registry, SheetSpec
from .names import define_cross_sheet_names
from .sheets import (bookings, calendar, clients, contracts, dashboard,
                     editing, equipment, expenses, gallery,
                     getting_started, leads, marketing, packages,
                     payments, profit, questionnaire, reports, reviews,
                     settings, shots, staff, timeline, travel)
from .writer import SheetWriter

#: register modules, in the order their specs are registered
REGISTER_MODULES = [
    clients, leads, bookings, packages, contracts, payments, timeline,
    shots, staff, editing, gallery, equipment, travel, marketing,
    expenses, profit, reviews,
]

#: tab order of the finished workbook
SHEET_ORDER: List[str] = [
    "Dashboard",
    "Getting Started",
    "Settings & Lists",
    "Client CRM",
    "Lead Tracker",
    "Booking Calendar",
    "Wedding Bookings",
    "Packages & Pricing",
    "Contracts & Documents",
    "Payments & Invoices",
    "Wedding Timeline",
    "Shot List Manager",
    "Photographer & Staff",
    "Editing Workflow",
    "Gallery & Delivery",
    "Equipment Inventory",
    "Travel & Mileage",
    "Marketing Calendar",
    "Expense Tracker",
    "Profit Analysis",
    "Client Questionnaire",
    "Reviews & Referrals",
    "Monthly & Annual Reports",
]

#: sheets that are not registers - name -> builder(book, fmts, cfg, reg)
PAGE_BUILDERS: Dict[str, Callable] = {
    "Dashboard": dashboard.build,
    "Getting Started": getting_started.build,
    "Settings & Lists": settings.build,
    "Booking Calendar": calendar.build,
    "Client Questionnaire": questionnaire.build,
    "Monthly & Annual Reports": reports.build,
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


def build_workbook(path: str = "Wedding Photography Spreadsheet.xlsx",
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

    # 3. Live lookup names - Booking IDs above all, because twelve tabs
    #    take their drop-down from it.
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
        "subject": "Wedding photography enquiries, bookings, packages, "
                   "payments, timelines, shot lists, editing, galleries "
                   "and profit",
        "author": cfg.author,
        "manager": cfg.author,
        "company": cfg.author,
        "category": "Small business spreadsheet template",
        "keywords": ("wedding photographer spreadsheet, photography "
                     "crm, wedding booking tracker, photography "
                     "pricing calculator, shot list template, wedding "
                     "timeline planner, photography invoice tracker, "
                     "client management, gallery delivery tracker, "
                     "photography business template"),
        "comments": (f"{cfg.product_name}. Created by {cfg.author}. "
                     f"Version {VERSION}. Formula cells are locked "
                     f"and every sheet is protected with the password "
                     f"'{cfg.password}'. A business organisation tool: "
                     f"it does not connect to any booking platform and "
                     f"does not generate or sign contracts."),
        "status": "Final",
        "created": _dt.datetime(cfg.today.year, cfg.today.month,
                                cfg.today.day),
    })
````

---

## 3. Shared column factories

The helpers every register uses: automatic ID columns, date and month keys, ranking columns, alert conditional formats and the lookup that joins a row back to its wedding.

### `wedding_photography_spreadsheet/sheets/__init__.py`

*0 lines*

````python

````

### `wedding_photography_spreadsheet/sheets/common.py`

*560 lines · 57 functions*

> Small building blocks shared by the register sheets.
>
> Everything a register needs more than once lives here: the
> standard columns, the conditional format recipes, the
> blank-row guards and the cross sheet lookups.  A register
> module should read like a description of its table, not
> like a formula library.

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
| `lookup_client(target, key, missing)` |  |
| `lookup_client_num(target, key)` |  |
| `lookup_booking(target, key, missing)` |  |
| `lookup_booking_num(target, key)` |  |
| `lookup_booking_num_zero(target, key)` |  |
| `booking_client_col(key)` | Hidden column carrying the Client ID behind a Booking ID |
| `lookup_package(target, key, missing)` |  |
| `lookup_package_num(target, key)` |  |
| `lookup_staff(target, key, missing)` |  |
| `lookup_staff_num(target, key)` |  |
| `lookup_lead(target, key, missing)` |  |
| `lookup_equipment(target, key, missing)` |  |
| `in_list(cell, values)` | `TRUE` when `cell` matches any of `values` |
| `is_holding(cell)` | The booking still occupies the date in the calendar |
| `is_earning(cell)` | The booking counts towards revenue |
| `is_shot(cell)` | The wedding day has happened, so editing and delivery apply |
| `is_open_lead(cell)` | A lead that is still live and still worth chasing |
| `is_won_lead(cell)` |  |
| `is_capital(cell)` | A capital purchase rather than a running cost |
| `is_direct_cost(cell)` | A cost that belongs to one wedding rather than to the business |
| `rank_desc(key, guard, header, width, note)` | Hidden dense rank, 1 = the biggest value in `key` |
| `rank_asc(key, guard, header, width, note)` | Hidden dense rank, 1 = the smallest value in `key` |
| `seq_key(parent, trigger)` | Hidden `1, 2, 3 ...` within a parent, plus a `parent|n` key |
| `package_value(column, package_cell)` | A figure from the editable package table on Packages & Pricing |
| `package_text(column, package_cell)` | As :func:`package_value` but for an inclusion written in words |
| `days_until(date_cell)` | Whole days from today to a date - negative once it has passed |
| `countdown_band(days_cell)` | Turn a day count into the band the dashboard groups on |
| `hours_between(start, end)` | Hours between two time cells, tolerating a finish after midnight |
| `sales_tax(amount)` | Sales tax / VAT on a package total |
| `monthly_equivalent(amount, frequency_cell)` | Turn any recurring cost into what it really costs each month |
| `mileage_cost(miles, allowance)` | Mileage charged at the configured rate, past the free allowance |

**Module constants:** `ROW_OFFSET` (HEADER_ROW).

````python
"""Small building blocks shared by the register sheets.

Everything a register needs more than once lives here: the
standard columns, the conditional format recipes, the
blank-row guards and the cross sheet lookups.  A register
module should read like a description of its table, not
like a formula library.
"""

from __future__ import annotations

from typing import Any, Dict, List

from ..config import (CAPITAL_CATEGORIES, DIRECT_CATEGORIES,
                      EARNING_STATUSES, HOLDING_STATUSES,
                      OPEN_LEAD_STATUSES, SHOT_STATUSES,
                      WON_LEAD_STATUSES)
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


def lookup_client(target: str, key: str = "client",
                  missing: str = "Not in Client CRM") -> str:
    return lookup("Client CRM", target, "id", key, missing)


def lookup_client_num(target: str, key: str = "client") -> str:
    return lookup_num("Client CRM", target, "id", key)


def lookup_booking(target: str, key: str = "booking",
                   missing: str = "Not in Wedding Bookings") -> str:
    return lookup("Wedding Bookings", target, "id", key, missing)


def lookup_booking_num(target: str, key: str = "booking") -> str:
    return lookup_num("Wedding Bookings", target, "id", key)


def lookup_booking_num_zero(target: str, key: str = "booking") -> str:
    return lookup_num_zero("Wedding Bookings", target, "id", key)


def booking_client_col(key: str = "booking") -> Col:
    """Hidden column carrying the Client ID behind a Booking ID.

    The Client CRM rolls payments, galleries and reviews up per couple.
    Storing the client against each row keeps those `COUNTIFS` one
    dimensional instead of a lookup inside a lookup inside a criteria
    argument, which Excel will not do.
    """
    return Col("client", "Client ID", 11, "id",
               formula=(f'=IF({{{key}}}{{r}}="","",'
                        f'IFERROR(INDEX([[Wedding Bookings.client]],'
                        f'MATCH({{{key}}}{{r}},'
                        f'[[Wedding Bookings.id]],0))&"",""))'),
               hidden=True)


def lookup_package(target: str, key: str = "package",
                   missing: str = "Not in Packages & Pricing") -> str:
    return lookup("Packages & Pricing", target, "name", key, missing)


def lookup_package_num(target: str, key: str = "package") -> str:
    return lookup_num("Packages & Pricing", target, "name", key)


def lookup_staff(target: str, key: str = "photographer",
                 missing: str = "Not in Photographer & Staff") -> str:
    return lookup("Photographer & Staff", target, "id", key, missing)


def lookup_staff_num(target: str, key: str = "photographer") -> str:
    return lookup_num("Photographer & Staff", target, "id", key)


def lookup_lead(target: str, key: str = "lead",
                missing: str = "Not in Lead Tracker") -> str:
    return lookup("Lead Tracker", target, "id", key, missing)


def lookup_equipment(target: str, key: str = "equipment",
                     missing: str = "Not in Equipment Inventory") -> str:
    return lookup("Equipment Inventory", target, "id", key, missing)


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
    """The booking still occupies the date in the calendar."""
    return in_list(cell, HOLDING_STATUSES)


def is_earning(cell: str) -> str:
    """The booking counts towards revenue."""
    return in_list(cell, EARNING_STATUSES)


def is_shot(cell: str) -> str:
    """The wedding day has happened, so editing and delivery apply."""
    return in_list(cell, SHOT_STATUSES)


def is_open_lead(cell: str) -> str:
    """A lead that is still live and still worth chasing."""
    return in_list(cell, OPEN_LEAD_STATUSES)


def is_won_lead(cell: str) -> str:
    return in_list(cell, WON_LEAD_STATUSES)


def is_capital(cell: str) -> str:
    """A capital purchase rather than a running cost."""
    return in_list(cell, CAPITAL_CATEGORIES)


def is_direct_cost(cell: str) -> str:
    """A cost that belongs to one wedding rather than to the business."""
    return in_list(cell, DIRECT_CATEGORIES)


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

def package_value(column: str, package_cell: str) -> str:
    """A figure from the editable package table on Packages & Pricing.

    An unknown package name is worth zero rather than breaking the row,
    because a photographer who types a new package should get a number
    they can correct, not an error across twenty columns.
    """
    return (f'IFERROR(INDEX({column},MATCH({package_cell},'
            f'PackageNameList,0))*1,0)')


def package_text(column: str, package_cell: str) -> str:
    """As :func:`package_value` but for an inclusion written in words."""
    return (f'IFERROR(INDEX({column},MATCH({package_cell},'
            f'PackageNameList,0))&"","")')


def days_until(date_cell: str) -> str:
    """Whole days from today to a date - negative once it has passed."""
    return f'({date_cell}-TodayDate)'


def countdown_band(days_cell: str) -> str:
    """Turn a day count into the band the dashboard groups on."""
    return (f'IF({days_cell}<0,"Past",'
            f'IF({days_cell}<=7,"This week",'
            f'IF({days_cell}<=30,"Next 30 days",'
            f'IF({days_cell}<=90,"1-3 months",'
            f'IF({days_cell}<=180,"3-6 months",'
            f'IF({days_cell}<=365,"6-12 months","Over a year"))))))')


def hours_between(start: str, end: str) -> str:
    """Hours between two time cells, tolerating a finish after midnight."""
    return (f'IF({start}="",0,IF({end}="",0,'
            f'ROUND((({end}-{start})+({end}<{start}))*24,2)))')


def sales_tax(amount: str) -> str:
    """Sales tax / VAT on a package total.

    Whether photography is taxable at all, and at what rate, is a local
    question, so the rate and the switch both live on Settings and the
    rate ships at zero.
    """
    return (f'IF(TaxOnPackages<>"Yes",0,ROUND({amount}*TaxRate,2))')


def monthly_equivalent(amount: str, frequency_cell: str) -> str:
    """Turn any recurring cost into what it really costs each month.

    A yearly insurance premium and a monthly software subscription are
    not comparable until both are expressed per month, and anybody
    looking at an expense breakdown is asking the per-month question.
    """
    return (f'IF({frequency_cell}="","",'
            f'IFERROR(INDEX(FrequencyPerYear,MATCH({frequency_cell},'
            f'FrequencyList,0))*{amount}/12,0))')


def mileage_cost(miles: str, allowance: str = "") -> str:
    """Mileage charged at the configured rate, past the free allowance."""
    free = allowance or "FreeTravelMiles"
    return f'ROUND(MAX(0,{miles}-{free})*MileageRate,2)'
````

---

## 4. Settings and the two front pages

Settings & Lists is the only tab a buyer has to fill in before anything else works: the studio, the prices, the turnaround promises, the targets and the 48 drop-down lists. The Dashboard is the centrepiece - money, this month, the next fortnight, and a matrix that counts every alert on eleven registers. Getting Started explains the Booking ID and is honest about what the workbook is not.

### `wedding_photography_spreadsheet/sheets/settings.py`

*655 lines · 2 functions*

> `Settings & Lists` - every number and every list the workbook runs on.
>
> Nothing in this file is hard coded to one country, one currency or one
> way of pricing a wedding.  Your deposit percentage, when the balance is
> due, what an extra hour costs, your mileage rate, how long a gallery
> stays up, your target margin and all fifty-odd drop-downs are set here,
> once, and every other tab follows.
>
> The money rules matter most.  A photographer in Portland charging 30%
> to hold a date and a photographer in Perth charging 50% are both right,
> and a spreadsheet that bakes in either one is wrong for the other.

| Defines | What it does |
|---------|--------------|
| `_month_keys(today, back, ahead)` | `yyyy-mm` keys around today, oldest first |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Settings & Lists'), `LAST_COL` (11), `DOCUMENT_TYPES` (14 items), `DOCUMENT_STATUSES` (8 items), `EQUIPMENT_STATUSES` (7 items), `GALLERY_PLATFORMS` (12 items), `MARKETING_STATUSES` (7 items), `REQUEST_SOURCES` (7 items), `REVIEW_ASK_METHODS` (7 items), `SHOT_GROUP_LABELS` (6 entries), `SHOT_WHEN` (10 items), `STAFF_STATUSES` (5 items), `STAFF_TYPES` (5 items), `TRAVEL_PURPOSES` (10 items) and 1 more.

````python
"""`Settings & Lists` - every number and every list the workbook runs on.

Nothing in this file is hard coded to one country, one currency or one
way of pricing a wedding.  Your deposit percentage, when the balance is
due, what an extra hour costs, your mileage rate, how long a gallery
stays up, your target margin and all fifty-odd drop-downs are set here,
once, and every other tab follows.

The money rules matter most.  A photographer in Portland charging 30%
to hold a date and a photographer in Perth charging 50% are both right,
and a spreadsheet that bakes in either one is wrong for the other.
"""

from __future__ import annotations

import datetime as _dt
from typing import List, Tuple

from ..config import (BOOKING_STATUSES, CLIENT_STATUSES, CONTENT_TYPES,
                      COUNTRIES, DELIVERY_PREFERENCES, DISCLAIMER,
                      EQUIPMENT_CATEGORIES, EQUIPMENT_CONDITIONS,
                      EXPENSE_CATEGORIES, FREQUENCY_PER_YEAR,
                      GALLERY_STATUSES, HONESTY_NOTE, ID_PREFIX_FIELDS,
                      LEAD_SOURCES, LEAD_STATUSES, MARKETING_PLATFORMS,
                      MONTH_NAMES, PACKAGE_TIERS, PAYMENT_METHODS,
                      PAYMENT_STAGES, PHOTO_STYLES, PRIORITY_LEVELS,
                      PRIVACY_NOTE, RECURRING_FREQUENCIES,
                      REVIEW_PLATFORMS, REVIEW_STATUSES, SHOT_LIST,
                      STAFF_ROLES, TAX_CATEGORIES, TIMELINE_EVENTS,
                      TRAVEL_MODES, VENUE_TYPES, WEDDING_STYLES, YES_NO,
                      Config, Palette)
from ..layout import Registry, col_letter, dynamic_name, quote_sheet
from ..writer import SheetWriter, number_check

NAME = "Settings & Lists"
LAST_COL = 11

#: lists the config does not already carry, defined where they are used
DOCUMENT_TYPES = [
    "Wedding contract", "Booking form", "Model release",
    "Venue permit", "Venue public liability", "Insurance certificate",
    "Questionnaire", "Timeline sign-off", "Shot list sign-off",
    "Invoice", "Receipt", "Album sign-off", "Second shooter agreement",
    "Other",
]
DOCUMENT_STATUSES = [
    "Not needed", "Outstanding", "Sent", "Viewed", "Signed",
    "Counter-signed", "Filed", "Cancelled",
]
EQUIPMENT_STATUSES = [
    "In use", "Spare", "Being repaired", "On loan", "Sold",
    "Lost or stolen", "Retired",
]
GALLERY_PLATFORMS = [
    "Pic-Time", "Pixieset", "ShootProof", "CloudSpot", "SmugMug",
    "Zenfolio", "Dropbox", "Google Drive", "WeTransfer",
    "Own website", "USB only", "Other",
]
MARKETING_STATUSES = [
    "Idea", "Drafted", "Scheduled", "Published", "Evergreen",
    "Archived", "Cancelled",
]
REQUEST_SOURCES = [
    "Couple asked for it", "Parent asked for it",
    "Wedding planner", "Venue", "Your standard list",
    "Your own idea", "Other",
]
REVIEW_ASK_METHODS = [
    "Email", "Text message", "In the gallery", "In person",
    "Thank-you card", "Automatic reminder", "Other",
]
SHOT_GROUP_LABELS = {
    "details": "Details", "prep": "Preparation", "couple": "Couple",
    "family": "Family", "ceremony": "Ceremony",
    "reception": "Reception",
}
SHOT_WHEN = [
    "Before the ceremony", "Getting ready", "Ceremony",
    "Straight after the ceremony", "Family groups", "Couple portraits",
    "Golden hour", "Reception", "Evening", "Any time",
]
STAFF_STATUSES = ["Active", "Occasional", "On leave", "Left", "Prospect"]
STAFF_TYPES = [
    "Owner", "Employee", "Self-employed contractor", "Agency",
    "Unpaid helper",
]
TRAVEL_PURPOSES = [
    "Wedding day", "Engagement shoot", "Venue visit", "Consultation",
    "Rehearsal", "Album delivery", "Equipment collection",
    "Training or workshop", "Networking", "Other",
]

#: name -> values.  Everything the workbook offers you in a drop-down.
LISTS: List[Tuple[str, List[str]]] = [
    ("YesNoList", list(YES_NO)),
    ("BookingStatusList", list(BOOKING_STATUSES)),
    ("ClientStatusList", list(CLIENT_STATUSES)),
    ("LeadStatusList", list(LEAD_STATUSES)),
    ("LeadSourceList", list(LEAD_SOURCES)),
    ("PriorityList", list(PRIORITY_LEVELS)),
    ("TierList", list(PACKAGE_TIERS)),
    ("VenueTypeList", list(VENUE_TYPES)),
    ("WeddingStyleList", list(WEDDING_STYLES)),
    ("PhotoStyleList", list(PHOTO_STYLES)),
    ("DeliveryPreferenceList", list(DELIVERY_PREFERENCES)),
    ("DocumentTypeList", list(DOCUMENT_TYPES)),
    ("DocumentStatusList", list(DOCUMENT_STATUSES)),
    ("PaymentStageList", list(PAYMENT_STAGES)),
    ("PaymentMethodList", list(PAYMENT_METHODS)),
    ("ShotGroupList", [SHOT_GROUP_LABELS[g[0]] for g in SHOT_LIST]),
    ("ShotWhenList", list(SHOT_WHEN)),
    ("RequestSourceList", list(REQUEST_SOURCES)),
    ("StaffRoleList", list(STAFF_ROLES)),
    ("StaffStatusList", list(STAFF_STATUSES)),
    ("StaffTypeList", list(STAFF_TYPES)),
    ("GalleryStatusList", list(GALLERY_STATUSES)),
    ("GalleryPlatformList", list(GALLERY_PLATFORMS)),
    ("EquipmentCategoryList", list(EQUIPMENT_CATEGORIES)),
    ("EquipmentStatusList", list(EQUIPMENT_STATUSES)),
    ("EquipmentConditionList", list(EQUIPMENT_CONDITIONS)),
    ("TravelModeList", list(TRAVEL_MODES)),
    ("TravelPurposeList", list(TRAVEL_PURPOSES)),
    ("MarketingPlatformList", list(MARKETING_PLATFORMS)),
    ("ContentTypeList", list(CONTENT_TYPES)),
    ("ContentStatusList", list(MARKETING_STATUSES)),
    ("ExpenseCategoryList", list(EXPENSE_CATEGORIES)),
    ("TaxCategoryList", list(TAX_CATEGORIES)),
    ("FrequencyList", list(RECURRING_FREQUENCIES)),
    ("ReviewStatusList", list(REVIEW_STATUSES)),
    ("ReviewPlatformList", list(REVIEW_PLATFORMS)),
    ("ReviewAskMethodList", list(REVIEW_ASK_METHODS)),
    ("CountryList", list(COUNTRIES)),
    ("MonthNameList", list(MONTH_NAMES)),
]


def _month_keys(today: _dt.date, back: int = 15, ahead: int = 21
                ) -> List[str]:
    """`yyyy-mm` keys around today, oldest first.

    Weddings are booked a long way ahead, so this window reaches
    further forward than it does back.
    """
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
             "Your studio, your prices, your rules, your lists. Set "
             "these once and every tab in the workbook follows.",
             LAST_COL)
    s.set_col(0, 0, 3)
    s.set_col(1, 2, 26)
    s.set_col(3, 3, 20)
    s.set_col(4, LAST_COL, 17)

    row = 4
    row = s.note(
        row,
        "White cells are yours to change; everything else is locked "
        "and calculated. Nothing here assumes a country, a currency "
        "or a way of pricing a wedding - set your own deposit rule, "
        "your own payment dates, your own mileage rate and your own "
        "targets, and the rest of the workbook follows.",
        LAST_COL, height=34)
    row += 1

    # --------------------------------------------------------- studio --
    row = s.section(row, "Your studio", LAST_COL)
    row = s.field(row, "Studio name", cfg.studio.name, name="StudioName",
                  value_end=5, help_col=7,
                  help_text="Printed at the top of every report and "
                            "every printable sheet you take to a "
                            "wedding.")
    row = s.field(row, "Photographer", cfg.studio.owner,
                  name="OwnerName", value_end=5, help_col=7)
    row = s.field(row, "Email", cfg.studio.email, value_end=5,
                  help_col=7)
    row = s.field(row, "Phone", cfg.studio.phone, value_end=5,
                  help_col=7)
    row = s.field(row, "Website", cfg.studio.website, value_end=5,
                  help_col=7)
    row = s.field(row, "Instagram", cfg.studio.instagram, value_end=5,
                  help_col=7)
    row = s.field(row, "Based in", f"{cfg.studio.city}, "
                  f"{cfg.studio.region}", value_end=5, help_col=7)
    row = s.field(row, "Country", cfg.studio.country, value_end=5,
                  help_col=7)
    row = s.field(row, "Currency symbol", cfg.currency,
                  name="CurrencySign",
                  help_text="Shown on screen. The cell formats follow "
                            "the symbol you build the file with.")
    row = s.field(row, "Currency code", cfg.currency_code)
    row = s.field(row, "Today", None, style="date", name="TodayDate",
                  mode="calc", formula="=TODAY()",
                  help_text="Live. Everything that says 'overdue', "
                            "'coming up' or 'this month' reads this "
                            "one cell.")
    row = s.field(row, "This month", None, style="text",
                  name="ThisMonthKey", mode="calc",
                  formula='=TEXT(TODAY(),"yyyy-mm")',
                  help_text="The month key every monthly total groups "
                            "on.")
    row = s.field(row, "This year", None, style="int",
                  name="ThisYear", mode="calc",
                  formula="=YEAR(TODAY())")
    row += 1

    # ------------------------------------------------- the money rules --
    row = s.section(row, "How you take the money", LAST_COL)
    row = s.note(
        row,
        "These four rules decide what every booking shows as a "
        "deposit, when the balance is due and when a payment counts "
        "as late. There is no standard - 30% on booking with the "
        "balance a fortnight out is common, 50/50 is common, and "
        "your contract is what actually matters. Put yours in.",
        LAST_COL, height=42)
    row += 1
    row = s.field(row, "Deposit to hold a date", cfg.deposit_percent,
                  style="percent", name="DepositPercent",
                  validate=number_check(0, 1),
                  help_text="A share of the total booking value. "
                            "Non-refundable deposits are normal in "
                            "wedding photography and should be "
                            "spelled out in your contract.")
    row = s.field(row, "Deposit due within", cfg.deposit_days_after_booking,
                  style="int", name="DepositDays",
                  validate=number_check(0, 365),
                  help_text="Days after you confirm the booking. "
                            "Until it arrives the date is not really "
                            "held.")
    row = s.field(row, "Second instalment", cfg.second_percent,
                  style="percent", name="SecondPercent",
                  validate=number_check(0, 1))
    row = s.field(row, "Second instalment due",
                  cfg.second_payment_days_before, style="int",
                  name="SecondPaymentDays", validate=number_check(0, 999),
                  help_text="Days before the wedding.")
    row = s.field(row, "Balance due", cfg.final_payment_days_before,
                  style="int", name="FinalPaymentDays",
                  validate=number_check(0, 365),
                  help_text="Days before the wedding. Collect the "
                            "balance before the day, never after - "
                            "chasing money from a couple on honeymoon "
                            "is a miserable job.")
    row = s.field(row, "Invoice payment terms", cfg.payment_terms_days,
                  style="int", name="PaymentTerms",
                  validate=number_check(0, 365))
    row = s.field(row, "Chase an unpaid invoice after",
                  cfg.payment_chase_days, style="int",
                  name="PaymentChaseDays", validate=number_check(1, 365))
    row += 1

    # ----------------------------------------------------- the pricing --
    row = s.section(row, "What the extras cost", LAST_COL)
    row = s.note(
        row,
        "Package prices live on Packages & Pricing, one row each. "
        "These are the charges a booking works out for itself: "
        "overtime, a second shooter, mileage past the free miles.",
        LAST_COL, height=30)
    row += 1
    row = s.field(row, "Extra hour on the day", cfg.extra_hour_rate,
                  style="money", name="ExtraHourRate",
                  validate=number_check(0, 100000),
                  help_text="Charged automatically when the coverage "
                            "times run past the package hours.")
    row = s.field(row, "Second shooter", cfg.second_shooter_rate,
                  style="money", name="SecondShooterRate",
                  validate=number_check(0, 100000),
                  help_text="What a second shooter costs you for a "
                            "wedding day.")
    row = s.field(row, "Assistant", cfg.assistant_rate, style="money",
                  name="AssistantRate", validate=number_check(0, 100000))
    row = s.field(row, "Engagement session", cfg.engagement_session_price,
                  style="money", name="EngagementPrice",
                  validate=number_check(0, 100000))
    row = s.field(row, "Engagement session costs you",
                  round(cfg.engagement_session_price * 0.35, 2),
                  style="money", name="EngagementCost",
                  validate=number_check(0, 100000),
                  help_text="Your time and travel for a shoot you have "
                            "included in a package.")
    row = s.field(row, "Album base price", cfg.album_base_price,
                  style="money", name="AlbumPrice",
                  validate=number_check(0, 100000))
    row = s.field(row, "Rush delivery fee", cfg.rush_delivery_fee,
                  style="money", name="RushFee",
                  validate=number_check(0, 100000))
    row = s.field(row, "Print credit costs you", 0.4, style="percent",
                  name="PrintCostRatio", validate=number_check(0, 1),
                  help_text="What a dollar of print credit actually "
                            "costs you at the lab. Packages use it to "
                            "cost their print allowance.")
    row = s.field(row, "Mileage rate", cfg.mileage_rate, style="money4",
                  name="MileageRate", validate=number_check(0, 100),
                  help_text="Per mile. Check your own tax authority's "
                            "approved rate each year - it changes, and "
                            "it is different in every country.")
    row = s.field(row, "Free miles included", cfg.free_travel_miles,
                  style="int", name="FreeTravelMiles",
                  validate=number_check(0, 10000),
                  help_text="Used when the package does not set its "
                            "own travel allowance.")
    row = s.field(row, "Flag a trip that loses more than", 150,
                  style="money", name="TravelLossLimit",
                  validate=number_check(0, 100000))
    row += 1

    # ---------------------------------------------------------- the tax --
    row = s.section(row, "Tax", LAST_COL)
    row = s.note(
        row,
        "Whether wedding photography is taxable where you are, at "
        "what rate, and whether you are even registered, are local "
        "questions with local answers. The rate ships at zero and the "
        "switch ships off, so the workbook does nothing until you "
        "tell it what is true for you.",
        LAST_COL, height=38)
    row += 1
    row = s.field(row, "Add tax to packages",
                  "Yes" if cfg.tax_on_packages else "No", style="flag",
                  name="TaxOnPackages",
                  validate={"validate": "list", "source": "=YesNoList"},
                  help_text="No by default. Turn it on only if you "
                            "charge sales tax or VAT on photography.")
    row = s.field(row, "Tax rate", cfg.tax_rate, style="percent",
                  name="TaxRate", validate=number_check(0, 1))
    row = s.field(row, "What it is called", cfg.tax_name,
                  name="TaxName", value_end=5, help_col=7)
    row = s.field(row, "Receipt needed above", 25, style="money",
                  name="ReceiptThreshold", validate=number_check(0, 100000),
                  help_text="Expenses over this with no receipt are "
                            "flagged. Your own rules may be stricter.")
    row += 1

    # ------------------------------------------------- the work after it --
    row = s.section(row, "Editing and delivery", LAST_COL)
    row = s.field(row, "Preview images within", cfg.preview_days,
                  style="int", name="PreviewDays",
                  validate=number_check(0, 365),
                  help_text="Days after the wedding. A handful of "
                            "images in the first week buys you weeks "
                            "of patience on the rest.")
    row = s.field(row, "Full gallery within", cfg.delivery_days,
                  style="int", name="DeliveryDays",
                  validate=number_check(1, 730),
                  help_text="Days after the wedding. Promise what you "
                            "can do in your worst month, not your "
                            "best.")
    row = s.field(row, "Rush delivery within", 14, style="int",
                  name="RushDeliveryDays", validate=number_check(1, 365))
    row = s.field(row, "Album within", cfg.album_days, style="int",
                  name="AlbumDays", validate=number_check(1, 730))
    row = s.field(row, "Gallery stays up for",
                  cfg.gallery_months_live * 30, style="int",
                  name="GalleryLiveDays", validate=number_check(30, 3650),
                  help_text="Days from delivery. Tell the couple this "
                            "date when you hand the gallery over.")
    row = s.field(row, "Editing hours a wedding",
                  cfg.editing_hours_per_wedding, style="number1",
                  name="EditingHours", validate=number_check(0, 500),
                  help_text="Your honest average. Packages cost their "
                            "editing on it and profit per hour depends "
                            "on it.")
    row = s.field(row, "Editing cost an hour", cfg.editing_cost_per_hour,
                  style="money", name="EditingRate",
                  validate=number_check(0, 10000),
                  help_text="What an hour of editing costs the "
                            "business - an outside editor's rate, or "
                            "what your own hour is worth.")
    row = s.field(row, "Admin hours a wedding", 6, style="number1",
                  name="AdminHoursPerWedding", validate=number_check(0, 200),
                  help_text="Enquiries, the consultation, the "
                            "contract, the invoices, the "
                            "questionnaire, the delivery email. Six "
                            "hours is a conservative guess.")
    row += 1

    # ----------------------------------------------------- when to nag --
    row = s.section(row, "When the workbook should nag you", LAST_COL)
    row = s.field(row, "Chase an enquiry after", cfg.lead_chase_days,
                  style="int", name="LeadChaseDays",
                  validate=number_check(1, 90),
                  help_text="Days with no reply from them. Wedding "
                            "enquiries go cold fast - most couples "
                            "contact several photographers the same "
                            "evening.")
    row = s.field(row, "Chase after a consultation",
                  cfg.consultation_chase_days, style="int",
                  name="ConsultationChaseDays", validate=number_check(1, 90))
    row = s.field(row, "Chase an unsigned document",
                  cfg.contract_warn_days, style="int",
                  name="DocumentChaseDays", validate=number_check(1, 180))
    row = s.field(row, "Chase a review after", cfg.review_chase_days,
                  style="int", name="ReviewChaseDays",
                  validate=number_check(1, 180))
    row = s.field(row, "'Coming up' means within", cfg.upcoming_days,
                  style="int", name="UpcomingDays",
                  validate=number_check(1, 365),
                  help_text="Days. Controls what the dashboard calls "
                            "an upcoming wedding.")
    row = s.field(row, "Warn about a warranty", cfg.warranty_warn_days,
                  style="int", name="WarrantyWarnDays",
                  validate=number_check(1, 365))
    row = s.field(row, "Service kit every", 12, style="int",
                  name="ServiceMonths", validate=number_check(1, 120),
                  help_text="Months between services for anything you "
                            "have put a service date against.")
    row = s.field(row, "Insure anything worth over", 500, style="money",
                  name="InsureAbove", validate=number_check(0, 1000000))
    row = s.field(row, "Equipment useful life", 5, style="number1",
                  name="EquipmentLifeYears", validate=number_check(1, 40),
                  help_text="Years. Used to spread what a camera cost "
                            "across the years you will use it.")
    row += 1

    # ------------------------------------------------------- the targets --
    row = s.section(row, "Your targets", LAST_COL)
    row = s.note(
        row,
        "The dashboard measures you against these, so make them real. "
        "A target you set because it sounded good makes every chart "
        "in the workbook useless.",
        LAST_COL, height=30)
    row += 1
    row = s.field(row, "Weddings a year", cfg.target_weddings_year,
                  style="int", name="TargetWeddings",
                  validate=number_check(1, 400))
    row = s.field(row, "Average booking value", cfg.target_booking_value,
                  style="money", name="TargetAverageValue",
                  validate=number_check(0, 1000000))
    row = s.field(row, "Revenue a year", cfg.target_revenue_year,
                  style="money", name="TargetRevenue",
                  validate=number_check(0, 100000000))
    row = s.field(row, "Profit margin", cfg.target_profit_margin,
                  style="percent", name="TargetMargin",
                  validate=number_check(0, 1),
                  help_text="Packages use this to work out what they "
                            "would have to sell for.")
    row = s.field(row, "Enquiry to booking", cfg.target_conversion,
                  style="percent", name="TargetConversion",
                  validate=number_check(0, 1))
    row = s.field(row, "Average review", cfg.target_rating,
                  style="rating", name="TargetRating",
                  validate=number_check(1, 5))
    row = s.field(row, "Lowest hourly rate you will accept", 45,
                  style="money", name="MinimumHourly",
                  validate=number_check(0, 10000),
                  help_text="Profit per hour across shooting, "
                            "editing, travel and admin. Anything under "
                            "this is flagged, on the package and on "
                            "the wedding.")
    row = s.field(row, "An hour of yours is worth", 60, style="money",
                  name="OwnerHourlyValue", validate=number_check(0, 10000),
                  help_text="Used to cost the hours you put into "
                            "marketing. Free marketing is not free.")
    row += 1

    # ------------------------------------------------ live totals ------
    row = s.section(row, "Worked out from your data", LAST_COL)
    row = s.note(
        row,
        "These three are calculated, not typed. Profit Analysis uses "
        "them to spread the year's overheads across the year's "
        "weddings.",
        LAST_COL, height=28)
    row += 1
    row = s.field(row, "Weddings this year", None, style="int",
                  name="WeddingsThisYear", mode="calc",
                  formula='=COUNTIFS([[Wedding Bookings.year]],'
                          'YEAR(TODAY()),'
                          '[[Wedding Bookings.earns]],"Yes")')
    row = s.field(row, "Overheads this year", None, style="money",
                  name="OverheadsThisYear", mode="calc",
                  formula='=SUMIFS([[Expense Tracker.amount]],'
                          '[[Expense Tracker.year]],YEAR(TODAY()),'
                          '[[Expense Tracker.is_overhead]],"Yes")')
    row = s.field(row, "Couples on the books", None, style="int",
                  name="ClientCount", mode="calc",
                  formula='=COUNTIF([[Client CRM.name]],"?*")')
    row += 1

    # ------------------------------------------- the timeline template --
    row = s.section(row, "Standard wedding day timeline", LAST_COL)
    row = s.note(
        row,
        "The Wedding Timeline tab builds a day plan from this table. "
        "Minutes are counted from the ceremony - negative is before "
        "it, positive is after. These are sensible starting points for "
        "a traditional day; change them to match how you actually "
        "work, and override anything on the wedding itself.",
        LAST_COL, height=44)
    row += 1
    head = row
    s.table_head(head, 1, ["Event", "Minutes from ceremony",
                           "How long", "Who covers it"],
                 widths=[30, 20, 14, 22])
    first = head + 1
    spare = 6
    for index, (_key, label, offset, length, role) in enumerate(
            TIMELINE_EVENTS):
        line = first + index
        s.write(line, 1, label, fmts.input("text"))
        s.write(line, 2, offset, fmts.input("int"))
        s.write(line, 3, length, fmts.input("int"))
        s.write(line, 4, role, fmts.input("text"))
    for index in range(len(TIMELINE_EVENTS),
                       len(TIMELINE_EVENTS) + spare):
        line = first + index
        for column, style in ((1, "text"), (2, "int"), (3, "int"),
                              (4, "text")):
            s.write(line, column, None, fmts.input(style))
    last = first + len(TIMELINE_EVENTS) + spare - 1
    for key, column in (("TimelineEventList", 1),
                        ("TimelineOffsetList", 2),
                        ("TimelineDurationList", 3),
                        ("TimelineRoleList", 4)):
        book.define_name(
            key, f"={sheet_ref}!${col_letter(column)}${first + 1}:"
                 f"${col_letter(column)}${last + 1}")
    row = last + 2

    # ------------------------------------------- recurring frequencies --
    row = s.section(row, "How often a recurring cost repeats", LAST_COL)
    row = s.note(
        row,
        "This is what turns any expense into a monthly equivalent, so "
        "a yearly insurance premium and a monthly subscription can sit "
        "in the same breakdown. 'One off' is zero because it is not a "
        "standing cost.",
        LAST_COL, height=30)
    row += 1
    freq_head = row
    s.table_head(freq_head, 1, ["How often", "Times a year"],
                 widths=[24, 16])
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

    # -------------------------------------------------------- prefixes --
    row = s.section(row, "ID prefixes", LAST_COL)
    row = s.note(
        row,
        "The letters in front of every automatic ID. Booking IDs are "
        "the ones that matter - they are the key the whole workbook "
        "joins on. Change the prefixes here if you want to, but do it "
        "before you have typed two hundred weddings, not after.",
        LAST_COL, height=32)
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
        s.set_col(column, column, 28)
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
        "The months the reports pick from. Plain text in yyyy-mm form "
        "so that every monthly total in the workbook groups on exactly "
        "the same thing. The window runs further forward than back, "
        "because weddings are booked a long way ahead.",
        LAST_COL, height=32)
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

    row = s.note(row, HONESTY_NOTE, LAST_COL, height=56)
    row += 1
    row = s.note(row, PRIVACY_NOTE, LAST_COL, height=48)
    row += 1
    row = s.note(row, DISCLAIMER, LAST_COL, height=96)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    s.protect()
    return s
````

### `wedding_photography_spreadsheet/sheets/dashboard.py`

*494 lines · 4 functions*

> `Dashboard` - the page that answers "how is the business doing?"
>
> Everything here is worked out from the registers, so nothing is ever
> out of date and there is nothing to refresh.  The questions are in the
> order a wedding photographer actually asks them: what is happening
> this week, what needs doing today, am I booked up, is the money in,
> how far behind is the editing, and which weddings are actually worth
> shooting.

| Defines | What it does |
|---------|--------------|
| `_month_key(n)` | The `yyyy-mm` key for n months before this one |
| `_by_rank(sheet, rank_key, field, n)` | The nth row of a register by one of its hidden rank columns |
| `_nth(sheet, trigger, field, n)` | Row n of a register, blank when that row is empty |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Dashboard'), `LAST_COL` (12), `MONTHS` (12), `ACTION_ROWS` (10), `PACKAGE_ROWS` (6), `EARNS` ('[[Wedding Bookings.earns]],"Yes"'), `HOLDS` ('[[Wedding Bookings.holds]],"Yes"'), `WED_NUM` ('[[Wedding Bookings.wed_num]]').

````python
"""`Dashboard` - the page that answers "how is the business doing?"

Everything here is worked out from the registers, so nothing is ever
out of date and there is nothing to refresh.  The questions are in the
order a wedding photographer actually asks them: what is happening
this week, what needs doing today, am I booked up, is the money in,
how far behind is the editing, and which weddings are actually worth
shooting.
"""

from __future__ import annotations

from ..config import MAX_WEDDINGS_SHOWN, Config, Palette
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Dashboard"
LAST_COL = 12
MONTHS = 12
ACTION_ROWS = 10
PACKAGE_ROWS = 6

EARNS = '[[Wedding Bookings.earns]],"Yes"'
HOLDS = '[[Wedding Bookings.holds]],"Yes"'
WED_NUM = "[[Wedding Bookings.wed_num]]"


def _month_key(n: int) -> str:
    """The `yyyy-mm` key for n months before this one."""
    if n == 0:
        return 'TEXT(TodayDate,"yyyy-mm")'
    return (f'TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{n},1),'
            f'"yyyy-mm")')


def _by_rank(sheet: str, rank_key: str, field: str, n: int) -> str:
    """The nth row of a register by one of its hidden rank columns."""
    return (f'IFERROR(INDEX([[{sheet}.{field}]],MATCH({n},'
            f'[[{sheet}.{rank_key}]],0)),"")')


def _nth(sheet: str, trigger: str, field: str, n: int) -> str:
    """Row n of a register, blank when that row is empty."""
    return (f'IFERROR(IF(INDEX([[{sheet}.{trigger}]],{n})="","",'
            f'INDEX([[{sheet}.{field}]],{n})),"")')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)

    s.banner("Wedding Photography Dashboard",
             "Your weddings, your money and your backlog - worked out "
             "from the tabs you have already filled in.",
             LAST_COL)
    s.set_col(0, 0, 24)
    for column in range(1, LAST_COL + 1):
        s.set_col(column, column, 13)

    # ----------------------------------------------------------- header --
    row = 4
    s.merge(row, 0, 0, "Studio", fmts.chrome("label_bold"))
    s.merge(row, 1, 3, None, fmts.calc("text"))
    s.formula(row, 1, "=StudioName", fmts.calc("text"))
    s.merge(row, 4, 4, "Month", fmts.chrome("label_right"))
    s.formula(row, 5, '=TEXT(TodayDate,"mmmm yyyy")', fmts.calc("text"))
    s.merge(row, 6, 6, "Couples", fmts.chrome("label_right"))
    s.formula(row, 7, "=ClientCount", fmts.calc("int"))
    s.merge(row, 8, 8, "Today", fmts.chrome("label_right"))
    s.formula(row, 9, "=TodayDate", fmts.calc("date"))
    s.merge(row, 10, LAST_COL, None, fmts.chrome("note"))
    s.formula(
        row, 10,
        '="Figures update the moment you type. Nothing to refresh."',
        fmts.chrome("note"))
    row += 2

    # ------------------------------------------------------- this week --
    row = s.section(row, "The next few weeks", LAST_COL)

    next_days = (f'COUNTIFS({WED_NUM},">="&TodayDate,{WED_NUM},'
                 f'"<="&TodayDate+UpcomingDays,{HOLDS})')
    s.kpi(row, 0, "Weddings coming up", f"={next_days}", "int",
          note="Inside your 'coming up' window from Settings")
    s.kpi(row, 2, "Next wedding in",
          '=IFERROR(INDEX([[Wedding Bookings.days_away]],'
          'MATCH(1,[[Wedding Bookings.upcoming_rank]],0)),"")',
          "int", note="Days away")
    s.kpi(row, 4, "This weekend",
          f'=COUNTIFS({WED_NUM},">="&TodayDate,{WED_NUM},'
          f'"<="&TodayDate+7,{HOLDS})', "int")
    s.kpi(row, 6, "Timelines not built",
          f'=COUNTIFS({WED_NUM},">="&TodayDate,{WED_NUM},'
          f'"<="&TodayDate+UpcomingDays,{HOLDS},'
          f'[[Wedding Bookings.timeline_lines]],0)', "int",
          note="Upcoming weddings with no running order yet")
    s.kpi(row, 8, "Questionnaires out",
          '=COUNTIFS([[Wedding Bookings.questionnaire_sent]],"?*",'
          '[[Wedding Bookings.questionnaire_back]],"",'
          f'{HOLDS})', "int")
    row = s.kpi(row, 10, "Nobody assigned",
                f'=COUNTIFS({WED_NUM},">="&TodayDate,{HOLDS},'
                f'[[Wedding Bookings.photographer]],"")', "int",
                note="Upcoming weddings with no lead photographer")

    s.kpi(row, 0, "Date clashes",
          '=COUNTIFS([[Wedding Bookings.clash]],"Yes")', "int",
          note="Two live weddings on one day")
    s.kpi(row, 2, "Photographer clashes",
          '=COUNTIFS([[Wedding Bookings.team_clash]],"Yes")', "int")
    s.kpi(row, 4, "Contracts unsigned",
          f'=COUNTIFS({HOLDS},[[Wedding Bookings.contract_done]],'
          f'"<>Yes",{WED_NUM},">="&TodayDate)', "int")
    s.kpi(row, 6, "Shot lists empty",
          f'=COUNTIFS({WED_NUM},">="&TodayDate,{WED_NUM},'
          f'"<="&TodayDate+UpcomingDays,{HOLDS},'
          f'[[Wedding Bookings.shot_lines]],0)', "int")
    s.kpi(row, 8, "Must-have shots missed",
          '=COUNTIFS([[Shot List Manager.must_have]],"Yes",'
          '[[Shot List Manager.status]],"Missed")', "int",
          note="Flagged as essential and never taken")
    row = s.kpi(row, 10, "Kit needing attention",
                '=COUNTIFS([[Equipment Inventory.alert]],"Due -*")+'
                'COUNTIFS([[Equipment Inventory.alert]],"Order -*")+'
                'COUNTIFS([[Equipment Inventory.alert]],"Check -*")',
                "int")
    row += 1

    # ------------------------------------------------------- the money --
    row = s.section(row, "The money", LAST_COL)
    s.kpi(row, 0, "Booked this year",
          '=SUMIFS([[Wedding Bookings.total_value]],'
          f'[[Wedding Bookings.year]],YEAR(TodayDate),{EARNS})',
          "money0", note="Value of every confirmed wedding dated this "
                         "year")
    s.kpi(row, 2, "Against target",
          '=IF(TargetRevenue<=0,"",'
          'SUMIFS([[Wedding Bookings.total_value]],'
          f'[[Wedding Bookings.year]],YEAR(TodayDate),{EARNS})/'
          'TargetRevenue)', "percent0")
    s.kpi(row, 4, "Collected",
          f'=SUMIFS([[Wedding Bookings.collected]],{EARNS})', "money0",
          note="Across every confirmed wedding, past and future")
    s.kpi(row, 6, "Still owed",
          f'=SUMIFS([[Wedding Bookings.outstanding]],{EARNS})',
          "money0")
    s.kpi(row, 8, "Overdue now",
          '=SUMIFS([[Payments & Invoices.balance]],'
          '[[Payments & Invoices.is_late]],"Yes")', "money0",
          note="Invoices past their due date")
    row = s.kpi(row, 10, "Invoices late",
                '=COUNTIFS([[Payments & Invoices.is_late]],"Yes")',
                "int")

    s.kpi(row, 0, "Average booking",
          f'=IF(COUNTIFS({EARNS})<=0,"",'
          f'SUMIFS([[Wedding Bookings.total_value]],{EARNS})/'
          f'COUNTIFS({EARNS}))', "money0")
    s.kpi(row, 2, "Average profit",
          f'=IF(COUNTIFS({EARNS})<=0,"",'
          f'SUMIFS([[Wedding Bookings.profit]],{EARNS})/'
          f'COUNTIFS({EARNS}))', "money0")
    s.kpi(row, 4, "Margin",
          f'=IF(SUMIFS([[Wedding Bookings.total_value]],{EARNS})<=0,"",'
          f'SUMIFS([[Wedding Bookings.profit]],{EARNS})/'
          f'SUMIFS([[Wedding Bookings.total_value]],{EARNS}))',
          "percent0", note="Target is on Settings")
    s.kpi(row, 6, "Expenses this year",
          '=SUMIFS([[Expense Tracker.amount]],'
          '[[Expense Tracker.year]],YEAR(TodayDate))', "money0")
    s.kpi(row, 8, "Overheads this year", "=OverheadsThisYear",
          "money0", note="Costs not tied to one wedding")
    row = s.kpi(row, 10, "Earned after delivery",
                '=SUM([[Gallery & Delivery.after_total]])', "money0",
                note="Prints and albums sold once the gallery was up")
    row += 1

    # --------------------------------------------------- what to do now --
    row = s.section(row, "What needs doing", LAST_COL)
    row = s.note(
        row,
        "Every register writes one plain-English line per row saying "
        "what it needs. These are the counts. Sort any tab by its "
        "'What to do' column to see the rows themselves.",
        LAST_COL, height=30)
    row += 1

    head = row
    s.table_head(head, 0, ["Where", "Urgent", "Overdue", "Chase or send",
                           "Add or check", "All clear", "Needs you"],
                 widths=[24, 13, 13, 14, 14, 13, 13])
    row = head + 1
    tabs = [
        ("Lead Tracker", "Lead Tracker"),
        ("Wedding Bookings", "Wedding Bookings"),
        ("Contracts & Documents", "Contracts & Documents"),
        ("Payments & Invoices", "Payments & Invoices"),
        ("Wedding Timeline", "Wedding Timeline"),
        ("Shot List Manager", "Shot List Manager"),
        ("Editing Workflow", "Editing Workflow"),
        ("Gallery & Delivery", "Gallery & Delivery"),
        ("Equipment Inventory", "Equipment Inventory"),
        ("Expense Tracker", "Expense Tracker"),
        ("Reviews & Referrals", "Reviews & Referrals"),
    ]
    for label, sheet in tabs:
        col = f"[[{sheet}.alert]]"
        s.write(row, 0, label, fmts.chrome("label_bold"))
        s.formula(row, 1, f'=COUNTIF({col},"Urgent -*")',
                  fmts.cell("int", "calc"))
        s.formula(row, 2, f'=COUNTIF({col},"Overdue -*")',
                  fmts.cell("int", "calc"))
        s.formula(row, 3,
                  f'=COUNTIF({col},"Chase -*")+COUNTIF({col},"Send -*")'
                  f'+COUNTIF({col},"Follow up -*")+'
                  f'COUNTIF({col},"Due -*")',
                  fmts.cell("int", "calc"))
        s.formula(row, 4,
                  f'=COUNTIF({col},"Add -*")+COUNTIF({col},"Check -*")+'
                  f'COUNTIF({col},"Action -*")+'
                  f'COUNTIF({col},"Missing -*")+'
                  f'COUNTIF({col},"Order -*")+'
                  f'COUNTIF({col},"Duplicate -*")+'
                  f'COUNTIF({col},"Loss -*")+'
                  f'COUNTIF({col},"Expired -*")',
                  fmts.cell("int", "calc"))
        s.formula(row, 5, f'=COUNTIF({col},"OK*")',
                  fmts.cell("int", "calc"))
        s.formula(row, 6,
                  f'=COUNTIF({col},"?*")-COUNTIF({col},"OK*")',
                  fmts.cell("int", "calc"))
        row += 1
    s.write(row, 0, "Everything", fmts.chrome("total_label"))
    first_tab, last_tab = head + 1, row - 1
    for column in range(1, 7):
        s.formula(row, column,
                  f"=SUM({chr(65 + column)}{first_tab + 1}:"
                  f"{chr(65 + column)}{last_tab + 1})",
                  fmts.chrome("total_value",
                              num_format=s.num_format("int")))
    row += 2

    # ------------------------------------------------- the next weddings --
    row = s.section(row, "The next weddings", LAST_COL)
    head = row
    s.table_head(head, 0, ["Couple", "Date", "Days", "Venue",
                           "Photographer", "Value", "Paid",
                           "Contract", "What to do"],
                 widths=[22, 12, 8, 20, 16, 12, 9, 13, 30])
    row = head + 1
    for n in range(1, MAX_WEDDINGS_SHOWN + 1):
        rank = (f'MATCH({n},[[Wedding Bookings.upcoming_rank]],0)')
        def pick(field, fmt_style, r=rank):
            return (f'=IFERROR(INDEX([[Wedding Bookings.{field}]],{r}),'
                    f'"")')
        s.formula(row, 0, pick("client_name", "text"),
                  fmts.cell("text", "calc"))
        s.formula(row, 1, pick("wedding_date", "date"),
                  fmts.cell("date", "calc"))
        s.formula(row, 2, pick("days_away", "int"),
                  fmts.cell("int", "calc"))
        s.formula(row, 3, pick("venue", "text"),
                  fmts.cell("text", "calc"))
        s.formula(row, 4, pick("photographer_name", "text"),
                  fmts.cell("text", "calc"))
        s.formula(row, 5, pick("total_value", "money"),
                  fmts.cell("money", "calc"))
        s.formula(row, 6, pick("paid_pct", "percent0"),
                  fmts.cell("percent0", "calc"))
        s.formula(row, 7, pick("contract_state", "status"),
                  fmts.cell("text", "calc"))
        s.formula(row, 8, pick("alert", "status"),
                  fmts.cell("text", "calc"))
        row += 1
    row += 1

    # ----------------------------------------------- post-production ----
    row = s.section(row, "Post-production", LAST_COL)
    s.kpi(row, 0, "Weddings in editing",
          '=COUNTIFS([[Editing Workflow.stage]],"<>Complete",'
          '[[Editing Workflow.booking]],"?*")', "int",
          note="Shot, not yet delivered")
    s.kpi(row, 2, "Galleries overdue",
          '=COUNTIFS([[Editing Workflow.late]],"Yes",'
          '[[Editing Workflow.stage]],"<>Complete")', "int")
    s.kpi(row, 4, "Not backed up",
          '=COUNTIFS([[Editing Workflow.backup]],"<>Yes",'
          '[[Editing Workflow.booking]],"?*")', "int",
          note="The only number on this page that cannot be fixed "
               "later")
    s.kpi(row, 6, "Average turnaround",
          '=IFERROR(AVERAGEIFS([[Editing Workflow.days_taken]],'
          '[[Editing Workflow.days_taken]],">0"),"")', "int",
          note="Days from wedding to finished gallery")
    s.kpi(row, 8, "You promise",
          "=DeliveryDays", "int", note="Days, from Settings")
    row = s.kpi(row, 10, "Delivered, never opened",
                '=COUNTIFS([[Gallery & Delivery.opened]],"No")', "int")

    s.kpi(row, 0, "Albums to send",
          '=COUNTIFS([[Gallery & Delivery.album_ordered]],"Yes",'
          '[[Gallery & Delivery.album_sent]],"")', "int")
    s.kpi(row, 2, "Galleries closing soon",
          '=COUNTIFS([[Gallery & Delivery.expires]],">="&TodayDate,'
          '[[Gallery & Delivery.expires]],"<="&TodayDate+30,'
          '[[Gallery & Delivery.expiry_warned]],"<>Yes")', "int",
          note="Expiring within a month with no warning sent")
    s.kpi(row, 4, "Reviews received",
          '=COUNTIFS([[Reviews & Referrals.has_review]],"Yes")', "int")
    s.kpi(row, 6, "Average review",
          '=IFERROR(AVERAGE([[Reviews & Referrals.rating]]),"")',
          "rating")
    s.kpi(row, 8, "Weddings from referrals",
          '=COUNTIFS([[Lead Tracker.referred_by]],"?*",'
          '[[Lead Tracker.is_won]],"Yes")', "int")
    row = s.kpi(row, 10, "Referral revenue",
                '=SUMIFS([[Lead Tracker.booked_value]],'
                '[[Lead Tracker.referred_by]],"?*",'
                '[[Lead Tracker.is_won]],"Yes")', "money0")
    row += 1

    # ------------------------------------------------------- enquiries --
    row = s.section(row, "Enquiries", LAST_COL)
    s.kpi(row, 0, "Open enquiries",
          '=COUNTIFS([[Lead Tracker.is_open]],"Yes")', "int")
    s.kpi(row, 2, "Needing a reply",
          '=COUNTIF([[Lead Tracker.alert]],"Urgent -*")+'
          'COUNTIF([[Lead Tracker.alert]],"Action -*")', "int")
    s.kpi(row, 4, "Overdue a chase",
          '=COUNTIFS([[Lead Tracker.overdue]],">0",'
          '[[Lead Tracker.is_open]],"Yes")', "int")
    s.kpi(row, 6, "Enquiries this year",
          '=COUNTIFS([[Lead Tracker.enq_num]],">="&'
          'DATE(YEAR(TodayDate),1,1))', "int")
    s.kpi(row, 8, "Booked from them",
          '=COUNTIFS([[Lead Tracker.enq_num]],">="&'
          'DATE(YEAR(TodayDate),1,1),'
          '[[Lead Tracker.is_won]],"Yes")', "int")
    row = s.kpi(row, 10, "Conversion",
                '=IF(COUNTIFS([[Lead Tracker.enq_num]],">="&'
                'DATE(YEAR(TodayDate),1,1))<=0,"",'
                'COUNTIFS([[Lead Tracker.enq_num]],">="&'
                'DATE(YEAR(TodayDate),1,1),'
                '[[Lead Tracker.is_won]],"Yes")/'
                'COUNTIFS([[Lead Tracker.enq_num]],">="&'
                'DATE(YEAR(TodayDate),1,1)))', "percent0",
                note="Target is on Settings")
    row += 1

    # --------------------------------------------------- the packages ---
    row = s.section(row, "Which packages are working", LAST_COL)
    head = row
    s.table_head(head, 0, ["Package", "Price", "Cost", "Margin",
                           "Your hourly rate", "Asked about", "Booked",
                           "Conversion", "Revenue", "What to do"],
                 widths=[20, 12, 12, 10, 15, 12, 10, 11, 13, 28])
    row = head + 1
    for n in range(1, PACKAGE_ROWS + 1):
        for col, (field, style) in enumerate([
                ("name", "text"), ("total_price", "money"),
                ("total_cost", "money"), ("margin", "percent0"),
                ("hourly", "money"), ("enquiries", "int"),
                ("sold", "int"), ("conversion", "percent0"),
                ("revenue", "money"), ("alert", "text")]):
            s.formula(row, col,
                      "=" + _nth("Packages & Pricing", "name", field, n),
                      fmts.cell(style, "calc"))
        row += 1
    row += 1

    # ------------------------------------------------ twelve months -----
    row = s.section(row, "The last twelve months", LAST_COL)
    head = row
    s.table_head(head, 0, ["Month", "Weddings", "Revenue booked",
                           "Money in", "Expenses", "Profit",
                           "Enquiries", "Booked", "Conversion",
                           "Against target"],
                 widths=[14, 11, 15, 13, 12, 12, 11, 10, 11, 14])
    row = head + 1
    first_month = row
    for n in range(MONTHS - 1, -1, -1):
        key = _month_key(n)
        s.formula(row, 0, f'={key}', fmts.cell("text", "calc"))
        s.formula(row, 1,
                  f'=COUNTIFS([[Wedding Bookings.month]],{key},{EARNS})',
                  fmts.cell("int", "calc"))
        s.formula(row, 2,
                  f'=SUMIFS([[Wedding Bookings.total_value]],'
                  f'[[Wedding Bookings.month]],{key},{EARNS})',
                  fmts.cell("money", "calc"))
        s.formula(row, 3,
                  f'=SUMIFS([[Payments & Invoices.received]],'
                  f'[[Payments & Invoices.month]],{key})',
                  fmts.cell("money", "calc"))
        s.formula(row, 4,
                  f'=SUMIFS([[Expense Tracker.amount]],'
                  f'[[Expense Tracker.month]],{key})',
                  fmts.cell("money", "calc"))
        s.formula(row, 5,
                  f'=SUMIFS([[Wedding Bookings.profit]],'
                  f'[[Wedding Bookings.month]],{key},{EARNS})-'
                  f'SUMIFS([[Expense Tracker.amount]],'
                  f'[[Expense Tracker.month]],{key},'
                  f'[[Expense Tracker.is_overhead]],"Yes")',
                  fmts.cell("money", "calc"))
        s.formula(row, 6,
                  f'=COUNTIFS([[Lead Tracker.month]],{key})',
                  fmts.cell("int", "calc"))
        s.formula(row, 7,
                  f'=COUNTIFS([[Lead Tracker.month]],{key},'
                  f'[[Lead Tracker.is_won]],"Yes")',
                  fmts.cell("int", "calc"))
        s.formula(row, 8,
                  f'=IF(COUNTIFS([[Lead Tracker.month]],{key})<=0,"",'
                  f'COUNTIFS([[Lead Tracker.month]],{key},'
                  f'[[Lead Tracker.is_won]],"Yes")/'
                  f'COUNTIFS([[Lead Tracker.month]],{key}))',
                  fmts.cell("percent0", "calc"))
        s.formula(row, 9,
                  f'=IF(TargetRevenue<=0,"",'
                  f'SUMIFS([[Wedding Bookings.total_value]],'
                  f'[[Wedding Bookings.month]],{key},{EARNS})/'
                  f'(TargetRevenue/12))',
                  fmts.cell("percent0", "calc"))
        row += 1
    last_month = row - 1
    s.write(row, 0, "Twelve months", fmts.chrome("total_label"))
    for column, style in ((1, "int"), (2, "money"), (3, "money"),
                          (4, "money"), (5, "money"), (6, "int"),
                          (7, "int")):
        letter = chr(65 + column)
        s.formula(row, column,
                  f"=SUM({letter}{first_month + 1}:{letter}"
                  f"{last_month + 1})",
                  fmts.chrome("total_value",
                              num_format=s.num_format(style)))
    s.formula(row, 8,
              f"=IF(G{row + 1}<=0,\"\",H{row + 1}/G{row + 1})",
              fmts.chrome("total_value",
                          num_format=s.num_format("percent0")))
    s.formula(row, 9,
              f'=IF(TargetRevenue<=0,"",C{row + 1}/TargetRevenue)',
              fmts.chrome("total_value",
                          num_format=s.num_format("percent0")))
    row += 2

    # ------------------------------------------------- best and worst ---
    row = s.section(row, "Best and worst", LAST_COL)
    head = row
    s.table_head(head, 0, ["", "Couple", "Date", "Package", "Value",
                           "Profit", "Margin", "Profit an hour"],
                 widths=[18, 20, 12, 16, 12, 12, 10, 14])
    row = head + 1
    for n in range(1, 4):
        s.write(row, 0, f"Most profitable #{n}",
                fmts.chrome("label_bold"))
        for col, (field, style) in enumerate([
                ("couple", "text"), ("wedding_date", "date"),
                ("package", "text"), ("revenue", "money"),
                ("net_profit", "money"), ("net_margin", "percent0"),
                ("per_hour", "money")], start=1):
            s.formula(row, col,
                      "=" + _by_rank("Profit Analysis",
                                     "net_profit_rank", field, n),
                      fmts.cell(style, "calc"))
        row += 1
    for n in range(1, 4):
        s.write(row, 0, f"Least profitable #{n}",
                fmts.chrome("label_bold"))
        for col, (field, style) in enumerate([
                ("couple", "text"), ("wedding_date", "date"),
                ("package", "text"), ("revenue", "money"),
                ("net_profit", "money"), ("net_margin", "percent0"),
                ("per_hour", "money")], start=1):
            s.formula(
                row, col,
                f'=IFERROR(INDEX([[Profit Analysis.{field}]],'
                f'MATCH(COUNTIF([[Profit Analysis.booking]],"?*")'
                f'-{n - 1},[[Profit Analysis.net_profit_rank]],0)),"")',
                fmts.cell(style, "calc"))
        row += 1
    row += 1

    row = s.note(
        row,
        "This workbook is a business organisation tool. It does not "
        "connect to any booking platform, gallery host or payment "
        "processor, it does not send email, and it does not generate "
        "or sign contracts. Every figure on this page comes from "
        "something you typed on one of the other tabs.",
        LAST_COL, height=44)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    s.protect()
    return s
````

### `wedding_photography_spreadsheet/sheets/getting_started.py`

*286 lines · 1 function*

> `Getting Started` - what to do first, and what this file is not.
>
> Twenty-three tabs is a lot to open cold.  This page says which five
> matter on day one, walks the five things the workbook is actually for,
> explains the one idea the whole file turns on - the Booking ID - and
> is honest about what a spreadsheet cannot do.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Getting Started'), `LAST_COL` (10).

````python
"""`Getting Started` - what to do first, and what this file is not.

Twenty-three tabs is a lot to open cold.  This page says which five
matter on day one, walks the five things the workbook is actually for,
explains the one idea the whole file turns on - the Booking ID - and
is honest about what a spreadsheet cannot do.
"""

from __future__ import annotations

from ..config import (CONTRACT_NOTE, DISCLAIMER, HONESTY_NOTE,
                      LISTING_TITLE, PRIVACY_NOTE, PRODUCT_NAME,
                      VERSION, Config, Palette)
from ..layout import Registry
from ..writer import SheetWriter

NAME = "Getting Started"
LAST_COL = 10


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.ink)

    s.banner(PRODUCT_NAME,
             "From the first enquiry to the final gallery - read this "
             "page once and the other twenty-two make sense.",
             LAST_COL)
    s.set_col(0, 0, 3)
    s.set_col(1, 1, 30)
    s.set_col(2, LAST_COL, 15)

    row = 4
    row = s.note(
        row,
        "Everything with a white background is yours to type in. "
        "Everything else is a formula, locked so you cannot "
        "accidentally overwrite it. If you want to change a locked "
        "cell, unprotect the sheet - the password is printed at the "
        "bottom of this page.",
        LAST_COL, height=34)
    row += 1

    # ---- the one idea ------------------------------------------------------
    row = s.section(row, "The one thing to understand", LAST_COL)
    row = s.note(
        row,
        "Every wedding gets a Booking ID - WED-0001, WED-0002 and so "
        "on - created for you on the Wedding Bookings tab. That ID is "
        "how the workbook knows that this contract, those three "
        "invoices, that timeline, that shot list, that editing job, "
        "that gallery and those two trips in the car all belong to "
        "the same wedding. Put the Booking ID on a row and it joins "
        "up. Leave it off and the row is an orphan: it still works, "
        "but nothing adds it up.",
        LAST_COL, height=58)
    row += 1
    row = s.bullets(row, [
        "Create the wedding first, on Wedding Bookings. Everything "
        "else refers back to it.",
        "Never type a couple's name where a Booking ID is asked for. "
        "Names are not unique, get spelled two ways, and change.",
        "The Booking ID drop-down on every other tab is live - a "
        "wedding you add this morning is in it this afternoon.",
        "The same idea runs for couples (Client CRM), packages "
        "(Packages & Pricing) and people (Photographer & Staff).",
    ], LAST_COL)
    row += 1

    # ---- the first hour ----------------------------------------------------
    row = s.section(row, "Your first hour", LAST_COL)
    steps = [
        ("1. Settings & Lists",
         "Put your studio name in, then your deposit rule, your "
         "payment dates, your mileage rate and your turnaround "
         "promise. Ten minutes here and every other tab is right."),
        ("2. Packages & Pricing",
         "Type the packages you actually sell, what each includes, "
         "and honestly what each costs you to deliver. The sheet "
         "works out your margin, your real hourly rate and what you "
         "would have to charge to hit your target."),
        ("3. Photographer & Staff",
         "Add yourself first, with 'Owner' set to Yes. Then anybody "
         "who seconds or edits for you, with their rate."),
        ("4. Client CRM",
         "Add the couples you already have booked. Just the names, "
         "the date and how they found you."),
        ("5. Wedding Bookings",
         "One row per wedding. Pick the couple, the date, the times "
         "and the package - the price, the deposit, the balance date "
         "and the delivery date all follow."),
        ("6. Payments & Invoices",
         "Add the instalments you have already taken so the money "
         "columns are true from the start."),
    ]
    for label, text in steps:
        s.write(row, 1, label, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, text, fmts.chrome("note"))
        s.row_height(row, 32)
        row += 1
    row += 1

    # ---- the five jobs -----------------------------------------------------
    row = s.section(row, "The five things this is for", LAST_COL)
    jobs = [
        ("An enquiry becomes a booking",
         "Lead Tracker -> Wedding Bookings",
         "Log the enquiry the day it lands. The tab tells you when "
         "to reply, when to chase and when to let it go. When they "
         "say yes, put the Client ID on the lead and create the "
         "wedding - the lead's value and your conversion rate follow "
         "automatically."),
        ("A package becomes money in the bank",
         "Packages & Pricing -> Wedding Bookings -> Payments",
         "The package sets the price. The booking adds overtime, "
         "add-ons and travel, works out the deposit and the balance "
         "date from your rules, and the payments tab chases anything "
         "late. The booking always shows what is still owed."),
        ("Planning the wedding day",
         "Wedding Timeline + Shot List Manager",
         "Build the running order from the ceremony time - the "
         "standard events fill themselves in. Write the family "
         "groups out by name. Print the Client Questionnaire and "
         "take it with you."),
        ("Editing becomes a delivered gallery",
         "Editing Workflow -> Gallery & Delivery",
         "Ten steps from backing up the cards to uploading the "
         "gallery. The promised date comes from your turnaround "
         "setting. After delivery the gallery tab walks you through "
         "the fortnight that follows - open it, ask for the review, "
         "chase the album, warn them before it closes."),
        ("Knowing what a wedding actually made",
         "Profit Analysis",
         "Type a Booking ID and see the whole wedding: what came in, "
         "what went out, how many hours it really took and what that "
         "works out at an hour. It is usually less than you think, "
         "and it is the number that should decide your prices."),
    ]
    for title, tabs, text in jobs:
        s.write(row, 1, title, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, tabs, fmts.chrome("subsection"))
        row += 1
        s.merge(row, 1, LAST_COL, text, fmts.chrome("note"))
        s.row_height(row, 42)
        row += 1
    row += 1

    # ---- how to read a row --------------------------------------------------
    row = s.section(row, "How to read any register", LAST_COL)
    row = s.bullets(row, [
        "Row 4 is the heading, row 5 is the first record. Everything "
        "below it is yours.",
        "White cells are inputs. Shaded cells are worked out and "
        "locked.",
        "The last column on every register is called 'What to do'. It "
        "writes one plain line per row - 'Chase - contract not signed "
        "back', 'Overdue - gallery past the promised date', 'OK'. "
        "Sort by it on a Monday morning.",
        "The Dashboard counts those lines, so the number next to a "
        "tab's name is the number of rows on that tab wanting "
        "something from you.",
        "Hidden columns hold the plumbing - month keys, sort keys, "
        "yes/no flags the formulas need. Leave them be; they are "
        "hidden because they are dull, not because they are secret.",
        "Need more rows? Select the last data row, copy it, and "
        "insert copied cells. The formulas come with it.",
    ], LAST_COL)
    row += 1

    # ---- what it is not ------------------------------------------------------
    row = s.section(row, "What this is not", LAST_COL)
    row = s.note(row, HONESTY_NOTE, LAST_COL, height=58)
    row += 1
    row = s.bullets(row, [
        "It is not connected to anything. No enquiry form, no "
        "calendar sync, no gallery host, no payment processor, no "
        "email. Every figure here is one you typed.",
        "It does not write or sign contracts. Contracts & Documents "
        "records which paperwork exists, when it went out and where "
        "the signed copy is kept - nothing more.",
        "It is not accounting software and it is not tax advice. "
        "Give your accountant the Expense Tracker and the Payments "
        "tab.",
        "It does not store passwords, card numbers or anything else "
        "you would not want in a file you email to your bookkeeper.",
        "It will not make you book more weddings. It will stop you "
        "losing the ones you have already booked, which is nearly as "
        "useful.",
    ], LAST_COL)
    row += 1
    row = s.note(row, CONTRACT_NOTE, LAST_COL, height=48)
    row += 1
    row = s.note(row, PRIVACY_NOTE, LAST_COL, height=48)
    row += 1

    # ---- the tabs ------------------------------------------------------------
    row = s.section(row, "Every tab, and what it is for", LAST_COL)
    head = row
    s.table_head(head, 1, ["Tab", "What it is for", "Fill in first?"],
                 widths=[26, 62, 14])
    row = head + 1
    tabs = [
        ("Dashboard", "Everything at a glance. Read-only.", "No"),
        ("Getting Started", "This page.", "No"),
        ("Settings & Lists",
         "Your studio, your prices, your rules, your drop-downs.",
         "Yes - first"),
        ("Client CRM", "One row per couple.", "Yes"),
        ("Lead Tracker",
         "Every enquiry, and whether you replied in time.", "Yes"),
        ("Booking Calendar",
         "Twelve months of dates with the booked ones coloured.",
         "No"),
        ("Wedding Bookings",
         "One row per wedding. The spine of the whole file.",
         "Yes - second"),
        ("Packages & Pricing",
         "What you sell, what it costs you, what it should cost.",
         "Yes"),
        ("Contracts & Documents",
         "Which paperwork exists and where it is kept.", "As you go"),
        ("Payments & Invoices",
         "Every instalment, due and received.", "Yes"),
        ("Wedding Timeline",
         "The running order of each wedding day.", "Per wedding"),
        ("Shot List Manager",
         "Every photograph promised, grouped and ticked off.",
         "Per wedding"),
        ("Photographer & Staff",
         "Who shoots, who edits, what they cost.", "Yes"),
        ("Editing Workflow",
         "Ten steps from memory card to finished gallery.",
         "Per wedding"),
        ("Gallery & Delivery",
         "Handing it over, and the month afterwards.",
         "Per wedding"),
        ("Equipment Inventory",
         "What you own, what it is worth, what needs looking at.",
         "When you can"),
        ("Travel & Mileage",
         "Every journey, what it cost and what to charge on.",
         "As you go"),
        ("Marketing Calendar",
         "What you posted and whether anybody booked.",
         "When you can"),
        ("Expense Tracker", "Everything the business pays out.",
         "As you go"),
        ("Profit Analysis",
         "What each wedding actually made. Read-only.", "No"),
        ("Client Questionnaire",
         "The printable form, filled in from the booking.", "No"),
        ("Reviews & Referrals",
         "Who you asked, what they wrote, what it led to.",
         "After delivery"),
        ("Monthly & Annual Reports",
         "The month and the year in numbers. Read-only.", "No"),
    ]
    for name, purpose, first in tabs:
        s.write(row, 1, name, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL - 1, purpose, fmts.chrome("note"))
        s.write(row, LAST_COL, first, fmts.chrome("note"))
        row += 1
    row += 1

    # ---- the small print -------------------------------------------------
    row = s.section(row, "The small print", LAST_COL)
    row = s.field(row, "Product", PRODUCT_NAME, mode="calc",
                  value_end=LAST_COL)
    row = s.field(row, "Listing title", LISTING_TITLE, mode="calc",
                  value_end=LAST_COL)
    row = s.field(row, "Version", VERSION, mode="calc")
    row = s.field(row, "Created by", cfg.author, mode="calc")
    row = s.field(row, "Sheet password", cfg.password, mode="calc",
                  help_text="Every sheet is protected so formulas "
                            "cannot be overwritten by accident. Use "
                            "this to unprotect one if you want to "
                            "change it.")
    row = s.field(row, "Example data",
                  "Included" if cfg.include_sample_data
                  else "Not included - this is the blank copy",
                  mode="calc", value_end=LAST_COL)
    row += 1
    row = s.note(row, DISCLAIMER, LAST_COL, height=96)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    s.protect()
    return s
````

---

## 5. Enquiry to confirmed booking

Workflow one. The lead tracker chases an enquiry on your own schedule and reports the win rate; the client CRM holds the couple; Wedding Bookings mints the Booking ID that every other tab joins to.

### `wedding_photography_spreadsheet/sheets/leads.py`

*228 lines · 1 function*

> Lead Tracker - every enquiry, and what happened to it.
>
> A lead is not a client.  Most enquiries never book, and the ones that
> do take weeks of back and forth to get there.  This tab exists to make
> sure nobody is left waiting for a reply, and to answer the only
> marketing question that matters: of the people who asked, how many
> said yes, and where did they come from?

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Lead Tracker').

````python
"""Lead Tracker - every enquiry, and what happened to it.

A lead is not a client.  Most enquiries never book, and the ones that
do take weeks of back and forth to get there.  This tab exists to make
sure nobody is left waiting for a reply, and to answer the only
marketing question that matters: of the people who asked, how many
said yes, and where did they come from?
"""

from __future__ import annotations

from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, date_num_col, id_col, is_open_lead,
                     is_won_lead, money_cf, month_col, priority_cf,
                     seq_num_col, status_cf)

NAME = "Lead Tracker"


def spec(cfg) -> SheetSpec:
    columns = [
        id_col("name", "LeadPrefix", digits=cfg.digits("lead"),
               header="Lead ID"),

        # ---- the enquiry --------------------------------------------------
        Col("name", "Enquiry name", 19, "text",
            note="Whoever wrote to you. Add them to the Client CRM once "
                 "they are worth keeping.",
            group="The enquiry"),
        Col("partner", "Partner name", 16, "text", group="The enquiry"),
        Col("email", "Email", 23, "text", group="The enquiry"),
        Col("phone", "Phone", 14, "text", group="The enquiry"),
        Col("client", "Client ID", 11, "id",
            validate=list_from("ClientList"),
            note="Link the enquiry to a Client CRM row once you have "
                 "created one. Leave blank for a cold enquiry.",
            group="The enquiry"),
        Col("enquiry_date", "Enquiry date", 12, "date",
            validate=date_check(), group="The enquiry"),
        Col("source", "Lead source", 15, "text",
            validate=list_from("LeadSourceList"),
            note="Where this enquiry actually came from. The marketing "
                 "dashboard is only as honest as this column.",
            group="The enquiry"),
        Col("campaign", "Came from post", 13, "id",
            validate=list_from("CampaignList"),
            note="The Post ID from the Marketing Calendar, when you "
                 "know which post they saw. Two seconds here is the "
                 "only way to ever learn which marketing works.",
            group="The enquiry"),
        Col("referred_by", "Referred by", 12, "id",
            validate=list_from("BookingList"),
            note="The Booking ID of the couple who sent them. Using "
                 "the ID rather than a name is what lets Reviews & "
                 "Referrals add up what a past couple has been worth.",
            group="The enquiry"),
        Col("referrer_name", "Referrer", 18, "text",
            formula=('=IF({referred_by}{r}="","",'
                     'IFERROR(INDEX([[Wedding Bookings.client_name]],'
                     'MATCH({referred_by}{r},'
                     '[[Wedding Bookings.id]],0))&"",""))'),
            group="The enquiry"),

        # ---- what they are asking for -------------------------------------
        Col("wedding_date", "Wedding date", 12, "date",
            validate=date_check(), group="What they want"),
        Col("venue", "Venue", 20, "text", group="What they want"),
        Col("guests", "Guests", 8, "int",
            validate=number_check(0, 2000), group="What they want"),
        Col("package", "Package asked about", 17, "text",
            validate=list_from("PackageList"), group="What they want"),
        Col("budget", "Their budget", 12, "money",
            validate=number_check(0, 1000000),
            note="What they said, not what you hope. An honest number "
                 "here makes the conversion rate mean something.",
            group="What they want"),

        # ---- where it got to ----------------------------------------------
        Col("status", "Lead status", 17, "status",
            validate=list_from("LeadStatusList"), cf=status_cf(),
            note="New until you reply. Booked only once the contract is "
                 "signed and the deposit has cleared.",
            group="Where it got to"),
        Col("priority", "Priority", 9, "status",
            validate=list_from("PriorityList"), cf=priority_cf(),
            group="Where it got to"),
        Col("replied", "First replied", 12, "date",
            validate=date_check(),
            note="The date you first answered. Fill this in - the "
                 "response-time figure on the dashboard is read from it.",
            group="Where it got to"),
        Col("consultation", "Consultation", 12, "date",
            validate=date_check(), group="Where it got to"),
        Col("proposal_sent", "Proposal sent", 12, "date",
            validate=date_check(), group="Where it got to"),
        Col("decision_date", "Booked or lost on", 13, "date",
            validate=date_check(), group="Where it got to"),
        Col("lost_reason", "Why it was lost", 22, "text",
            note="Price, date gone, no reply, went with a friend. Six "
                 "months of this column is worth more than any course.",
            group="Where it got to"),
        Col("follow_up", "Follow up on", 12, "date",
            validate=date_check(),
            note="The next time you intend to chase. Leave it blank and "
                 "the workbook will work one out for you.",
            group="Where it got to"),
        Col("notes", "Notes", 30, "wrap", group="Where it got to"),

        # ---- hidden plumbing ----------------------------------------------
        month_col("enquiry_date", "month", "Enquiry month"),
        date_num_col("enquiry_date", "enq_num", "Enquiry date number"),
        date_num_col("follow_up", "follow_num", "Follow up number"),
        seq_num_col("name"),

        # ---- what the workbook works out ----------------------------------
        Col("is_open", "Still live", 9, "flag",
            formula=('=IF({name}{r}="","",'
                     f'IF({is_open_lead("{status}{r}")},"Yes","No"))'),
            hidden=True, group="How it is going"),
        Col("is_won", "Booked", 9, "flag",
            formula=('=IF({name}{r}="","",'
                     f'IF({is_won_lead("{status}{r}")},"Yes","No"))'),
            hidden=True, group="How it is going"),
        Col("response_hours", "Days to reply", 11, "number1",
            formula=('=IF({name}{r}="","",IF({enquiry_date}{r}="","",'
                     'IF({replied}{r}="","",'
                     'ROUND({replied}{r}-{enquiry_date}{r},1))))'),
            note="Couples book the photographer who answered first far "
                 "more often than the one who answered best.",
            group="How it is going"),
        Col("age_days", "Days old", 9, "int",
            formula=('=IF({name}{r}="","",IF({enquiry_date}{r}="","",'
                     'IF({decision_date}{r}<>"",'
                     '{decision_date}{r}-{enquiry_date}{r},'
                     'TodayDate-{enquiry_date}{r})))'),
            group="How it is going"),
        Col("silent_days", "Days since contact", 13, "int",
            formula=('=IF({name}{r}="","",'
                     'IF({is_open}{r}<>"Yes","",'
                     'TodayDate-MAX({enquiry_date}{r}+0,{replied}{r}+0,'
                     '{consultation}{r}+0,{proposal_sent}{r}+0)))'),
            note="How long this couple has been waiting to hear from "
                 "you. The one number on the tab worth sorting by.",
            group="How it is going"),
        Col("chase_on", "Chase by", 12, "date",
            formula=('=IF({name}{r}="","",IF({is_open}{r}<>"Yes","",'
                     'IF({follow_up}{r}<>"",{follow_up}{r},'
                     'IF({proposal_sent}{r}<>"",'
                     '{proposal_sent}{r}+ConsultationChaseDays,'
                     'IF({replied}{r}<>"",{replied}{r}+LeadChaseDays,'
                     'IF({enquiry_date}{r}="","",'
                     '{enquiry_date}{r}+LeadChaseDays))))))'),
            note="Your own follow-up date if you set one, otherwise the "
                 "chase windows from Settings.",
            group="How it is going"),
        Col("overdue", "Days overdue", 11, "int",
            formula=('=IF({name}{r}="","",IF({chase_on}{r}="","",'
                     'MAX(0,TodayDate-{chase_on}{r})))'),
            group="How it is going"),
        Col("lead_time", "Months to wedding", 13, "number1",
            formula=('=IF({name}{r}="","",IF({wedding_date}{r}="","",'
                     'ROUND(({wedding_date}{r}-{enquiry_date}{r})/30.44,'
                     '1)))'),
            group="How it is going"),
        Col("booked_value", "Booked for", 12, "money",
            formula=('=IF({name}{r}="","",IF({client}{r}="","",'
                     'IF({is_won}{r}<>"Yes","",'
                     'SUMIFS([[Wedding Bookings.total_value]],'
                     '[[Wedding Bookings.client]],{client}{r}))))'),
            cf=money_cf(),
            note="What this enquiry turned into, read back from Wedding "
                 "Bookings through the Client ID.",
            group="How it is going"),
        Col("value_gap", "Against their budget", 15, "money",
            formula=('=IF({name}{r}="","",IF({booked_value}{r}="","",'
                     'IF({budget}{r}<=0,"",'
                     '{booked_value}{r}-{budget}{r})))'),
            note="Positive means they spent more than they first said "
                 "they would, which is what a good proposal does.",
            group="How it is going"),

        Col("alert", "What to do", 30, "status",
            formula=(
                     '=IF({name}{r}="","",IF({is_open}{r}<>"Yes",'
                     'IF({is_won}{r}="Yes",IF({client}{r}="",'
                     '"Add - create the client record","OK"),"OK"),'
                     'IF({enquiry_date}{r}="","Add - date the enquiry",'
                     'IF({replied}{r}="",'
                     'IF(TodayDate-{enquiry_date}{r}>=LeadChaseDays,'
                     '"Urgent - never replied to","Action - reply today"),'
                     'IF({overdue}{r}>14,"Overdue - chase or mark it lost",'
                     'IF({overdue}{r}>0,"Follow up - overdue chase",'
                     'IF({status}{r}="Proposal sent",'
                     '"Follow up - proposal out, no answer",'
                     'IF({wedding_date}{r}<>"",'
                     'IF({wedding_date}{r}-TodayDate<30,'
                     '"Urgent - wedding is close and not booked","OK"),'
                     '"OK"))))))))'),
            cf=alert_cf(), group="How it is going"),
    ]

    return SheetSpec(
        name=NAME,
        title="Lead Tracker",
        subtitle=("Every enquiry, how fast you answered it and whether "
                  "it turned into a wedding"),
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=None,
        freeze_cols=2,
        samples=sample("leads", cfg),
        help_text=(
            "One row per enquiry, including the ones that go nowhere - "
            "especially those. A conversion rate calculated only from "
            "the couples who booked is not a conversion rate. Set the "
            "status honestly and the Dashboard will tell you where your "
            "weddings really come from."),
        extra_notes=[
            "'Days since contact' is the column to sort by on a Monday "
            "morning. Anything over your chase window is a couple who "
            "is, right now, reading somebody else's reply.",
            "'Why it was lost' takes ten seconds to fill in and is the "
            "most useful thing on the tab after a year. Price and "
            "'date already gone' are very different problems.",
        ],
    )
````

### `wedding_photography_spreadsheet/sheets/clients.py`

*198 lines · 1 function*

> Client CRM - one row per couple, from first enquiry to final gallery.
>
> This is the people half of the join.  A client row carries who they
> are, how to reach them and what they said they wanted; everything that
> happens to their wedding is read back from the other tabs through the
> Booking ID.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Client CRM').

````python
"""Client CRM - one row per couple, from first enquiry to final gallery.

This is the people half of the join.  A client row carries who they
are, how to reach them and what they said they wanted; everything that
happens to their wedding is read back from the other tabs through the
Booking ID.
"""

from __future__ import annotations

from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from
from .common import (alert_cf, blank, date_num_col, id_col, month_col,
                     money_cf, owing_cf, priority_cf, rank_desc, seq_num_col,
                     status_cf, yesno, zero)

NAME = "Client CRM"


def spec(cfg) -> SheetSpec:
    columns = [
        id_col("name", "ClientPrefix", digits=cfg.digits("client"),
               header="Client ID",
               note="Automatic. This is the ID the Lead Tracker and "
                    "Wedding Bookings point back at."),

        # ---- who they are ------------------------------------------------
        Col("name", "Client name", 20, "text",
            note="The person you deal with. One row per couple, not one "
                 "per person.",
            group="The couple"),
        Col("partner", "Partner name", 18, "text", group="The couple"),
        Col("email", "Email", 24, "text", group="The couple"),
        Col("phone", "Phone", 15, "text", group="The couple"),
        Col("country", "Country", 14, "text",
            validate=list_from("CountryList"), group="The couple"),
        Col("status", "Client status", 14, "status",
            validate=list_from("ClientStatusList"),
            cf=status_cf(),
            note="Enquiry while you are still talking, Booked once the "
                 "contract is signed, Complete when the gallery is out "
                 "and the review is in.",
            group="The couple"),
        Col("priority", "Priority", 10, "status",
            validate=list_from("PriorityList"), cf=priority_cf(),
            group="The couple"),

        # ---- the wedding -------------------------------------------------
        Col("wedding_date", "Wedding date", 13, "date",
            validate=date_check(),
            note="Leave blank for an enquiry with no date yet.",
            group="The wedding"),
        Col("venue", "Venue", 22, "text", group="The wedding"),
        Col("venue_city", "Venue town or city", 16, "text",
            group="The wedding"),
        Col("source", "How they found you", 16, "text",
            validate=list_from("LeadSourceList"),
            note="The single most valuable column on this tab - it is "
                 "what tells you where to spend your marketing time.",
            group="The wedding"),
        Col("referred_by", "Referred by", 18, "text",
            note="If a past client or a venue sent them, name them here "
                 "and the Reviews & Referrals tab will credit it.",
            group="The wedding"),
        Col("first_contact", "First contact", 12, "date",
            validate=date_check(), group="The wedding"),

        # ---- what they want ----------------------------------------------
        Col("style", "Wedding style", 15, "text",
            validate=list_from("WeddingStyleList"), group="What they want"),
        Col("photo_style", "Photography style", 17, "text",
            validate=list_from("PhotoStyleList"), group="What they want"),
        Col("delivery_pref", "Delivery preference", 17, "text",
            validate=list_from("DeliveryPreferenceList"),
            group="What they want"),
        Col("notes", "Notes", 34, "wrap",
            note="What they are like, what they care about, what they are "
                 "worried about. Read this before every call.",
            group="What they want"),
        Col("marketing_ok", "Happy to be featured", 12, "flag",
            validate=yesno("Did they agree you may post their photographs?"),
            note="Ask, record the answer, and do not post until it says "
                 "Yes. A signed contract clause is not the same as a "
                 "couple who want their faces on Instagram.",
            group="What they want"),

        # ---- hidden plumbing ---------------------------------------------
        month_col("wedding_date", "wedding_month", "Wedding month"),
        date_num_col("wedding_date", "wed_num", "Wedding date number"),
        seq_num_col("name"),

        # ---- what the rest of the workbook knows --------------------------
        Col("bookings", "Weddings", 9, "int",
            formula=zero("name",
                         'COUNTIFS([[Wedding Bookings.client]],{id}{r})'),
            note="Counted from Wedding Bookings. A couple can have more "
                 "than one - an engagement shoot and the wedding.",
            group="What they are worth"),
        Col("booked_value", "Booked value", 13, "money",
            formula=zero("name",
                         'SUMIFS([[Wedding Bookings.total_value]],'
                         '[[Wedding Bookings.client]],{id}{r})'),
            cf=money_cf(), group="What they are worth"),
        Col("collected", "Collected", 12, "money",
            formula=zero("name",
                         'SUMIFS([[Payments & Invoices.received]],'
                         '[[Payments & Invoices.client]],{id}{r})'),
            cf=money_cf(), group="What they are worth"),
        Col("outstanding", "Still owed", 12, "money",
            formula=blank("name", '{booked_value}{r}-{collected}{r}'),
            cf=owing_cf(),
            note="Booked value less everything received. The number you "
                 "check before the wedding day, not after it.",
            group="What they are worth"),
        Col("next_date", "Next wedding", 13, "date",
            formula=('=IF({name}{r}="","",'
                     'IF(SUMPRODUCT(MAX(([[Wedding Bookings.client]]='
                     '{id}{r})*([[Wedding Bookings.wed_num]]>=TodayDate)*'
                     '[[Wedding Bookings.seq_num]]))<=0,"",'
                     'IFERROR(INDEX([[Wedding Bookings.wedding_date]],'
                     'SUMPRODUCT(MAX(([[Wedding Bookings.client]]={id}{r})*'
                     '([[Wedding Bookings.wed_num]]>=TodayDate)*'
                     '[[Wedding Bookings.seq_num]]))),"")))'),
            note="The next wedding of theirs that has not happened yet.",
            group="What they are worth"),
        Col("days_away", "Days away", 10, "int",
            formula=('=IF({name}{r}="","",IF({next_date}{r}="","",'
                     '{next_date}{r}-TodayDate))'),
            group="What they are worth"),
        Col("delivered", "Gallery delivered", 13, "date",
            formula=('=IF({name}{r}="","",'
                     'IF(SUMPRODUCT(MAX(([[Gallery & Delivery.client]]='
                     '{id}{r})*[[Gallery & Delivery.delivered_num]]))<=0,"",'
                     'SUMPRODUCT(MAX(([[Gallery & Delivery.client]]='
                     '{id}{r})*[[Gallery & Delivery.delivered_num]]))))'),
            group="What they are worth"),
        Col("review", "Review", 11, "text",
            formula=('=IF({name}{r}="","",'
                     'IFERROR(INDEX([[Reviews & Referrals.status]],'
                     'MATCH({id}{r},[[Reviews & Referrals.client]],0))&"",'
                     '"Not requested"))'),
            group="What they are worth"),
        Col("referrals", "Referrals sent", 11, "int",
            formula=zero("name",
                         'COUNTIFS([[Client CRM.referred_by]],{name}{r})'),
            note="How many other couples named this one as the reason "
                 "they got in touch.",
            group="What they are worth"),
        rank_desc("booked_value", header="Value rank",
                  note="1 = the couple who has booked the most with you."),

        Col("alert", "What to do", 30, "status",
            formula=('=IF({name}{r}="","",'
                     'IF({status}{r}="Cancelled","",'
                     'IF({name}{r}="","Add - name this client",'
                     'IF({email}{r}="","Missing - no email address",'
                     'IF({status}{r}="Enquiry",'
                     'IF({wedding_date}{r}="","Follow up - no date yet",'
                     '"Follow up - enquiry not yet booked"),'
                     'IF({outstanding}{r}>0.004,'
                     'IF({days_away}{r}<>"",'
                     'IF({days_away}{r}<0,"Overdue - wedding done and '
                     'still owing","Due - balance before the wedding"),'
                     '"Due - balance outstanding"),'
                     'IF({status}{r}="Delivered",'
                     'IF({review}{r}="Received","OK",'
                     '"Send - ask for the review"),"OK")))))))'),
            cf=alert_cf(),
            note="One line per client. Everything on the dashboard's "
                 "action list is counted from this column.",
            group="What they are worth"),
    ]

    return SheetSpec(
        name=NAME,
        title="Client CRM",
        subtitle=("Every couple who has ever got in touch, what they "
                  "booked and what they still owe"),
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=None,
        freeze_cols=2,
        samples=sample("clients", cfg),
        help_text=(
            "One row per couple. The Client ID in column A is what the "
            "Lead Tracker and Wedding Bookings point back at, so never "
            "retype a client's name into another tab - point at the ID "
            "and the name, the venue and the date follow."),
        extra_notes=[
            "'How they found you' is the column that pays for this "
            "workbook. Fill it in every single time, even for a walk-in, "
            "and the Dashboard will tell you which channel is actually "
            "worth your Sunday evenings.",
            "Record only what you need. No card numbers, no bank details "
            "and no ID scans belong in a spreadsheet.",
        ],
    )
````

### `wedding_photography_spreadsheet/sheets/bookings.py`

*479 lines · 1 function*

> Wedding Bookings - one row per wedding, and the spine of the file.
>
> The Booking ID generated here is what the timeline, the shot list, the
> contracts, the payments, the editing workflow, the gallery, the travel
> log and the profit analysis all point back at.  Nothing else in the
> workbook is allowed to identify a wedding by the couple's name.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Wedding Bookings').

````python
"""Wedding Bookings - one row per wedding, and the spine of the file.

The Booking ID generated here is what the timeline, the shot list, the
contracts, the payments, the editing workflow, the gallery, the travel
log and the profit analysis all point back at.  Nothing else in the
workbook is allowed to identify a wedding by the couple's name.
"""

from __future__ import annotations

from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check, time_check
from .common import (alert_cf, countdown_band, date_num_col,
                     flag_no_bad_cf, hours_between, id_col, is_earning,
                     is_holding, is_shot, lookup_client, lookup_package_num,
                     lookup_staff, money_cf, month_col, owing_cf,
                     rank_desc, sales_tax, seq_num_col, status_cf,
                     weekday_col, year_col, zero)

NAME = "Wedding Bookings"


def spec(cfg) -> SheetSpec:
    columns = [
        id_col("client", "BookingPrefix", digits=cfg.digits("booking"),
               header="Booking ID",
               note="Automatic, and the most important value in the "
                    "workbook. Nine other tabs join to it."),

        # ---- who and when -------------------------------------------------
        Col("client", "Client ID", 11, "id",
            validate=list_from("ClientList"),
            note="Pick the couple from the Client CRM. Everything else "
                 "about them follows from this one cell.",
            group="The wedding"),
        Col("client_name", "Couple", 20, "text",
            formula=lookup_client("name"), group="The wedding"),
        Col("partner_name", "Partner", 16, "text",
            formula=lookup_client("partner"), group="The wedding"),
        Col("wedding_date", "Wedding date", 13, "date",
            validate=date_check(),
            note="The one date the whole workbook counts from.",
            group="The wedding"),
        weekday_col("wedding_date", header="Day"),
        Col("status", "Booking status", 14, "status",
            validate=list_from("BookingStatusList"), cf=status_cf(),
            note="Pencilled in holds the date but earns nothing. "
                 "Confirmed means signed and deposit paid.",
            group="The wedding"),
        Col("confirmed_date", "Confirmed on", 12, "date",
            validate=date_check(), group="The wedding"),

        # ---- the day ------------------------------------------------------
        Col("ceremony_time", "Ceremony", 10, "time",
            validate=time_check(),
            note="The anchor for the whole wedding-day timeline - every "
                 "default event time is an offset from this.",
            group="The day"),
        Col("reception_time", "Reception", 10, "time",
            validate=time_check(), group="The day"),
        Col("photo_start", "Coverage from", 11, "time",
            validate=time_check(), group="The day"),
        Col("photo_end", "Coverage to", 11, "time",
            validate=time_check(),
            note="If you finish after midnight the workbook still gets "
                 "the hours right.",
            group="The day"),
        Col("venue", "Main venue", 20, "text", group="The day"),
        Col("venue_type", "Venue type", 13, "text",
            validate=list_from("VenueTypeList"), group="The day"),
        Col("prep_location", "Getting ready at", 20, "text",
            group="The day"),
        Col("ceremony_location", "Ceremony at", 20, "text",
            group="The day"),
        Col("reception_location", "Reception at", 20, "text",
            group="The day"),
        Col("guests", "Guests", 8, "int",
            validate=number_check(0, 2000), group="The day"),

        # ---- the team -----------------------------------------------------
        Col("photographer", "Lead photographer", 16, "id",
            validate=list_from("StaffList"), group="The team"),
        Col("photographer_name", "Lead name", 16, "text",
            formula=lookup_staff("name"), group="The team"),
        Col("second", "Second shooter", 14, "id",
            validate=list_from("StaffList"),
            note="Leave blank if you are shooting alone. The cost "
                 "follows from the rate on their staff row.",
            group="The team"),
        Col("second_name", "Second name", 15, "text",
            formula=lookup_staff("name", key="second", missing=""),
            group="The team"),

        # ---- the money ----------------------------------------------------
        Col("package", "Package", 17, "text",
            validate=list_from("PackageList"),
            note="Pick from Packages & Pricing. The price, the hours "
                 "and the inclusions all come from that tab.",
            group="The money"),
        Col("package_price", "Package price", 13, "money",
            formula=lookup_package_num("total_price"),
            cf=money_cf(), group="The money"),
        Col("package_hours", "Hours included", 12, "number1",
            formula=lookup_package_num("hours"), group="The money"),
        Col("coverage_hours", "Hours booked", 12, "number1",
            formula=('=IF({client}{r}="","",'
                     + hours_between("{photo_start}{r}", "{photo_end}{r}")
                     + ")"),
            note="Worked out from the coverage times, so overtime "
                 "cannot be forgotten about.",
            group="The money"),
        Col("extra_hours", "Extra hours", 10, "number1",
            formula=('=IF({client}{r}="",0,'
                     'MAX(0,{coverage_hours}{r}-{package_hours}{r}))'),
            group="The money"),
        Col("extra_hour_fee", "Extra hour charge", 14, "money",
            formula=('=IF({client}{r}="",0,'
                     'ROUND({extra_hours}{r}*ExtraHourRate,2))'),
            group="The money"),
        Col("addons", "Other add-ons", 13, "money",
            validate=number_check(0, 1000000),
            note="Albums, extra prints, a drone, a rehearsal dinner. "
                 "Type the total you agreed.",
            group="The money"),
        Col("travel_fee", "Travel charged", 13, "money",
            validate=number_check(0, 100000),
            note="Leave blank to let the Travel & Mileage tab work it "
                 "out from the distance and your mileage rate.",
            group="The money"),
        Col("travel_used", "Travel used", 12, "money",
            formula=('=IF({client}{r}="",0,IF({travel_fee}{r}<>"",'
                     '{travel_fee}{r},'
                     'SUMIFS([[Travel & Mileage.chargeable]],'
                     '[[Travel & Mileage.booking]],{id}{r})))'),
            hidden=True, group="The money"),
        Col("discount", "Discount", 11, "money",
            validate=number_check(0, 1000000), group="The money"),
        Col("subtotal", "Subtotal", 12, "money",
            formula=('=IF({client}{r}="",0,'
                     'ROUND({package_price}{r}+{extra_hour_fee}{r}+'
                     'IF({addons}{r}="",0,{addons}{r})+{travel_used}{r}-'
                     'IF({discount}{r}="",0,{discount}{r}),2))'),
            group="The money"),
        Col("tax", "Tax", 11, "money",
            formula=('=IF({client}{r}="",0,'
                     + sales_tax("{subtotal}{r}") + ")"),
            group="The money"),
        Col("total_value", "Total booking value", 15, "money",
            formula=('=IF({client}{r}="",0,'
                     'ROUND({subtotal}{r}+{tax}{r},2))'),
            cf=money_cf(),
            note="What this wedding is worth. Every revenue figure in "
                 "the workbook is a sum of this column.",
            group="The money"),

        # ---- getting paid --------------------------------------------------
        Col("deposit_due_amount", "Deposit due", 12, "money",
            formula=('=IF({client}{r}="",0,'
                     'ROUND({total_value}{r}*DepositPercent,2))'),
            group="Getting paid"),
        Col("deposit_due_date", "Deposit by", 12, "date",
            formula=('=IF({client}{r}="","",'
                     'IF({confirmed_date}{r}="","",'
                     '{confirmed_date}{r}+DepositDays))'),
            group="Getting paid"),
        Col("final_due_date", "Final payment by", 14, "date",
            formula=('=IF({client}{r}="","",IF({wedding_date}{r}="","",'
                     '{wedding_date}{r}-FinalPaymentDays))'),
            note="Your own rule from Settings, counted back from the "
                 "wedding day.",
            group="Getting paid"),
        Col("collected", "Collected", 12, "money",
            formula=zero("client",
                         'SUMIFS([[Payments & Invoices.received]],'
                         '[[Payments & Invoices.booking]],{id}{r})'),
            cf=money_cf(), group="Getting paid"),
        Col("outstanding", "Outstanding", 12, "money",
            formula=('=IF({client}{r}="",0,'
                     'ROUND({total_value}{r}-{collected}{r},2))'),
            cf=owing_cf(), group="Getting paid"),
        Col("paid_pct", "Paid", 8, "percent0",
            formula=('=IF({client}{r}="","",IF({total_value}{r}<=0,"",'
                     '{collected}{r}/{total_value}{r}))'),
            group="Getting paid"),
        Col("deposit_paid", "Deposit in", 10, "flag",
            formula=('=IF({client}{r}="","",'
                     'IF({collected}{r}>={deposit_due_amount}{r},'
                     'IF({deposit_due_amount}{r}<=0,"n/a","Yes"),"No"))'),
            cf=flag_no_bad_cf(), group="Getting paid"),
        Col("payment_state", "Payment status", 14, "status",
            formula=('=IF({client}{r}="","",'
                     'IF({total_value}{r}<=0,"Nothing to pay",'
                     'IF({outstanding}{r}<=0.004,"Paid in full",'
                     'IF({collected}{r}<=0,'
                     'IF({final_due_date}{r}<>"",'
                     'IF(TodayDate>{final_due_date}{r},"Overdue","Due"),'
                     '"Due"),'
                     'IF({final_due_date}{r}<>"",'
                     'IF(TodayDate>{final_due_date}{r},"Overdue",'
                     '"Part paid"),"Part paid")))))'),
            cf=status_cf(), group="Getting paid"),

        # ---- the questionnaire ---------------------------------------------
        Col("must_haves", "Must-have photographs", 30, "wrap",
            note="The shots they will be upset not to have. Read this "
                 "the night before, not on the morning.",
            group="What they told you"),
        Col("family_list", "Family group list", 30, "wrap",
            note="Write the groups out by name. 'Bride's family' is "
                 "how a forty-minute family session happens.",
            group="What they told you"),
        Col("key_guests", "Important guests", 24, "wrap",
            group="What they told you"),
        Col("traditions", "Traditions and rituals", 24, "wrap",
            note="Anything cultural or religious you need to know "
                 "about, and when in the day it happens.",
            group="What they told you"),
        Col("restrictions", "Venue restrictions", 24, "wrap",
            note="No flash in the church, no photographs during the "
                 "ceremony, a balcony you may not stand on.",
            group="What they told you"),
        Col("sensitivities", "Family sensitivities", 24, "wrap",
            note="Divorced parents, a recent bereavement, somebody who "
                 "must not be in a group shot with somebody else.",
            group="What they told you"),
        Col("special_requests", "Special requests", 24, "wrap",
            group="What they told you"),
        Col("questionnaire_sent", "Questionnaire sent", 13, "date",
            validate=date_check(), group="What they told you"),
        Col("questionnaire_back", "Questionnaire back", 13, "date",
            validate=date_check(), group="What they told you"),
        Col("notes", "Notes", 28, "wrap", group="What they told you"),

        # ---- hidden plumbing ------------------------------------------------
        month_col("wedding_date", "month", "Wedding month"),
        year_col("wedding_date"),
        Col("season", "Season", 9, "text",
            formula=('=IF({wedding_date}{r}="","",'
                     'IF(MONTH({wedding_date}{r})<=2,"Winter",'
                     'IF(MONTH({wedding_date}{r})<=5,"Spring",'
                     'IF(MONTH({wedding_date}{r})<=8,"Summer",'
                     'IF(MONTH({wedding_date}{r})<=11,"Autumn",'
                     '"Winter")))))'),
            hidden=True),
        Col("month_seq", "Line in the month", 9, "int",
            formula=('=IF({client}{r}="","",IF({earns}{r}<>"Yes","",'
                     'COUNTIFS({month_first}:{month}{r},{month}{r},'
                     '{earns_first}:{earns}{r},"Yes")))'),
            hidden=True),
        Col("month_key", "Month key", 14, "text",
            formula=('=IF({month_seq}{r}="","",'
                     '{month}{r}&"|"&{month_seq}{r})'),
            hidden=True,
            note="Lets Monthly & Annual Reports list the weddings in "
                 "a month without an array formula."),
        date_num_col("wedding_date", "wed_num", "Wedding date number"),
        seq_num_col("client"),
        Col("holds", "Holds the date", 9, "flag",
            formula=('=IF({client}{r}="","",'
                     f'IF({is_holding("{status}{r}")},"Yes","No"))'),
            hidden=True),
        Col("earns", "Counts as revenue", 9, "flag",
            formula=('=IF({client}{r}="","",'
                     f'IF({is_earning("{status}{r}")},"Yes","No"))'),
            hidden=True),
        Col("timeline_lines", "Timeline lines", 9, "int",
            formula=('=IF({client}{r}="",0,'
                     'COUNTIFS([[Wedding Timeline.booking]],{id}{r}))'),
            hidden=True),
        Col("shot_lines", "Shot list lines", 9, "int",
            formula=('=IF({client}{r}="",0,'
                     'COUNTIFS([[Shot List Manager.booking]],'
                     '{id}{r}))'),
            hidden=True),
        Col("upcoming_key", "Upcoming key", 12, "number",
            formula=('=IF({client}{r}="",9999999999,'
                     'IF({holds}{r}<>"Yes",9999999999,'
                     'IF({wed_num}{r}<TodayDate,9999999999,'
                     '{wed_num}{r}*10000+ROW()-4)))'),
            hidden=True),
        Col("upcoming_rank", "Next up", 9, "int",
            formula=('=IF({upcoming_key}{r}>=9999999999,"",'
                     'COUNTIFS({upcoming_key_first}:'
                     '{upcoming_key_last},"<"&{upcoming_key}{r})+1)'),
            hidden=True,
            note="1 = the next wedding in the diary. The Dashboard "
                 "and the Booking Calendar read it."),
        Col("shot", "Wedding happened", 9, "flag",
            formula=('=IF({client}{r}="","",'
                     f'IF({is_shot("{status}{r}")},"Yes","No"))'),
            hidden=True),

        # ---- where it is up to ----------------------------------------------
        Col("days_away", "Days away", 10, "int",
            formula=('=IF({client}{r}="","",IF({wedding_date}{r}="","",'
                     '{wedding_date}{r}-TodayDate))'),
            group="Where it is up to"),
        Col("countdown", "When", 13, "text",
            formula=('=IF({client}{r}="","",IF({days_away}{r}="","",'
                     + countdown_band("{days_away}{r}") + "))"),
            group="Where it is up to"),
        Col("clash", "Date clash", 10, "flag",
            formula=('=IF({client}{r}="","",IF({wedding_date}{r}="","",'
                     'IF({holds}{r}<>"Yes","No",'
                     'IF(COUNTIFS([[Wedding Bookings.wed_num]],'
                     '{wed_num}{r},[[Wedding Bookings.holds]],"Yes")>1,'
                     '"Yes","No"))))'),
            cf=flag_no_bad_cf(),
            note="Two live weddings on the same day. Almost always a "
                 "mistake, and much cheaper to find here than in June.",
            group="Where it is up to"),
        Col("team_clash", "Photographer clash", 13, "flag",
            formula=('=IF({client}{r}="","",IF({photographer}{r}="","",'
                     'IF({holds}{r}<>"Yes","No",'
                     'IF(COUNTIFS([[Wedding Bookings.wed_num]],'
                     '{wed_num}{r},[[Wedding Bookings.photographer]],'
                     '{photographer}{r},[[Wedding Bookings.holds]],'
                     '"Yes")>1,"Yes","No"))))'),
            cf=flag_no_bad_cf(),
            note="The same photographer booked twice on one day.",
            group="Where it is up to"),
        Col("contract_done", "Contract settled", 11, "flag",
            formula=('=IF({client}{r}="","",'
                     'IFERROR(INDEX([[Contracts & Documents.is_done]],'
                     'MATCH({id}{r},'
                     '[[Contracts & Documents.booking]],0))&"","No"))'),
            hidden=True, group="Where it is up to"),
        Col("contract_state", "Contract", 13, "status",
            formula=('=IF({client}{r}="","",'
                     'IFERROR(INDEX([[Contracts & Documents.status]],'
                     'MATCH({id}{r},[[Contracts & Documents.booking]],0))'
                     '&"","Not sent"))'),
            cf=status_cf(), group="Where it is up to"),
        Col("shots_done", "Shot list", 9, "percent0",
            formula=('=IF({client}{r}="","",'
                     'IF(COUNTIFS([[Shot List Manager.booking]],{id}{r})'
                     '<=0,"",'
                     'COUNTIFS([[Shot List Manager.booking]],{id}{r},'
                     '[[Shot List Manager.done]],"Yes")/'
                     'COUNTIFS([[Shot List Manager.booking]],{id}{r})))'),
            group="Where it is up to"),
        Col("editing_state", "Editing", 14, "status",
            formula=('=IF({client}{r}="","",'
                     'IFERROR(INDEX([[Editing Workflow.stage]],'
                     'MATCH({id}{r},[[Editing Workflow.booking]],0))&"",'
                     'IF({shot}{r}="Yes","Not started","")))'),
            cf=status_cf(), group="Where it is up to"),
        Col("gallery_state", "Gallery", 13, "status",
            formula=('=IF({client}{r}="","",'
                     'IFERROR(INDEX([[Gallery & Delivery.status]],'
                     'MATCH({id}{r},[[Gallery & Delivery.booking]],0))&"",'
                     'IF({shot}{r}="Yes","Not created","")))'),
            cf=status_cf(), group="Where it is up to"),
        Col("deliver_by", "Deliver by", 12, "date",
            formula=('=IF({client}{r}="","",IF({wedding_date}{r}="","",'
                     'IF({shot}{r}<>"Yes","",'
                     '{wedding_date}{r}+DeliveryDays)))'),
            note="Your promised turnaround from Settings, counted from "
                 "the wedding day.",
            group="Where it is up to"),
        Col("delivery_overdue", "Days late", 10, "int",
            formula=('=IF({client}{r}="","",IF({deliver_by}{r}="","",'
                     'IF({gallery_state}{r}="Delivered","",'
                     'MAX(0,TodayDate-{deliver_by}{r}))))'),
            group="Where it is up to"),

        # ---- what it made ----------------------------------------------------
        Col("direct_costs", "Direct costs", 12, "money",
            formula=('=IF({client}{r}="",0,'
                     'SUMIFS([[Expense Tracker.amount]],'
                     '[[Expense Tracker.booking]],{id}{r})+'
                     'SUMIFS([[Travel & Mileage.total_cost]],'
                     '[[Travel & Mileage.booking]],{id}{r}))'),
            group="What it made"),
        Col("team_cost", "Team cost", 11, "money",
            formula=('=IF({client}{r}="",0,'
                     'IF({second}{r}="",0,SecondShooterRate))'),
            note="The second shooter at the rate on Settings. Anything "
                 "else you paid out goes on the Expense Tracker "
                 "against this Booking ID.",
            group="What it made"),
        Col("editing_cost", "Editing cost", 12, "money",
            formula=('=IF({client}{r}="",0,'
                     'IFERROR(INDEX([[Editing Workflow.cost]],'
                     'MATCH({id}{r},[[Editing Workflow.booking]],0))*1,'
                     '0))'),
            group="What it made"),
        Col("total_cost", "Total cost", 12, "money",
            formula=('=IF({client}{r}="",0,'
                     '{direct_costs}{r}+{team_cost}{r}+'
                     '{editing_cost}{r})'),
            group="What it made"),
        Col("profit", "Profit", 12, "money",
            formula=('=IF({client}{r}="",0,'
                     'ROUND({total_value}{r}-{tax}{r}-'
                     '{total_cost}{r},2))'),
            cf=money_cf(),
            note="What the wedding actually made: the booking value "
                 "less the tax you only collected, less everything you "
                 "paid out against it.",
            group="What it made"),
        Col("margin", "Margin", 9, "percent0",
            formula=('=IF({client}{r}="","",IF({total_value}{r}<=0,"",'
                     '{profit}{r}/{total_value}{r}))'),
            group="What it made"),
        Col("profit_hour", "Profit an hour", 12, "money",
            formula=('=IF({client}{r}="","",'
                     'IF({coverage_hours}{r}+EditingHours<=0,"",'
                     'ROUND({profit}{r}/({coverage_hours}{r}+'
                     'EditingHours),2)))'),
            note="Coverage plus your assumed editing hours. The number "
                 "that decides whether a cheap wedding was worth it.",
            group="What it made"),
        rank_desc("profit", guard='{client}{r}=""', header="Profit rank",
                  note="1 = your most profitable wedding."),

        Col("alert", "What to do", 32, "status",
            formula=(
                '=IF({client}{r}="","",'
                'IF({status}{r}="Cancelled","",'
                'IF({wedding_date}{r}="","Add - no wedding date",'
                'IF({clash}{r}="Yes","Duplicate - two weddings this day",'
                'IF({team_clash}{r}="Yes",'
                '"Duplicate - photographer booked twice",'
                'IF({package}{r}="","Add - no package chosen",'
                'IF({contract_state}{r}="Not sent",'
                '"Send - contract not out yet",'
                'IF({days_away}{r}>=0,'
                'IF({contract_done}{r}<>"Yes",'
                '"Chase - contract not signed back",'
                'IF({payment_state}{r}="Overdue",'
                '"Overdue - payment past due",'
                'IF({days_away}{r}<=UpcomingDays,'
                'IF({questionnaire_back}{r}="",'
                '"Action - questionnaire not back",'
                'IF({photographer}{r}="","Add - nobody assigned",'
                '"OK - coming up")),"OK"))),'
                'IF({payment_state}{r}="Overdue",'
                '"Overdue - wedding done and still owing",'
                'IF({delivery_overdue}{r}>0,'
                '"Overdue - gallery past the promised date",'
                'IF({gallery_state}{r}="Delivered","OK",'
                '"Action - in post-production")))))))))))'),
            cf=alert_cf(),
            note="One line per wedding, read in priority order. The "
                 "Dashboard's action list counts this column.",
            group="What it made"),
    ]

    return SheetSpec(
        name=NAME,
        title="Wedding Bookings",
        subtitle=("Every wedding you have taken on - the day, the team, "
                  "the money and where it has got to"),
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=None,
        freeze_cols=3,
        samples=sample("bookings", cfg),
        help_text=(
            "One row per wedding. The Booking ID in column A is the key "
            "the whole workbook turns on - the timeline, the shot list, "
            "the contract, the invoices, the editing job, the gallery, "
            "the travel and the profit analysis all point at it. Create "
            "the booking here first, then use the ID everywhere else."),
        extra_notes=[
            "'Date clash' and 'Photographer clash' read every other row "
            "on this tab. A double booking found in the spreadsheet is "
            "an awkward phone call; one found by the couple is a "
            "refund and a review.",
            "Coverage hours are worked out from the two time cells, so "
            "overtime bills itself. A finish after midnight is handled.",
            "Profit here is the booking value less tax you merely "
            "collected and less everything logged against the Booking "
            "ID - the second shooter, the album, the travel and the "
            "editing. Profit Analysis breaks the same figure down.",
        ],
    )
````

---

## 6. Package to payment collection

Workflow two. What a package costs you to deliver and earns per hour, the documents that have to come back signed, and the deposit, instalments and balance with an overdue count in days.

### `wedding_photography_spreadsheet/sheets/packages.py`

*270 lines · 1 function*

> Packages & Pricing - what you sell, what it costs you, what it earns.
>
> This is the pricing calculator at the heart of the product.  Change the
> hours on a package and the cost of delivering it, the margin, the
> effective hourly rate and the suggested price all move with it, and
> every booking already sold on that package re-prices itself.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Packages & Pricing').

````python
"""Packages & Pricing - what you sell, what it costs you, what it earns.

This is the pricing calculator at the heart of the product.  Change the
hours on a package and the cost of delivering it, the margin, the
effective hourly rate and the suggested price all move with it, and
every booking already sold on that package re-prices itself.
"""

from __future__ import annotations

from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import (alert_cf, bar_cf, blank, loss_cf, money_cf, rank_desc,
                     status_cf, yesno, zero)

NAME = "Packages & Pricing"


def spec(cfg) -> SheetSpec:
    columns = [
        Col("name", "Package name", 20, "id",
            note="Type the name exactly as you want it to appear on "
                 "the booking drop-down and on your invoices.",
            group="The package"),
        Col("tier", "Tier", 11, "text",
            validate=list_from("TierList"), group="The package"),
        Col("active", "Still selling", 10, "flag",
            validate=yesno("Retired packages stay here so old bookings "
                           "keep their price, but drop off the list of "
                           "what you offer."),
            group="The package"),
        Col("summary", "One-line summary", 30, "wrap",
            group="The package"),

        # ---- what the couple gets -----------------------------------------
        Col("hours", "Coverage hours", 12, "number1",
            validate=number_check(0, 24),
            note="Hours on the day. Anything past this bills at your "
                 "extra-hour rate on the booking automatically.",
            group="What they get"),
        Col("photographers", "Photographers", 12, "int",
            validate=number_check(1, 6), group="What they get"),
        Col("engagement", "Engagement shoot", 12, "flag",
            validate=yesno(), group="What they get"),
        Col("images", "Edited images", 11, "int",
            validate=number_check(0, 5000), group="What they get"),
        Col("album", "Album", 14, "text", group="What they get"),
        Col("print_credit", "Print credit", 11, "money",
            validate=number_check(0, 100000), group="What they get"),
        Col("travel_allowance", "Travel included", 12, "int",
            validate=number_check(0, 5000),
            note="Miles included before you start charging mileage.",
            group="What they get"),
        Col("inclusions", "Everything included", 34, "wrap",
            note="The list you paste into a proposal. Write it once "
                 "here rather than re-typing it for every enquiry.",
            group="What they get"),

        # ---- the price ----------------------------------------------------
        Col("base_price", "List price", 12, "money",
            validate=number_check(0, 1000000), cf=money_cf(),
            group="The price"),
        Col("standard_discount", "Standard discount", 13, "money",
            validate=number_check(0, 1000000),
            note="An off-season or weekday reduction you are happy to "
                 "give without being asked.",
            group="The price"),
        Col("total_price", "Price charged", 13, "money",
            formula=zero("name",
                         'ROUND({base_price}{r}-'
                         'IF({standard_discount}{r}="",0,'
                         '{standard_discount}{r}),2)'),
            cf=money_cf(),
            note="What a booking on this package is actually worth. "
                 "Every booking reads this cell.",
            group="The price"),

        # ---- what it costs you ---------------------------------------------
        Col("second_cost", "Second shooter", 12, "money",
            formula=zero("name",
                         'IF({photographers}{r}<=1,0,'
                         '({photographers}{r}-1)*SecondShooterRate)'),
            group="What it costs you"),
        Col("album_cost", "Album cost", 11, "money",
            validate=number_check(0, 100000),
            note="What your lab charges you, not what you charge the "
                 "couple.",
            group="What it costs you"),
        Col("print_cost", "Print cost", 11, "money",
            formula=zero("name",
                         'ROUND(IF({print_credit}{r}="",0,'
                         '{print_credit}{r})*PrintCostRatio,2)'),
            group="What it costs you"),
        Col("engagement_cost", "Engagement cost", 13, "money",
            formula=zero("name",
                         'IF({engagement}{r}<>"Yes",0,'
                         'EngagementCost)'),
            group="What it costs you"),
        Col("editing_hours", "Editing hours", 12, "number1",
            validate=number_check(0, 400),
            note="Leave blank to use the studio default from Settings. "
                 "Be honest here - underestimating editing is how a "
                 "package that looks profitable is not.",
            group="What it costs you"),
        Col("editing_hours_used", "Editing hours used", 13, "number1",
            formula=zero("name",
                         'IF({editing_hours}{r}="",EditingHours,'
                         '{editing_hours}{r})'),
            hidden=True, group="What it costs you"),
        Col("editing_cost", "Editing cost", 12, "money",
            formula=zero("name",
                         'ROUND({editing_hours_used}{r}*EditingRate,2)'),
            group="What it costs you"),
        Col("other_cost", "Other costs", 11, "money",
            validate=number_check(0, 100000),
            note="Insurance apportioned, a sample album, anything else "
                 "this package specifically costs you.",
            group="What it costs you"),
        Col("total_cost", "Cost to deliver", 13, "money",
            formula=zero("name",
                         'ROUND({second_cost}{r}+'
                         'IF({album_cost}{r}="",0,{album_cost}{r})+'
                         '{print_cost}{r}+{engagement_cost}{r}+'
                         '{editing_cost}{r}+'
                         'IF({other_cost}{r}="",0,{other_cost}{r}),2)'),
            group="What it costs you"),

        # ---- is it worth selling --------------------------------------------
        Col("profit", "Profit each", 12, "money",
            formula=zero("name",
                         'ROUND({total_price}{r}-{total_cost}{r},2)'),
            cf=loss_cf(), group="Is it worth selling"),
        Col("margin", "Margin", 9, "percent0",
            formula=blank("name",
                          'IF({total_price}{r}<=0,"",'
                          '{profit}{r}/{total_price}{r})'),
            cf=bar_cf(), group="Is it worth selling"),
        Col("total_hours", "Your hours", 11, "number1",
            formula=zero("name",
                         '{hours}{r}+{editing_hours_used}{r}+'
                         'AdminHoursPerWedding'),
            note="Shooting, editing and the admin every wedding takes. "
                 "The only honest basis for an hourly rate.",
            group="Is it worth selling"),
        Col("hourly", "Your hourly rate", 13, "money",
            formula=blank("name",
                          'IF({total_hours}{r}<=0,"",'
                          'ROUND({profit}{r}/{total_hours}{r},2))'),
            cf=loss_cf(),
            note="Profit divided by every hour the package really "
                 "takes. Compare it with what you would accept from an "
                 "employer before you discount anything.",
            group="Is it worth selling"),
        Col("target_price", "Price for target margin", 15, "money",
            formula=blank("name",
                          'IF(TargetMargin>=1,"",'
                          'ROUND({total_cost}{r}/(1-TargetMargin),2))'),
            note="What this package would have to sell for to hit the "
                 "margin you set on Settings.",
            group="Is it worth selling"),
        Col("price_gap", "Under-priced by", 13, "money",
            formula=zero("name",
                         'IF({target_price}{r}="",0,'
                         'ROUND(MAX(0,{target_price}{r}-'
                         '{total_price}{r}),2))'),
            cf=loss_cf(), group="Is it worth selling"),
        Col("meets_target", "Hits target", 10, "flag",
            formula=blank("name",
                          'IF({total_price}{r}<=0,"",'
                          'IF({margin}{r}>=TargetMargin,"Yes","No"))'),
            cf=status_cf(), group="Is it worth selling"),

        # ---- how it is actually selling ---------------------------------------
        Col("sold", "Weddings sold", 11, "int",
            formula=zero("name",
                         'COUNTIFS([[Wedding Bookings.package]],'
                         '{name}{r},[[Wedding Bookings.earns]],"Yes")'),
            group="How it sells"),
        Col("sold_year", "Sold this year", 11, "int",
            formula=zero("name",
                         'COUNTIFS([[Wedding Bookings.package]],'
                         '{name}{r},[[Wedding Bookings.earns]],"Yes",'
                         '[[Wedding Bookings.year]],YEAR(TodayDate))'),
            group="How it sells"),
        Col("enquiries", "Asked about", 11, "int",
            formula=zero("name",
                         'COUNTIFS([[Lead Tracker.package]],{name}{r})'),
            group="How it sells"),
        Col("conversion", "Enquiry to booking", 13, "percent0",
            formula=blank("name",
                          'IF({enquiries}{r}<=0,"",'
                          'MIN(1,{sold}{r}/{enquiries}{r}))'),
            cf=bar_cf(),
            note="How often an enquiry about this package turns into a "
                 "wedding. A package nobody books may be priced wrong, "
                 "or may simply be the one that makes the next one "
                 "look reasonable.",
            group="How it sells"),
        Col("revenue", "Revenue booked", 13, "money",
            formula=zero("name",
                         'SUMIFS([[Wedding Bookings.total_value]],'
                         '[[Wedding Bookings.package]],{name}{r},'
                         '[[Wedding Bookings.earns]],"Yes")'),
            cf=money_cf(), group="How it sells"),
        Col("profit_booked", "Profit booked", 13, "money",
            formula=zero("name",
                         'SUMIFS([[Wedding Bookings.profit]],'
                         '[[Wedding Bookings.package]],{name}{r},'
                         '[[Wedding Bookings.earns]],"Yes")'),
            cf=loss_cf(), group="How it sells"),
        Col("share", "Share of bookings", 12, "percent0",
            formula=blank("name",
                          'IF(COUNTIFS([[Wedding Bookings.earns]],'
                          '"Yes")<=0,"",{sold}{r}/'
                          'COUNTIFS([[Wedding Bookings.earns]],"Yes"))'),
            group="How it sells"),
        rank_desc("revenue", header="Revenue rank",
                  note="1 = the package that has brought in the most."),
        rank_desc("sold", header="Popularity rank"),

        Col("alert", "What to do", 32, "status",
            formula=(
                '=IF({name}{r}="","",'
                'IF({active}{r}<>"Yes","Retired - kept for old bookings",'
                'IF({base_price}{r}="","Add - no price set",'
                'IF({hours}{r}="","Add - no coverage hours",'
                'IF({profit}{r}<0,"Loss - costs more than it sells for",'
                'IF({price_gap}{r}>0,'
                '"Check - under your target margin",'
                'IF({hourly}{r}<>"",'
                'IF({hourly}{r}<MinimumHourly,'
                '"Check - poor hourly rate",'
                'IF({enquiries}{r}>=5,'
                'IF({conversion}{r}<0.1,'
                '"Check - plenty of interest, nobody books it","OK"),'
                '"OK")),"OK")))))))'),
            cf=alert_cf(), group="How it sells"),
    ]

    return SheetSpec(
        name=NAME,
        title="Packages & Pricing",
        subtitle=("What you sell, what each one really costs you to "
                  "deliver, and what it ought to cost"),
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=None,
        freeze_cols=2,
        samples=sample("packages", cfg),
        help_text=(
            "Fill in what the package includes and what it costs you to "
            "deliver, and the sheet works out the margin, your real "
            "hourly rate and the price you would need to charge to hit "
            "your target. Bookings read the 'Price charged' column, so "
            "changing a price here re-prices every wedding sold on it."),
        extra_notes=[
            "'Your hours' is coverage plus editing plus the admin every "
            "wedding takes. A package that looks fine on margin can pay "
            "you less per hour than the venue's bar staff, and this is "
            "the column that says so.",
            "Costs here are per wedding. Fixed costs - insurance, "
            "software, your website - belong on the Expense Tracker and "
            "come out of the business as a whole, not out of one "
            "package.",
            "'Price for target margin' uses cost / (1 - margin), which "
            "is the only correct way round. Adding a percentage to cost "
            "gives a smaller margin than you intended.",
        ],
    )
````

### `wedding_photography_spreadsheet/sheets/contracts.py`

*151 lines · 1 function*

> Contracts & Documents - a register of paperwork, not a contract tool.
>
> This tab records which documents exist for a wedding, where they are
> kept and whether they have come back signed.  It does not draft them,
> it does not sign them and it is not a substitute for a contract written
> for your jurisdiction.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Contracts & Documents').

````python
"""Contracts & Documents - a register of paperwork, not a contract tool.

This tab records which documents exist for a wedding, where they are
kept and whether they have come back signed.  It does not draft them,
it does not sign them and it is not a substitute for a contract written
for your jurisdiction.
"""

from __future__ import annotations

from ..config import CONTRACT_NOTE
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from
from .common import (alert_cf, blank, date_num_col, id_col, lookup_booking,
                     lookup_booking_num, status_cf, yesno, zero)

NAME = "Contracts & Documents"


def spec(cfg) -> SheetSpec:
    columns = [
        id_col("booking", "ContractPrefix", digits=cfg.digits("contract"),
               header="Document ID"),

        Col("booking", "Booking ID", 11, "id",
            validate=list_from("BookingList"),
            note="Which wedding this document belongs to. One booking "
                 "can have several documents.",
            group="The document"),
        Col("couple", "Couple", 20, "text",
            formula=lookup_booking("client_name"), group="The document"),
        Col("wedding_date", "Wedding date", 12, "date",
            formula=lookup_booking("wedding_date", missing=""),
            group="The document"),
        Col("doc_type", "Document", 18, "text",
            validate=list_from("DocumentTypeList"),
            note="Contract, model release, venue permit, invoice, "
                 "insurance certificate - whatever the wedding needs.",
            group="The document"),
        Col("title", "Title or reference", 26, "text",
            group="The document"),
        Col("required", "Needed", 9, "flag",
            validate=yesno("No for optional paperwork, so it stops "
                           "showing up as missing."),
            group="The document"),

        # ---- where it is up to ----------------------------------------------
        Col("status", "Status", 15, "status",
            validate=list_from("DocumentStatusList"), cf=status_cf(),
            group="Where it is up to"),
        Col("sent_date", "Sent", 11, "date",
            validate=date_check(), group="Where it is up to"),
        Col("signed_date", "Signed or received", 13, "date",
            validate=date_check(), group="Where it is up to"),
        Col("expires", "Expires", 11, "date",
            validate=date_check(),
            note="For anything with a shelf life - a venue permit, a "
                 "certificate of insurance.",
            group="Where it is up to"),
        Col("stored_at", "Where it is kept", 26, "text",
            note="A folder name, a cloud link, a filing cabinet "
                 "drawer. The workbook does not hold the file itself.",
            group="Where it is up to"),
        Col("signers", "Who signs", 18, "text",
            group="Where it is up to"),
        Col("notes", "Notes", 28, "wrap", group="Where it is up to"),

        # ---- the chase --------------------------------------------------------
        date_num_col("wedding_date", "wed_num", "Wedding date number"),
        Col("is_done", "Complete", 9, "flag",
            formula=blank("booking",
                          'IF({signed_date}{r}<>"","Yes",'
                          'IF({status}{r}="Signed","Yes",'
                          'IF({status}{r}="Counter-signed","Yes",'
                          'IF({status}{r}="Filed","Yes",'
                          'IF({status}{r}="Not needed","Yes","No")))))'),
            hidden=True),
        Col("days_out", "Days since sent", 11, "int",
            formula=blank("booking",
                          'IF({sent_date}{r}="","",'
                          'IF({is_done}{r}="Yes","",'
                          'TodayDate-{sent_date}{r}))'),
            group="The chase"),
        Col("turnaround", "Days to sign", 11, "int",
            formula=blank("booking",
                          'IF({sent_date}{r}="","",'
                          'IF({signed_date}{r}="","",'
                          '{signed_date}{r}-{sent_date}{r}))'),
            group="The chase"),
        Col("days_to_wedding", "Days to wedding", 13, "int",
            formula=blank("booking",
                          'IF({wedding_date}{r}="","",'
                          '{wedding_date}{r}-TodayDate)'),
            group="The chase"),
        Col("booking_value", "Booking value", 13, "money",
            formula=zero("booking", lookup_booking_num("total_value")[1:]),
            note="What is riding on this piece of paper.",
            group="The chase"),

        Col("alert", "What to do", 32, "status",
            formula=(
                '=IF({booking}{r}="","",'
                'IF({required}{r}="No","",'
                'IF({is_done}{r}="Yes",'
                'IF({expires}{r}="","OK",'
                'IF({expires}{r}<TodayDate,"Expired - needs renewing",'
                'IF({expires}{r}-TodayDate<=30,'
                '"Due - expires within a month","OK"))),'
                'IF({sent_date}{r}="",'
                'IF({days_to_wedding}{r}<>"",'
                'IF({days_to_wedding}{r}<=30,'
                '"Urgent - not sent and the wedding is close",'
                '"Send - not sent yet"),"Send - not sent yet"),'
                'IF({days_to_wedding}{r}<>"",'
                'IF({days_to_wedding}{r}<=14,'
                '"Urgent - still unsigned, wedding is next",'
                'IF({days_out}{r}>DocumentChaseDays,'
                '"Chase - sent a while ago, nothing back","OK")),'
                'IF({days_out}{r}>DocumentChaseDays,'
                '"Chase - sent a while ago, nothing back","OK"))))))'),
            cf=alert_cf(), group="The chase"),
    ]

    return SheetSpec(
        name=NAME,
        title="Contracts & Documents",
        subtitle=("A register of the paperwork for each wedding - what "
                  "exists, where it is kept and what is still missing"),
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=None,
        freeze_cols=3,
        samples=sample("contracts", cfg),
        help_text=(
            "One row per document per wedding. Record what you sent, "
            "when it came back and where the signed copy lives. The "
            "booking row shows the status of its contract, and this "
            "tab chases anything unsigned as the wedding gets closer."),
        extra_notes=[
            CONTRACT_NOTE,
            "This is a tracking register. It does not draft documents, "
            "it cannot sign anything, and it is not legal advice. Have "
            "your contract written or checked by somebody qualified in "
            "your country - a wedding contract that has never been "
            "read by a lawyer is worth what you paid for it.",
            "The workbook stores a location, not the file. Keep the "
            "signed copies wherever you normally keep them and put the "
            "folder name or link in 'Where it is kept'.",
        ],
    )
````

### `wedding_photography_spreadsheet/sheets/payments.py`

*176 lines · 1 function*

> Payments & Invoices - every instalment, due and received.
>
> Wedding work is paid in pieces over a year or more, which is why a
> photographer can be busy, fully booked and still short of money in
> March.  One row per instalment, joined to the wedding by Booking ID.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Payments & Invoices').

````python
"""Payments & Invoices - every instalment, due and received.

Wedding work is paid in pieces over a year or more, which is why a
photographer can be busy, fully booked and still short of money in
March.  One row per instalment, joined to the wedding by Booking ID.
"""

from __future__ import annotations

from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, booking_client_col, blank, date_num_col, id_col, lookup_booking,
                     month_col, money_cf, owing_cf, status_cf, year_col,
                     zero)

NAME = "Payments & Invoices"


def spec(cfg) -> SheetSpec:
    columns = [
        id_col("booking", "InvoicePrefix", digits=cfg.digits("invoice"),
               header="Invoice ID"),

        Col("booking", "Booking ID", 11, "id",
            validate=list_from("BookingList"),
            note="The wedding this instalment belongs to. The booking "
                 "adds up everything received against its ID.",
            group="The invoice"),
        Col("couple", "Couple", 20, "text",
            formula=lookup_booking("client_name"), group="The invoice"),
        booking_client_col(),
        Col("wedding_date", "Wedding date", 12, "date",
            formula=lookup_booking("wedding_date", missing=""),
            group="The invoice"),
        Col("stage", "Stage", 14, "text",
            validate=list_from("PaymentStageList"),
            note="Deposit, second instalment, balance, an add-on "
                 "bought later, or a refund.",
            group="The invoice"),
        Col("reference", "Your reference", 15, "text",
            note="Your own invoice number if you raise them elsewhere.",
            group="The invoice"),
        Col("description", "What it is for", 26, "wrap",
            group="The invoice"),

        # ---- the money --------------------------------------------------------
        Col("amount", "Amount due", 12, "money",
            validate=number_check(-1000000, 1000000), cf=money_cf(),
            note="Enter a refund as a negative number so the totals "
                 "stay right.",
            group="The money"),
        Col("due_date", "Due", 11, "date",
            validate=date_check(), group="The money"),
        Col("received", "Received", 12, "money",
            validate=number_check(-1000000, 1000000),
            note="What actually arrived. A part payment is fine - "
                 "leave the amount due alone and the balance shows.",
            group="The money"),
        Col("received_date", "Received on", 12, "date",
            validate=date_check(), group="The money"),
        Col("method", "How", 13, "text",
            validate=list_from("PaymentMethodList"), group="The money"),
        Col("fee", "Processing fee", 12, "money",
            validate=number_check(0, 100000),
            note="What the card processor or bank took. Small per "
                 "payment, not small over a year.",
            group="The money"),
        Col("net", "Net to you", 12, "money",
            formula=zero("booking",
                         'ROUND(IF({received}{r}="",0,{received}{r})-'
                         'IF({fee}{r}="",0,{fee}{r}),2)'),
            cf=money_cf(), group="The money"),
        Col("balance", "Still due", 11, "money",
            formula=zero("booking",
                         'ROUND(IF({amount}{r}="",0,{amount}{r})-'
                         'IF({received}{r}="",0,{received}{r}),2)'),
            cf=owing_cf(), group="The money"),

        # ---- where it stands ---------------------------------------------------
        Col("status", "Status", 13, "status",
            formula=blank("booking",
                          'IF({amount}{r}="","Draft",'
                          'IF({balance}{r}<=0.004,"Paid",'
                          'IF(IF({received}{r}="",0,{received}{r})>0,'
                          'IF({due_date}{r}="","Part paid",'
                          'IF(TodayDate>{due_date}{r},"Overdue",'
                          '"Part paid")),'
                          'IF({due_date}{r}="","Unpaid",'
                          'IF(TodayDate>{due_date}{r},"Overdue",'
                          '"Unpaid")))))'),
            cf=status_cf(), group="Where it stands"),
        Col("days_overdue", "Days overdue", 11, "int",
            formula=zero("booking",
                         'IF({due_date}{r}="",0,'
                         'IF({balance}{r}<=0.004,0,'
                         'MAX(0,TodayDate-{due_date}{r})))'),
            group="Where it stands"),
        Col("days_to_pay", "Days taken", 10, "int",
            formula=blank("booking",
                          'IF({due_date}{r}="","",'
                          'IF({received_date}{r}="","",'
                          '{received_date}{r}-{due_date}{r}))'),
            note="Negative means they paid early. Worth knowing who "
                 "does.",
            group="Where it stands"),
        Col("reminders", "Reminders sent", 11, "int",
            validate=number_check(0, 50), group="Where it stands"),
        Col("last_reminder", "Last reminder", 12, "date",
            validate=date_check(), group="Where it stands"),
        Col("notes", "Notes", 26, "wrap", group="Where it stands"),

        # ---- plumbing -----------------------------------------------------------
        month_col("received_date", "month", "Month received"),
        year_col("received_date"),
        date_num_col("due_date", "due_num", "Due date number"),
        date_num_col("received_date", "paid_num", "Paid date number"),
        Col("is_paid", "Paid flag", 9, "flag",
            formula=blank("booking",
                          'IF({balance}{r}<=0.004,"Yes","No")'),
            hidden=True),
        Col("is_late", "Late flag", 9, "flag",
            formula=blank("booking",
                          'IF({days_overdue}{r}>0,"Yes","No")'),
            hidden=True),

        Col("alert", "What to do", 32, "status",
            formula=(
                '=IF({booking}{r}="","",'
                'IF({amount}{r}="","Add - no amount on this invoice",'
                'IF({balance}{r}<=0.004,"OK",'
                'IF({due_date}{r}="","Add - no due date",'
                'IF({days_overdue}{r}>30,'
                '"Urgent - over a month overdue",'
                'IF({days_overdue}{r}>0,'
                'IF({reminders}{r}>=2,'
                '"Urgent - chased twice, still unpaid",'
                '"Overdue - send a reminder"),'
                'IF({due_date}{r}-TodayDate<=7,'
                '"Due - payment due within the week",'
                'IF({wedding_date}{r}<>"",'
                'IF({wedding_date}{r}-TodayDate<=FinalPaymentDays,'
                '"Chase - wedding is close and this is unpaid","OK"),'
                '"OK"))))))))'),
            cf=alert_cf(),
            note="The one column to sort by on a Monday morning.",
            group="Where it stands"),
    ]

    return SheetSpec(
        name=NAME,
        title="Payments & Invoices",
        subtitle=("Every instalment for every wedding - what is due, "
                  "what has arrived and what is late"),
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=None,
        freeze_cols=3,
        samples=sample("payments", cfg),
        help_text=(
            "One row per instalment. A typical wedding has three: a "
            "deposit to hold the date, a second payment some months "
            "out and the balance before the day. Put the Booking ID on "
            "every row and the wedding adds them up for you."),
        extra_notes=[
            "The booking row works out what the deposit and the "
            "balance should be from your rules on Settings. This tab "
            "records what you actually invoiced and what actually "
            "arrived - the two are allowed to differ, and the "
            "difference is the point.",
            "Record processing fees. Two per cent on a $2,500 wedding "
            "is $50, and across twenty-four weddings it is a lens.",
            "Enter refunds and chargebacks as negative amounts rather "
            "than deleting the original row, so the history survives.",
        ],
    )
````

---

## 7. The wedding day itself

Workflow three, and the part most spreadsheets skip. The timeline is built from offsets against the ceremony time, so moving the ceremony moves the day. The shot list holds the must-haves and the family groups, and the questionnaire prints them for the couple to confirm. The calendar shows twelve months of squares and turns red when two couples want the same Saturday.

### `wedding_photography_spreadsheet/sheets/timeline.py`

*215 lines · 1 function*

> Wedding Timeline - the running order of a wedding day, hour by hour.
>
> Pick the event and the sheet knows roughly when it happens relative to
> the ceremony, so a full day's timeline is a column of drop-downs rather
> than an hour of arithmetic.  Every default can be overridden, because
> no two weddings run the same.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Wedding Timeline').

````python
"""Wedding Timeline - the running order of a wedding day, hour by hour.

Pick the event and the sheet knows roughly when it happens relative to
the ceremony, so a full day's timeline is a column of drop-downs rather
than an hour of arithmetic.  Every default can be overridden, because
no two weddings run the same.
"""

from __future__ import annotations

from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check, time_check
from .common import (alert_cf, blank, date_num_col, id_col, lookup_booking,
                     lookup_staff, status_cf, yesno, zero)

NAME = "Wedding Timeline"


def spec(cfg) -> SheetSpec:
    columns = [
        id_col("booking", "TimelinePrefix", digits=cfg.digits("timeline"),
               header="Line ID"),

        Col("booking", "Booking ID", 11, "id",
            validate=list_from("BookingList"),
            note="Which wedding this line belongs to. Add one row per "
                 "event - the day plan for a wedding is all the rows "
                 "sharing a Booking ID.",
            group="The wedding"),
        Col("couple", "Couple", 18, "text",
            formula=lookup_booking("client_name"), group="The wedding"),
        Col("wedding_date", "Wedding date", 12, "date",
            formula=lookup_booking("wedding_date", missing=""),
            group="The wedding"),
        Col("ceremony", "Ceremony at", 11, "time",
            formula=lookup_booking("ceremony_time", missing=""),
            note="Taken from the booking. Change it there and every "
                 "line of the day moves with it.",
            group="The wedding"),

        # ---- the event --------------------------------------------------------
        Col("event", "Event", 24, "text",
            validate=list_from("TimelineEventList"),
            note="Choosing a standard event fills in a sensible time "
                 "and length. Type over either if this wedding is "
                 "different.",
            group="The event"),
        Col("label", "Your wording", 24, "text",
            note="Optional. What you want printed on the day sheet if "
                 "'Reception entrance' is not what they call it.",
            group="The event"),
        Col("shown", "Shows as", 24, "text",
            formula=blank("booking",
                          'IF({label}{r}<>"",{label}{r},{event}{r})'),
            group="The event"),
        Col("offset_default", "Standard offset", 12, "int",
            formula=zero("booking",
                         'IFERROR(INDEX(TimelineOffsetList,'
                         'MATCH({event}{r},TimelineEventList,0))*1,0)'),
            hidden=True, group="The event"),
        Col("offset", "Minutes from ceremony", 15, "int",
            validate=number_check(-900, 900),
            note="Negative is before the ceremony, positive is after. "
                 "Leave blank to use the standard offset.",
            group="The event"),
        Col("offset_used", "Offset used", 11, "int",
            formula=zero("booking",
                         'IF({offset}{r}="",{offset_default}{r},'
                         '{offset}{r})'),
            hidden=True, group="The event"),
        Col("time_override", "Fixed time", 10, "time",
            validate=time_check(),
            note="For anything with a time of its own - a church slot, "
                 "a sunset, the venue's curfew.",
            group="The event"),
        Col("start", "Starts", 10, "time",
            formula=blank("booking",
                          'IF({time_override}{r}<>"",{time_override}{r},'
                          'IF({ceremony}{r}="","",'
                          'MOD({ceremony}{r}+{offset_used}{r}/1440,1)))'),
            note="Worked out from the ceremony time unless you have "
                 "fixed it.",
            group="The event"),
        Col("duration_default", "Standard length", 12, "int",
            formula=zero("booking",
                         'IFERROR(INDEX(TimelineDurationList,'
                         'MATCH({event}{r},TimelineEventList,0))*1,0)'),
            hidden=True, group="The event"),
        Col("duration", "Minutes", 9, "int",
            validate=number_check(0, 900), group="The event"),
        Col("duration_used", "Length used", 10, "int",
            formula=zero("booking",
                         'IF({duration}{r}="",{duration_default}{r},'
                         '{duration}{r})'),
            hidden=True, group="The event"),
        Col("finish", "Ends", 10, "time",
            formula=blank("booking",
                          'IF({start}{r}="","",'
                          'MOD({start}{r}+{duration_used}{r}/1440,1))'),
            group="The event"),

        # ---- who and where -----------------------------------------------------
        Col("location", "Where", 20, "text", group="Who and where"),
        Col("role", "Who covers it", 16, "text",
            formula=blank("booking",
                          'IFERROR(INDEX(TimelineRoleList,'
                          'MATCH({event}{r},TimelineEventList,0))&"",'
                          '"Lead photographer")'),
            group="Who and where"),
        Col("staff", "Assigned to", 12, "id",
            validate=list_from("StaffList"), group="Who and where"),
        Col("staff_name", "Name", 16, "text",
            formula=lookup_staff("name", key="staff", missing=""),
            group="Who and where"),
        Col("critical", "Must not miss", 11, "flag",
            validate=yesno("The moments there is no second take of - "
                           "the vows, the first kiss, the speeches."),
            group="Who and where"),
        Col("notes", "Notes", 28, "wrap", group="Who and where"),

        # ---- on the day ---------------------------------------------------------
        Col("done", "Done", 8, "flag",
            validate=yesno("Tick it off on the day, or afterwards "
                           "when you are checking nothing was missed."),
            group="On the day"),
        Col("actual", "Actually at", 11, "time",
            validate=time_check(), group="On the day"),
        Col("drift", "Minutes late", 11, "int",
            formula=blank("booking",
                          'IF({actual}{r}="","",IF({start}{r}="","",'
                          'ROUND(MOD({actual}{r}-{start}{r}+0.5,1)'
                          '*1440-720,0)))'),
            note="How far off the plan the day ran. Positive is late. "
                 "Useful next time you are asked how long the family "
                 "photographs really take.",
            group="On the day"),
        Col("status", "Status", 13, "status",
            formula=blank("booking",
                          'IF({done}{r}="Yes","Done",'
                          'IF({wedding_date}{r}="","Planned",'
                          'IF({wedding_date}{r}<TodayDate,"Missed",'
                          '"Planned")))'),
            cf=status_cf(), group="On the day"),

        # ---- plumbing ------------------------------------------------------------
        date_num_col("wedding_date", "wed_num", "Wedding date number"),
        Col("start_num", "Start number", 10, "number",
            formula=zero("booking", 'IF({start}{r}="",0,{start}{r})'),
            hidden=True),
        Col("sort_key", "Sort key", 12, "number",
            formula=zero("booking",
                         'ROUND({start_num}{r}*100000,0)+ROW()-4'),
            hidden=True),
        Col("day_order", "Order", 8, "int",
            formula=blank("booking",
                          'COUNTIFS([[Wedding Timeline.booking]],'
                          '{booking}{r},[[Wedding Timeline.sort_key]],'
                          '"<"&{sort_key}{r})+1'),
            note="Where this line falls in the running order for its "
                 "wedding. The printable day sheet reads it.",
            group="On the day"),
        Col("day_key", "Day key", 14, "text",
            formula=blank("booking",
                          '{booking}{r}&"|"&{day_order}{r}'),
            hidden=True),

        Col("alert", "What to do", 30, "status",
            formula=(
                '=IF({booking}{r}="","",'
                'IF({done}{r}="Yes","",'
                'IF({event}{r}="","Add - no event chosen",'
                'IF({start}{r}="","Check - no time worked out",'
                'IF({wedding_date}{r}="","Check - booking has no date",'
                'IF({wedding_date}{r}<TodayDate,'
                'IF({critical}{r}="Yes",'
                '"Check - a must-not-miss moment is unticked",'
                '"Check - not ticked off"),'
                'IF({wedding_date}{r}-TodayDate<=7,'
                'IF({staff}{r}="","Add - nobody assigned yet","OK"),'
                '"OK")))))))'),
            cf=alert_cf(), group="On the day"),
    ]

    return SheetSpec(
        name=NAME,
        title="Wedding Timeline",
        subtitle=("The running order of each wedding day, built from "
                  "the ceremony time"),
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=None,
        freeze_cols=3,
        samples=sample("timeline", cfg),
        help_text=(
            "Add one row per event and pick the Booking ID it belongs "
            "to. Choosing a standard event sets a time relative to the "
            "ceremony and a sensible length, so a twenty-line day plan "
            "takes a couple of minutes. Override anything that differs "
            "and the rest of the day keeps up."),
        extra_notes=[
            "Change the ceremony time on the booking and every "
            "calculated line on this tab moves with it. This is the "
            "reason to let the sheet work the times out rather than "
            "typing them: ceremony times move, and they move late.",
            "'Must not miss' is for the moments with no second take. "
            "Anything flagged and left unticked after the wedding "
            "comes back as an alert, which is a much better way to "
            "find out than an email from the couple.",
            "'Minutes late' builds up real numbers for how long each "
            "part of a day actually takes. After a dozen weddings you "
            "will know whether your forty-minute family slot has ever "
            "once taken forty minutes.",
        ],
    )
````

### `wedding_photography_spreadsheet/sheets/shots.py`

*202 lines · 1 function*

> Shot List Manager - the list you work from, grouped and ticked off.
>
> The family-group section is the one that saves a wedding.  Written out
> by name beforehand it takes twenty minutes; improvised on the day with
> a hundred guests watching, it takes an hour and somebody's grandmother
> gets left out of it.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Shot List Manager').

````python
"""Shot List Manager - the list you work from, grouped and ticked off.

The family-group section is the one that saves a wedding.  Written out
by name beforehand it takes twenty minutes; improvised on the day with
a hundred guests watching, it takes an hour and somebody's grandmother
gets left out of it.
"""

from __future__ import annotations

from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, time_check
from .common import (alert_cf, blank, date_num_col, id_col, lookup_booking,
                     priority_cf, status_cf, yesno)

NAME = "Shot List Manager"


def spec(cfg) -> SheetSpec:
    columns = [
        id_col("booking", "ShotPrefix", digits=cfg.digits("shot"),
               header="Shot ID"),

        Col("booking", "Booking ID", 11, "id",
            validate=list_from("BookingList"),
            note="The wedding this shot belongs to. Copy a standard "
                 "list down, change the Booking ID and you have a shot "
                 "list for the next wedding.",
            group="The wedding"),
        Col("couple", "Couple", 18, "text",
            formula=lookup_booking("client_name"), group="The wedding"),
        Col("wedding_date", "Wedding date", 12, "date",
            formula=lookup_booking("wedding_date", missing=""),
            group="The wedding"),

        # ---- the shot ----------------------------------------------------------
        Col("group_name", "Group", 16, "text",
            validate=list_from("ShotGroupList"),
            note="Details, preparation, couple, family, ceremony, "
                 "reception. Grouping is what makes the list usable "
                 "while you are holding a camera.",
            group="The shot"),
        Col("shot", "Shot", 32, "text",
            note="Be specific. 'Family' is not a shot; 'Bride with "
                 "her mother and both sisters' is.",
            group="The shot"),
        Col("people", "Who is in it", 26, "wrap",
            note="Names, not relationships. On the day you will be "
                 "calling these out across a room.",
            group="The shot"),
        Col("priority", "Priority", 11, "text",
            validate=list_from("PriorityList"), cf=priority_cf(),
            group="The shot"),
        Col("must_have", "Must have", 10, "flag",
            validate=yesno("The shots the couple specifically asked "
                           "for. These are chased separately."),
            group="The shot"),
        Col("requested_by", "Asked for by", 14, "text",
            validate=list_from("RequestSourceList"),
            note="Whether this came from the couple, a parent or your "
                 "own standard list. Useful when there are forty "
                 "family groups and somebody has to be told no.",
            group="The shot"),
        Col("when", "When in the day", 15, "text",
            validate=list_from("ShotWhenList"), group="The shot"),
        Col("location", "Where", 18, "text", group="The shot"),
        Col("notes", "Notes", 26, "wrap", group="The shot"),

        # ---- on the day ----------------------------------------------------------
        Col("done", "Taken", 8, "flag",
            validate=yesno(), cf=status_cf(), group="On the day"),
        Col("taken_at", "Taken at", 10, "time",
            validate=time_check(), group="On the day"),
        Col("skipped", "Skipped", 9, "flag",
            validate=yesno("Yes when it could not happen - the light "
                           "went, the person left, the rain won. Stops "
                           "it being chased as missing."),
            group="On the day"),
        Col("skip_reason", "Why not", 22, "wrap", group="On the day"),
        Col("delivered", "In the gallery", 11, "flag",
            validate=yesno(), group="On the day"),

        # ---- plumbing -------------------------------------------------------------
        date_num_col("wedding_date", "wed_num", "Wedding date number"),
        Col("open", "Outstanding", 9, "flag",
            formula=blank("booking",
                          'IF({done}{r}="Yes","No",'
                          'IF({skipped}{r}="Yes","No","Yes"))'),
            hidden=True),
        Col("print_flag", "On the printed list", 9, "int",
            formula=('=IF({booking}{r}="",0,'
                     'IF({group_name}{r}="Family",1,'
                     'IF({must_have}{r}="Yes",1,0)))'),
            hidden=True),
        Col("print_seq", "Printed line", 9, "int",
            formula=('=IF({booking}{r}="","",'
                     'IF({print_flag}{r}=0,"",'
                     'COUNTIFS({booking_first}:{booking}{r},'
                     '{booking}{r},{print_flag_first}:'
                     '{print_flag}{r},1)))'),
            hidden=True),
        Col("print_key", "Printed key", 14, "text",
            formula=('=IF({print_seq}{r}="","",'
                     '{booking}{r}&"|"&{print_seq}{r})'),
            hidden=True,
            note="Family groups and must-have shots, numbered within "
                 "each wedding. The Client Questionnaire prints them "
                 "in this order."),
        Col("group_total", "In this group", 11, "int",
            formula=blank("booking",
                          'COUNTIFS([[Shot List Manager.booking]],'
                          '{booking}{r},'
                          '[[Shot List Manager.group_name]],'
                          '{group_name}{r})'),
            group="How it is going"),
        Col("group_done", "Group taken", 11, "int",
            formula=blank("booking",
                          'COUNTIFS([[Shot List Manager.booking]],'
                          '{booking}{r},'
                          '[[Shot List Manager.group_name]],'
                          '{group_name}{r},'
                          '[[Shot List Manager.done]],"Yes")'),
            group="How it is going"),
        Col("group_pct", "Group complete", 12, "percent0",
            formula=blank("booking",
                          'IF({group_total}{r}<=0,"",'
                          '{group_done}{r}/{group_total}{r})'),
            group="How it is going"),
        Col("wedding_pct", "Wedding complete", 13, "percent0",
            formula=blank("booking",
                          'IF(COUNTIFS([[Shot List Manager.booking]],'
                          '{booking}{r})<=0,"",'
                          'COUNTIFS([[Shot List Manager.booking]],'
                          '{booking}{r},[[Shot List Manager.done]],'
                          '"Yes")/'
                          'COUNTIFS([[Shot List Manager.booking]],'
                          '{booking}{r}))'),
            group="How it is going"),
        Col("status", "Status", 12, "status",
            formula=blank("booking",
                          'IF({done}{r}="Yes","Taken",'
                          'IF({skipped}{r}="Yes","Skipped",'
                          'IF({wedding_date}{r}="","To do",'
                          'IF({wedding_date}{r}<TodayDate,"Missed",'
                          '"To do"))))'),
            cf=status_cf(), group="How it is going"),

        Col("alert", "What to do", 30, "status",
            formula=(
                '=IF({booking}{r}="","",'
                'IF({shot}{r}="","Add - no shot described",'
                'IF({done}{r}="Yes",'
                'IF({delivered}{r}="No",'
                '"Check - taken but not in the gallery","OK"),'
                'IF({skipped}{r}="Yes",'
                'IF({skip_reason}{r}="","Add - say why it was missed",'
                '"OK"),'
                'IF({wedding_date}{r}="","OK",'
                'IF({wedding_date}{r}<TodayDate,'
                'IF({must_have}{r}="Yes",'
                '"Urgent - a must-have shot was never taken",'
                '"Check - not taken and not skipped"),'
                'IF({wedding_date}{r}-TodayDate<=7,'
                'IF({people}{r}="",'
                'IF({group_name}{r}="Family",'
                '"Add - family shot with no names","OK"),"OK"),'
                '"OK")))))))'),
            cf=alert_cf(), group="How it is going"),
    ]

    return SheetSpec(
        name=NAME,
        title="Shot List Manager",
        subtitle=("Every photograph you have promised, grouped the way "
                  "you will shoot it and ticked off as you go"),
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=None,
        freeze_cols=3,
        samples=sample("shots", cfg),
        help_text=(
            "One row per shot, tagged with the Booking ID. Build a "
            "standard list once, then copy it down and change the "
            "Booking ID for each new wedding. The booking row shows "
            "how much of the list is done, and anything flagged as a "
            "must-have but never taken comes back as an alert."),
        extra_notes=[
            "Write family groups out by name. A list that says "
            "'Bride's family' turns into fifteen minutes of asking who "
            "else should be in it while two hundred guests wait for "
            "their dinner.",
            "'Asked for by' is worth filling in. When the groom's "
            "mother produces a list of nineteen extra groups on the "
            "morning, it helps to know which of the forty you already "
            "have came from the couple themselves.",
            "'Skipped' with a reason is not a failure - the light "
            "goes, people leave, it rains. Recording it stops the shot "
            "being chased and gives you something to say if you are "
            "ever asked about it.",
        ],
    )
````

### `wedding_photography_spreadsheet/sheets/staff.py`

*222 lines · 1 function*

> Photographer & Staff - who shoots, who edits, and what they cost.
>
> Small studios and solo photographers with a regular second shooter both
> need the same thing: a rate on file, a count of who is working which
> weekend, and an honest figure for what the team costs against what the
> weddings bring in.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Photographer & Staff').

````python
"""Photographer & Staff - who shoots, who edits, and what they cost.

Small studios and solo photographers with a regular second shooter both
need the same thing: a rate on file, a count of who is working which
weekend, and an honest figure for what the team costs against what the
weddings bring in.
"""

from __future__ import annotations

from ..layout import Col, SheetSpec, min_if
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, id_col, money_cf, rank_desc,
                     rating_cf, status_cf, yesno, zero)

NAME = "Photographer & Staff"


def spec(cfg) -> SheetSpec:
    columns = [
        id_col("name", "StaffPrefix", digits=cfg.digits("staff"),
               header="Staff ID",
               note="Automatic. Bookings, the timeline and the editing "
                    "workflow all assign work by this ID."),

        Col("name", "Name", 20, "id", group="Who"),
        Col("role", "Role", 15, "text",
            validate=list_from("StaffRoleList"), group="Who"),
        Col("status", "Status", 12, "status",
            validate=list_from("StaffStatusList"), cf=status_cf(),
            group="Who"),
        Col("type", "Employment", 13, "text",
            validate=list_from("StaffTypeList"),
            note="Whether somebody is an employee or a contractor is a "
                 "legal question in your country, not a spreadsheet "
                 "one. This column records your answer, it does not "
                 "decide it.",
            group="Who"),
        Col("phone", "Phone", 15, "text", group="Who"),
        Col("email", "Email", 24, "text", group="Who"),
        Col("started", "Started", 11, "date",
            validate=date_check(), group="Who"),

        # ---- what they cost -------------------------------------------------
        Col("day_rate", "Wedding day rate", 13, "money",
            validate=number_check(0, 100000),
            note="What you pay them for a full wedding. Leave blank "
                 "for yourself - your pay is the profit.",
            group="What they cost"),
        Col("hour_rate", "Hourly rate", 11, "money",
            validate=number_check(0, 10000), group="What they cost"),
        Col("edit_rate", "Editing rate", 11, "money",
            validate=number_check(0, 10000), group="What they cost"),
        Col("is_owner", "Owner", 8, "flag",
            validate=yesno("Yes for you. Owner time is not counted as "
                           "a cost, because paying yourself a wage out "
                           "of your own profit counts it twice."),
            group="What they cost"),

        # ---- what they can do -----------------------------------------------
        Col("can_lead", "Can lead", 9, "flag",
            validate=yesno(), group="What they can do"),
        Col("can_second", "Can second", 10, "flag",
            validate=yesno(), group="What they can do"),
        Col("can_edit", "Can edit", 9, "flag",
            validate=yesno(), group="What they can do"),
        Col("specialism", "Specialism", 18, "text",
            group="What they can do"),
        Col("kit", "Own kit", 20, "wrap", group="What they can do"),
        Col("insured", "Insured", 9, "flag",
            validate=yesno("Their own public liability cover, if they "
                           "carry it. Yours is on the Expense Tracker."),
            group="What they can do"),
        Col("insurance_expiry", "Insurance to", 12, "date",
            validate=date_check(), group="What they can do"),

        # ---- how much they work ---------------------------------------------
        Col("leading", "Leading", 9, "int",
            formula=zero("name",
                         'COUNTIFS([[Wedding Bookings.photographer]],'
                         '{id}{r},[[Wedding Bookings.holds]],"Yes")'),
            group="How much they work"),
        Col("seconding", "Seconding", 10, "int",
            formula=zero("name",
                         'COUNTIFS([[Wedding Bookings.second]],{id}{r},'
                         '[[Wedding Bookings.holds]],"Yes")'),
            group="How much they work"),
        Col("weddings", "Weddings", 9, "int",
            formula=zero("name", '{leading}{r}+{seconding}{r}'),
            group="How much they work"),
        Col("upcoming", "Still to come", 11, "int",
            formula=zero("name",
                         'COUNTIFS([[Wedding Bookings.photographer]],'
                         '{id}{r},[[Wedding Bookings.holds]],"Yes",'
                         '[[Wedding Bookings.wed_num]],">="&TodayDate)+'
                         'COUNTIFS([[Wedding Bookings.second]],{id}{r},'
                         '[[Wedding Bookings.holds]],"Yes",'
                         '[[Wedding Bookings.wed_num]],">="&TodayDate)'),
            group="How much they work"),
        Col("next_num", "Next wedding number", 12, "number",
            formula=zero("name", min_if(
                "[[Wedding Bookings.wed_num]]",
                '([[Wedding Bookings.photographer]]={id}{r})'
                '+([[Wedding Bookings.second]]={id}{r})>0',
                '[[Wedding Bookings.holds]]="Yes"',
                '[[Wedding Bookings.wed_num]]>=TodayDate')),
            hidden=True, group="How much they work"),
        Col("next_wedding", "Next wedding", 12, "date",
            formula=blank("name",
                          'IF({next_num}{r}>=2958465,"",'
                          'IF({next_num}{r}<=0,"",{next_num}{r}))'),
            group="How much they work"),
        Col("shoot_hours", "Hours on the day", 13, "number1",
            formula=zero("name",
                         'SUMIFS([[Wedding Bookings.coverage_hours]],'
                         '[[Wedding Bookings.photographer]],{id}{r},'
                         '[[Wedding Bookings.holds]],"Yes")'),
            group="How much they work"),
        Col("edit_jobs", "Editing jobs", 11, "int",
            formula=zero("name",
                         'COUNTIFS([[Editing Workflow.editor]],'
                         '{id}{r})'),
            group="How much they work"),
        Col("edit_hours", "Editing hours", 12, "number1",
            formula=zero("name",
                         'SUMIFS([[Editing Workflow.hours]],'
                         '[[Editing Workflow.editor]],{id}{r})'),
            group="How much they work"),

        # ---- what it comes to -------------------------------------------------
        Col("shoot_pay", "Shoot pay", 12, "money",
            formula=zero("name",
                         'IF({is_owner}{r}="Yes",0,'
                         'ROUND({weddings}{r}*'
                         'IF({day_rate}{r}="",0,{day_rate}{r}),2))'),
            group="What it comes to"),
        Col("edit_pay", "Editing pay", 11, "money",
            formula=zero("name",
                         'IF({is_owner}{r}="Yes",0,'
                         'ROUND({edit_hours}{r}*'
                         'IF({edit_rate}{r}="",0,{edit_rate}{r}),2))'),
            group="What it comes to"),
        Col("total_pay", "Total pay", 12, "money",
            formula=zero("name", '{shoot_pay}{r}+{edit_pay}{r}'),
            cf=money_cf(), group="What it comes to"),
        Col("revenue_touched", "Revenue worked on", 14, "money",
            formula=zero("name",
                         'SUMIFS([[Wedding Bookings.total_value]],'
                         '[[Wedding Bookings.photographer]],{id}{r},'
                         '[[Wedding Bookings.earns]],"Yes")'),
            cf=money_cf(),
            note="Only weddings they led. A second shooter has not "
                 "brought the booking in.",
            group="What it comes to"),
        Col("pay_ratio", "Pay as share", 11, "percent0",
            formula=blank("name",
                          'IF({revenue_touched}{r}<=0,"",'
                          '{total_pay}{r}/{revenue_touched}{r})'),
            cf=bar_cf(), group="What it comes to"),
        rank_desc("weddings", header="Busiest"),

        # ---- how they are doing -----------------------------------------------
        Col("rating", "Average review", 11, "rating",
            formula=blank("name",
                          'IFERROR(AVERAGEIFS('
                          '[[Reviews & Referrals.rating]],'
                          '[[Reviews & Referrals.photographer]],'
                          '{id}{r}),"")'),
            cf=rating_cf(), group="How they are doing"),
        Col("late_jobs", "Late deliveries", 11, "int",
            formula=zero("name",
                         'COUNTIFS([[Editing Workflow.editor]],{id}{r},'
                         '[[Editing Workflow.late]],"Yes")'),
            group="How they are doing"),
        Col("notes", "Notes", 28, "wrap", group="How they are doing"),

        Col("alert", "What to do", 30, "status",
            formula=(
                '=IF({name}{r}="","",'
                'IF({status}{r}<>"Active","",'
                'IF({insured}{r}="Yes",'
                'IF({insurance_expiry}{r}<>"",'
                'IF({insurance_expiry}{r}<TodayDate,'
                '"Expired - their insurance has run out",'
                'IF({insurance_expiry}{r}-TodayDate<=30,'
                '"Due - insurance expires within a month","OK")),'
                '"Check - insured but no expiry date"),'
                'IF({late_jobs}{r}>=2,'
                '"Check - more than one late delivery",'
                'IF({day_rate}{r}="",'
                'IF({is_owner}{r}="Yes","OK",'
                '"Add - no day rate on file"),"OK")))))'),
            cf=alert_cf(), group="How they are doing"),
    ]

    return SheetSpec(
        name=NAME,
        title="Photographer & Staff",
        subtitle=("Everybody who shoots or edits for you, what they "
                  "cost and how much they are working"),
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=None,
        freeze_cols=2,
        samples=sample("staff", cfg),
        help_text=(
            "Add yourself as the first row with 'Owner' set to Yes. "
            "Bookings pick a lead photographer and an optional second "
            "shooter from this list, the editing workflow picks an "
            "editor, and the pay columns add up what the team costs "
            "against the weddings they worked on."),
        extra_notes=[
            "Owner rows are costed at zero on purpose. Your time is "
            "paid out of the profit, and counting a notional wage as a "
            "cost as well would understate every wedding twice over.",
            "Whether somebody is an employee or a self-employed "
            "contractor is a legal test in your country, and getting "
            "it wrong is expensive. The 'Employment' column records "
            "what you have decided; it does not decide it for you.",
        ],
    )
````

### `wedding_photography_spreadsheet/sheets/calendar.py`

*186 lines · 1 function*

> `Booking Calendar` - twelve months of dates, at a glance.
>
> A wedding diary is mostly a question of which dates have gone.  This
> tab draws a real calendar for any twelve months you choose and colours
> every day you are holding, so "are you free on the 14th of June?" is a
> glance rather than a scroll through two hundred rows.
>
> Weeks start on Saturday, because that is the day most weddings happen.
>
> It is drawn by formula from the Wedding Bookings tab.  It does not
> sync with a phone, a website or an enquiry form: a date is taken here
> only because you typed a booking there.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Booking Calendar'), `LAST_COL` (24), `MONTHS` (12), `WEEK_ROWS` (6), `PER_ROW` (3), `BLOCK` (8), `WEEKDAYS` (7 items), `HOLDS` ('[[Wedding Bookings.holds]],"Yes"'), `WED_NUM` ('[[Wedding Bookings.wed_num]]').

````python
"""`Booking Calendar` - twelve months of dates, at a glance.

A wedding diary is mostly a question of which dates have gone.  This
tab draws a real calendar for any twelve months you choose and colours
every day you are holding, so "are you free on the 14th of June?" is a
glance rather than a scroll through two hundred rows.

Weeks start on Saturday, because that is the day most weddings happen.

It is drawn by formula from the Wedding Bookings tab.  It does not
sync with a phone, a website or an enquiry form: a date is taken here
only because you typed a booking there.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, col_letter
from ..writer import SheetWriter, cf_formula, date_check, number_check

NAME = "Booking Calendar"
LAST_COL = 24
MONTHS = 12
WEEK_ROWS = 6
PER_ROW = 3                      # month blocks across the page
BLOCK = 8                        # 7 day columns + 1 spacer

#: Saturday first - a wedding photographer's week
WEEKDAYS = ["Sat", "Sun", "Mon", "Tue", "Wed", "Thu", "Fri"]

HOLDS = '[[Wedding Bookings.holds]],"Yes"'
WED_NUM = "[[Wedding Bookings.wed_num]]"


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.accent)

    s.banner("Booking Calendar",
             "Twelve months of dates with every wedding you are "
             "holding marked. Change the starting month and the whole "
             "year redraws.",
             LAST_COL)
    s.set_col(0, 0, 3)
    for column in range(1, LAST_COL + 1):
        s.set_col(column, column, 6)

    row = 4
    row = s.note(
        row,
        "Weeks begin on Saturday. A day counts as taken when a "
        "booking on that date has a status that holds the diary - "
        "pencilled in counts, cancelled does not. Two live weddings "
        "on one day turn red here and are flagged on the booking too.",
        LAST_COL, height=34)
    row += 1

    row = s.field(row, "Calendar starts", cfg.today.replace(day=1),
                  style="date", name="CalendarStart",
                  validate=date_check(),
                  help_text="Type the first of any month. The twelve "
                            "months below follow on from it.",
                  help_col=8)
    row = s.field(row, "Months to draw", MONTHS, style="int",
                  name="CalendarMonths", validate=number_check(1, 12),
                  help_text="Up to twelve.", help_col=8)
    row += 1

    # ---- legend -----------------------------------------------------------
    s.write(row, 1, "Free", fmts.chrome("legend_input"))
    s.merge(row, 2, 3, "Held", fmts.cf("good"))
    s.merge(row, 4, 6, "Two weddings", fmts.cf("bad"))
    s.merge(row, 7, LAST_COL,
            "Hold a date by giving the booking a status of Pencilled "
            "in or better. Names are in the list underneath.",
            fmts.chrome("note"))
    row += 2

    # ---- twelve month blocks -----------------------------------------------
    grid_top = row
    for index in range(MONTHS):
        band, slot = divmod(index, PER_ROW)
        column = 1 + slot * BLOCK
        top = grid_top + band * (WEEK_ROWS + 5)

        first_of_month = (f'DATE(YEAR(CalendarStart),'
                          f'MONTH(CalendarStart)+{index},1)')
        shown = f'{index + 1}<=CalendarMonths'

        s.merge(top, column, column + 6, None, fmts.chrome("subsection"))
        s.formula(top, column,
                  f'=IF(NOT({shown}),"",'
                  f'TEXT({first_of_month},"mmmm yyyy"))',
                  fmts.chrome("subsection"))

        s.merge(top + 1, column, column + 6, None, fmts.chrome("note"))
        s.formula(
            top + 1, column,
            f'=IF(NOT({shown}),"",'
            f'COUNTIFS({WED_NUM},">="&{first_of_month},'
            f'{WED_NUM},"<="&EOMONTH({first_of_month},0),{HOLDS})'
            f'&" booked")',
            fmts.chrome("note"))

        head = top + 2
        for offset, label in enumerate(WEEKDAYS):
            s.write(head, column + offset, label, fmts.chrome("header"))

        # the Saturday on or before the 1st (serial 6 is a Saturday)
        origin = (f'({first_of_month}-MOD({first_of_month}-6,7))')
        for week in range(WEEK_ROWS):
            line = head + 1 + week
            for offset in range(7):
                serial = f"({origin}+{week * 7 + offset})"
                s.formula(
                    line, column + offset,
                    f'=IF(NOT({shown}),"",'
                    f'IF(MONTH({serial})<>MONTH({first_of_month}),"",'
                    f'DAY({serial})))',
                    fmts.cell("int", "calc"))

        # one pair of conditional rules per block, anchored on its own
        # top-left cell so the relative reference walks the grid
        anchor = f"{col_letter(column)}{head + 2}"
        rng = (f"{anchor}:{col_letter(column + 6)}"
               f"{head + WEEK_ROWS + 1}")
        serial_of_cell = (f'DATE(YEAR(CalendarStart),'
                          f'MONTH(CalendarStart)+{index},{anchor})')
        clash = reg.render(
            f'AND(ISNUMBER({anchor}),'
            f'COUNTIFS({WED_NUM},{serial_of_cell},{HOLDS})>1)')
        held = reg.render(
            f'AND(ISNUMBER({anchor}),'
            f'COUNTIFS({WED_NUM},{serial_of_cell},{HOLDS})>0)')
        s.cf(rng, cf_formula(clash, "bad"))
        s.cf(rng, cf_formula(held, "good"))

    bands = (MONTHS + PER_ROW - 1) // PER_ROW
    row = grid_top + bands * (WEEK_ROWS + 5) + 1

    # ---- the same thing as a list ------------------------------------------
    row = s.section(row, "Every date you are holding", LAST_COL)
    row = s.note(
        row,
        "The diary as a list, soonest first, straight from the "
        "Wedding Bookings tab. This is the one to read down the phone "
        "when somebody asks about a date.",
        LAST_COL, height=28)
    row += 1

    head = row
    s.table_head(head, 1, ["#", "Date", "Day", "Couple", "Status",
                           "Venue", "Photographer", "Ceremony",
                           "Hours", "Value", "Days away",
                           "What to do"],
                 widths=[4, 12, 7, 20, 13, 18, 15, 9, 8, 11, 10, 28])
    row = head + 1
    shown_rows = max(12, cfg.scale_rows(60))
    for n in range(1, shown_rows + 1):
        rank = f'MATCH({n},[[Wedding Bookings.upcoming_rank]],0)'
        s.write(row, 1, n, fmts.chrome("day_number"))
        for col, (field, style) in enumerate([
                ("wedding_date", "date"), ("weekday", "text"),
                ("client_name", "text"), ("status", "status"),
                ("venue", "text"), ("photographer_name", "text"),
                ("ceremony_time", "time"),
                ("coverage_hours", "number1"),
                ("total_value", "money"), ("days_away", "int"),
                ("alert", "text")], start=2):
            s.formula(row, col,
                      f'=IFERROR(INDEX([[Wedding Bookings.{field}]],'
                      f'{rank}),"")',
                      fmts.cell(style, "calc"))
        row += 1
    row += 1

    row = s.note(
        row,
        "This calendar is drawn by formula. It does not connect to "
        "Google Calendar, an enquiry form or a booking platform, and "
        "nothing you do here changes a booking - edit the wedding on "
        "the Wedding Bookings tab and this redraws itself.",
        LAST_COL, height=32)

    s.finish(freeze_row=4, gridlines=False, fit_width=False)
    s.protect()
    return s
````

### `wedding_photography_spreadsheet/sheets/questionnaire.py`

*228 lines · 2 functions*

> `Client Questionnaire` - the printable form, and the answers.
>
> Pick a Booking ID at the top and the whole sheet fills in from what
> you already know, leaving only the questions that need the couple's
> answers.  Print it, email it, or work through it on a call.
>
> The answers live on the booking, not here, so whatever is written down
> travels with the wedding to the shot list, the timeline and the day
> sheet.  This tab is the form, not the filing cabinet.

| Defines | What it does |
|---------|--------------|
| `_b(field)` | A field from the chosen booking |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Client Questionnaire'), `LAST_COL` (9), `SHOT_ROWS` (22), `TIMELINE_ROWS` (20).

````python
"""`Client Questionnaire` - the printable form, and the answers.

Pick a Booking ID at the top and the whole sheet fills in from what
you already know, leaving only the questions that need the couple's
answers.  Print it, email it, or work through it on a call.

The answers live on the booking, not here, so whatever is written down
travels with the wedding to the shot list, the timeline and the day
sheet.  This tab is the form, not the filing cabinet.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry
from ..writer import SheetWriter, list_from

NAME = "Client Questionnaire"
LAST_COL = 9
SHOT_ROWS = 22
TIMELINE_ROWS = 20


def _b(field: str) -> str:
    """A field from the chosen booking."""
    return (f'IFERROR(INDEX([[Wedding Bookings.{field}]],'
            f'MATCH(QuestionnaireBooking,'
            f'[[Wedding Bookings.id]],0)),"")')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME,
                    tab_color=Palette.primary_soft)

    s.banner("Client Questionnaire",
             "Everything you need to know before the day, on one "
             "page. Pick a wedding and the rest fills itself in.",
             LAST_COL)
    s.set_col(0, 0, 3)
    s.set_col(1, 1, 26)
    s.set_col(2, 4, 20)
    s.set_col(5, LAST_COL, 18)

    row = 4
    row = s.field(row, "Wedding", None, style="text",
                  name="QuestionnaireBooking",
                  validate=list_from("BookingList"),
                  value_end=4, help_col=6,
                  help_text="Pick the Booking ID. Everything below "
                            "follows from this one cell.")
    row += 1

    # ---- what you already know -------------------------------------------
    row = s.section(row, "What you already know", LAST_COL)
    row = s.note(
        row,
        "Taken from the booking - there is no point asking a couple "
        "for information you have already written down.",
        LAST_COL, height=26)
    row += 1

    known = [
        ("Couple", "client_name", "text"),
        ("Partner", "partner_name", "text"),
        ("Wedding date", "wedding_date", "date"),
        ("Day of the week", "weekday", "text"),
        ("Ceremony", "ceremony_time", "time"),
        ("Reception", "reception_time", "time"),
        ("Coverage", "coverage_hours", "number1"),
        ("Main venue", "venue", "text"),
        ("Getting ready at", "prep_location", "text"),
        ("Ceremony at", "ceremony_location", "text"),
        ("Reception at", "reception_location", "text"),
        ("Guests", "guests", "int"),
        ("Package", "package", "text"),
        ("Lead photographer", "photographer_name", "text"),
        ("Second shooter", "second_name", "text"),
        ("Days to go", "days_away", "int"),
    ]
    half = (len(known) + 1) // 2
    top = row
    for index, (label, field, style) in enumerate(known):
        column = 1 if index < half else 5
        line = top + (index if index < half else index - half)
        s.write(line, column, label, fmts.chrome("label_bold"))
        s.merge(line, column + 1, column + 2, None,
                fmts.cell(style, "calc"))
        s.formula(line, column + 1, "=" + _b(field),
                  fmts.cell(style, "calc"))
    row = top + half + 1

    # ---- the questions -----------------------------------------------------
    row = s.section(row, "The questions", LAST_COL)
    row = s.note(
        row,
        "Write the answers straight onto the booking - these cells "
        "show what is recorded there. Keeping the answers on the "
        "wedding rather than on this form is what lets the shot list, "
        "the timeline and the day sheet all use them.",
        LAST_COL, height=32)
    row += 1

    questions = [
        ("Must-have photographs", "must_haves",
         "The shots they would be upset not to have. Ask for five, "
         "not fifty."),
        ("Family groups, by name", "family_list",
         "Names, not relationships. This is the question that saves "
         "forty minutes on the day."),
        ("Important guests", "key_guests",
         "Anybody travelling a long way, anybody elderly, anybody who "
         "must be photographed."),
        ("Traditions and rituals", "traditions",
         "Anything cultural or religious, and roughly when in the day "
         "it happens."),
        ("Venue restrictions", "restrictions",
         "No flash, no photographs during the ceremony, a balcony you "
         "may not stand on. Ask the venue as well as the couple."),
        ("Family sensitivities", "sensitivities",
         "Divorced parents, a recent bereavement, two people who must "
         "not be in a group shot together. Far better known now."),
        ("Anything else", "special_requests",
         "A surprise, a dog, a drone, a second dress, a late-night "
         "sparkler exit."),
    ]
    for label, field, help_text in questions:
        s.write(row, 1, label, fmts.chrome("label_bold"))
        s.merge(row, 2, 5, None, fmts.cell("wrap", "calc"))
        s.formula(row, 2, "=" + _b(field), fmts.cell("wrap", "calc"))
        s.merge(row, 6, LAST_COL, help_text, fmts.chrome("note"))
        s.row_height(row, 34)
        row += 1
    row += 1

    s.write(row, 1, "Questionnaire sent", fmts.chrome("label_bold"))
    s.formula(row, 2, "=" + _b("questionnaire_sent"),
              fmts.cell("date", "calc"))
    s.write(row, 3, "Answers back", fmts.chrome("label_right"))
    s.formula(row, 4, "=" + _b("questionnaire_back"),
              fmts.cell("date", "calc"))
    s.merge(row, 5, LAST_COL, None, fmts.chrome("note"))
    s.formula(
        row, 5,
        '=IF(QuestionnaireBooking="","Pick a wedding above.",'
        'IF(' + _b("questionnaire_back") + '<>"",'
        '"Answers are in - read them again the week before.",'
        'IF(' + _b("questionnaire_sent") + '="",'
        '"Not sent yet. Send it once the contract is signed.",'
        '"Sent, nothing back. Chase it - the day plan depends on '
        'it.")))',
        fmts.chrome("note"))
    row += 2

    # ---- the day plan ------------------------------------------------------
    row = s.section(row, "The day as it stands", LAST_COL)
    row = s.note(
        row,
        "Built from the Wedding Timeline tab for this Booking ID, in "
        "time order. Print this page and the couple can see the "
        "running order you are working to.",
        LAST_COL, height=28)
    row += 1

    head = row
    s.table_head(head, 1, ["Time", "Until", "What is happening",
                           "Where", "Who", "Must not miss"],
                 widths=[9, 9, 30, 22, 18, 12])
    row = head + 1
    for n in range(1, TIMELINE_ROWS + 1):
        key = f'QuestionnaireBooking&"|"&{n}'
        for col, (field, style) in enumerate([
                ("start", "time"), ("finish", "time"),
                ("shown", "text"), ("location", "text"),
                ("staff_name", "text"), ("critical", "flag")],
                start=1):
            s.formula(
                row, col,
                f'=IF(QuestionnaireBooking="","",'
                f'IFERROR(INDEX([[Wedding Timeline.{field}]],'
                f'MATCH({key},[[Wedding Timeline.day_key]],0)),""))',
                fmts.cell(style, "calc"))
        row += 1
    row += 1

    # ---- the family groups -------------------------------------------------
    row = s.section(row, "Family and group photographs", LAST_COL)
    row = s.note(
        row,
        "Every shot on the Shot List Manager for this wedding that is "
        "either a family group or flagged as a must-have. Hand this "
        "to whoever is rounding people up - a bridesmaid, an usher, "
        "the best man. It works far better than you shouting.",
        LAST_COL, height=32)
    row += 1

    head = row
    s.table_head(head, 1, ["#", "Group", "Shot", "Who is in it",
                           "Asked for by", "Must have", "Taken"],
                 widths=[4, 14, 30, 26, 15, 11, 9])
    row = head + 1
    for n in range(1, SHOT_ROWS + 1):
        s.write(row, 1, n, fmts.chrome("day_number"))
        rank = (f'MATCH(QuestionnaireBooking&"|"&{n},'
                f'[[Shot List Manager.print_key]],0)')
        for col, (field, style) in enumerate([
                ("group_name", "text"), ("shot", "text"),
                ("people", "wrap"), ("requested_by", "text"),
                ("must_have", "flag"), ("done", "flag")], start=2):
            s.formula(
                row, col,
                f'=IF(QuestionnaireBooking="","",'
                f'IFERROR(INDEX([[Shot List Manager.{field}]],'
                f'{rank}),""))',
                fmts.cell(style, "calc"))
        row += 1
    row += 1

    row = s.note(
        row,
        "Print this page before every wedding. Reading the answers "
        "back the week before is the whole point of asking the "
        "questions six months earlier.",
        LAST_COL, height=26)

    s.finish(freeze_row=5, gridlines=False, fit_width=True,
             print_area=f"A1:{chr(65 + LAST_COL)}{row + 2}")
    s.protect()
    return s
````

---

## 8. Editing to final gallery delivery

Workflow four. Ten editing steps against a due date derived from the turnaround you promised, then the gallery: delivery, views, downloads, album orders, after-sales and the expiry date that is worth warning a couple about.

### `wedding_photography_spreadsheet/sheets/editing.py`

*258 lines · 1 function*

> Editing Workflow - one row per wedding, ten steps from card to gallery.
>
> The backlog is where wedding photography businesses quietly fail.  The
> shooting is finished, the money is mostly in, and three weddings sit
> unedited while the next one is being booked.  This tab makes the
> backlog visible and puts a date on it.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Editing Workflow'), `STEP_KEYS` ([f'step_{key}' for key, _ in EDITING_STEPS]).

````python
"""Editing Workflow - one row per wedding, ten steps from card to gallery.

The backlog is where wedding photography businesses quietly fail.  The
shooting is finished, the money is mostly in, and three weddings sit
unedited while the next one is being booked.  This tab makes the
backlog visible and puts a date on it.
"""

from __future__ import annotations

from ..config import EDITING_STEPS
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, date_num_col, id_col,
                     lookup_booking, lookup_staff, money_cf, status_cf,
                     yesno, zero)

NAME = "Editing Workflow"

#: the ten steps become ten tick columns, and the stage is derived
STEP_KEYS = [f"step_{key}" for key, _ in EDITING_STEPS]


def spec(cfg) -> SheetSpec:
    columns = [
        id_col("booking", "EditPrefix", digits=cfg.digits("edit"),
               header="Job ID"),

        Col("booking", "Booking ID", 11, "id",
            validate=list_from("BookingList"),
            note="One row per wedding. The booking shows the stage "
                 "this job has reached.",
            group="The job"),
        Col("couple", "Couple", 18, "text",
            formula=lookup_booking("client_name"), group="The job"),
        Col("wedding_date", "Wedding date", 12, "date",
            formula=lookup_booking("wedding_date", missing=""),
            group="The job"),
        Col("editor", "Editor", 11, "id",
            validate=list_from("StaffList"), group="The job"),
        Col("editor_name", "Editor name", 16, "text",
            formula=lookup_staff("name", key="editor", missing=""),
            group="The job"),
        Col("outsourced", "Outsourced", 10, "flag",
            validate=yesno("Yes when an outside editor does it. The "
                           "cost goes on the Expense Tracker against "
                           "the same Booking ID."),
            group="The job"),
        Col("priority", "Priority", 11, "text",
            validate=list_from("PriorityList"), group="The job"),
        Col("rush", "Rush job", 9, "flag",
            validate=yesno("Yes if they paid the rush fee. The "
                           "promised date tightens automatically."),
            group="The job"),
    ]

    # ---- the ten steps ------------------------------------------------------
    for key, label in EDITING_STEPS:
        columns.append(
            Col(f"step_{key}", label, 13, "flag",
                validate=yesno(), cf=status_cf(), group="The ten steps"))

    done_sum = "+".join(f'({{step_{key}}}{{r}}="Yes")' for key, _ in EDITING_STEPS)
    steps = len(EDITING_STEPS)

    # the stage name is the last ticked step, so it can never disagree
    # with the ticks themselves
    stage = '""'
    for key, label in EDITING_STEPS:
        stage = f'IF({{step_{key}}}{{r}}="Yes","{label}",{stage})'

    columns += [
        Col("steps_done", "Steps done", 10, "int",
            formula=zero("booking", done_sum), group="How far along"),
        Col("progress", "Progress", 10, "percent0",
            formula=blank("booking", f"{{steps_done}}{{r}}/{steps}"),
            cf=bar_cf(), group="How far along"),
        Col("stage", "Stage", 20, "status",
            formula=blank("booking",
                          f'IF({{steps_done}}{{r}}<=0,"Not started",'
                          f'IF({{steps_done}}{{r}}>={steps},"Complete",'
                          f'{stage}))'),
            cf=status_cf(),
            note="Always the last step you ticked, so the stage and "
                 "the ticks can never disagree.",
            group="How far along"),

        # ---- dates ------------------------------------------------------------
        Col("started", "Started", 11, "date",
            validate=date_check(), group="Dates"),
        Col("preview_due", "Preview due", 12, "date",
            formula=blank("booking",
                          'IF({wedding_date}{r}="","",'
                          '{wedding_date}{r}+PreviewDays)'),
            group="Dates"),
        Col("preview_sent", "Preview sent", 12, "date",
            validate=date_check(), group="Dates"),
        Col("due", "Gallery promised", 14, "date",
            formula=blank("booking",
                          'IF({wedding_date}{r}="","",'
                          'IF({rush}{r}="Yes",'
                          '{wedding_date}{r}+RushDeliveryDays,'
                          '{wedding_date}{r}+DeliveryDays))'),
            note="Your standard turnaround from Settings, or the "
                 "shorter rush promise if they paid for it.",
            group="Dates"),
        Col("finished", "Finished", 11, "date",
            validate=date_check(), group="Dates"),
        Col("days_taken", "Days taken", 10, "int",
            formula=blank("booking",
                          'IF({finished}{r}="","",'
                          'IF({wedding_date}{r}="","",'
                          '{finished}{r}-{wedding_date}{r}))'),
            group="Dates"),
        Col("days_left", "Days left", 10, "int",
            formula=blank("booking",
                          'IF({due}{r}="","",'
                          'IF({finished}{r}<>"","",'
                          '{due}{r}-TodayDate))'),
            note="Negative means you are past the date you promised.",
            group="Dates"),
        Col("late", "Late", 8, "flag",
            formula=blank("booking",
                          'IF({due}{r}="","No",'
                          'IF({finished}{r}<>"",'
                          'IF({finished}{r}>{due}{r},"Yes","No"),'
                          'IF(TodayDate>{due}{r},"Yes","No")))'),
            cf=status_cf(), group="Dates"),

        # ---- the work -----------------------------------------------------------
        Col("shot_count", "Photos taken", 11, "int",
            validate=number_check(0, 100000), group="The work"),
        Col("kept", "Kept after culling", 13, "int",
            validate=number_check(0, 100000), group="The work"),
        Col("delivered_count", "Delivered", 10, "int",
            validate=number_check(0, 100000), group="The work"),
        Col("promised_count", "Package promises", 13, "int",
            formula=zero("booking",
                         'IFERROR(INDEX('
                         '[[Packages & Pricing.images]],'
                         'MATCH(INDEX([[Wedding Bookings.package]],'
                         'MATCH({booking}{r},'
                         '[[Wedding Bookings.id]],0)),'
                         '[[Packages & Pricing.name]],0))*1,0)'),
            note="How many edited images the package says they get.",
            group="The work"),
        Col("short", "Short by", 9, "int",
            formula=zero("booking",
                         'IF({delivered_count}{r}="",0,'
                         'MAX(0,{promised_count}{r}-'
                         '{delivered_count}{r}))'),
            group="The work"),
        Col("keep_rate", "Keep rate", 10, "percent0",
            formula=blank("booking",
                          'IF({shot_count}{r}="","",'
                          'IF({shot_count}{r}<=0,"",'
                          'IF({kept}{r}="","",'
                          '{kept}{r}/{shot_count}{r})))'),
            group="The work"),
        Col("hours", "Hours spent", 11, "number1",
            validate=number_check(0, 500),
            note="Be honest. This is the number that decides whether "
                 "your packages are priced properly.",
            group="The work"),
        Col("hours_expected", "Hours expected", 12, "number1",
            formula=zero("booking", "EditingHours"), hidden=True,
            group="The work"),
        Col("hours_over", "Hours over", 10, "number1",
            formula=zero("booking",
                         'IF({hours}{r}="",0,'
                         'ROUND({hours}{r}-{hours_expected}{r},2))'),
            group="The work"),
        Col("rate", "Editing rate", 11, "money",
            formula=zero("booking",
                         'IF({editor}{r}="",EditingRate,'
                         'IFERROR(INDEX('
                         '[[Photographer & Staff.edit_rate]],'
                         'MATCH({editor}{r},'
                         '[[Photographer & Staff.id]],0))*1,'
                         'EditingRate))'),
            hidden=True, group="The work"),
        Col("cost", "Editing cost", 12, "money",
            formula=zero("booking",
                         'ROUND(IF({hours}{r}="",0,{hours}{r})*'
                         '{rate}{r},2)'),
            cf=money_cf(),
            note="Hours times the editor's rate. The booking subtracts "
                 "this from the wedding's profit.",
            group="The work"),
        Col("backup", "Backed up twice", 12, "flag",
            validate=yesno("Two copies in two places. The one thing "
                           "on this tab you cannot undo."),
            cf=status_cf(), group="The work"),
        Col("notes", "Notes", 26, "wrap", group="The work"),

        date_num_col("wedding_date", "wed_num", "Wedding date number"),
        date_num_col("due", "due_num", "Due date number"),

        Col("alert", "What to do", 32, "status",
            formula=(
                '=IF({booking}{r}="","",'
                'IF({steps_done}{r}>=' + str(steps) + ','
                'IF({short}{r}>0,'
                '"Check - fewer images than the package promises",'
                '"OK"),'
                'IF({backup}{r}<>"Yes",'
                '"Urgent - not backed up in two places",'
                'IF({wedding_date}{r}="","Check - booking has no date",'
                'IF({wedding_date}{r}>TodayDate,"OK - not shot yet",'
                'IF({late}{r}="Yes",'
                'IF({days_left}{r}<>"",'
                'IF({days_left}{r}<-14,'
                '"Urgent - two weeks past the promised date",'
                '"Overdue - past the date you promised"),'
                '"Overdue - past the date you promised"),'
                'IF({preview_sent}{r}="",'
                'IF({preview_due}{r}<>"",'
                'IF(TodayDate>{preview_due}{r},'
                '"Action - preview images are overdue",'
                '"OK - in progress"),"OK - in progress"),'
                'IF({days_left}{r}<>"",'
                'IF({days_left}{r}<=7,'
                '"Action - gallery due within the week","OK"),'
                '"OK"))))))))'),
            cf=alert_cf(), group="The work"),
    ]

    return SheetSpec(
        name=NAME,
        title="Editing Workflow",
        subtitle=("Every wedding between the last frame and the "
                  "finished gallery, and how late it is"),
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=None,
        freeze_cols=3,
        samples=sample("editing", cfg),
        help_text=(
            "One row per wedding. Tick the ten steps as you work "
            "through them and the stage, the progress bar and the "
            "booking all keep up on their own. The promised date "
            "comes from your turnaround setting, counted from the "
            "wedding day."),
        extra_notes=[
            "Tick 'Backed up twice' before anything else, and mean "
            "it. Two copies in two places. Everything else on this "
            "tab can be redone; this cannot.",
            "Log the hours honestly. A wedding that took twenty-six "
            "hours to edit instead of eighteen is not a disaster, but "
            "a package priced on eighteen hours that always takes "
            "twenty-six is.",
            "A backlog of three unedited weddings in August is "
            "normal. A backlog of three unedited weddings in January "
            "is a business problem, and this tab is where you see the "
            "difference.",
        ],
    )
````

### `wedding_photography_spreadsheet/sheets/gallery.py`

*213 lines · 1 function*

> Gallery & Delivery - handing the photographs over, and what happens next.
>
> Delivery is not the end of the job.  Galleries expire, albums get
> ordered months later, print sales happen only if somebody mentions
> them, and the review you want is easiest to ask for in the week the
> couple first sees their photographs.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Gallery & Delivery').

````python
"""Gallery & Delivery - handing the photographs over, and what happens next.

Delivery is not the end of the job.  Galleries expire, albums get
ordered months later, print sales happen only if somebody mentions
them, and the review you want is easiest to ask for in the week the
couple first sees their photographs.
"""

from __future__ import annotations

from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, booking_client_col, blank, date_num_col, id_col, lookup_booking,
                     money_cf, status_cf, yesno, zero)

NAME = "Gallery & Delivery"


def spec(cfg) -> SheetSpec:
    columns = [
        id_col("booking", "GalleryPrefix", digits=cfg.digits("gallery"),
               header="Gallery ID"),

        Col("booking", "Booking ID", 11, "id",
            validate=list_from("BookingList"),
            note="One row per wedding gallery.", group="The gallery"),
        Col("couple", "Couple", 18, "text",
            formula=lookup_booking("client_name"), group="The gallery"),
        booking_client_col(),
        Col("wedding_date", "Wedding date", 12, "date",
            formula=lookup_booking("wedding_date", missing=""),
            group="The gallery"),
        Col("platform", "Where it lives", 15, "text",
            validate=list_from("GalleryPlatformList"),
            group="The gallery"),
        Col("gallery_name", "Gallery name", 20, "text",
            group="The gallery"),
        Col("link", "Link", 26, "link",
            note="Paste the address. The workbook stores the link, "
                 "not the photographs.",
            group="The gallery"),
        Col("has_password", "Password set", 11, "flag",
            validate=yesno("Yes or No only. Never type the password "
                           "itself into a spreadsheet you email to "
                           "anybody."),
            cf=status_cf(), group="The gallery"),

        # ---- delivery ------------------------------------------------------
        Col("status", "Status", 14, "status",
            validate=list_from("GalleryStatusList"), cf=status_cf(),
            group="Delivery"),
        Col("promised", "Promised by", 12, "date",
            formula=blank("booking",
                          'IFERROR(INDEX([[Editing Workflow.due]],'
                          'MATCH({booking}{r},'
                          '[[Editing Workflow.booking]],0)),'
                          'IF({wedding_date}{r}="","",'
                          '{wedding_date}{r}+DeliveryDays))'),
            group="Delivery"),
        Col("delivered_on", "Delivered", 11, "date",
            validate=date_check(), group="Delivery"),
        Col("days_late", "Days late", 10, "int",
            formula=zero("booking",
                         'IF({promised}{r}="",0,'
                         'IF({delivered_on}{r}<>"",'
                         'MAX(0,{delivered_on}{r}-{promised}{r}),'
                         'MAX(0,TodayDate-{promised}{r})))'),
            group="Delivery"),
        Col("images", "Images in it", 11, "int",
            validate=number_check(0, 100000), group="Delivery"),
        Col("expires", "Gallery expires", 13, "date",
            formula=blank("booking",
                          'IF({delivered_on}{r}="","",'
                          '{delivered_on}{r}+GalleryLiveDays)'),
            note="How long you keep it online, from Settings. Tell "
                 "the couple this date - it is the single most common "
                 "thing a photographer gets an angry email about.",
            group="Delivery"),
        Col("expiry_warned", "Expiry warning sent", 14, "flag",
            validate=yesno(), group="Delivery"),
        Col("archived", "Archived", 10, "flag",
            validate=yesno("Taken offline but kept on your own "
                           "storage."),
            group="Delivery"),

        # ---- what the couple did -------------------------------------------
        Col("first_viewed", "First opened", 12, "date",
            validate=date_check(), group="What they did"),
        Col("views", "Views", 8, "int",
            validate=number_check(0, 1000000), group="What they did"),
        Col("downloads", "Downloads", 10, "int",
            validate=number_check(0, 1000000), group="What they did"),
        Col("favourites", "Favourites picked", 13, "int",
            validate=number_check(0, 100000), group="What they did"),
        Col("usb_sent", "USB or print box sent", 14, "flag",
            validate=yesno(), group="What they did"),
        Col("opened", "Have they looked", 12, "flag",
            formula=blank("booking",
                          'IF({delivered_on}{r}="","",'
                          'IF({first_viewed}{r}<>"","Yes",'
                          'IF(IF({views}{r}="",0,{views}{r})>0,"Yes",'
                          '"No")))'),
            cf=status_cf(), group="What they did"),

        # ---- what it earned afterwards -------------------------------------
        Col("album_ordered", "Album ordered", 12, "flag",
            validate=yesno(), group="Afterwards"),
        Col("album_due", "Album due", 11, "date",
            formula=blank("booking",
                          'IF({album_ordered}{r}<>"Yes","",'
                          'IF({delivered_on}{r}="","",'
                          '{delivered_on}{r}+AlbumDays))'),
            group="Afterwards"),
        Col("album_sent", "Album sent", 11, "date",
            validate=date_check(), group="Afterwards"),
        Col("print_sales", "Print sales", 11, "money",
            validate=number_check(0, 100000),
            note="Money the gallery made after delivery. For a lot of "
                 "photographers this is the most under-collected "
                 "revenue in the business.",
            group="Afterwards"),
        Col("extra_sales", "Other extras", 11, "money",
            validate=number_check(0, 100000), group="Afterwards"),
        Col("after_total", "Earned after delivery", 14, "money",
            formula=zero("booking",
                         'IF({print_sales}{r}="",0,{print_sales}{r})+'
                         'IF({extra_sales}{r}="",0,{extra_sales}{r})'),
            cf=money_cf(), group="Afterwards"),
        Col("review_asked", "Review asked for", 13, "date",
            validate=date_check(),
            note="Ask in the week they first open the gallery. It is "
                 "the only week they will ever be this pleased with "
                 "you.",
            group="Afterwards"),
        Col("review_in", "Review received", 12, "flag",
            formula=blank("booking",
                          'IF(COUNTIFS('
                          '[[Reviews & Referrals.booking]],'
                          '{booking}{r})>0,"Yes","No")'),
            cf=status_cf(), group="Afterwards"),
        Col("notes", "Notes", 26, "wrap", group="Afterwards"),

        date_num_col("wedding_date", "wed_num", "Wedding date number"),
        date_num_col("delivered_on", "delivered_num", "Delivered number"),
        Col("is_delivered", "Delivered flag", 9, "flag",
            formula=blank("booking",
                          'IF({delivered_on}{r}<>"","Yes","No")'),
            hidden=True),

        Col("alert", "What to do", 32, "status",
            formula=(
                '=IF({booking}{r}="","",'
                'IF({delivered_on}{r}="",'
                'IF({promised}{r}="","Check - no promised date",'
                'IF(TodayDate>{promised}{r},'
                '"Overdue - past the date you promised",'
                'IF({promised}{r}-TodayDate<=7,'
                '"Due - gallery due within the week","OK"))),'
                'IF({has_password}{r}<>"Yes",'
                '"Check - gallery has no password",'
                'IF({opened}{r}="No",'
                'IF(TodayDate-{delivered_on}{r}>7,'
                '"Chase - delivered, never opened","OK"),'
                'IF({album_ordered}{r}="Yes",'
                'IF({album_sent}{r}="",'
                'IF({album_due}{r}<>"",'
                'IF(TodayDate>{album_due}{r},'
                '"Overdue - album past its date",'
                '"Action - album still to send"),'
                '"Action - album still to send"),"OK"),'
                'IF({review_in}{r}="No",'
                'IF({review_asked}{r}="",'
                'IF(TodayDate-{delivered_on}{r}<=30,'
                '"Action - ask for a review now","OK"),"OK"),'
                'IF({expires}{r}<>"",'
                'IF({expiry_warned}{r}<>"Yes",'
                'IF({expires}{r}-TodayDate<=30,'
                '"Send - warn them the gallery is closing","OK"),'
                '"OK"),"OK")))))))'),
            cf=alert_cf(), group="Afterwards"),
    ]

    return SheetSpec(
        name=NAME,
        title="Gallery & Delivery",
        subtitle=("Handing the photographs over, and everything that "
                  "should happen in the month afterwards"),
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=None,
        freeze_cols=3,
        samples=sample("gallery", cfg),
        help_text=(
            "One row per wedding gallery. The promised date comes from "
            "the editing job, the expiry date from your own setting, "
            "and the alert column walks you through the fortnight "
            "after delivery - open the gallery, ask for the review, "
            "chase the album, warn them before it closes."),
        extra_notes=[
            "Never type a gallery password into this workbook. The "
            "'Password set' column is a Yes or No, because a "
            "spreadsheet gets emailed to accountants and left open on "
            "laptops in coffee shops.",
            "Tell the couple the expiry date when you deliver, and "
            "warn them again a month before. Galleries closing with "
            "nothing downloaded is the most avoidable complaint in "
            "wedding photography.",
            "'Earned after delivery' is money most photographers never "
            "count. Prints, extra albums, a parents' copy - all of it "
            "comes from the gallery being open and somebody asking.",
        ],
    )
````

---

## 9. What each wedding actually made

Workflow five. Equipment, travel, marketing and every other cost, tagged to a wedding or to the business, feeding a profit analysis that reports margin and profit per hour - then the reviews that bring the next booking, and the monthly and annual reports.

### `wedding_photography_spreadsheet/sheets/equipment.py`

*206 lines · 1 function*

> Equipment Inventory - what you own, what it is worth and what it owes you.
>
> Cameras are the second biggest cost in a wedding photography business
> after your own time, and the one most often treated as a hobby
> purchase.  This tab puts a cost-per-wedding on every body and lens, and
> keeps the serial numbers an insurer will ask for.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Equipment Inventory').

````python
"""Equipment Inventory - what you own, what it is worth and what it owes you.

Cameras are the second biggest cost in a wedding photography business
after your own time, and the one most often treated as a hobby
purchase.  This tab puts a cost-per-wedding on every body and lens, and
keeps the serial numbers an insurer will ask for.
"""

from __future__ import annotations

from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, lookup_staff, money_cf,
                     rank_desc, status_cf, yesno, zero)

NAME = "Equipment Inventory"


def spec(cfg) -> SheetSpec:
    columns = [
        id_col("item", "EquipmentPrefix", digits=cfg.digits("equipment"),
               header="Item ID"),

        Col("item", "Item", 22, "id", group="What it is"),
        Col("category", "Category", 15, "text",
            validate=list_from("EquipmentCategoryList"),
            group="What it is"),
        Col("make", "Make and model", 22, "text", group="What it is"),
        Col("serial", "Serial number", 18, "text",
            note="The first thing an insurer or the police will ask "
                 "for, and the hardest thing to find afterwards.",
            group="What it is"),
        Col("status", "Status", 13, "status",
            validate=list_from("EquipmentStatusList"), cf=status_cf(),
            group="What it is"),
        Col("condition", "Condition", 12, "text",
            validate=list_from("EquipmentConditionList"),
            group="What it is"),
        Col("assigned", "Assigned to", 12, "id",
            validate=list_from("StaffList"), group="What it is"),
        Col("assigned_name", "Who has it", 16, "text",
            formula=lookup_staff("name", key="assigned", missing=""),
            group="What it is"),
        Col("location", "Kept where", 16, "text", group="What it is"),

        # ---- what it cost -----------------------------------------------------
        Col("bought", "Bought", 11, "date",
            validate=date_check(), group="What it cost"),
        Col("price", "Price paid", 12, "money",
            validate=number_check(0, 1000000), cf=money_cf(),
            group="What it cost"),
        Col("life_years", "Useful life", 10, "number1",
            validate=number_check(0, 40),
            note="How many years you expect to get out of it. Leave "
                 "blank to use the default from Settings.",
            group="What it cost"),
        Col("life_used", "Life used", 10, "number1",
            formula=zero("item",
                         'IF({life_years}{r}="",EquipmentLifeYears,'
                         '{life_years}{r})'),
            hidden=True, group="What it cost"),
        Col("age_years", "Age in years", 11, "number1",
            formula=blank("item",
                          'IF({bought}{r}="","",'
                          'ROUND((TodayDate-{bought}{r})/365.25,2))'),
            group="What it cost"),
        Col("annual_cost", "Cost a year", 11, "money",
            formula=zero("item",
                         'IF({life_used}{r}<=0,0,'
                         'ROUND(IF({price}{r}="",0,{price}{r})/'
                         '{life_used}{r},2))'),
            note="Spread over its useful life. This, not the sticker "
                 "price, is what the camera costs your business each "
                 "year.",
            group="What it cost"),
        Col("written_down", "Written down to", 13, "money",
            formula=zero("item",
                         'IF({price}{r}="",0,'
                         'IF({age_years}{r}="",{price}{r},'
                         'ROUND(MAX(0,{price}{r}-{annual_cost}{r}*'
                         '{age_years}{r}),2)))'),
            note="A rough straight-line figure for your own planning. "
                 "Your accountant's depreciation schedule is the one "
                 "that counts at tax time.",
            group="What it cost"),
        Col("weddings_since", "Weddings since", 12, "int",
            formula=zero("item",
                         'IF({bought}{r}="",0,'
                         'COUNTIFS([[Wedding Bookings.wed_num]],'
                         '">="&{bought}{r},'
                         '[[Wedding Bookings.wed_num]],'
                         '"<="&TodayDate,'
                         '[[Wedding Bookings.earns]],"Yes"))'),
            group="What it cost"),
        Col("cost_per_wedding", "Cost a wedding", 12, "money",
            formula=blank("item",
                          'IF({weddings_since}{r}<=0,"",'
                          'ROUND(IF({price}{r}="",0,{price}{r})/'
                          '{weddings_since}{r},2))'),
            note="What it has cost you per wedding so far. A lens "
                 "that has shot forty weddings has earned its keep; "
                 "one that has shot two has not.",
            group="What it cost"),
        rank_desc("price", guard='{item}{r}=""', header="Value rank"),

        # ---- looking after it --------------------------------------------------
        Col("insured", "Insured", 9, "flag",
            validate=yesno(), cf=status_cf(), group="Looking after it"),
        Col("insured_value", "Insured for", 12, "money",
            validate=number_check(0, 1000000), group="Looking after it"),
        Col("under_insured", "Short by", 11, "money",
            formula=zero("item",
                         'IF({insured}{r}<>"Yes",0,'
                         'MAX(0,IF({price}{r}="",0,{price}{r})-'
                         'IF({insured_value}{r}="",0,'
                         '{insured_value}{r})))'),
            group="Looking after it"),
        Col("warranty_to", "Warranty to", 12, "date",
            validate=date_check(), group="Looking after it"),
        Col("serviced", "Last serviced", 12, "date",
            validate=date_check(), group="Looking after it"),
        Col("service_due", "Service due", 12, "date",
            formula=blank("item",
                          'IF({serviced}{r}="","",'
                          '{serviced}{r}+ServiceMonths*30)'),
            group="Looking after it"),
        Col("shutter", "Shutter count", 12, "int",
            validate=number_check(0, 2000000),
            note="For camera bodies. Most shutters are rated between "
                 "150,000 and 500,000 actuations.",
            group="Looking after it"),
        Col("shutter_limit", "Rated for", 11, "int",
            validate=number_check(0, 2000000), group="Looking after it"),
        Col("shutter_used", "Shutter life used", 13, "percent0",
            formula=blank("item",
                          'IF({shutter_limit}{r}="","",'
                          'IF({shutter_limit}{r}<=0,"",'
                          'IF({shutter}{r}="","",'
                          '{shutter}{r}/{shutter_limit}{r})))'),
            group="Looking after it"),
        Col("backup_for", "Backs up", 16, "text",
            note="Which item this is the spare for. Shooting a "
                 "wedding without a second body is a decision, not an "
                 "accident.",
            group="Looking after it"),
        Col("replace_cost", "Replacement cost", 13, "money",
            validate=number_check(0, 1000000), group="Looking after it"),
        Col("notes", "Notes", 24, "wrap", group="Looking after it"),

        Col("alert", "What to do", 32, "status",
            formula=(
                '=IF({item}{r}="","",'
                'IF({status}{r}="Sold","",'
                'IF({status}{r}="Lost or stolen",'
                '"Check - write it off and tell your insurer",'
                'IF({status}{r}="Being repaired",'
                '"Action - away for repair, check your cover",'
                'IF({serial}{r}="","Add - no serial number recorded",'
                'IF({insured}{r}<>"Yes",'
                'IF(IF({price}{r}="",0,{price}{r})>=InsureAbove,'
                '"Check - valuable and not insured","OK"),'
                'IF({under_insured}{r}>0,'
                '"Check - insured for less than it cost",'
                'IF({shutter_used}{r}<>"",'
                'IF({shutter_used}{r}>=0.8,'
                '"Order - shutter near the end of its life",'
                'IF({service_due}{r}<>"",'
                'IF(TodayDate>{service_due}{r},'
                '"Due - service overdue","OK"),"OK")),'
                'IF({service_due}{r}<>"",'
                'IF(TodayDate>{service_due}{r},'
                '"Due - service overdue","OK"),"OK")))))))))'),
            cf=alert_cf(), group="Looking after it"),
    ]

    return SheetSpec(
        name=NAME,
        title="Equipment Inventory",
        subtitle=("Every body, lens, light and card you own - what it "
                  "is worth, what it costs you and when it needs "
                  "looking at"),
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=None,
        freeze_cols=2,
        samples=sample("equipment", cfg),
        help_text=(
            "List everything you would have to replace if the car were "
            "broken into. Serial numbers, what you paid and what it is "
            "insured for. The sheet works out what each item costs you "
            "per year and per wedding, and flags anything expensive "
            "that is not covered."),
        extra_notes=[
            "Record serial numbers now, while the box is in front of "
            "you. Nobody has ever found a serial number quickly on "
            "the day they needed it.",
            "'Cost a wedding' is the honest way to think about kit. A "
            "$2,400 lens across forty weddings is sixty dollars a "
            "wedding; across three it is eight hundred.",
            "The written-down figure here is straight-line and rough, "
            "for your own planning only. Capital allowances and "
            "depreciation for tax are a matter for your accountant "
            "and your country's rules.",
        ],
    )
````

### `wedding_photography_spreadsheet/sheets/travel.py`

*221 lines · 1 function*

> Travel & Mileage - getting to the wedding, and what it costs.
>
> Mileage is the deduction photographers most often lose, because the
> record has to be made at the time and nobody wants to make it at
> eleven at night in a car park.  One row, four cells, and the rest is
> worked out.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Travel & Mileage').

````python
"""Travel & Mileage - getting to the wedding, and what it costs.

Mileage is the deduction photographers most often lose, because the
record has to be made at the time and nobody wants to make it at
eleven at night in a car park.  One row, four cells, and the rest is
worked out.
"""

from __future__ import annotations

from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, date_num_col, id_col,
                     mileage_cost, money_cf, month_col, status_cf,
                     year_col, yesno, zero)

NAME = "Travel & Mileage"


def spec(cfg) -> SheetSpec:
    columns = [
        id_col("date", "TravelPrefix", digits=cfg.digits("travel"),
               header="Trip ID"),

        Col("date", "Date", 11, "date",
            validate=date_check(),
            note="The day you travelled. Log it the same week - "
                 "reconstructing mileage from memory in April is how "
                 "the deduction gets lost.",
            group="The trip"),
        Col("booking", "Booking ID", 11, "id",
            validate=list_from("BookingList"),
            note="Leave blank for travel that is not for one "
                 "particular wedding - a trade show, a bank visit.",
            group="The trip"),
        Col("couple", "Couple", 18, "text",
            formula=blank("date",
                          'IF({booking}{r}="","",'
                          'IFERROR(INDEX('
                          '[[Wedding Bookings.client_name]],'
                          'MATCH({booking}{r},'
                          '[[Wedding Bookings.id]],0))&"",""))'),
            group="The trip"),
        Col("purpose", "Purpose", 16, "text",
            validate=list_from("TravelPurposeList"), group="The trip"),
        Col("from_place", "From", 20, "text", group="The trip"),
        Col("to_place", "To", 20, "text", group="The trip"),
        Col("mode", "How", 12, "text",
            validate=list_from("TravelModeList"), group="The trip"),

        # ---- the distance -----------------------------------------------------
        Col("one_way", "Miles one way", 12, "number1",
            validate=number_check(0, 10000), group="The distance"),
        Col("return_trip", "Return", 9, "flag",
            validate=yesno("Yes doubles the distance. Almost always "
                           "yes, and almost always forgotten."),
            group="The distance"),
        Col("miles", "Total miles", 11, "number1",
            formula=zero("date",
                         'ROUND(IF({one_way}{r}="",0,{one_way}{r})*'
                         'IF({return_trip}{r}="Yes",2,1),1)'),
            group="The distance"),
        Col("scout_trips", "Extra trips", 10, "int",
            validate=number_check(0, 50),
            note="Venue visits, a rehearsal, dropping an album off. "
                 "Each one is the same distance again.",
            group="The distance"),
        Col("all_miles", "Miles in all", 11, "number1",
            formula=zero("date",
                         'ROUND({miles}{r}*(1+'
                         'IF({scout_trips}{r}="",0,'
                         '{scout_trips}{r})),1)'),
            group="The distance"),

        # ---- the cost ---------------------------------------------------------
        Col("mileage", "Mileage cost", 12, "money",
            formula=zero("date",
                         'ROUND({all_miles}{r}*MileageRate,2)'),
            note="At the rate on Settings. Check your own country's "
                 "approved figure each tax year - it changes.",
            group="The cost"),
        Col("fuel", "Fuel instead", 11, "money",
            validate=number_check(0, 100000),
            note="If you claim actual fuel rather than a mileage "
                 "rate, put it here and the mileage figure is "
                 "ignored. Most countries let you do one or the "
                 "other, not both.",
            group="The cost"),
        Col("tolls", "Tolls", 10, "money",
            validate=number_check(0, 100000), group="The cost"),
        Col("parking", "Parking", 10, "money",
            validate=number_check(0, 100000), group="The cost"),
        Col("accommodation", "Accommodation", 13, "money",
            validate=number_check(0, 100000),
            note="For a wedding far enough away that driving home at "
                 "one in the morning is not sensible.",
            group="The cost"),
        Col("other_cost", "Other", 10, "money",
            validate=number_check(0, 100000), group="The cost"),
        Col("total_cost", "Total cost", 12, "money",
            formula=zero("date",
                         'ROUND(IF({fuel}{r}<>"",{fuel}{r},'
                         '{mileage}{r})+'
                         'IF({tolls}{r}="",0,{tolls}{r})+'
                         'IF({parking}{r}="",0,{parking}{r})+'
                         'IF({accommodation}{r}="",0,'
                         '{accommodation}{r})+'
                         'IF({other_cost}{r}="",0,{other_cost}{r}),2)'),
            cf=money_cf(),
            note="What the trip cost you. The wedding subtracts this "
                 "from its profit.",
            group="The cost"),

        # ---- charging it on ----------------------------------------------------
        Col("allowance", "Miles included", 11, "int",
            formula=zero("date",
                         'IF({booking}{r}="",0,'
                         'IFERROR(INDEX('
                         '[[Packages & Pricing.travel_allowance]],'
                         'MATCH(INDEX([[Wedding Bookings.package]],'
                         'MATCH({booking}{r},'
                         '[[Wedding Bookings.id]],0)),'
                         '[[Packages & Pricing.name]],0))*1,'
                         'FreeTravelMiles))'),
            note="Free miles from the package the couple booked.",
            group="Charging it on"),
        Col("billable_miles", "Miles to charge", 12, "number1",
            formula=zero("date",
                         'IF({booking}{r}="",0,'
                         'ROUND(MAX(0,{all_miles}{r}-'
                         '{allowance}{r}),1))'),
            group="Charging it on"),
        Col("suggested", "Suggested charge", 13, "money",
            formula=zero("date",
                         'IF({booking}{r}="",0,'
                         + mileage_cost("{all_miles}{r}",
                                        "{allowance}{r}") + ")"),
            note="Mileage past the package allowance, plus nothing "
                 "else. Tolls and accommodation are a conversation, "
                 "not an automatic charge.",
            group="Charging it on"),
        Col("charge_override", "Charge instead", 12, "money",
            validate=number_check(0, 100000), group="Charging it on"),
        Col("chargeable", "Charged on", 12, "money",
            formula=zero("date",
                         'IF({charge_override}{r}<>"",'
                         '{charge_override}{r},{suggested}{r})'),
            cf=money_cf(),
            note="What the booking adds to the couple's bill. Set "
                 "'Charge instead' to zero to absorb a trip.",
            group="Charging it on"),
        Col("net_cost", "Costs you", 11, "money",
            formula=zero("date",
                         'ROUND({total_cost}{r}-{chargeable}{r},2)'),
            group="Charging it on"),
        Col("hours", "Hours travelling", 12, "number1",
            validate=number_check(0, 100),
            note="Driving time is working time, even though nobody "
                 "pays you for it. Profit per hour on the booking is "
                 "kinder than it should be if you leave this blank.",
            group="Charging it on"),
        Col("receipt", "Receipt kept", 11, "flag",
            validate=yesno(), cf=status_cf(), group="Charging it on"),
        Col("notes", "Notes", 24, "wrap", group="Charging it on"),

        month_col("date"),
        year_col("date"),
        date_num_col("date", "date_num", "Date number"),

        Col("alert", "What to do", 30, "status",
            formula=(
                '=IF({date}{r}="","",'
                'IF({all_miles}{r}<=0,'
                'IF({total_cost}{r}<=0,"Add - no distance or cost",'
                '"OK"),'
                'IF({from_place}{r}="",'
                '"Add - no start point for the journey",'
                'IF({to_place}{r}="","Add - no destination",'
                'IF({receipt}{r}<>"Yes",'
                'IF({total_cost}{r}-{mileage}{r}>0,'
                '"Missing - costs claimed with no receipt","OK"),'
                'IF({booking}{r}="","OK",'
                'IF({net_cost}{r}>TravelLossLimit,'
                '"Loss - this trip cost far more than you charged",'
                '"OK")))))))'),
            cf=alert_cf(), group="Charging it on"),
    ]

    return SheetSpec(
        name=NAME,
        title="Travel & Mileage",
        subtitle=("Every journey for every wedding - the distance, "
                  "the cost and what you can charge on"),
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=None,
        freeze_cols=3,
        samples=sample("travel", cfg),
        help_text=(
            "Four cells per trip: the date, the Booking ID, the "
            "one-way distance and whether it was a return. Everything "
            "else is worked out - the mileage at your rate, the free "
            "miles the package includes and what is fair to charge on "
            "top."),
        extra_notes=[
            "Mileage rates and what counts as business travel are set "
            "by your own tax authority and change from year to year. "
            "The rate lives on Settings so you can keep it current; "
            "check it against the official figure each April.",
            "Travel between your home and a regular place of work is "
            "usually not deductible, and a wedding venue usually is, "
            "but the line is drawn differently in different "
            "countries. If you are not sure, ask your accountant "
            "rather than this spreadsheet.",
            "Log the hours. A wedding ninety minutes away is three "
            "hours of unpaid driving on top of a ten-hour day, and "
            "that is the difference between a good booking and a "
            "tiring one.",
        ],
    )
````

### `wedding_photography_spreadsheet/sheets/marketing.py`

*186 lines · 1 function*

> Marketing Calendar - what you posted, and whether anybody booked.
>
> Most photography marketing advice stops at "post consistently".  This
> tab goes one step further and ties posts to the enquiries and the
> weddings that came out of them, so you can tell the difference between
> the content that fills your diary and the content that fills an hour.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Marketing Calendar').

````python
"""Marketing Calendar - what you posted, and whether anybody booked.

Most photography marketing advice stops at "post consistently".  This
tab goes one step further and ties posts to the enquiries and the
weddings that came out of them, so you can tell the difference between
the content that fills your diary and the content that fills an hour.
"""

from __future__ import annotations

from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, date_num_col, id_col, money_cf,
                     month_col, rank_desc, status_cf, year_col, yesno, zero)

NAME = "Marketing Calendar"


def spec(cfg) -> SheetSpec:
    columns = [
        id_col("date", "MarketingPrefix", digits=cfg.digits("marketing"),
               header="Post ID"),

        Col("date", "Planned for", 11, "date",
            validate=date_check(), group="The post"),
        Col("platform", "Platform", 14, "text",
            validate=list_from("MarketingPlatformList"),
            group="The post"),
        Col("type", "Type", 14, "text",
            validate=list_from("ContentTypeList"), group="The post"),
        Col("title", "Title or hook", 26, "text", group="The post"),
        Col("content", "What it says", 30, "wrap", group="The post"),
        Col("booking", "Wedding featured", 13, "id",
            validate=list_from("BookingList"),
            note="Which wedding the photographs came from. Check the "
                 "couple were happy for their photographs to be used "
                 "before you post them.",
            group="The post"),
        Col("permission", "Permission given", 13, "flag",
            validate=yesno("The couple agreed their photographs can "
                           "be shown publicly. Worth a line in the "
                           "contract and a tick here."),
            cf=status_cf(), group="The post"),
        Col("status", "Status", 13, "status",
            validate=list_from("ContentStatusList"), cf=status_cf(),
            group="The post"),
        Col("posted_on", "Actually posted", 13, "date",
            validate=date_check(), group="The post"),

        # ---- what it did -----------------------------------------------------
        Col("reach", "Reach", 10, "int",
            validate=number_check(0, 100000000), group="What it did"),
        Col("engagement", "Likes and comments", 14, "int",
            validate=number_check(0, 10000000), group="What it did"),
        Col("saves", "Saves or shares", 13, "int",
            validate=number_check(0, 10000000), group="What it did"),
        Col("clicks", "Clicks through", 12, "int",
            validate=number_check(0, 10000000), group="What it did"),
        Col("engagement_rate", "Engagement rate", 13, "percent0",
            formula=blank("date",
                          'IF({reach}{r}="","",IF({reach}{r}<=0,"",'
                          '(IF({engagement}{r}="",0,'
                          '{engagement}{r})+'
                          'IF({saves}{r}="",0,{saves}{r}))/'
                          '{reach}{r}))'),
            group="What it did"),
        Col("enquiries", "Enquiries from it", 13, "int",
            formula=zero("date",
                         'COUNTIFS([[Lead Tracker.campaign]],'
                         '{id}{r})'),
            note="Counted from the Lead Tracker, where a lead names "
                 "this Post ID as the campaign it came from.",
            group="What it did"),
        Col("manual_enquiries", "Counted by hand", 12, "int",
            validate=number_check(0, 1000),
            note="For enquiries you know came from a post but never "
                 "tagged.",
            group="What it did"),
        Col("all_enquiries", "Enquiries in all", 13, "int",
            formula=zero("date",
                         '{enquiries}{r}+'
                         'IF({manual_enquiries}{r}="",0,'
                         '{manual_enquiries}{r})'),
            group="What it did"),
        Col("bookings", "Weddings booked", 13, "int",
            formula=zero("date",
                         'COUNTIFS([[Lead Tracker.campaign]],{id}{r},'
                         '[[Lead Tracker.is_won]],"Yes")'),
            group="What it did"),
        Col("revenue", "Revenue from it", 13, "money",
            formula=zero("date",
                         'SUMIFS([[Lead Tracker.booked_value]],'
                         '[[Lead Tracker.campaign]],{id}{r},'
                         '[[Lead Tracker.is_won]],"Yes")'),
            cf=money_cf(), group="What it did"),

        # ---- what it cost -----------------------------------------------------
        Col("spend", "Spent on it", 11, "money",
            validate=number_check(0, 1000000),
            note="Advertising, a directory listing, a styled shoot. "
                 "Leave blank for anything that only cost you time.",
            group="What it cost"),
        Col("hours", "Hours it took", 11, "number1",
            validate=number_check(0, 500), group="What it cost"),
        Col("time_cost", "Your time is worth", 13, "money",
            formula=zero("date",
                         'ROUND(IF({hours}{r}="",0,{hours}{r})*'
                         'OwnerHourlyValue,2)'),
            note="Hours times what an hour of yours is worth, from "
                 "Settings. Free marketing is not free.",
            group="What it cost"),
        Col("true_cost", "Real cost", 11, "money",
            formula=zero("date",
                         'ROUND(IF({spend}{r}="",0,{spend}{r})+'
                         '{time_cost}{r},2)'),
            group="What it cost"),
        Col("roi", "Return", 10, "number1",
            formula=blank("date",
                          'IF({true_cost}{r}<=0,"",'
                          'ROUND({revenue}{r}/{true_cost}{r},2))'),
            note="Revenue for every pound or dollar it really cost. "
                 "Under 1 means it lost money once your time is "
                 "counted.",
            group="What it cost"),
        Col("cost_per_enquiry", "Cost an enquiry", 12, "money",
            formula=blank("date",
                          'IF({all_enquiries}{r}<=0,"",'
                          'ROUND({true_cost}{r}/'
                          '{all_enquiries}{r},2))'),
            group="What it cost"),
        rank_desc("revenue", guard='{date}{r}=""', header="Best posts"),
        Col("notes", "Notes", 24, "wrap", group="What it cost"),

        month_col("date"),
        year_col("date"),
        date_num_col("date", "date_num", "Date number"),

        Col("alert", "What to do", 30, "status",
            formula=(
                '=IF({date}{r}="","",'
                'IF({status}{r}="Published",'
                'IF({booking}{r}<>"",'
                'IF({permission}{r}<>"Yes",'
                '"Urgent - posted without recorded permission",'
                'IF({reach}{r}="","Add - no figures logged yet",'
                '"OK")),'
                'IF({reach}{r}="","Add - no figures logged yet","OK")),'
                'IF({status}{r}="Cancelled","",'
                'IF({date}{r}<TodayDate,'
                '"Overdue - planned but never posted",'
                'IF({date}{r}-TodayDate<=3,'
                'IF({content}{r}="","Action - nothing written yet",'
                '"Due - posting within three days"),"OK")))))'),
            cf=alert_cf(), group="What it cost"),
    ]

    return SheetSpec(
        name=NAME,
        title="Marketing Calendar",
        subtitle=("What you plan to post, what you actually posted, "
                  "and which of it brought weddings in"),
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=None,
        freeze_cols=3,
        samples=sample("marketing", cfg),
        help_text=(
            "Plan posts ahead, then come back and fill in what "
            "happened. Put the Post ID into a lead's campaign column "
            "and the enquiries, the weddings and the revenue find "
            "their own way back here."),
        extra_notes=[
            "The 'Real cost' column counts your time at the hourly "
            "value you set on Settings. Four hours editing a blog "
            "post is not free marketing; it is the most expensive "
            "marketing you do.",
            "Tag leads with the Post ID. It takes two seconds at "
            "enquiry and is the only way to ever know which half of "
            "your marketing works.",
            "Check permission before posting a wedding. Most couples "
            "are delighted; the ones who are not tend to be very not, "
            "and a line in your contract plus a tick in this column "
            "settles it.",
        ],
    )
````

### `wedding_photography_spreadsheet/sheets/expenses.py`

*180 lines · 1 function*

> Expense Tracker - everything the business pays out.
>
> Costs that can be pinned to a wedding carry its Booking ID and come
> straight off that wedding's profit.  Everything else - insurance,
> software, the website - is an overhead the whole year has to carry,
> and is treated as one.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Expense Tracker').

````python
"""Expense Tracker - everything the business pays out.

Costs that can be pinned to a wedding carry its Booking ID and come
straight off that wedding's profit.  Everything else - insurance,
software, the website - is an overhead the whole year has to carry,
and is treated as one.
"""

from __future__ import annotations

from ..config import DISCLAIMER
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, date_num_col, id_col, is_capital,
                     is_direct_cost, money_cf, month_col, monthly_equivalent,
                     status_cf, year_col, yesno, zero)

NAME = "Expense Tracker"


def spec(cfg) -> SheetSpec:
    columns = [
        id_col("date", "ExpensePrefix", digits=cfg.digits("expense"),
               header="Expense ID"),

        Col("date", "Date", 11, "date",
            validate=date_check(), group="What you paid"),
        Col("supplier", "Paid to", 20, "text", group="What you paid"),
        Col("category", "Category", 18, "text",
            validate=list_from("ExpenseCategoryList"),
            note="The category decides whether it comes off one "
                 "wedding or off the year as a whole.",
            group="What you paid"),
        Col("description", "What it was", 28, "wrap",
            group="What you paid"),
        Col("booking", "Booking ID", 11, "id",
            validate=list_from("BookingList"),
            note="Only for costs caused by one wedding - a second "
                 "shooter, an album, parking at the venue. Leave "
                 "blank for anything the whole business uses.",
            group="What you paid"),
        Col("amount", "Amount", 12, "money",
            validate=number_check(-1000000, 1000000), cf=money_cf(),
            note="What left your account, tax included.",
            group="What you paid"),
        Col("tax_amount", "Tax included", 11, "money",
            validate=number_check(0, 1000000),
            note="The sales tax or VAT part, if you can reclaim it.",
            group="What you paid"),
        Col("net", "Net of tax", 11, "money",
            formula=zero("date",
                         'ROUND(IF({amount}{r}="",0,{amount}{r})-'
                         'IF({tax_amount}{r}="",0,{tax_amount}{r}),2)'),
            group="What you paid"),
        Col("method", "How paid", 13, "text",
            validate=list_from("PaymentMethodList"),
            group="What you paid"),
        Col("receipt", "Receipt kept", 11, "flag",
            validate=yesno("Photograph it now. Thermal receipts fade "
                           "to blank in about a year, which is "
                           "exactly when you are asked for them."),
            cf=status_cf(), group="What you paid"),
        Col("reference", "Reference", 14, "text",
            group="What you paid"),

        # ---- how it is treated -------------------------------------------------
        Col("is_direct", "Wedding cost", 11, "flag",
            formula=blank("date",
                          'IF({booking}{r}<>"","Yes",'
                          f'IF({is_direct_cost("{category}{r}")},'
                          '"Yes","No"))'),
            note="A cost caused by one wedding. The booking subtracts "
                 "it from that wedding's profit.",
            group="How it is treated"),
        Col("is_overhead", "Overhead", 10, "flag",
            formula=blank("date",
                          'IF({is_direct}{r}="Yes","No","Yes")'),
            group="How it is treated"),
        Col("is_capital", "Equipment buy", 11, "flag",
            formula=blank("date",
                          f'IF({is_capital("{category}{r}")},'
                          '"Yes","No")'),
            note="Kit you will use for years rather than something "
                 "consumed now. Usually claimed differently.",
            group="How it is treated"),
        Col("recurring", "Repeats", 12, "text",
            validate=list_from("FrequencyList"),
            note="Subscriptions, insurance, your website. Leave blank "
                 "for a one-off.",
            group="How it is treated"),
        Col("monthly", "A month", 11, "money",
            formula=blank("date",
                          'IF({recurring}{r}="","",'
                          + monthly_equivalent("{amount}{r}",
                                               "{recurring}{r}") + ")"),
            note="What a repeating cost really works out at each "
                 "month. A yearly insurance premium and a monthly "
                 "subscription are not comparable until both are "
                 "expressed this way.",
            group="How it is treated"),
        Col("annual", "A year", 11, "money",
            formula=zero("date",
                         'IF({monthly}{r}="",0,'
                         'ROUND({monthly}{r}*12,2))'),
            group="How it is treated"),

        # ---- tax ---------------------------------------------------------------
        Col("business_pct", "Business use", 11, "percent0",
            validate=number_check(0, 1),
            note="For anything used privately too - a phone, a car, "
                 "a room at home. Leave blank to mean entirely "
                 "business.",
            group="Tax"),
        Col("claimable", "Claimable part", 12, "money",
            formula=zero("date",
                         'ROUND({net}{r}*'
                         'IF({business_pct}{r}="",1,'
                         '{business_pct}{r}),2)'),
            note="Your own working figure. What is actually "
                 "deductible is a matter for your accountant and your "
                 "country's rules.",
            group="Tax"),
        Col("tax_category", "Tax category", 16, "text",
            validate=list_from("TaxCategoryList"), group="Tax"),
        Col("notes", "Notes", 24, "wrap", group="Tax"),

        month_col("date"),
        year_col("date"),
        date_num_col("date", "date_num", "Date number"),

        Col("alert", "What to do", 30, "status",
            formula=(
                '=IF({date}{r}="","",'
                'IF({amount}{r}="","Add - no amount",'
                'IF({category}{r}="","Add - no category",'
                'IF({receipt}{r}<>"Yes",'
                'IF({amount}{r}>=ReceiptThreshold,'
                '"Missing - no receipt for a large payment",'
                '"Missing - no receipt"),'
                'IF({is_direct}{r}="Yes",'
                'IF({booking}{r}="",'
                '"Add - wedding cost with no Booking ID","OK"),'
                'IF({business_pct}{r}<>"",'
                'IF({business_pct}{r}<=0,'
                '"Check - nothing claimable on this one","OK"),'
                '"OK"))))))'),
            cf=alert_cf(), group="Tax"),
    ]

    return SheetSpec(
        name=NAME,
        title="Expense Tracker",
        subtitle=("Everything the business pays out, split between "
                  "the weddings that caused it and the overheads that "
                  "did not"),
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=None,
        freeze_cols=3,
        samples=sample("expenses", cfg),
        help_text=(
            "One row per payment. Put a Booking ID on anything caused "
            "by a particular wedding and it comes off that wedding's "
            "profit; leave it blank and it is treated as an overhead "
            "the whole year carries."),
        extra_notes=[
            DISCLAIMER,
            "Photograph receipts the day you get them. Thermal paper "
            "fades to blank in about a year, which is roughly when "
            "somebody asks to see it.",
            "Equipment is flagged separately because most tax systems "
            "treat buying a camera differently from buying memory "
            "cards. Which rules apply to you is a question for your "
            "accountant.",
            "The repeating costs are the ones that creep. Four "
            "subscriptions at nineteen a month is nine hundred a year "
            "and nobody ever notices it leaving.",
        ],
    )
````

### `wedding_photography_spreadsheet/sheets/profit.py`

*277 lines · 2 functions*

> Profit Analysis - what each wedding actually made, line by line.
>
> The booking row gives a profit figure.  This tab shows the working:
> where the money came from, where it went, how many hours it really
> took and what that works out at an hour.  Type a Booking ID and every
> other cell fills itself in.

| Defines | What it does |
|---------|--------------|
| `_b(field)` | A numeric field from the booking, zero when it is missing |
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Profit Analysis').

````python
"""Profit Analysis - what each wedding actually made, line by line.

The booking row gives a profit figure.  This tab shows the working:
where the money came from, where it went, how many hours it really
took and what that works out at an hour.  Type a Booking ID and every
other cell fills itself in.
"""

from __future__ import annotations

from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from
from .common import (alert_cf, bar_cf, blank, date_num_col, id_col,
                     lookup_booking, lookup_booking_num_zero, loss_cf,
                     money_cf, month_col, rank_desc, status_cf, year_col,
                     zero)

NAME = "Profit Analysis"


def _b(field: str) -> str:
    """A numeric field from the booking, zero when it is missing."""
    return lookup_booking_num_zero(field)[1:]


def spec(cfg) -> SheetSpec:
    columns = [
        id_col("booking", "ProfitPrefix", digits=cfg.digits("profit"),
               header="Line ID"),

        Col("booking", "Booking ID", 11, "id",
            validate=list_from("BookingList"),
            note="Type or pick a Booking ID. Everything to the right "
                 "fills itself in - this tab is read-only working.",
            group="The wedding"),
        Col("couple", "Couple", 18, "text",
            formula=lookup_booking("client_name"), group="The wedding"),
        Col("wedding_date", "Wedding date", 12, "date",
            formula=lookup_booking("wedding_date", missing=""),
            group="The wedding"),
        Col("package", "Package", 16, "text",
            formula=lookup_booking("package", missing=""),
            group="The wedding"),
        Col("status", "Status", 13, "status",
            formula=lookup_booking("status", missing=""),
            cf=status_cf(), group="The wedding"),
        Col("venue_type", "Venue type", 12, "text",
            formula=lookup_booking("venue_type", missing=""),
            group="The wedding"),
        Col("season", "Season", 9, "text",
            formula=lookup_booking("season", missing=""),
            group="The wedding"),

        # ---- what came in -----------------------------------------------------
        Col("package_price", "Package", 12, "money",
            formula=zero("booking", _b("package_price")),
            group="What came in"),
        Col("extra_hours", "Extra hours", 11, "money",
            formula=zero("booking", _b("extra_hour_fee")),
            group="What came in"),
        Col("addons", "Add-ons", 11, "money",
            formula=zero("booking", _b("addons")),
            group="What came in"),
        Col("travel_charged", "Travel charged", 12, "money",
            formula=zero("booking", _b("travel_used")),
            group="What came in"),
        Col("discount", "Discount given", 12, "money",
            formula=zero("booking", _b("discount")),
            cf=loss_cf(), group="What came in"),
        Col("after_sales", "Prints and albums", 13, "money",
            formula=zero("booking",
                         'SUMIFS([[Gallery & Delivery.after_total]],'
                         '[[Gallery & Delivery.booking]],'
                         '{booking}{r})'),
            note="Money the gallery earned after delivery. The "
                 "booking does not count it; the year should.",
            group="What came in"),
        Col("tax_collected", "Tax collected", 11, "money",
            formula=zero("booking", _b("tax")),
            note="Never yours. Subtracted before profit.",
            group="What came in"),
        Col("revenue", "Revenue", 12, "money",
            formula=zero("booking",
                         'ROUND({package_price}{r}+{extra_hours}{r}+'
                         '{addons}{r}+{travel_charged}{r}-'
                         '{discount}{r}+{after_sales}{r},2)'),
            cf=money_cf(),
            note="Everything the wedding brought in, before tax and "
                 "before costs.",
            group="What came in"),
        Col("collected", "Actually received", 13, "money",
            formula=zero("booking", _b("collected")),
            group="What came in"),
        Col("outstanding", "Still owed", 11, "money",
            formula=zero("booking", _b("outstanding")),
            cf=loss_cf(), group="What came in"),

        # ---- what went out -----------------------------------------------------
        Col("second_cost", "Second shooter", 12, "money",
            formula=zero("booking", _b("team_cost")),
            group="What went out"),
        Col("editing_cost", "Editing", 11, "money",
            formula=zero("booking", _b("editing_cost")),
            group="What went out"),
        Col("travel_cost", "Travel", 11, "money",
            formula=zero("booking",
                         'SUMIFS([[Travel & Mileage.total_cost]],'
                         '[[Travel & Mileage.booking]],{booking}{r})'),
            group="What went out"),
        Col("album_cost", "Albums and prints", 13, "money",
            formula=zero("booking",
                         'SUMIFS([[Expense Tracker.amount]],'
                         '[[Expense Tracker.booking]],{booking}{r},'
                         '[[Expense Tracker.category]],'
                         '"Albums & prints")'),
            group="What went out"),
        Col("other_direct", "Other wedding costs", 14, "money",
            formula=zero("booking",
                         'SUMIFS([[Expense Tracker.amount]],'
                         '[[Expense Tracker.booking]],{booking}{r})-'
                         '{album_cost}{r}'),
            group="What went out"),
        Col("overhead_share", "Share of overheads", 14, "money",
            formula=zero("booking",
                         'IF(WeddingsThisYear<=0,0,'
                         'ROUND(OverheadsThisYear/'
                         'WeddingsThisYear,2))'),
            note="The year's overheads spread evenly across the "
                 "year's weddings. Rough, but far better than "
                 "pretending insurance and software are free.",
            group="What went out"),
        Col("direct_total", "Direct costs", 12, "money",
            formula=zero("booking",
                         'ROUND({second_cost}{r}+{editing_cost}{r}+'
                         '{travel_cost}{r}+{album_cost}{r}+'
                         '{other_direct}{r},2)'),
            group="What went out"),
        Col("total_cost", "All costs", 12, "money",
            formula=zero("booking",
                         'ROUND({direct_total}{r}+'
                         '{overhead_share}{r},2)'),
            group="What went out"),

        # ---- what it made --------------------------------------------------------
        Col("gross_profit", "Profit before overheads", 16, "money",
            formula=zero("booking",
                         'ROUND({revenue}{r}-{tax_collected}{r}-'
                         '{direct_total}{r},2)'),
            cf=loss_cf(), group="What it made"),
        Col("gross_margin", "Gross margin", 11, "percent0",
            formula=blank("booking",
                          'IF({revenue}{r}<=0,"",'
                          '{gross_profit}{r}/{revenue}{r})'),
            cf=bar_cf(), group="What it made"),
        Col("net_profit", "Profit after overheads", 16, "money",
            formula=zero("booking",
                         'ROUND({revenue}{r}-{tax_collected}{r}-'
                         '{total_cost}{r},2)'),
            cf=loss_cf(),
            note="The real number. What this wedding left you once "
                 "its share of running the business is taken out.",
            group="What it made"),
        Col("net_margin", "Net margin", 10, "percent0",
            formula=blank("booking",
                          'IF({revenue}{r}<=0,"",'
                          '{net_profit}{r}/{revenue}{r})'),
            cf=bar_cf(), group="What it made"),

        # ---- what it took ---------------------------------------------------------
        Col("shoot_hours", "Hours shooting", 12, "number1",
            formula=zero("booking", _b("coverage_hours")),
            group="What it took"),
        Col("edit_hours", "Hours editing", 12, "number1",
            formula=zero("booking",
                         'IFERROR(INDEX([[Editing Workflow.hours]],'
                         'MATCH({booking}{r},'
                         '[[Editing Workflow.booking]],0))*1,'
                         'EditingHours)'),
            group="What it took"),
        Col("travel_hours", "Hours travelling", 13, "number1",
            formula=zero("booking",
                         'SUMIFS([[Travel & Mileage.hours]],'
                         '[[Travel & Mileage.booking]],{booking}{r})'),
            group="What it took"),
        Col("admin_hours", "Hours on admin", 12, "number1",
            formula=zero("booking", "AdminHoursPerWedding"),
            group="What it took"),
        Col("total_hours", "Hours in all", 11, "number1",
            formula=zero("booking",
                         'ROUND({shoot_hours}{r}+{edit_hours}{r}+'
                         '{travel_hours}{r}+{admin_hours}{r},2)'),
            note="Shooting, editing, driving and admin. The whole "
                 "job, not just the day.",
            group="What it took"),
        Col("per_hour", "Profit an hour", 12, "money",
            formula=blank("booking",
                          'IF({total_hours}{r}<=0,"",'
                          'ROUND({net_profit}{r}/'
                          '{total_hours}{r},2))'),
            cf=loss_cf(),
            note="The number worth knowing before you agree to the "
                 "next discount.",
            group="What it took"),
        Col("vs_target", "Against target", 12, "money",
            formula=zero("booking",
                         'ROUND({net_profit}{r}-'
                         '(TargetAverageValue*TargetMargin),2)'),
            cf=loss_cf(), group="What it took"),
        Col("beats_hourly", "Beats your rate", 12, "flag",
            formula=blank("booking",
                          'IF({per_hour}{r}="","",'
                          'IF({per_hour}{r}>=MinimumHourly,'
                          '"Yes","No"))'),
            cf=status_cf(), group="What it took"),
        rank_desc("net_profit", guard='{booking}{r}=""',
                  header="Profit rank"),
        rank_desc("per_hour", guard='{booking}{r}=""',
                  header="Hourly rank"),

        month_col("wedding_date", "month", "Wedding month"),
        year_col("wedding_date"),
        date_num_col("wedding_date", "wed_num", "Wedding date number"),

        Col("verdict", "What it tells you", 34, "status",
            formula=(
                '=IF({booking}{r}="","",'
                'IF({revenue}{r}<=0,"Check - no revenue on this one",'
                'IF({net_profit}{r}<0,'
                '"Loss - this wedding cost more than it earned",'
                'IF({outstanding}{r}>0,'
                '"Unpaid - profitable on paper, not yet in the bank",'
                'IF({per_hour}{r}<>"",'
                'IF({per_hour}{r}<MinimumHourly,'
                '"Check - below the hourly rate you set",'
                'IF({net_margin}{r}<>"",'
                'IF({net_margin}{r}<TargetMargin,'
                '"Check - under your target margin",'
                'IF({discount}{r}>0,'
                '"OK - good wedding, and you discounted it",'
                '"OK - this is the kind you want more of")),'
                '"OK")),"OK"))))) '.strip()),
            cf=alert_cf(), group="What it took"),
    ]

    return SheetSpec(
        name=NAME,
        title="Profit Analysis",
        subtitle=("What each wedding really made once travel, "
                  "editing, the second shooter and a share of the "
                  "overheads come out"),
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=None,
        freeze_cols=3,
        samples=sample("profit", cfg),
        help_text=(
            "Type a Booking ID in column B and the whole line fills "
            "itself in. Nothing else on this tab is typed - it is the "
            "working behind the profit figure on the booking, "
            "including the hours and the overheads the booking "
            "leaves out."),
        extra_notes=[
            "'Profit after overheads' spreads the year's running "
            "costs evenly across the year's weddings. It is a rough "
            "apportionment, and it is still far closer to the truth "
            "than pretending insurance, software and your website "
            "cost nothing.",
            "'Profit an hour' counts shooting, editing, driving and "
            "admin. It is usually a third of what photographers think "
            "it is, and it is the single most useful number in this "
            "workbook.",
            "A wedding can be profitable and still be a problem if it "
            "is unpaid. The verdict column says so rather than "
            "letting a good margin hide a bad debt.",
        ],
    )
````

### `wedding_photography_spreadsheet/sheets/reviews.py`

*179 lines · 1 function*

> Reviews & Referrals - the cheapest marketing there is, tracked.
>
> A wedding photographer's next booking usually comes from the last
> couple.  This tab records who was asked, who wrote something, where it
> went, and which weddings came back out of it.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |

**Module constants:** `NAME` ('Reviews & Referrals').

````python
"""Reviews & Referrals - the cheapest marketing there is, tracked.

A wedding photographer's next booking usually comes from the last
couple.  This tab records who was asked, who wrote something, where it
went, and which weddings came back out of it.
"""

from __future__ import annotations

from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, booking_client_col, blank, date_num_col, id_col, lookup_booking,
                     money_cf, month_col, rating_cf, score_validation,
                     status_cf, year_col, yesno, zero)

NAME = "Reviews & Referrals"


def spec(cfg) -> SheetSpec:
    columns = [
        id_col("booking", "ReviewPrefix", digits=cfg.digits("review"),
               header="Review ID"),

        Col("booking", "Booking ID", 11, "id",
            validate=list_from("BookingList"), group="Who"),
        Col("couple", "Couple", 18, "text",
            formula=lookup_booking("client_name"), group="Who"),
        booking_client_col(),
        Col("wedding_date", "Wedding date", 12, "date",
            formula=lookup_booking("wedding_date", missing=""),
            group="Who"),
        Col("photographer", "Photographer", 12, "id",
            formula=lookup_booking("photographer", missing=""),
            note="Taken from the booking, so the staff tab can show "
                 "an average review per photographer.",
            group="Who"),

        # ---- the ask ----------------------------------------------------------
        Col("asked_on", "Asked on", 11, "date",
            validate=date_check(),
            note="Ask in the fortnight after they first open the "
                 "gallery. Later than that and the answer is usually "
                 "silence.",
            group="The ask"),
        Col("how_asked", "How you asked", 14, "text",
            validate=list_from("ReviewAskMethodList"), group="The ask"),
        Col("reminders", "Reminders", 10, "int",
            validate=number_check(0, 20), group="The ask"),
        Col("status", "Status", 13, "status",
            validate=list_from("ReviewStatusList"), cf=status_cf(),
            group="The ask"),

        # ---- the review --------------------------------------------------------
        Col("received_on", "Received", 11, "date",
            validate=date_check(), group="The review"),
        Col("platform", "Where", 14, "text",
            validate=list_from("ReviewPlatformList"),
            group="The review"),
        Col("rating", "Rating", 9, "rating",
            validate=score_validation(5), cf=rating_cf(),
            group="The review"),
        Col("headline", "Headline", 24, "text", group="The review"),
        Col("quote", "What they said", 34, "wrap",
            note="Paste the useful sentence. This is the one you "
                 "will want for a listing, a website or a post, and "
                 "hunting for it later takes longer than you think.",
            group="The review"),
        Col("can_use", "May you quote it", 13, "flag",
            validate=yesno(), cf=status_cf(), group="The review"),
        Col("used_on", "Used where", 18, "text", group="The review"),
        Col("responded", "You replied", 11, "flag",
            validate=yesno("Reply to every review, including - "
                           "especially - the ones that are not five "
                           "stars."),
            group="The review"),
        Col("days_to_review", "Days to answer", 12, "int",
            formula=blank("booking",
                          'IF({asked_on}{r}="","",'
                          'IF({received_on}{r}="","",'
                          '{received_on}{r}-{asked_on}{r}))'),
            group="The review"),

        # ---- what it led to -------------------------------------------------------
        Col("referrals", "People referred", 12, "int",
            validate=number_check(0, 100),
            note="How many couples they have sent your way.",
            group="What it led to"),
        Col("referred_leads", "Enquiries from them", 14, "int",
            formula=zero("booking",
                         'COUNTIFS([[Lead Tracker.referred_by]],'
                         '{booking}{r})'),
            note="Counted from the Lead Tracker, where an enquiry "
                 "names this Booking ID as the couple who sent them.",
            group="What it led to"),
        Col("referred_bookings", "Weddings from them", 14, "int",
            formula=zero("booking",
                         'COUNTIFS([[Lead Tracker.referred_by]],'
                         '{booking}{r},[[Lead Tracker.is_won]],"Yes")'),
            group="What it led to"),
        Col("referred_value", "Worth to you", 12, "money",
            formula=zero("booking",
                         'SUMIFS([[Lead Tracker.booked_value]],'
                         '[[Lead Tracker.referred_by]],{booking}{r},'
                         '[[Lead Tracker.is_won]],"Yes")'),
            cf=money_cf(),
            note="What the weddings they sent you were worth. It is "
                 "usually the largest number on this tab.",
            group="What it led to"),
        Col("thanked", "Thanked them", 11, "flag",
            validate=yesno("A card, a print, a discount on their "
                           "album. Somebody who sends you a wedding "
                           "should hear from you."),
            group="What it led to"),
        Col("notes", "Notes", 24, "wrap", group="What it led to"),

        month_col("received_on", "month", "Month received"),
        year_col("received_on"),
        date_num_col("received_on", "received_num", "Received number"),
        Col("has_review", "Got one", 9, "flag",
            formula=blank("booking",
                          'IF({received_on}{r}<>"","Yes","No")'),
            hidden=True),

        Col("alert", "What to do", 32, "status",
            formula=(
                '=IF({booking}{r}="","",'
                'IF({received_on}{r}="",'
                'IF({asked_on}{r}="","Send - never asked for one",'
                'IF({status}{r}="Declined","",'
                'IF(TodayDate-{asked_on}{r}>ReviewChaseDays,'
                'IF({reminders}{r}>=2,"Check - asked twice, let it go",'
                '"Follow up - asked a while ago, nothing back"),'
                '"OK"))),'
                'IF({rating}{r}<>"",'
                'IF({rating}{r}<=3,'
                'IF({responded}{r}<>"Yes",'
                '"Urgent - poor review with no reply",'
                '"Check - see what went wrong"),'
                'IF({responded}{r}<>"Yes","Action - reply to it",'
                'IF({can_use}{r}="Yes",'
                'IF({used_on}{r}="",'
                '"Action - good quote going unused","OK"),'
                'IF({referred_bookings}{r}>0,'
                'IF({thanked}{r}<>"Yes",'
                '"Send - they sent you a wedding, thank them",'
                '"OK"),"OK")))),'
                '"Add - review received but no rating")))'),
            cf=alert_cf(), group="What it led to"),
    ]

    return SheetSpec(
        name=NAME,
        title="Reviews & Referrals",
        subtitle=("Who you asked, what they wrote, and the weddings "
                  "that came back out of it"),
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=None,
        freeze_cols=3,
        samples=sample("reviews", cfg),
        help_text=(
            "One row per couple you asked. The gallery tab tells you "
            "when to ask; this one tracks what came back and keeps "
            "the quotes somewhere you can find them. Tag new leads "
            "with the Booking ID of whoever referred them and the "
            "value finds its way back here."),
        extra_notes=[
            "Ask once, remind once, then stop. A third email does "
            "not produce a review, it produces a slightly worse "
            "memory of you.",
            "Reply to every review. A thoughtful reply to a three "
            "star review reads better to the next couple than the "
            "three stars read badly.",
            "The referral column is usually where the real money is. "
            "A couple who sends you two weddings is worth more than "
            "most of your advertising, and costs a card and a print.",
        ],
    )
````

### `wedding_photography_spreadsheet/sheets/reports.py`

*386 lines · 3 functions*

> `Monthly & Annual Reports` - the year in numbers, and the year before.
>
> Two reports on one tab.  The monthly one picks a month and shows what
> came in, what went out and what was left.  The annual one shows twelve
> months side by side, this year against last, with the figures an
> accountant asks for and the figures that tell you whether to put your
> prices up.

| Defines | What it does |
|---------|--------------|
| `_sum_categories(categories, extra)` | Total of several expense categories, with an optional filter |
| `_month_filter(prefix)` |  |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Monthly & Annual Reports'), `LAST_COL` (15), `EARNS` ('[[Wedding Bookings.earns]],"Yes"'), `WED_NUM` ('[[Wedding Bookings.wed_num]]'), `COST_GROUPS` (12 items).

````python
"""`Monthly & Annual Reports` - the year in numbers, and the year before.

Two reports on one tab.  The monthly one picks a month and shows what
came in, what went out and what was left.  The annual one shows twelve
months side by side, this year against last, with the figures an
accountant asks for and the figures that tell you whether to put your
prices up.
"""

from __future__ import annotations

from ..config import MONTH_NAMES, Config, Palette
from ..layout import Registry, col_letter
from ..writer import SheetWriter, number_check

NAME = "Monthly & Annual Reports"
LAST_COL = 15

EARNS = '[[Wedding Bookings.earns]],"Yes"'
WED_NUM = "[[Wedding Bookings.wed_num]]"

#: expense groupings the annual report breaks costs down by
COST_GROUPS = [
    ("Second shooters and assistants",
     ["Second shooters", "Assistants"]),
    ("Albums, prints and postage",
     ["Albums & prints", "Packaging & postage"]),
    ("Editing paid out", ["Editing services"]),
    ("Travel, mileage and staying over",
     ["Travel & mileage", "Accommodation", "Meals on shoot"]),
    ("Equipment bought",
     ["Camera equipment", "Lens purchases", "Lighting & grip",
      "Memory & storage", "Editing computer"]),
    ("Equipment repair and hire",
     ["Equipment repair", "Equipment rental"]),
    ("Software and website",
     ["Software subscriptions", "Website & hosting"]),
    ("Marketing", ["Marketing & advertising", "Client gifts"]),
    ("Insurance, licences and professional fees",
     ["Insurance", "Accounting & legal", "Licences & permits"]),
    ("Studio, office and training",
     ["Studio rent", "Office expenses", "Education & training"]),
    ("Bank and card fees", ["Bank & card fees"]),
    ("Everything else", ["Miscellaneous", "Tax & national insurance"]),
]


def _sum_categories(categories, extra: str = "") -> str:
    """Total of several expense categories, with an optional filter."""
    parts = [f'SUMIFS([[Expense Tracker.amount]],'
             f'[[Expense Tracker.category]],"{c}"{extra})'
             for c in categories]
    return "+".join(parts)


def _month_filter(prefix: str = "Report") -> str:
    return (f'[[Expense Tracker.month]],{prefix}Key')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.deep)

    s.banner("Monthly & Annual Reports",
             "One month in detail, then the whole year beside the "
             "year before it.",
             LAST_COL)
    s.set_col(0, 0, 3)
    s.set_col(1, 1, 34)
    for column in range(2, LAST_COL + 1):
        s.set_col(column, column, 13)

    # ================================================= the monthly report ==
    row = 4
    row = s.section(row, "Monthly report", LAST_COL)
    row = s.field(row, "Month", cfg.today.month, style="int",
                  name="ReportMonth", validate=number_check(1, 12),
                  help_text="1 to 12.", help_col=6)
    row = s.field(row, "Year", cfg.today.year, style="int",
                  name="ReportYear", validate=number_check(2000, 2100),
                  help_text="Both reports read these two cells.",
                  help_col=6)
    row = s.field(row, "Reporting on", None, style="text",
                  name="ReportKey", mode="calc",
                  formula='=TEXT(DATE(ReportYear,ReportMonth,1),'
                          '"yyyy-mm")',
                  help_text="The month key everything groups on.",
                  help_col=6)
    row = s.field(row, "From", None, style="date", name="ReportFrom",
                  mode="calc",
                  formula="=DATE(ReportYear,ReportMonth,1)")
    row = s.field(row, "To", None, style="date", name="ReportTo",
                  mode="calc",
                  formula="=EOMONTH(DATE(ReportYear,ReportMonth,1),0)")
    row += 1

    month_cost = _month_filter()

    s.kpi(row, 1, "Weddings shot",
          f'=COUNTIFS([[Wedding Bookings.month]],ReportKey,{EARNS})',
          "int")
    s.kpi(row, 3, "Value of them",
          f'=SUMIFS([[Wedding Bookings.total_value]],'
          f'[[Wedding Bookings.month]],ReportKey,{EARNS})', "money0")
    s.kpi(row, 5, "Money actually in",
          '=SUMIFS([[Payments & Invoices.received]],'
          '[[Payments & Invoices.month]],ReportKey)', "money0",
          note="Payments received in the month, whatever wedding "
               "they were for")
    s.kpi(row, 7, "Processing fees",
          '=SUMIFS([[Payments & Invoices.fee]],'
          '[[Payments & Invoices.month]],ReportKey)', "money0")
    s.kpi(row, 9, "Spent",
          '=SUMIFS([[Expense Tracker.amount]],'
          '[[Expense Tracker.month]],ReportKey)', "money0")
    row = s.kpi(row, 11, "Left over",
                '=SUMIFS([[Payments & Invoices.received]],'
                '[[Payments & Invoices.month]],ReportKey)-'
                'SUMIFS([[Payments & Invoices.fee]],'
                '[[Payments & Invoices.month]],ReportKey)-'
                'SUMIFS([[Expense Tracker.amount]],'
                '[[Expense Tracker.month]],ReportKey)', "money0",
                note="Cash in less cash out. Not the same as profit "
                     "on the weddings")

    s.kpi(row, 1, "Enquiries",
          '=COUNTIFS([[Lead Tracker.month]],ReportKey)', "int")
    s.kpi(row, 3, "Booked from them",
          '=COUNTIFS([[Lead Tracker.month]],ReportKey,'
          '[[Lead Tracker.is_won]],"Yes")', "int")
    s.kpi(row, 5, "Conversion",
          '=IF(COUNTIFS([[Lead Tracker.month]],ReportKey)<=0,"",'
          'COUNTIFS([[Lead Tracker.month]],ReportKey,'
          '[[Lead Tracker.is_won]],"Yes")/'
          'COUNTIFS([[Lead Tracker.month]],ReportKey))', "percent0")
    s.kpi(row, 7, "Galleries delivered",
          '=COUNTIFS([[Gallery & Delivery.delivered_num]],'
          '">="&ReportFrom,'
          '[[Gallery & Delivery.delivered_num]],"<="&ReportTo)', "int")
    s.kpi(row, 9, "Miles driven",
          '=SUMIFS([[Travel & Mileage.all_miles]],'
          '[[Travel & Mileage.month]],ReportKey)', "number1")
    row = s.kpi(row, 11, "Reviews in",
                '=COUNTIFS([[Reviews & Referrals.month]],ReportKey)',
                "int")
    row += 1

    # ---- where the month's money went ---------------------------------
    row = s.subsection(row, "Where the month's money went", LAST_COL)
    head = row
    s.table_head(head, 1, ["Cost", "This month", "Share",
                           "Same month last year", "Change"],
                 widths=[34, 14, 10, 20, 12])
    row = head + 1
    first_cost = row
    for label, categories in COST_GROUPS:
        s.write(row, 1, label, fmts.chrome("label"))
        s.formula(row, 2, "=" + _sum_categories(categories,
                                                f",{month_cost}"),
                  fmts.cell("money", "calc"))
        s.formula(row, 3,
                  f'=IF(SUMIFS([[Expense Tracker.amount]],'
                  f'[[Expense Tracker.month]],ReportKey)<=0,"",'
                  f'C{row + 1}/SUMIFS([[Expense Tracker.amount]],'
                  f'[[Expense Tracker.month]],ReportKey))',
                  fmts.cell("percent0", "calc"))
        last_year = ('[[Expense Tracker.month]],'
                     'TEXT(DATE(ReportYear-1,ReportMonth,1),"yyyy-mm")')
        s.formula(row, 4,
                  "=" + _sum_categories(categories, f",{last_year}"),
                  fmts.cell("money", "calc"))
        s.formula(row, 5,
                  f'=IF(E{row + 1}<=0,"",C{row + 1}/E{row + 1}-1)',
                  fmts.cell("percent0", "calc"))
        row += 1
    last_cost = row - 1
    s.write(row, 1, "Everything", fmts.chrome("total_label"))
    for column, style in ((2, "money"), (4, "money")):
        letter = col_letter(column)
        s.formula(row, column,
                  f"=SUM({letter}{first_cost + 1}:{letter}"
                  f"{last_cost + 1})",
                  fmts.chrome("total_value",
                              num_format=s.num_format(style)))
    s.formula(row, 5,
              f'=IF(E{row + 1}<=0,"",C{row + 1}/E{row + 1}-1)',
              fmts.chrome("total_value",
                          num_format=s.num_format("percent0")))
    row += 2

    # ---- the weddings in the month ---------------------------------------
    row = s.subsection(row, "The weddings in this month", LAST_COL)
    head = row
    s.table_head(head, 1, ["Date", "Couple", "Package", "Venue",
                           "Value", "Collected", "Outstanding",
                           "Profit", "Margin", "Status"],
                 widths=[12, 20, 16, 18, 12, 12, 12, 12, 9, 14])
    row = head + 1
    for n in range(1, 13):
        rank = (f'MATCH(ReportKey&"|"&{n},'
                f'[[Wedding Bookings.month_key]],0)')
        for col, (field, style) in enumerate([
                ("wedding_date", "date"), ("client_name", "text"),
                ("package", "text"), ("venue", "text"),
                ("total_value", "money"), ("collected", "money"),
                ("outstanding", "money"), ("profit", "money"),
                ("margin", "percent0"), ("status", "status")],
                start=1):
            s.formula(row, col,
                      f'=IFERROR(INDEX([[Wedding Bookings.{field}]],'
                      f'{rank}),"")',
                      fmts.cell(style, "calc"))
        row += 1
    row += 2

    # ================================================== the annual report ==
    row = s.section(row, "Annual report", LAST_COL)
    row = s.note(
        row,
        "Twelve months of the year chosen above, with last year "
        "underneath for comparison. Revenue is counted against the "
        "month the wedding happens, which is how a photographer "
        "thinks about a year; cash received is counted when it "
        "actually arrived, which is how a bank does.",
        LAST_COL, height=38)
    row += 1

    head = row
    s.table_head(head, 1,
                 ["Month"] + list(MONTH_NAMES) + ["Year"],
                 widths=[16] + [9] * 12 + [14])
    row = head + 1

    def month_row(label, template, style, total=True):
        nonlocal row
        s.write(row, 1, label, fmts.chrome("label"))
        for month in range(1, 13):
            key = f'TEXT(DATE(ReportYear,{month},1),"yyyy-mm")'
            s.formula(row, 1 + month, "=" + template.format(key=key),
                      fmts.cell(style, "calc"))
        if total:
            s.formula(row, 14,
                      f"=SUM(C{row + 1}:N{row + 1})",
                      fmts.chrome("total_value",
                                  num_format=s.num_format(style)))
        else:
            s.write(row, 14, None, fmts.cell(style, "calc"))
        row += 1

    month_row("Weddings",
              f'COUNTIFS([[Wedding Bookings.month]],{{key}},{EARNS})',
              "int")
    revenue_row = row
    month_row("Revenue booked",
              f'SUMIFS([[Wedding Bookings.total_value]],'
              f'[[Wedding Bookings.month]],{{key}},{EARNS})', "money")
    month_row("Money received",
              'SUMIFS([[Payments & Invoices.received]],'
              '[[Payments & Invoices.month]],{key})', "money")
    month_row("Direct costs",
              f'SUMIFS([[Wedding Bookings.total_cost]],'
              f'[[Wedding Bookings.month]],{{key}},{EARNS})', "money")
    month_row("Overheads",
              'SUMIFS([[Expense Tracker.amount]],'
              '[[Expense Tracker.month]],{key},'
              '[[Expense Tracker.is_overhead]],"Yes")', "money")
    month_row("Profit",
              f'SUMIFS([[Wedding Bookings.profit]],'
              f'[[Wedding Bookings.month]],{{key}},{EARNS})-'
              f'SUMIFS([[Expense Tracker.amount]],'
              f'[[Expense Tracker.month]],{{key}},'
              f'[[Expense Tracker.is_overhead]],"Yes")', "money")
    month_row("Enquiries",
              'COUNTIFS([[Lead Tracker.month]],{key})', "int")
    month_row("Enquiries booked",
              'COUNTIFS([[Lead Tracker.month]],{key},'
              '[[Lead Tracker.is_won]],"Yes")', "int")
    s.write(row, 1, "Galleries delivered", fmts.chrome("label"))
    for month in range(1, 13):
        s.formula(row, 1 + month,
                  f'=COUNTIFS([[Gallery & Delivery.delivered_num]],'
                  f'">="&DATE(ReportYear,{month},1),'
                  f'[[Gallery & Delivery.delivered_num]],"<="&'
                  f'EOMONTH(DATE(ReportYear,{month},1),0))',
                  fmts.cell("int", "calc"))
    s.formula(row, 14, f"=SUM(C{row + 1}:N{row + 1})",
              fmts.chrome("total_value",
                          num_format=s.num_format("int")))
    row += 1
    month_row("Miles",
              'SUMIFS([[Travel & Mileage.all_miles]],'
              '[[Travel & Mileage.month]],{key})', "number1")

    s.write(row, 1, "Last year's revenue", fmts.chrome("label"))
    for month in range(1, 13):
        key = f'TEXT(DATE(ReportYear-1,{month},1),"yyyy-mm")'
        s.formula(row, 1 + month,
                  f'=SUMIFS([[Wedding Bookings.total_value]],'
                  f'[[Wedding Bookings.month]],{key},{EARNS})',
                  fmts.cell("money", "calc"))
    s.formula(row, 14, f"=SUM(C{row + 1}:N{row + 1})",
              fmts.chrome("total_value",
                          num_format=s.num_format("money")))
    last_year_row = row
    row += 1

    s.write(row, 1, "Change on last year", fmts.chrome("label"))
    for month in range(1, 13):
        letter = col_letter(1 + month)
        s.formula(row, 1 + month,
                  f'=IF({letter}{last_year_row + 1}<=0,"",'
                  f'{letter}{revenue_row + 1}/'
                  f'{letter}{last_year_row + 1}-1)',
                  fmts.cell("percent0", "calc"))
    s.formula(row, 14,
              f'=IF(O{last_year_row + 1}<=0,"",'
              f'O{revenue_row + 1}/O{last_year_row + 1}-1)',
              fmts.chrome("total_value",
                          num_format=s.num_format("percent0")))
    row += 2

    # ---- the year in one glance -------------------------------------------
    row = s.subsection(row, "The year in one glance", LAST_COL)
    year_earns = f'[[Wedding Bookings.year]],ReportYear,{EARNS}'
    s.kpi(row, 1, "Weddings",
          f'=COUNTIFS({year_earns})', "int")
    s.kpi(row, 3, "Against target",
          f'=IF(TargetWeddings<=0,"",COUNTIFS({year_earns})/'
          f'TargetWeddings)', "percent0")
    s.kpi(row, 5, "Revenue",
          f'=SUMIFS([[Wedding Bookings.total_value]],{year_earns})',
          "money0")
    s.kpi(row, 7, "Average wedding",
          f'=IF(COUNTIFS({year_earns})<=0,"",'
          f'SUMIFS([[Wedding Bookings.total_value]],{year_earns})/'
          f'COUNTIFS({year_earns}))', "money0")
    s.kpi(row, 9, "Profit",
          f'=SUMIFS([[Wedding Bookings.profit]],{year_earns})-'
          f'SUMIFS([[Expense Tracker.amount]],'
          f'[[Expense Tracker.year]],ReportYear,'
          f'[[Expense Tracker.is_overhead]],"Yes")', "money0")
    row = s.kpi(row, 11, "Margin",
                f'=IF(SUMIFS([[Wedding Bookings.total_value]],'
                f'{year_earns})<=0,"",'
                f'(SUMIFS([[Wedding Bookings.profit]],{year_earns})-'
                f'SUMIFS([[Expense Tracker.amount]],'
                f'[[Expense Tracker.year]],ReportYear,'
                f'[[Expense Tracker.is_overhead]],"Yes"))/'
                f'SUMIFS([[Wedding Bookings.total_value]],'
                f'{year_earns}))', "percent0")

    s.kpi(row, 1, "Hours shooting",
          f'=SUMIFS([[Wedding Bookings.coverage_hours]],{year_earns})',
          "number1")
    s.kpi(row, 3, "Hours editing",
          '=SUMIFS([[Editing Workflow.hours]],'
          '[[Editing Workflow.booking]],"?*")', "number1",
          note="All recorded editing, not only this year's")
    s.kpi(row, 5, "Equipment owned",
          '=SUM([[Equipment Inventory.price]])', "money0",
          note="What it cost, not what it is worth now")
    s.kpi(row, 7, "Insured value",
          '=SUMIFS([[Equipment Inventory.insured_value]],'
          '[[Equipment Inventory.insured]],"Yes")', "money0")
    s.kpi(row, 9, "Claimable expenses",
          '=SUMIFS([[Expense Tracker.claimable]],'
          '[[Expense Tracker.year]],ReportYear)', "money0",
          note="Your own working figure - not tax advice")
    row = s.kpi(row, 11, "Tax collected",
                f'=SUMIFS([[Wedding Bookings.tax]],{year_earns})',
                "money0",
                note="Only yours to pass on, never to spend")
    row += 1

    row = s.note(
        row,
        "These figures are a management summary built from what you "
        "typed, not a set of accounts. Which costs are deductible, "
        "how equipment is written down and when revenue is "
        "recognised all depend on your country's rules and on how "
        "your business is set up. Give your accountant the Expense "
        "Tracker and the Payments tab, not this page.",
        LAST_COL, height=48)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    s.protect()
    return s
````

---

## 10. The quality tools

Not shipped with the workbook, but the reason it can be trusted: a static validator that checks the Booking ID join in both directions and enforces the promises the listing makes, a bracket and placeholder checker, a specification checker and a sample data checker that both run before the file is written, and a recalculation harness that really evaluates every formula.

### `tools/validate_wedding_workbook.py`

*1,019 lines · 17 functions*

> Static checks on a generated Wedding Photographer Business Manager.
>
> Run it with any Python that has openpyxl installed::
>
>     python tools/validate_wedding_workbook.py output/wedding.xlsx
>
> It never opens Excel - it parses the XML and reports anything that would
> break for a buyer: unresolved template tokens, formulas pointing at the
> wrong sheet, defined names that do not exist, missing sheet protection,
> formula cells that are not locked, input cells that are, and - the one
> that matters most here - any break in the Booking ID chain that holds
> the CRM, the payments, the timeline, the shot list, the editing queue
> and the gallery together.

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
| `check_chain(wb, problems, notes)` | The five workflows only hold while every one of these does |
| `check_booking_join(wb, problems, notes)` | The Booking ID is the join key |
| `check_promises(wb, problems, notes)` | Three promises this product makes to the buyer in writing |
| `check(path)` |  |

**Module constants:** `FUNCTION_RE` (re.compile(...)), `NAME_RE` (re.compile(...)), `BARE_COL_RE` (re.compile(...)), `KNOWN_FUNCTIONS` ({'IF', 'IFERROR', 'SUM', 'SUMIF', 'SUMIFS', 'COUNT', 'COU...), `MODERN_ONLY` ({'XLOOKUP', 'IFS', 'MAXIFS', 'MINIFS', 'TEXTJOIN', 'CONCA...), `PRECEDENCE_OK` (0 entries), `TIME_DIVISOR_RE` (re.compile(...)), `RANGE_RE` (re.compile(...)), `SEEDED_INPUTS` ({'Booking Calendar!C10', 'Client Questionnaire!C10', 'Mon...), `CHAIN_CHECKS` (79 items), `PAGE_FORMULAS` (7 items), `PAGE_INPUTS` (31 items), `REQUIRED_LIST_NAMES` (48 items).

````python
"""Static checks on a generated Wedding Photographer Business Manager.

Run it with any Python that has openpyxl installed::

    python tools/validate_wedding_workbook.py output/wedding.xlsx

It never opens Excel - it parses the XML and reports anything that would
break for a buyer: unresolved template tokens, formulas pointing at the
wrong sheet, defined names that do not exist, missing sheet protection,
formula cells that are not locked, input cells that are, and - the one
that matters most here - any break in the Booking ID chain that holds
the CRM, the payments, the timeline, the shot list, the editing queue
and the gallery together.
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
from wedding_photography_spreadsheet.build import (                # noqa: E402
    SHEET_ORDER as EXPECTED_SHEETS)

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

#: Dividing by one of these is a unit conversion, not a mis-bracket.
#: A wedding timeline is built as `ceremony_time + minutes/1440` -
#: a date serial plus a fraction of a day - and bracketing it as
#: `(ceremony_time + minutes)/1440` would be the actual bug.
TIME_DIVISOR_RE = re.compile(r"/(?:1440|86400|24|60)\b")


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
        if TIME_DIVISOR_RE.search(argument):
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

#: value cells that are deliberately left unlocked even though a
#: formula sits in them - the buyer is meant to type over the top
SEEDED_INPUTS = {
    "Booking Calendar!C10",       # the month the calendar opens on
    "Client Questionnaire!C10",   # the wedding being printed
    "Monthly & Annual Reports!C10",   # the month being reported on
}

#: Every link below is a formula in row 5 of a register.  Together
#: they are the product: break one and a buyer gets a blank cell
#: instead of an answer, with nothing to tell them why.
#:
#: The Booking ID (WED-nnnn) is the join key.  Nine registers reach
#: back to Wedding Bookings through it and Wedding Bookings reaches
#: forward to six of them, so most of what follows is that one
#: decision being checked in both directions.
CHAIN_CHECKS = [
    # ---- Wedding Bookings: the hub every other tab joins to -------------
    ("Wedding Bookings", "id", ("text", "BookingPrefix")),
    ("Wedding Bookings", "client_name", ("sheet", "Client CRM")),
    ("Wedding Bookings", "package_price", ("sheet", "Packages & Pricing")),
    ("Wedding Bookings", "photographer_name", ("sheet",
                                               "Photographer & Staff")),
    ("Wedding Bookings", "collected", ("sheet", "Payments & Invoices")),
    ("Wedding Bookings", "contract_state", ("sheet",
                                            "Contracts & Documents")),
    ("Wedding Bookings", "travel_used", ("sheet", "Travel & Mileage")),
    ("Wedding Bookings", "editing_cost", ("sheet", "Editing Workflow")),
    ("Wedding Bookings", "subtotal", ("column", "Wedding Bookings",
                                      "package_price")),
    ("Wedding Bookings", "total_value", ("column", "Wedding Bookings",
                                         "subtotal")),
    ("Wedding Bookings", "total_cost", ("column", "Wedding Bookings",
                                        "team_cost")),
    ("Wedding Bookings", "profit", ("column", "Wedding Bookings",
                                    "total_cost")),
    ("Wedding Bookings", "margin", ("column", "Wedding Bookings",
                                    "total_value")),
    ("Wedding Bookings", "outstanding", ("column", "Wedding Bookings",
                                         "collected")),
    ("Wedding Bookings", "paid_pct", ("column", "Wedding Bookings",
                                      "collected")),
    ("Wedding Bookings", "team_cost", ("text", "SecondShooterRate")),
    ("Wedding Bookings", "clash", ("column", "Wedding Bookings",
                                   "wed_num")),
    ("Wedding Bookings", "team_clash", ("column", "Wedding Bookings",
                                        "photographer")),
    ("Wedding Bookings", "alert", ("text", "UpcomingDays")),
    ("Wedding Bookings", "upcoming_rank", ("column", "Wedding Bookings",
                                           "upcoming_key")),
    ("Wedding Bookings", "month_key", ("column", "Wedding Bookings",
                                       "month_seq")),
    ("Wedding Bookings", "timeline_lines", ("sheet", "Wedding Timeline")),
    ("Wedding Bookings", "shot_lines", ("sheet", "Shot List Manager")),
    # ---- enquiry to booking: workflow one -------------------------------
    ("Lead Tracker", "id", ("text", "LeadPrefix")),
    ("Lead Tracker", "booked_value", ("sheet", "Wedding Bookings")),
    ("Lead Tracker", "is_won", ("column", "Lead Tracker", "status")),
    ("Lead Tracker", "overdue", ("text", "TodayDate")),
    ("Lead Tracker", "alert", ("text", "LeadChaseDays")),
    ("Client CRM", "id", ("text", "ClientPrefix")),
    # ---- package to payment: workflow two -------------------------------
    ("Packages & Pricing", "total_price", ("column", "Packages & Pricing",
                                           "base_price")),
    ("Packages & Pricing", "total_cost", ("column", "Packages & Pricing",
                                          "album_cost")),
    ("Packages & Pricing", "margin", ("column", "Packages & Pricing",
                                      "total_price")),
    ("Packages & Pricing", "hourly", ("column", "Packages & Pricing",
                                      "total_hours")),
    ("Packages & Pricing", "sold", ("sheet", "Wedding Bookings")),
    ("Packages & Pricing", "revenue", ("sheet", "Wedding Bookings")),
    ("Packages & Pricing", "conversion", ("column", "Packages & Pricing",
                                          "sold")),
    ("Packages & Pricing", "alert", ("text", "MinimumHourly")),
    ("Payments & Invoices", "id", ("text", "InvoicePrefix")),
    ("Payments & Invoices", "client", ("sheet", "Wedding Bookings")),
    ("Payments & Invoices", "balance", ("column", "Payments & Invoices",
                                        "received")),
    ("Payments & Invoices", "is_late", ("column", "Payments & Invoices",
                                        "days_overdue")),
    ("Payments & Invoices", "alert", ("text", "TodayDate")),
    ("Contracts & Documents", "is_done", ("column",
                                          "Contracts & Documents",
                                          "signed_date")),
    ("Contracts & Documents", "alert", ("text", "DocumentChaseDays")),
    # ---- the wedding day: workflow three --------------------------------
    ("Wedding Timeline", "start", ("column", "Wedding Timeline",
                                   "offset_used")),
    ("Wedding Timeline", "finish", ("column", "Wedding Timeline",
                                    "start")),
    ("Wedding Timeline", "staff_name", ("sheet", "Photographer & Staff")),
    ("Wedding Timeline", "day_key", ("column", "Wedding Timeline",
                                     "day_order")),
    ("Wedding Timeline", "alert", ("text", "TodayDate")),
    ("Shot List Manager", "status", ("column", "Shot List Manager",
                                     "done")),
    ("Shot List Manager", "print_key", ("column", "Shot List Manager",
                                        "print_seq")),
    ("Shot List Manager", "alert", ("text", "TodayDate")),
    # ---- editing to delivery: workflow four -----------------------------
    ("Editing Workflow", "stage", ("column", "Editing Workflow",
                                   "step_backed_up")),
    ("Editing Workflow", "cost", ("column", "Editing Workflow", "hours")),
    ("Editing Workflow", "due", ("text", "DeliveryDays")),
    ("Editing Workflow", "late", ("column", "Editing Workflow", "due")),
    ("Editing Workflow", "days_taken", ("column", "Editing Workflow",
                                        "finished")),
    ("Editing Workflow", "alert", ("text", "TodayDate")),
    ("Gallery & Delivery", "client", ("sheet", "Wedding Bookings")),
    ("Gallery & Delivery", "expires", ("text", "GalleryLiveDays")),
    ("Gallery & Delivery", "after_total", ("column", "Gallery & Delivery",
                                           "print_sales")),
    ("Gallery & Delivery", "alert", ("text", "TodayDate")),
    # ---- money per wedding: workflow five -------------------------------
    ("Profit Analysis", "couple", ("sheet", "Wedding Bookings")),
    ("Profit Analysis", "net_profit", ("column", "Profit Analysis",
                                       "total_cost")),
    ("Profit Analysis", "net_margin", ("column", "Profit Analysis",
                                       "revenue")),
    ("Profit Analysis", "per_hour", ("column", "Profit Analysis",
                                     "net_profit")),
    ("Profit Analysis", "net_profit_rank", ("column", "Profit Analysis",
                                            "net_profit")),
    ("Expense Tracker", "is_overhead", ("column", "Expense Tracker",
                                        "is_direct")),
    ("Expense Tracker", "alert", ("text", "ReceiptThreshold")),
    ("Travel & Mileage", "all_miles", ("column", "Travel & Mileage",
                                       "miles")),
    ("Travel & Mileage", "chargeable", ("column", "Travel & Mileage",
                                        "suggested")),
    ("Travel & Mileage", "total_cost", ("column", "Travel & Mileage",
                                        "mileage")),
    ("Travel & Mileage", "alert", ("text", "TravelLossLimit")),
    # ---- the rest of the business ---------------------------------------
    ("Equipment Inventory", "alert", ("text", "InsureAbove")),
    ("Marketing Calendar", "id", ("text", "MarketingPrefix")),
    ("Marketing Calendar", "alert", ("text", "TodayDate")),
    ("Reviews & Referrals", "client", ("sheet", "Wedding Bookings")),
    ("Reviews & Referrals", "has_review", ("column", "Reviews & Referrals",
                                           "received_on")),
    ("Reviews & Referrals", "alert", ("text", "ReviewChaseDays")),
]

#: named cells that must still hold a live formula - a name pointing
#: at an empty cell means a page quietly stopped working
PAGE_FORMULAS = [
    ("TodayDate", "settings: the live date every alert reads"),
    ("ClientCount", "settings: couples on the books"),
    ("OverheadsThisYear", "settings: overheads totalled from the expense tracker"),
    ("WeddingsThisYear", "settings: weddings counted from the bookings register"),
    ("ReportKey", "reports: the month key the totals group on"),
    ("ReportFrom", "reports: the first day of the month reported on"),
    ("ReportTo", "reports: the last day of the month reported on"),
]

#: cells a buyer has to be able to type in
PAGE_INPUTS = [
    ("StudioName", "settings: the studio name"),
    ("DepositPercent", "settings: the deposit that holds a date"),
    ("FinalPaymentDays", "settings: when the balance falls due"),
    ("ExtraHourRate", "settings: the charge for an extra hour"),
    ("SecondShooterRate", "settings: what a second shooter costs"),
    ("EngagementCost", "settings: what an engagement shoot costs"),
    ("MileageRate", "settings: the mileage rate"),
    ("FreeTravelMiles", "settings: miles included before charging"),
    ("EditingHours", "settings: honest editing hours a wedding"),
    ("EditingRate", "settings: what an hour of editing costs"),
    ("OwnerHourlyValue", "settings: what your own hour is worth"),
    ("DeliveryDays", "settings: the gallery turnaround you promise"),
    ("PreviewDays", "settings: the sneak peek turnaround"),
    ("AlbumDays", "settings: the album turnaround"),
    ("GalleryLiveDays", "settings: how long a gallery stays up"),
    ("TargetMargin", "settings: the margin you are aiming at"),
    ("TargetWeddings", "settings: weddings a year you want"),
    ("TargetAverageValue", "settings: the average booking you want"),
    ("MinimumHourly", "settings: the hourly rate you refuse to go under"),
    ("TaxRate", "settings: sales tax or VAT on packages"),
    ("TaxOnPackages", "settings: whether tax applies at all"),
    ("LeadChaseDays", "settings: when an unanswered enquiry warns"),
    ("UpcomingDays", "settings: the wedding-soon window"),
    ("ReviewChaseDays", "settings: when to chase a review"),
    ("BookingPrefix", "settings: the Booking ID prefix"),
    ("ClientPrefix", "settings: the Client ID prefix"),
    ("CalendarStart", "calendar: the first month shown in the grid"),
    ("CalendarMonths", "calendar: how many months the grid shows"),
    ("QuestionnaireBooking", "questionnaire: the wedding being printed"),
    ("ReportMonth", "reports: the month being reported on"),
    ("ReportYear", "reports: the year being reported on"),
]

#: every drop down on the workbook reads one of these - a missing one
#: means a blank list and a column nobody can fill in
REQUIRED_LIST_NAMES = [
    "BookingList", "ClientList", "PackageList", "StaffList",
    "CampaignList", "YesNoList", "BookingStatusList",
    "ClientStatusList", "LeadStatusList", "LeadSourceList",
    "PriorityList", "TierList", "VenueTypeList", "WeddingStyleList",
    "PhotoStyleList", "DeliveryPreferenceList", "DocumentTypeList",
    "DocumentStatusList", "PaymentStageList", "PaymentMethodList",
    "ShotGroupList", "ShotWhenList", "RequestSourceList",
    "StaffRoleList", "StaffStatusList", "StaffTypeList",
    "GalleryStatusList", "GalleryPlatformList",
    "EquipmentCategoryList", "EquipmentStatusList",
    "EquipmentConditionList", "TravelModeList", "TravelPurposeList",
    "MarketingPlatformList", "ContentTypeList", "ContentStatusList",
    "ExpenseCategoryList", "TaxCategoryList", "FrequencyList",
    "ReviewStatusList", "ReviewPlatformList", "ReviewAskMethodList",
    "CountryList", "TimelineEventList", "TimelineOffsetList",
    "TimelineDurationList", "TimelineRoleList", "MonthNameList",
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
    from wedding_photography_spreadsheet.build import REGISTER_MODULES
    from wedding_photography_spreadsheet.config import Config
    cfg = Config()
    return {module.NAME: module.spec(cfg) for module in REGISTER_MODULES}


def check_chain(wb, problems: list, notes: list) -> None:
    """The five workflows only hold while every one of these does.

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


def check_booking_join(wb, problems: list, notes: list) -> None:
    """The Booking ID is the join key. Verify it in both directions.

    This is the one technical decision the whole product rests on: a
    unique WED-nnnn per wedding, generated once on Wedding Bookings and
    chosen from a drop-down everywhere else.  Without it the payments
    tab, the timeline, the shot list, the editing queue and the gallery
    are five unrelated lists that happen to be in the same file.

    Three things have to hold:

    1. Wedding Bookings really does mint the ID from BookingPrefix;
    2. every register that joins back does so through a column
       validated against BookingList, so a buyer cannot mistype it;
    3. that drop-down resolves to a real defined name.
    """
    specs = _specs()
    hub = specs.get("Wedding Bookings")
    if hub is None:
        problems.append("Wedding Bookings register has gone")
        return
    try:
        column = hub.index("id") + 1
    except Exception:                                  # noqa: B902
        problems.append("Wedding Bookings has no id column - there is "
                        "no join key left")
        return
    formula = wb["Wedding Bookings"].cell(row=5, column=column).value
    if not isinstance(formula, str) or "BookingPrefix" not in formula:
        problems.append("Wedding Bookings.id no longer builds the "
                        "Booking ID from BookingPrefix")

    defined = set(wb.defined_names)
    for ws in wb.worksheets:
        defined |= set(ws.defined_names)
    if "BookingList" not in defined:
        problems.append("BookingList is missing - every Booking ID "
                        "drop-down in the workbook is now empty")

    joined = []
    for name, spec in specs.items():
        if name == "Wedding Bookings":
            continue
        col = next((c for c in spec.columns if c.key == "booking"), None)
        if col is None:
            continue
        joined.append(name)
        source = (col.validate or {}).get("source")
        if not (isinstance(source, str)
                and source.lstrip("=") == "BookingList"):
            problems.append(
                f"{name}.booking is not a BookingList drop-down, so a "
                f"buyer can mistype the Booking ID and silently "
                f"disconnect the row")
    if len(joined) < 9:
        problems.append(f"only {len(joined)} registers join through a "
                        f"Booking ID; the design calls for at least 9")
    notes.append(f"booking ID join: {len(joined)} registers joined to "
                 f"Wedding Bookings through BookingList")


def check_promises(wb, problems: list, notes: list) -> None:
    """Three promises this product makes to the buyer in writing.

    1. It is an organiser, not an integration: it must never claim to
       sync with a booking platform or a calendar service.
    2. It tracks documents; it does not pretend to generate or execute
       legally binding contracts.
    3. The accounting figures carry a disclaimer pointing at a real
       accountant.
    """
    banned_sync = ("syncs with", "synchronise with", "synchronize with",
                   "live feed", "api connection", "automatically imports",
                   "real-time booking", "connects to your calendar")
    banned_legal = ("legally binding", "legally-binding",
                    "e-signature", "esignature", "electronic signature",
                    "court", "solicitor-approved", "lawyer-approved")
    sync_hits = legal_hits = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                value = cell.value
                if not isinstance(value, str) or value.startswith("="):
                    continue
                low = value.lower()
                negated = ("not " in low or "never " in low
                           or "cannot" in low or "does not" in low
                           or "no " in low)
                for phrase in banned_sync:
                    if phrase in low and not negated:
                        sync_hits += 1
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: claims a "
                            f"live booking connection -> {value[:70]}")
                for phrase in banned_legal:
                    if phrase in low and not negated:
                        legal_hits += 1
                        problems.append(
                            f"{ws.title}!{cell.coordinate}: claims to "
                            f"produce a binding contract -> "
                            f"{value[:70]}")

    wanted = {
        "organiser": ("organis", "organiz"),
        "not a contract tool": ("not a contract", "does not replace",
                                "track documents", "tracks documents",
                                "not legal advice"),
        "accountant": ("accountant",),
    }
    found = {key: False for key in wanted}
    for title in ("Getting Started", "Settings & Lists",
                  "Monthly & Annual Reports", "Contracts & Documents"):
        if title not in wb.sheetnames:
            continue
        for row in wb[title].iter_rows():
            for cell in row:
                value = cell.value
                if not isinstance(value, str):
                    continue
                low = value.lower()
                for key, needles in wanted.items():
                    if any(n in low for n in needles):
                        found[key] = True
    for key, ok in found.items():
        if not ok:
            problems.append(f"the workbook never states the "
                            f"'{key}' honesty note")
    notes.append(f"product promises: {sync_hits} false sync claims, "
                 f"{legal_hits} binding-contract claims, honesty notes "
                 f"{sum(found.values())}/{len(found)} present")


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
    check_booking_join(wb, problems, notes)
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
              else "output/Wedding Photographer Business Manager.xlsx")
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

### `tools/check_wedding_specs.py`

*170 lines · 5 functions*

> Fast structural sweep over the wedding product's register specs.
>
> Runs in well under a second and catches, before a workbook is ever
> built, the four mistakes that are expensive to find later:
>
> 1. a formula whose brackets do not balance (concatenated nested IFs
>    lose a closing paren silently);
> 2. a `{token}` that names no column on that sheet;
> 3. `a+b/c` where `(a+b)/c` was meant;
> 4. a COUNTIFS/SUMIFS criteria range that carries its own test.
>
>     python3 tools/check_wedding_specs.py

| Defines | What it does |
|---------|--------------|
| `_strings_out(text)` |  |
| `_args(text, start)` | Split the argument list of a call whose '(' is at `start` |
| `collapsed_criteria(formula)` |  |
| `precedence_slip(formula)` | `a+b/c` at the top level of a division - almost always a bug |
| `main()` |  |

**Module constants:** `TOKEN` (re.compile(...)), `CALL` (re.compile(...)), `ENGINE_TOKENS` ({'r', 'r_prev', 'first', 'last'}), `TIME_DIVISORS` (4 items).

````python
#!/usr/bin/env python3
"""Fast structural sweep over the wedding product's register specs.

Runs in well under a second and catches, before a workbook is ever
built, the four mistakes that are expensive to find later:

1. a formula whose brackets do not balance (concatenated nested IFs
   lose a closing paren silently);
2. a `{token}` that names no column on that sheet;
3. `a+b/c` where `(a+b)/c` was meant;
4. a COUNTIFS/SUMIFS criteria range that carries its own test.

    python3 tools/check_wedding_specs.py
"""

from __future__ import annotations

import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from wedding_photography_spreadsheet.config import Config      # noqa: E402

TOKEN = re.compile(r"\{([a-z_][a-z0-9_]*)\}")
CALL = re.compile(r"\b(COUNTIFS|SUMIFS|AVERAGEIFS)\s*\(")

#: tokens the engine substitutes that are not column keys
ENGINE_TOKENS = {"r", "r_prev", "first", "last"}


def _strings_out(text: str) -> str:
    return re.sub(r'"[^"]*"', '""', text)


def _args(text: str, start: int):
    """Split the argument list of a call whose '(' is at `start`."""
    depth, out, cur, i = 0, [], "", start
    while i < len(text):
        ch = text[i]
        if ch == "(":
            depth += 1
            if depth == 1:
                i += 1
                continue
        elif ch == ")":
            depth -= 1
            if depth == 0:
                out.append(cur)
                return out, i
        if depth == 1 and ch == ",":
            out.append(cur)
            cur = ""
        else:
            cur += ch
        i += 1
    return out, i


def collapsed_criteria(formula: str) -> str:
    body = _strings_out(formula)
    for match in CALL.finditer(body):
        name = match.group(1)
        args, _ = _args(body, match.end() - 1)
        first = 1 if name in ("SUMIFS", "AVERAGEIFS") else 0
        if (len(args) - first) % 2:
            return f"{name}: odd number of range/criteria arguments"
        for text in args[first::2]:
            if re.search(r"[<>=]", text):
                return f"{name}: criteria range carries a test -> {text[:40]}"
    return ""


#: Dividing by one of these is a unit conversion, not a mis-bracket.
#: A wedding timeline is `ceremony_time + minutes/1440` - a date serial
#: plus a fraction of a day - and bracketing it as
#: `(ceremony_time + minutes)/1440` would be the actual bug.
TIME_DIVISORS = ("1440", "86400", "24", "60")


def precedence_slip(formula: str) -> str:
    """`a+b/c` at the top level of a division - almost always a bug."""
    body = _strings_out(formula)
    for match in re.finditer(r"/", body):
        tail = body[match.start() + 1:match.start() + 8]
        if re.match(r"(%s)\b" % "|".join(TIME_DIVISORS), tail):
            continue
        # walk left over the numerator at depth 0
        i, depth = match.start() - 1, 0
        while i >= 0:
            ch = body[i]
            if ch == ")":
                depth += 1
            elif ch == "(":
                if depth == 0:
                    break
                depth -= 1
            elif depth == 0 and ch in "+-" and i > 0 and \
                    body[i - 1] not in "(,*/+-<>=&":
                return body[max(0, i - 25):match.start() + 12]
            i -= 1
    return ""


def main() -> int:
    from wedding_photography_spreadsheet.build import REGISTER_MODULES

    cfg = Config()
    problems = 0
    total_cols = 0
    total_formulas = 0

    for module in REGISTER_MODULES:
        spec = module.spec(cfg)
        keys = {c.key for c in spec.columns}
        formulas = [c for c in spec.columns if c.formula]
        total_cols += len(spec.columns)
        total_formulas += len(formulas)

        for column in formulas:
            text = column.formula
            if not isinstance(text, str):
                problems += 1
                print(f"{spec.name}.{column.key}: formula is "
                      f"{type(text).__name__}, not a string")
                continue
            where = f"{spec.name}.{column.key}"

            body = _strings_out(text)
            if body.count("(") != body.count(")"):
                problems += 1
                print(f"{where}: brackets {body.count('(')} open vs "
                      f"{body.count(')')} close")
            if text.count('"') % 2:
                problems += 1
                print(f"{where}: odd number of quotes")
            if text != text.strip():
                problems += 1
                print(f"{where}: stray whitespace around the formula")

            for token in set(TOKEN.findall(text)):
                if token in ENGINE_TOKENS or token in keys:
                    continue
                if token.endswith("_first") and token[:-6] in keys:
                    continue
                if token.endswith("_last") and token[:-5] in keys:
                    continue
                problems += 1
                print(f"{where}: {{{token}}} is not a column on this sheet")

            slip = collapsed_criteria(text)
            if slip:
                problems += 1
                print(f"{where}: {slip}")
            slip = precedence_slip(text)
            if slip:
                problems += 1
                print(f"{where}: a+b/c where (a+b)/c was meant -> {slip}")

        print(f"  {spec.name:26s} {len(spec.columns):3d} cols, "
              f"{len(formulas):3d} formulas, {spec.rows:5d} rows")

    print(f"\n{len(REGISTER_MODULES)} registers, {total_cols} columns, "
          f"{total_formulas} formula columns, {problems} problems")
    return 1 if problems else 0


if __name__ == "__main__":
    raise SystemExit(main())
````

### `tools/check_wedding_sample.py`

*146 lines · 3 functions*

> Validate the wedding product's example data before it is built.
>
> Four checks, every one of which has caught a real bug in this repo:
>
> 1. every key in a sample row is an *input* column on that register -
>    a value written into a formula column is silently dropped and takes
>    the formula for that row with it;
> 2. every value that feeds a drop-down exists in that drop-down's list -
>    otherwise Excel flags the cell the moment a buyer clicks it (199 of
>    these shipped undetected in an earlier product);
> 3. every ID that feeds a live list (Booking ID, Client ID, Staff ID,
>    package name, campaign ID) actually exists on the register it points
>    at - a Booking ID typo makes every downstream lookup return blank;
> 4. the rows fit in the register.
>
>     python3 tools/check_wedding_sample.py

| Defines | What it does |
|---------|--------------|
| `main()` |  |
| `_prefix_for(cfg, sheet)` |  |
| `_digits_for(cfg, sheet)` |  |

**Module constants:** `ID_PATTERN` (re.compile(...)), `DYNAMIC_LISTS` (CROSS_SHEET_NAMES).

````python
#!/usr/bin/env python3
"""Validate the wedding product's example data before it is built.

Four checks, every one of which has caught a real bug in this repo:

1. every key in a sample row is an *input* column on that register -
   a value written into a formula column is silently dropped and takes
   the formula for that row with it;
2. every value that feeds a drop-down exists in that drop-down's list -
   otherwise Excel flags the cell the moment a buyer clicks it (199 of
   these shipped undetected in an earlier product);
3. every ID that feeds a live list (Booking ID, Client ID, Staff ID,
   package name, campaign ID) actually exists on the register it points
   at - a Booking ID typo makes every downstream lookup return blank;
4. the rows fit in the register.

    python3 tools/check_wedding_sample.py
"""

from __future__ import annotations

import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from wedding_photography_spreadsheet.config import (Config,          # noqa: E402
                                                    TIMELINE_EVENTS)
from wedding_photography_spreadsheet.build import REGISTER_MODULES   # noqa: E402
from wedding_photography_spreadsheet.names import CROSS_SHEET_NAMES  # noqa: E402
from wedding_photography_spreadsheet.sheets.settings import LISTS    # noqa: E402

ID_PATTERN = re.compile(r"^[A-Z]{3}-\d+$")

#: which register and ID column each live drop-down reads from
DYNAMIC_LISTS = CROSS_SHEET_NAMES


def main() -> int:
    cfg = Config()
    lists = {name: set(values) for name, values in LISTS}
    # the Settings blocks that are written row by row rather than as
    # a plain list literal
    lists["TimelineEventList"] = {label for _k, label, *_r
                                  in TIMELINE_EVENTS}
    lists["TimelineRoleList"] = {role for *_r, role in TIMELINE_EVENTS}

    specs = {}
    for module in REGISTER_MODULES:
        spec = module.spec(cfg)
        specs[spec.name] = spec

    # Build the live contents of every dynamic list (BookingList and
    # friends) out of the sample data itself, so a Booking ID typo in
    # Payments is caught here rather than by a buyer wondering why a
    # row is blank.
    live: dict = {}
    for list_name, (sheet, key) in DYNAMIC_LISTS.items():
        spec = specs.get(sheet)
        if spec is None:
            continue
        column = next((c for c in spec.columns if c.key == key), None)
        values = set()
        if column is not None and column.formula:
            # an auto-numbered ID column: PREFIX-0001 per sample row
            prefix = _prefix_for(cfg, sheet)
            if prefix:
                values = {f"{prefix}{n:0{_digits_for(cfg, sheet)}d}"
                          for n in range(1, len(spec.samples) + 1)}
        else:
            values = {str(row.get(key, "")) for row in spec.samples
                      if row.get(key, "")}
        live[list_name] = values

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
                if not isinstance(source, str):
                    continue
                if not source.startswith("="):
                    # an inline list, e.g. ["Yes", "No"]
                    continue
                name = source.lstrip("=")
                if value in ("", None):
                    continue
                if name in lists:
                    if value not in lists[name]:
                        problems += 1
                        print(f"{spec.name}.{key} row {index}: "
                              f"{value!r} is not in {name}")
                elif name in live:
                    if value not in live[name]:
                        problems += 1
                        print(f"{spec.name}.{key} row {index}: "
                              f"{value!r} is not a live {name} value")
                else:
                    problems += 1
                    print(f"{spec.name}.{key} row {index}: unknown "
                          f"list {name}")

        print(f"  {spec.name:24s} {len(rows):4d} rows / {spec.rows:4d}, "
              f"{len(inputs)} input columns")

    print(f"\n{total_rows} sample rows checked, {problems} problems")
    return 1 if problems else 0


def _prefix_for(cfg, sheet: str) -> str:
    key = {"Wedding Bookings": "booking", "Client CRM": "client",
           "Photographer & Staff": "staff",
           "Marketing Calendar": "marketing"}.get(sheet)
    return cfg.prefix(key) if key else ""


def _digits_for(cfg, sheet: str) -> int:
    key = {"Wedding Bookings": "booking", "Client CRM": "client",
           "Photographer & Staff": "staff",
           "Marketing Calendar": "marketing"}.get(sheet)
    return cfg.digits(key) if key else 4


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

#### Client CRM

Every couple who has ever got in touch, what they booked and what they still owe

33 columns, 300 rows of capacity, 18 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Client ID | `id` | id | formula |
| B | Client name | `name` | text | you type it |
| C | Partner name | `partner` | text | you type it |
| D | Email | `email` | text | you type it |
| E | Phone | `phone` | text | you type it |
| F | Country | `country` | text | drop down `CountryList` |
| G | Client status | `status` | status | drop down `ClientStatusList` |
| H | Priority | `priority` | status | drop down `PriorityList` |
| I | Wedding date | `wedding_date` | date | you type it (checked) |
| J | Venue | `venue` | text | you type it |
| K | Venue town or city | `venue_city` | text | you type it |
| L | How they found you | `source` | text | drop down `LeadSourceList` |
| M | Referred by | `referred_by` | text | you type it |
| N | First contact | `first_contact` | date | you type it (checked) |
| O | Wedding style | `style` | text | drop down `WeddingStyleList` |
| P | Photography style | `photo_style` | text | drop down `PhotoStyleList` |
| Q | Delivery preference | `delivery_pref` | text | drop down `DeliveryPreferenceList` |
| R | Notes | `notes` | wrap | you type it |
| S | Happy to be featured | `marketing_ok` | flag | you type it (checked) |
| T | Wedding month | `wedding_month` | text | formula, hidden |
| U | Wedding date number | `wed_num` | number | formula, hidden |
| V | Row number | `seq_num` | int | formula, hidden |
| W | Weddings | `bookings` | int | formula |
| X | Booked value | `booked_value` | money | formula |
| Y | Collected | `collected` | money | formula |
| Z | Still owed | `outstanding` | money | formula |
| AA | Next wedding | `next_date` | date | formula |
| AB | Days away | `days_away` | int | formula |
| AC | Gallery delivered | `delivered` | date | formula |
| AD | Review | `review` | text | formula |
| AE | Referrals sent | `referrals` | int | formula |
| AF | Value rank | `booked_value_rank` | int | formula, hidden |
| AG | What to do | `alert` | status | formula |

#### Lead Tracker

Every enquiry, how fast you answered it and whether it turned into a wedding

40 columns, 400 rows of capacity, 23 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Lead ID | `id` | id | formula |
| B | Enquiry name | `name` | text | you type it |
| C | Partner name | `partner` | text | you type it |
| D | Email | `email` | text | you type it |
| E | Phone | `phone` | text | you type it |
| F | Client ID | `client` | id | drop down `ClientList` |
| G | Enquiry date | `enquiry_date` | date | you type it (checked) |
| H | Lead source | `source` | text | drop down `LeadSourceList` |
| I | Came from post | `campaign` | id | drop down `CampaignList` |
| J | Referred by | `referred_by` | id | drop down `BookingList` |
| K | Referrer | `referrer_name` | text | formula |
| L | Wedding date | `wedding_date` | date | you type it (checked) |
| M | Venue | `venue` | text | you type it |
| N | Guests | `guests` | int | you type it (checked) |
| O | Package asked about | `package` | text | drop down `PackageList` |
| P | Their budget | `budget` | money | you type it (checked) |
| Q | Lead status | `status` | status | drop down `LeadStatusList` |
| R | Priority | `priority` | status | drop down `PriorityList` |
| S | First replied | `replied` | date | you type it (checked) |
| T | Consultation | `consultation` | date | you type it (checked) |
| U | Proposal sent | `proposal_sent` | date | you type it (checked) |
| V | Booked or lost on | `decision_date` | date | you type it (checked) |
| W | Why it was lost | `lost_reason` | text | you type it |
| X | Follow up on | `follow_up` | date | you type it (checked) |
| Y | Notes | `notes` | wrap | you type it |
| Z | Enquiry month | `month` | text | formula, hidden |
| AA | Enquiry date number | `enq_num` | number | formula, hidden |
| AB | Follow up number | `follow_num` | number | formula, hidden |
| AC | Row number | `seq_num` | int | formula, hidden |
| AD | Still live | `is_open` | flag | formula, hidden |
| AE | Booked | `is_won` | flag | formula, hidden |
| AF | Days to reply | `response_hours` | number1 | formula |
| AG | Days old | `age_days` | int | formula |
| AH | Days since contact | `silent_days` | int | formula |
| AI | Chase by | `chase_on` | date | formula |
| AJ | Days overdue | `overdue` | int | formula |
| AK | Months to wedding | `lead_time` | number1 | formula |
| AL | Booked for | `booked_value` | money | formula |
| AM | Against their budget | `value_gap` | money | formula |
| AN | What to do | `alert` | status | formula |

#### Wedding Bookings

Every wedding you have taken on - the day, the team, the money and where it has got to

87 columns, 200 rows of capacity, 30 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Booking ID | `id` | id | formula |
| B | Client ID | `client` | id | drop down `ClientList` |
| C | Couple | `client_name` | text | formula |
| D | Partner | `partner_name` | text | formula |
| E | Wedding date | `wedding_date` | date | you type it (checked) |
| F | Day | `weekday` | text | formula |
| G | Booking status | `status` | status | drop down `BookingStatusList` |
| H | Confirmed on | `confirmed_date` | date | you type it (checked) |
| I | Ceremony | `ceremony_time` | time | you type it (checked) |
| J | Reception | `reception_time` | time | you type it (checked) |
| K | Coverage from | `photo_start` | time | you type it (checked) |
| L | Coverage to | `photo_end` | time | you type it (checked) |
| M | Main venue | `venue` | text | you type it |
| N | Venue type | `venue_type` | text | drop down `VenueTypeList` |
| O | Getting ready at | `prep_location` | text | you type it |
| P | Ceremony at | `ceremony_location` | text | you type it |
| Q | Reception at | `reception_location` | text | you type it |
| R | Guests | `guests` | int | you type it (checked) |
| S | Lead photographer | `photographer` | id | drop down `StaffList` |
| T | Lead name | `photographer_name` | text | formula |
| U | Second shooter | `second` | id | drop down `StaffList` |
| V | Second name | `second_name` | text | formula |
| W | Package | `package` | text | drop down `PackageList` |
| X | Package price | `package_price` | money | formula |
| Y | Hours included | `package_hours` | number1 | formula |
| Z | Hours booked | `coverage_hours` | number1 | formula |
| AA | Extra hours | `extra_hours` | number1 | formula |
| AB | Extra hour charge | `extra_hour_fee` | money | formula |
| AC | Other add-ons | `addons` | money | you type it (checked) |
| AD | Travel charged | `travel_fee` | money | you type it (checked) |
| AE | Travel used | `travel_used` | money | formula, hidden |
| AF | Discount | `discount` | money | you type it (checked) |
| AG | Subtotal | `subtotal` | money | formula |
| AH | Tax | `tax` | money | formula |
| AI | Total booking value | `total_value` | money | formula |
| AJ | Deposit due | `deposit_due_amount` | money | formula |
| AK | Deposit by | `deposit_due_date` | date | formula |
| AL | Final payment by | `final_due_date` | date | formula |
| AM | Collected | `collected` | money | formula |
| AN | Outstanding | `outstanding` | money | formula |
| AO | Paid | `paid_pct` | percent0 | formula |
| AP | Deposit in | `deposit_paid` | flag | formula |
| AQ | Payment status | `payment_state` | status | formula |
| AR | Must-have photographs | `must_haves` | wrap | you type it |
| AS | Family group list | `family_list` | wrap | you type it |
| AT | Important guests | `key_guests` | wrap | you type it |
| AU | Traditions and rituals | `traditions` | wrap | you type it |
| AV | Venue restrictions | `restrictions` | wrap | you type it |
| AW | Family sensitivities | `sensitivities` | wrap | you type it |
| AX | Special requests | `special_requests` | wrap | you type it |
| AY | Questionnaire sent | `questionnaire_sent` | date | you type it (checked) |
| AZ | Questionnaire back | `questionnaire_back` | date | you type it (checked) |
| BA | Notes | `notes` | wrap | you type it |
| BB | Wedding month | `month` | text | formula, hidden |
| BC | Year | `year` | int | formula, hidden |
| BD | Season | `season` | text | formula, hidden |
| BE | Line in the month | `month_seq` | int | formula, hidden |
| BF | Month key | `month_key` | text | formula, hidden |
| BG | Wedding date number | `wed_num` | number | formula, hidden |
| BH | Row number | `seq_num` | int | formula, hidden |
| BI | Holds the date | `holds` | flag | formula, hidden |
| BJ | Counts as revenue | `earns` | flag | formula, hidden |
| BK | Timeline lines | `timeline_lines` | int | formula, hidden |
| BL | Shot list lines | `shot_lines` | int | formula, hidden |
| BM | Upcoming key | `upcoming_key` | number | formula, hidden |
| BN | Next up | `upcoming_rank` | int | formula, hidden |
| BO | Wedding happened | `shot` | flag | formula, hidden |
| BP | Days away | `days_away` | int | formula |
| BQ | When | `countdown` | text | formula |
| BR | Date clash | `clash` | flag | formula |
| BS | Photographer clash | `team_clash` | flag | formula |
| BT | Contract settled | `contract_done` | flag | formula, hidden |
| BU | Contract | `contract_state` | status | formula |
| BV | Shot list | `shots_done` | percent0 | formula |
| BW | Editing | `editing_state` | status | formula |
| BX | Gallery | `gallery_state` | status | formula |
| BY | Deliver by | `deliver_by` | date | formula |
| BZ | Days late | `delivery_overdue` | int | formula |
| CA | Direct costs | `direct_costs` | money | formula |
| CB | Team cost | `team_cost` | money | formula |
| CC | Editing cost | `editing_cost` | money | formula |
| CD | Total cost | `total_cost` | money | formula |
| CE | Profit | `profit` | money | formula |
| CF | Margin | `margin` | percent0 | formula |
| CG | Profit an hour | `profit_hour` | money | formula |
| CH | Profit rank | `profit_rank` | int | formula, hidden |
| CI | What to do | `alert` | status | formula |

#### Packages & Pricing

What you sell, what each one really costs you to deliver, and what it ought to cost

41 columns, 30 rows of capacity, 17 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Package name | `name` | id | you type it |
| B | Tier | `tier` | text | drop down `TierList` |
| C | Still selling | `active` | flag | you type it (checked) |
| D | One-line summary | `summary` | wrap | you type it |
| E | Coverage hours | `hours` | number1 | you type it (checked) |
| F | Photographers | `photographers` | int | you type it (checked) |
| G | Engagement shoot | `engagement` | flag | you type it (checked) |
| H | Edited images | `images` | int | you type it (checked) |
| I | Album | `album` | text | you type it |
| J | Print credit | `print_credit` | money | you type it (checked) |
| K | Travel included | `travel_allowance` | int | you type it (checked) |
| L | Everything included | `inclusions` | wrap | you type it |
| M | List price | `base_price` | money | you type it (checked) |
| N | Standard discount | `standard_discount` | money | you type it (checked) |
| O | Price charged | `total_price` | money | formula |
| P | Second shooter | `second_cost` | money | formula |
| Q | Album cost | `album_cost` | money | you type it (checked) |
| R | Print cost | `print_cost` | money | formula |
| S | Engagement cost | `engagement_cost` | money | formula |
| T | Editing hours | `editing_hours` | number1 | you type it (checked) |
| U | Editing hours used | `editing_hours_used` | number1 | formula, hidden |
| V | Editing cost | `editing_cost` | money | formula |
| W | Other costs | `other_cost` | money | you type it (checked) |
| X | Cost to deliver | `total_cost` | money | formula |
| Y | Profit each | `profit` | money | formula |
| Z | Margin | `margin` | percent0 | formula |
| AA | Your hours | `total_hours` | number1 | formula |
| AB | Your hourly rate | `hourly` | money | formula |
| AC | Price for target margin | `target_price` | money | formula |
| AD | Under-priced by | `price_gap` | money | formula |
| AE | Hits target | `meets_target` | flag | formula |
| AF | Weddings sold | `sold` | int | formula |
| AG | Sold this year | `sold_year` | int | formula |
| AH | Asked about | `enquiries` | int | formula |
| AI | Enquiry to booking | `conversion` | percent0 | formula |
| AJ | Revenue booked | `revenue` | money | formula |
| AK | Profit booked | `profit_booked` | money | formula |
| AL | Share of bookings | `share` | percent0 | formula |
| AM | Revenue rank | `revenue_rank` | int | formula, hidden |
| AN | Popularity rank | `sold_rank` | int | formula, hidden |
| AO | What to do | `alert` | status | formula |

#### Contracts & Documents

A register of the paperwork for each wedding - what exists, where it is kept and what is still missing

21 columns, 200 rows of capacity, 11 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Document ID | `id` | id | formula |
| B | Booking ID | `booking` | id | drop down `BookingList` |
| C | Couple | `couple` | text | formula |
| D | Wedding date | `wedding_date` | date | formula |
| E | Document | `doc_type` | text | drop down `DocumentTypeList` |
| F | Title or reference | `title` | text | you type it |
| G | Needed | `required` | flag | you type it (checked) |
| H | Status | `status` | status | drop down `DocumentStatusList` |
| I | Sent | `sent_date` | date | you type it (checked) |
| J | Signed or received | `signed_date` | date | you type it (checked) |
| K | Expires | `expires` | date | you type it (checked) |
| L | Where it is kept | `stored_at` | text | you type it |
| M | Who signs | `signers` | text | you type it |
| N | Notes | `notes` | wrap | you type it |
| O | Wedding date number | `wed_num` | number | formula, hidden |
| P | Complete | `is_done` | flag | formula, hidden |
| Q | Days since sent | `days_out` | int | formula |
| R | Days to sign | `turnaround` | int | formula |
| S | Days to wedding | `days_to_wedding` | int | formula |
| T | Booking value | `booking_value` | money | formula |
| U | What to do | `alert` | status | formula |

#### Payments & Invoices

Every instalment for every wedding - what is due, what has arrived and what is late

29 columns, 600 rows of capacity, 13 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Invoice ID | `id` | id | formula |
| B | Booking ID | `booking` | id | drop down `BookingList` |
| C | Couple | `couple` | text | formula |
| D | Client ID | `client` | id | formula, hidden |
| E | Wedding date | `wedding_date` | date | formula |
| F | Stage | `stage` | text | drop down `PaymentStageList` |
| G | Your reference | `reference` | text | you type it |
| H | What it is for | `description` | wrap | you type it |
| I | Amount due | `amount` | money | you type it (checked) |
| J | Due | `due_date` | date | you type it (checked) |
| K | Received | `received` | money | you type it (checked) |
| L | Received on | `received_date` | date | you type it (checked) |
| M | How | `method` | text | drop down `PaymentMethodList` |
| N | Processing fee | `fee` | money | you type it (checked) |
| O | Net to you | `net` | money | formula |
| P | Still due | `balance` | money | formula |
| Q | Status | `status` | status | formula |
| R | Days overdue | `days_overdue` | int | formula |
| S | Days taken | `days_to_pay` | int | formula |
| T | Reminders sent | `reminders` | int | you type it (checked) |
| U | Last reminder | `last_reminder` | date | you type it (checked) |
| V | Notes | `notes` | wrap | you type it |
| W | Month received | `month` | text | formula, hidden |
| X | Year | `year` | int | formula, hidden |
| Y | Due date number | `due_num` | number | formula, hidden |
| Z | Paid date number | `paid_num` | number | formula, hidden |
| AA | Paid flag | `is_paid` | flag | formula, hidden |
| AB | Late flag | `is_late` | flag | formula, hidden |
| AC | What to do | `alert` | status | formula |

#### Wedding Timeline

The running order of each wedding day, built from the ceremony time

33 columns, 900 rows of capacity, 12 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Line ID | `id` | id | formula |
| B | Booking ID | `booking` | id | drop down `BookingList` |
| C | Couple | `couple` | text | formula |
| D | Wedding date | `wedding_date` | date | formula |
| E | Ceremony at | `ceremony` | time | formula |
| F | Event | `event` | text | drop down `TimelineEventList` |
| G | Your wording | `label` | text | you type it |
| H | Shows as | `shown` | text | formula |
| I | Standard offset | `offset_default` | int | formula, hidden |
| J | Minutes from ceremony | `offset` | int | you type it (checked) |
| K | Offset used | `offset_used` | int | formula, hidden |
| L | Fixed time | `time_override` | time | you type it (checked) |
| M | Starts | `start` | time | formula |
| N | Standard length | `duration_default` | int | formula, hidden |
| O | Minutes | `duration` | int | you type it (checked) |
| P | Length used | `duration_used` | int | formula, hidden |
| Q | Ends | `finish` | time | formula |
| R | Where | `location` | text | you type it |
| S | Who covers it | `role` | text | formula |
| T | Assigned to | `staff` | id | drop down `StaffList` |
| U | Name | `staff_name` | text | formula |
| V | Must not miss | `critical` | flag | you type it (checked) |
| W | Notes | `notes` | wrap | you type it |
| X | Done | `done` | flag | you type it (checked) |
| Y | Actually at | `actual` | time | you type it (checked) |
| Z | Minutes late | `drift` | int | formula |
| AA | Status | `status` | status | formula |
| AB | Wedding date number | `wed_num` | number | formula, hidden |
| AC | Start number | `start_num` | number | formula, hidden |
| AD | Sort key | `sort_key` | number | formula, hidden |
| AE | Order | `day_order` | int | formula |
| AF | Day key | `day_key` | text | formula, hidden |
| AG | What to do | `alert` | status | formula |

#### Shot List Manager

Every photograph you have promised, grouped the way you will shoot it and ticked off as you go

29 columns, 1400 rows of capacity, 15 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Shot ID | `id` | id | formula |
| B | Booking ID | `booking` | id | drop down `BookingList` |
| C | Couple | `couple` | text | formula |
| D | Wedding date | `wedding_date` | date | formula |
| E | Group | `group_name` | text | drop down `ShotGroupList` |
| F | Shot | `shot` | text | you type it |
| G | Who is in it | `people` | wrap | you type it |
| H | Priority | `priority` | text | drop down `PriorityList` |
| I | Must have | `must_have` | flag | you type it (checked) |
| J | Asked for by | `requested_by` | text | drop down `RequestSourceList` |
| K | When in the day | `when` | text | drop down `ShotWhenList` |
| L | Where | `location` | text | you type it |
| M | Notes | `notes` | wrap | you type it |
| N | Taken | `done` | flag | you type it (checked) |
| O | Taken at | `taken_at` | time | you type it (checked) |
| P | Skipped | `skipped` | flag | you type it (checked) |
| Q | Why not | `skip_reason` | wrap | you type it |
| R | In the gallery | `delivered` | flag | you type it (checked) |
| S | Wedding date number | `wed_num` | number | formula, hidden |
| T | Outstanding | `open` | flag | formula, hidden |
| U | On the printed list | `print_flag` | int | formula, hidden |
| V | Printed line | `print_seq` | int | formula, hidden |
| W | Printed key | `print_key` | text | formula, hidden |
| X | In this group | `group_total` | int | formula |
| Y | Group taken | `group_done` | int | formula |
| Z | Group complete | `group_pct` | percent0 | formula |
| AA | Wedding complete | `wedding_pct` | percent0 | formula |
| AB | Status | `status` | status | formula |
| AC | What to do | `alert` | status | formula |

#### Photographer & Staff

Everybody who shoots or edits for you, what they cost and how much they are working

38 columns, 40 rows of capacity, 19 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Staff ID | `id` | id | formula |
| B | Name | `name` | id | you type it |
| C | Role | `role` | text | drop down `StaffRoleList` |
| D | Status | `status` | status | drop down `StaffStatusList` |
| E | Employment | `type` | text | drop down `StaffTypeList` |
| F | Phone | `phone` | text | you type it |
| G | Email | `email` | text | you type it |
| H | Started | `started` | date | you type it (checked) |
| I | Wedding day rate | `day_rate` | money | you type it (checked) |
| J | Hourly rate | `hour_rate` | money | you type it (checked) |
| K | Editing rate | `edit_rate` | money | you type it (checked) |
| L | Owner | `is_owner` | flag | you type it (checked) |
| M | Can lead | `can_lead` | flag | you type it (checked) |
| N | Can second | `can_second` | flag | you type it (checked) |
| O | Can edit | `can_edit` | flag | you type it (checked) |
| P | Specialism | `specialism` | text | you type it |
| Q | Own kit | `kit` | wrap | you type it |
| R | Insured | `insured` | flag | you type it (checked) |
| S | Insurance to | `insurance_expiry` | date | you type it (checked) |
| T | Leading | `leading` | int | formula |
| U | Seconding | `seconding` | int | formula |
| V | Weddings | `weddings` | int | formula |
| W | Still to come | `upcoming` | int | formula |
| X | Next wedding number | `next_num` | number | formula, hidden |
| Y | Next wedding | `next_wedding` | date | formula |
| Z | Hours on the day | `shoot_hours` | number1 | formula |
| AA | Editing jobs | `edit_jobs` | int | formula |
| AB | Editing hours | `edit_hours` | number1 | formula |
| AC | Shoot pay | `shoot_pay` | money | formula |
| AD | Editing pay | `edit_pay` | money | formula |
| AE | Total pay | `total_pay` | money | formula |
| AF | Revenue worked on | `revenue_touched` | money | formula |
| AG | Pay as share | `pay_ratio` | percent0 | formula |
| AH | Busiest | `weddings_rank` | int | formula, hidden |
| AI | Average review | `rating` | rating | formula |
| AJ | Late deliveries | `late_jobs` | int | formula |
| AK | Notes | `notes` | wrap | you type it |
| AL | What to do | `alert` | status | formula |

#### Editing Workflow

Every wedding between the last frame and the finished gallery, and how late it is

46 columns, 200 rows of capacity, 24 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Job ID | `id` | id | formula |
| B | Booking ID | `booking` | id | drop down `BookingList` |
| C | Couple | `couple` | text | formula |
| D | Wedding date | `wedding_date` | date | formula |
| E | Editor | `editor` | id | drop down `StaffList` |
| F | Editor name | `editor_name` | text | formula |
| G | Outsourced | `outsourced` | flag | you type it (checked) |
| H | Priority | `priority` | text | drop down `PriorityList` |
| I | Rush job | `rush` | flag | you type it (checked) |
| J | Memory cards backed up | `step_backed_up` | flag | you type it (checked) |
| K | Photos imported | `step_imported` | flag | you type it (checked) |
| L | Culling complete | `step_culled` | flag | you type it (checked) |
| M | Selection complete | `step_selected` | flag | you type it (checked) |
| N | Editing complete | `step_edited` | flag | you type it (checked) |
| O | Colour correction done | `step_colour` | flag | you type it (checked) |
| P | Retouching done | `step_retouched` | flag | you type it (checked) |
| Q | Album designed | `step_album_designed` | flag | you type it (checked) |
| R | Client preview sent | `step_preview_sent` | flag | you type it (checked) |
| S | Final gallery uploaded | `step_final_uploaded` | flag | you type it (checked) |
| T | Steps done | `steps_done` | int | formula |
| U | Progress | `progress` | percent0 | formula |
| V | Stage | `stage` | status | formula |
| W | Started | `started` | date | you type it (checked) |
| X | Preview due | `preview_due` | date | formula |
| Y | Preview sent | `preview_sent` | date | you type it (checked) |
| Z | Gallery promised | `due` | date | formula |
| AA | Finished | `finished` | date | you type it (checked) |
| AB | Days taken | `days_taken` | int | formula |
| AC | Days left | `days_left` | int | formula |
| AD | Late | `late` | flag | formula |
| AE | Photos taken | `shot_count` | int | you type it (checked) |
| AF | Kept after culling | `kept` | int | you type it (checked) |
| AG | Delivered | `delivered_count` | int | you type it (checked) |
| AH | Package promises | `promised_count` | int | formula |
| AI | Short by | `short` | int | formula |
| AJ | Keep rate | `keep_rate` | percent0 | formula |
| AK | Hours spent | `hours` | number1 | you type it (checked) |
| AL | Hours expected | `hours_expected` | number1 | formula, hidden |
| AM | Hours over | `hours_over` | number1 | formula |
| AN | Editing rate | `rate` | money | formula, hidden |
| AO | Editing cost | `cost` | money | formula |
| AP | Backed up twice | `backup` | flag | you type it (checked) |
| AQ | Notes | `notes` | wrap | you type it |
| AR | Wedding date number | `wed_num` | number | formula, hidden |
| AS | Due date number | `due_num` | number | formula, hidden |
| AT | What to do | `alert` | status | formula |

#### Gallery & Delivery

Handing the photographs over, and everything that should happen in the month afterwards

36 columns, 200 rows of capacity, 21 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Gallery ID | `id` | id | formula |
| B | Booking ID | `booking` | id | drop down `BookingList` |
| C | Couple | `couple` | text | formula |
| D | Client ID | `client` | id | formula, hidden |
| E | Wedding date | `wedding_date` | date | formula |
| F | Where it lives | `platform` | text | drop down `GalleryPlatformList` |
| G | Gallery name | `gallery_name` | text | you type it |
| H | Link | `link` | link | you type it |
| I | Password set | `has_password` | flag | you type it (checked) |
| J | Status | `status` | status | drop down `GalleryStatusList` |
| K | Promised by | `promised` | date | formula |
| L | Delivered | `delivered_on` | date | you type it (checked) |
| M | Days late | `days_late` | int | formula |
| N | Images in it | `images` | int | you type it (checked) |
| O | Gallery expires | `expires` | date | formula |
| P | Expiry warning sent | `expiry_warned` | flag | you type it (checked) |
| Q | Archived | `archived` | flag | you type it (checked) |
| R | First opened | `first_viewed` | date | you type it (checked) |
| S | Views | `views` | int | you type it (checked) |
| T | Downloads | `downloads` | int | you type it (checked) |
| U | Favourites picked | `favourites` | int | you type it (checked) |
| V | USB or print box sent | `usb_sent` | flag | you type it (checked) |
| W | Have they looked | `opened` | flag | formula |
| X | Album ordered | `album_ordered` | flag | you type it (checked) |
| Y | Album due | `album_due` | date | formula |
| Z | Album sent | `album_sent` | date | you type it (checked) |
| AA | Print sales | `print_sales` | money | you type it (checked) |
| AB | Other extras | `extra_sales` | money | you type it (checked) |
| AC | Earned after delivery | `after_total` | money | formula |
| AD | Review asked for | `review_asked` | date | you type it (checked) |
| AE | Review received | `review_in` | flag | formula |
| AF | Notes | `notes` | wrap | you type it |
| AG | Wedding date number | `wed_num` | number | formula, hidden |
| AH | Delivered number | `delivered_num` | number | formula, hidden |
| AI | Delivered flag | `is_delivered` | flag | formula, hidden |
| AJ | What to do | `alert` | status | formula |

#### Equipment Inventory

Every body, lens, light and card you own - what it is worth, what it costs you and when it needs looking at

33 columns, 150 rows of capacity, 20 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Item ID | `id` | id | formula |
| B | Item | `item` | id | you type it |
| C | Category | `category` | text | drop down `EquipmentCategoryList` |
| D | Make and model | `make` | text | you type it |
| E | Serial number | `serial` | text | you type it |
| F | Status | `status` | status | drop down `EquipmentStatusList` |
| G | Condition | `condition` | text | drop down `EquipmentConditionList` |
| H | Assigned to | `assigned` | id | drop down `StaffList` |
| I | Who has it | `assigned_name` | text | formula |
| J | Kept where | `location` | text | you type it |
| K | Bought | `bought` | date | you type it (checked) |
| L | Price paid | `price` | money | you type it (checked) |
| M | Useful life | `life_years` | number1 | you type it (checked) |
| N | Life used | `life_used` | number1 | formula, hidden |
| O | Age in years | `age_years` | number1 | formula |
| P | Cost a year | `annual_cost` | money | formula |
| Q | Written down to | `written_down` | money | formula |
| R | Weddings since | `weddings_since` | int | formula |
| S | Cost a wedding | `cost_per_wedding` | money | formula |
| T | Value rank | `price_rank` | int | formula, hidden |
| U | Insured | `insured` | flag | you type it (checked) |
| V | Insured for | `insured_value` | money | you type it (checked) |
| W | Short by | `under_insured` | money | formula |
| X | Warranty to | `warranty_to` | date | you type it (checked) |
| Y | Last serviced | `serviced` | date | you type it (checked) |
| Z | Service due | `service_due` | date | formula |
| AA | Shutter count | `shutter` | int | you type it (checked) |
| AB | Rated for | `shutter_limit` | int | you type it (checked) |
| AC | Shutter life used | `shutter_used` | percent0 | formula |
| AD | Backs up | `backup_for` | text | you type it |
| AE | Replacement cost | `replace_cost` | money | you type it (checked) |
| AF | Notes | `notes` | wrap | you type it |
| AG | What to do | `alert` | status | formula |

#### Travel & Mileage

Every journey for every wedding - the distance, the cost and what you can charge on

33 columns, 300 rows of capacity, 18 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Trip ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Booking ID | `booking` | id | drop down `BookingList` |
| D | Couple | `couple` | text | formula |
| E | Purpose | `purpose` | text | drop down `TravelPurposeList` |
| F | From | `from_place` | text | you type it |
| G | To | `to_place` | text | you type it |
| H | How | `mode` | text | drop down `TravelModeList` |
| I | Miles one way | `one_way` | number1 | you type it (checked) |
| J | Return | `return_trip` | flag | you type it (checked) |
| K | Total miles | `miles` | number1 | formula |
| L | Extra trips | `scout_trips` | int | you type it (checked) |
| M | Miles in all | `all_miles` | number1 | formula |
| N | Mileage cost | `mileage` | money | formula |
| O | Fuel instead | `fuel` | money | you type it (checked) |
| P | Tolls | `tolls` | money | you type it (checked) |
| Q | Parking | `parking` | money | you type it (checked) |
| R | Accommodation | `accommodation` | money | you type it (checked) |
| S | Other | `other_cost` | money | you type it (checked) |
| T | Total cost | `total_cost` | money | formula |
| U | Miles included | `allowance` | int | formula |
| V | Miles to charge | `billable_miles` | number1 | formula |
| W | Suggested charge | `suggested` | money | formula |
| X | Charge instead | `charge_override` | money | you type it (checked) |
| Y | Charged on | `chargeable` | money | formula |
| Z | Costs you | `net_cost` | money | formula |
| AA | Hours travelling | `hours` | number1 | you type it (checked) |
| AB | Receipt kept | `receipt` | flag | you type it (checked) |
| AC | Notes | `notes` | wrap | you type it |
| AD | Month | `month` | text | formula, hidden |
| AE | Year | `year` | int | formula, hidden |
| AF | Date number | `date_num` | number | formula, hidden |
| AG | What to do | `alert` | status | formula |

#### Marketing Calendar

What you plan to post, what you actually posted, and which of it brought weddings in

32 columns, 400 rows of capacity, 17 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Post ID | `id` | id | formula |
| B | Planned for | `date` | date | you type it (checked) |
| C | Platform | `platform` | text | drop down `MarketingPlatformList` |
| D | Type | `type` | text | drop down `ContentTypeList` |
| E | Title or hook | `title` | text | you type it |
| F | What it says | `content` | wrap | you type it |
| G | Wedding featured | `booking` | id | drop down `BookingList` |
| H | Permission given | `permission` | flag | you type it (checked) |
| I | Status | `status` | status | drop down `ContentStatusList` |
| J | Actually posted | `posted_on` | date | you type it (checked) |
| K | Reach | `reach` | int | you type it (checked) |
| L | Likes and comments | `engagement` | int | you type it (checked) |
| M | Saves or shares | `saves` | int | you type it (checked) |
| N | Clicks through | `clicks` | int | you type it (checked) |
| O | Engagement rate | `engagement_rate` | percent0 | formula |
| P | Enquiries from it | `enquiries` | int | formula |
| Q | Counted by hand | `manual_enquiries` | int | you type it (checked) |
| R | Enquiries in all | `all_enquiries` | int | formula |
| S | Weddings booked | `bookings` | int | formula |
| T | Revenue from it | `revenue` | money | formula |
| U | Spent on it | `spend` | money | you type it (checked) |
| V | Hours it took | `hours` | number1 | you type it (checked) |
| W | Your time is worth | `time_cost` | money | formula |
| X | Real cost | `true_cost` | money | formula |
| Y | Return | `roi` | number1 | formula |
| Z | Cost an enquiry | `cost_per_enquiry` | money | formula |
| AA | Best posts | `revenue_rank` | int | formula, hidden |
| AB | Notes | `notes` | wrap | you type it |
| AC | Month | `month` | text | formula, hidden |
| AD | Year | `year` | int | formula, hidden |
| AE | Date number | `date_num` | number | formula, hidden |
| AF | What to do | `alert` | status | formula |

#### Expense Tracker

Everything the business pays out, split between the weddings that caused it and the overheads that did not

26 columns, 600 rows of capacity, 14 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Expense ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Paid to | `supplier` | text | you type it |
| D | Category | `category` | text | drop down `ExpenseCategoryList` |
| E | What it was | `description` | wrap | you type it |
| F | Booking ID | `booking` | id | drop down `BookingList` |
| G | Amount | `amount` | money | you type it (checked) |
| H | Tax included | `tax_amount` | money | you type it (checked) |
| I | Net of tax | `net` | money | formula |
| J | How paid | `method` | text | drop down `PaymentMethodList` |
| K | Receipt kept | `receipt` | flag | you type it (checked) |
| L | Reference | `reference` | text | you type it |
| M | Wedding cost | `is_direct` | flag | formula |
| N | Overhead | `is_overhead` | flag | formula |
| O | Equipment buy | `is_capital` | flag | formula |
| P | Repeats | `recurring` | text | drop down `FrequencyList` |
| Q | A month | `monthly` | money | formula |
| R | A year | `annual` | money | formula |
| S | Business use | `business_pct` | percent0 | you type it (checked) |
| T | Claimable part | `claimable` | money | formula |
| U | Tax category | `tax_category` | text | drop down `TaxCategoryList` |
| V | Notes | `notes` | wrap | you type it |
| W | Month | `month` | text | formula, hidden |
| X | Year | `year` | int | formula, hidden |
| Y | Date number | `date_num` | number | formula, hidden |
| Z | What to do | `alert` | status | formula |

#### Profit Analysis

What each wedding really made once travel, editing, the second shooter and a share of the overheads come out

44 columns, 200 rows of capacity, 1 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Line ID | `id` | id | formula |
| B | Booking ID | `booking` | id | drop down `BookingList` |
| C | Couple | `couple` | text | formula |
| D | Wedding date | `wedding_date` | date | formula |
| E | Package | `package` | text | formula |
| F | Status | `status` | status | formula |
| G | Venue type | `venue_type` | text | formula |
| H | Season | `season` | text | formula |
| I | Package | `package_price` | money | formula |
| J | Extra hours | `extra_hours` | money | formula |
| K | Add-ons | `addons` | money | formula |
| L | Travel charged | `travel_charged` | money | formula |
| M | Discount given | `discount` | money | formula |
| N | Prints and albums | `after_sales` | money | formula |
| O | Tax collected | `tax_collected` | money | formula |
| P | Revenue | `revenue` | money | formula |
| Q | Actually received | `collected` | money | formula |
| R | Still owed | `outstanding` | money | formula |
| S | Second shooter | `second_cost` | money | formula |
| T | Editing | `editing_cost` | money | formula |
| U | Travel | `travel_cost` | money | formula |
| V | Albums and prints | `album_cost` | money | formula |
| W | Other wedding costs | `other_direct` | money | formula |
| X | Share of overheads | `overhead_share` | money | formula |
| Y | Direct costs | `direct_total` | money | formula |
| Z | All costs | `total_cost` | money | formula |
| AA | Profit before overheads | `gross_profit` | money | formula |
| AB | Gross margin | `gross_margin` | percent0 | formula |
| AC | Profit after overheads | `net_profit` | money | formula |
| AD | Net margin | `net_margin` | percent0 | formula |
| AE | Hours shooting | `shoot_hours` | number1 | formula |
| AF | Hours editing | `edit_hours` | number1 | formula |
| AG | Hours travelling | `travel_hours` | number1 | formula |
| AH | Hours on admin | `admin_hours` | number1 | formula |
| AI | Hours in all | `total_hours` | number1 | formula |
| AJ | Profit an hour | `per_hour` | money | formula |
| AK | Against target | `vs_target` | money | formula |
| AL | Beats your rate | `beats_hourly` | flag | formula |
| AM | Profit rank | `net_profit_rank` | int | formula, hidden |
| AN | Hourly rank | `per_hour_rank` | int | formula, hidden |
| AO | Wedding month | `month` | text | formula, hidden |
| AP | Year | `year` | int | formula, hidden |
| AQ | Wedding date number | `wed_num` | number | formula, hidden |
| AR | What it tells you | `verdict` | status | formula |

#### Reviews & Referrals

Who you asked, what they wrote, and the weddings that came back out of it

30 columns, 200 rows of capacity, 16 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Review ID | `id` | id | formula |
| B | Booking ID | `booking` | id | drop down `BookingList` |
| C | Couple | `couple` | text | formula |
| D | Client ID | `client` | id | formula, hidden |
| E | Wedding date | `wedding_date` | date | formula |
| F | Photographer | `photographer` | id | formula |
| G | Asked on | `asked_on` | date | you type it (checked) |
| H | How you asked | `how_asked` | text | drop down `ReviewAskMethodList` |
| I | Reminders | `reminders` | int | you type it (checked) |
| J | Status | `status` | status | drop down `ReviewStatusList` |
| K | Received | `received_on` | date | you type it (checked) |
| L | Where | `platform` | text | drop down `ReviewPlatformList` |
| M | Rating | `rating` | rating | you type it (checked) |
| N | Headline | `headline` | text | you type it |
| O | What they said | `quote` | wrap | you type it |
| P | May you quote it | `can_use` | flag | you type it (checked) |
| Q | Used where | `used_on` | text | you type it |
| R | You replied | `responded` | flag | you type it (checked) |
| S | Days to answer | `days_to_review` | int | formula |
| T | People referred | `referrals` | int | you type it (checked) |
| U | Enquiries from them | `referred_leads` | int | formula |
| V | Weddings from them | `referred_bookings` | int | formula |
| W | Worth to you | `referred_value` | money | formula |
| X | Thanked them | `thanked` | flag | you type it (checked) |
| Y | Notes | `notes` | wrap | you type it |
| Z | Month received | `month` | text | formula, hidden |
| AA | Year | `year` | int | formula, hidden |
| AB | Received number | `received_num` | number | formula, hidden |
| AC | Got one | `has_review` | flag | formula, hidden |
| AD | What to do | `alert` | status | formula |

---

*12,666 lines of Python in 41 files.*
