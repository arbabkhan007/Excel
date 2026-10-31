# Daycare & Childcare Business Management Spreadsheet — source code

Complete source code of the generator that writes the `.xlsx` product.

Writes a 23 tab `.xlsx` workbook with over 143,000 formulas: enrolment with automatic child IDs, an age in months and an age band; families that hold several children and one balance; a waiting list with a queue position; a daily register with late arrivals, missing check-outs and an authorised-pickup check; tuition on six different billing bases with sibling discounts and subsidies; a charge ledger that adds up into printable invoices and receipts; daily reports, meals, allergies and medication, incident forms; staff with certificate countdowns, a rota that catches a double shift, payroll with employer costs; room occupancy with a live staff-to-child ratio alert; expenses, supplies, parent conversations and a document renewal tracker; and monthly and annual reports. Document author is **Novality Store**; every formula cell is locked and every sheet protected with the password **`premium`**, while the cells you are meant to type in stay open.

Nothing is hard coded to one country's licensing. The age bands, the adults each band needs and the maximum group size are a table on Settings, so a setting in Ohio, Ontario or Oxfordshire types in its own rules and every ratio alert in the workbook follows. The tax rate and the two switches that say what it applies to, the fees, the late pickup charge, the overtime threshold and every drop-down list are settings, not assumptions.

Pure Python, one dependency. Tested on XlsxWriter 3.2.0 and on the much older 0.3.0, so it does not rely on anything recent. The workbook itself is Excel 2010 safe — no dynamic arrays, no `XLOOKUP`, no `IFS`, nothing that needs array entry — and it opens in Google Sheets, LibreOffice and Numbers.

```bash
python build_daycare_workbook.py --help

  -o, --output PATH     where to write the file
  --blank               no example data, just the empty workbook
  --no-protect          leave the sheets unlocked
  --password TEXT       protection password (default: premium)
  --author TEXT         document author  (default: Novality Store)
  --centre TEXT         your setting's name on the cover
  --owner TEXT          the owner or director's name
  --currency SYMBOL     currency symbol  (default: $)
  --currency-code CODE  three letter code (default: USD)
  --tax PERCENT         tax rate, if childcare is taxed near you
  --tax-on WHAT         both, tuition, extras or neither
  --capacity N          licensed places for the whole setting
  --open-days N         days open a week
  --open-hour N         opening time, e.g. 7 for 7am
  --close-hour N        closing time, e.g. 18 for 6pm
  --weeks-open N        trading weeks a year
  --full-time-week AMT  the standard full-time weekly fee
  --basis TEXT          the billing basis new children start on
  --sibling-discount P  sibling discount as a percentage
  --hourly AMOUNT       default staff hourly rate
  --burden PERCENT      employer costs on top of wages
  --overtime-after N    hours a week before overtime starts
  --target AMOUNT       monthly revenue target
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
  * [`build_daycare_workbook.py`](#build_daycare_workbookpy)
* [2. The package](#2-the-package)
  * [`daycare_business_spreadsheet/__init__.py`](#daycare_business_spreadsheet__init__py)
  * [`daycare_business_spreadsheet/__main__.py`](#daycare_business_spreadsheet__main__py)
  * [`daycare_business_spreadsheet/cli.py`](#daycare_business_spreadsheetclipy)
  * [`daycare_business_spreadsheet/config.py`](#daycare_business_spreadsheetconfigpy)
  * [`daycare_business_spreadsheet/layout.py`](#daycare_business_spreadsheetlayoutpy)
  * [`daycare_business_spreadsheet/formats.py`](#daycare_business_spreadsheetformatspy)
  * [`daycare_business_spreadsheet/writer.py`](#daycare_business_spreadsheetwriterpy)
  * [`daycare_business_spreadsheet/names.py`](#daycare_business_spreadsheetnamespy)
  * [`daycare_business_spreadsheet/build.py`](#daycare_business_spreadsheetbuildpy)
  * [`daycare_business_spreadsheet/sample_data.py`](#daycare_business_spreadsheetsample_datapy)
* [3. The children and the families](#3-the-children-and-the-families)
  * [`daycare_business_spreadsheet/sheets/__init__.py`](#daycare_business_spreadsheetsheets__init__py)
  * [`daycare_business_spreadsheet/sheets/common.py`](#daycare_business_spreadsheetsheetscommonpy)
  * [`daycare_business_spreadsheet/sheets/children.py`](#daycare_business_spreadsheetsheetschildrenpy)
  * [`daycare_business_spreadsheet/sheets/parents.py`](#daycare_business_spreadsheetsheetsparentspy)
  * [`daycare_business_spreadsheet/sheets/waitlist.py`](#daycare_business_spreadsheetsheetswaitlistpy)
* [4. The register](#4-the-register)
  * [`daycare_business_spreadsheet/sheets/attendance.py`](#daycare_business_spreadsheetsheetsattendancepy)
* [5. The money](#5-the-money)
  * [`daycare_business_spreadsheet/sheets/tuition.py`](#daycare_business_spreadsheetsheetstuitionpy)
  * [`daycare_business_spreadsheet/sheets/invoices.py`](#daycare_business_spreadsheetsheetsinvoicespy)
  * [`daycare_business_spreadsheet/sheets/expenses.py`](#daycare_business_spreadsheetsheetsexpensespy)
* [6. The day a parent reads about](#6-the-day-a-parent-reads-about)
  * [`daycare_business_spreadsheet/sheets/daily_reports.py`](#daycare_business_spreadsheetsheetsdaily_reportspy)
  * [`daycare_business_spreadsheet/sheets/meals.py`](#daycare_business_spreadsheetsheetsmealspy)
  * [`daycare_business_spreadsheet/sheets/medical.py`](#daycare_business_spreadsheetsheetsmedicalpy)
  * [`daycare_business_spreadsheet/sheets/incidents.py`](#daycare_business_spreadsheetsheetsincidentspy)
* [7. The team](#7-the-team)
  * [`daycare_business_spreadsheet/sheets/staff.py`](#daycare_business_spreadsheetsheetsstaffpy)
  * [`daycare_business_spreadsheet/sheets/schedule.py`](#daycare_business_spreadsheetsheetsschedulepy)
  * [`daycare_business_spreadsheet/sheets/payroll.py`](#daycare_business_spreadsheetsheetspayrollpy)
* [8. Rooms, stock and paperwork](#8-rooms-stock-and-paperwork)
  * [`daycare_business_spreadsheet/sheets/classrooms.py`](#daycare_business_spreadsheetsheetsclassroomspy)
  * [`daycare_business_spreadsheet/sheets/supplies.py`](#daycare_business_spreadsheetsheetssuppliespy)
  * [`daycare_business_spreadsheet/sheets/communication.py`](#daycare_business_spreadsheetsheetscommunicationpy)
  * [`daycare_business_spreadsheet/sheets/documents.py`](#daycare_business_spreadsheetsheetsdocumentspy)
* [9. The pages](#9-the-pages)
  * [`daycare_business_spreadsheet/sheets/dashboard.py`](#daycare_business_spreadsheetsheetsdashboardpy)
  * [`daycare_business_spreadsheet/sheets/getting_started.py`](#daycare_business_spreadsheetsheetsgetting_startedpy)
  * [`daycare_business_spreadsheet/sheets/settings.py`](#daycare_business_spreadsheetsheetssettingspy)
  * [`daycare_business_spreadsheet/sheets/monthly.py`](#daycare_business_spreadsheetsheetsmonthlypy)
  * [`daycare_business_spreadsheet/sheets/annual.py`](#daycare_business_spreadsheetsheetsannualpy)
* [10. The quality tools](#10-the-quality-tools)
  * [`tools/validate_daycare_workbook.py`](#toolsvalidate_daycare_workbookpy)
  * [`tools/check_parens.py`](#toolscheck_parenspy)
  * [`tools/check_daycare_sample.py`](#toolscheck_daycare_samplepy)
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

`build_daycare_workbook.py` parses the command line into a `Config`, then
`daycare_business_spreadsheet.build.build_workbook(path, cfg)` does six things in order:

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
| 1 | **Dashboard** | page | `daycare_business_spreadsheet/sheets/dashboard.py` | - | - |
| 2 | **Getting Started** | page | `daycare_business_spreadsheet/sheets/getting_started.py` | - | - |
| 3 | **Settings & Lists** | page | `daycare_business_spreadsheet/sheets/settings.py` | - | - |
| 4 | **Child Enrollment** | register | `daycare_business_spreadsheet/sheets/children.py` | 73 | 150 |
| 5 | **Parent Database** | register | `daycare_business_spreadsheet/sheets/parents.py` | 51 | 200 |
| 6 | **Waitlist** | register | `daycare_business_spreadsheet/sheets/waitlist.py` | 36 | 120 |
| 7 | **Attendance** | register | `daycare_business_spreadsheet/sheets/attendance.py` | 35 | 2000 |
| 8 | **Tuition & Fees** | register | `daycare_business_spreadsheet/sheets/tuition.py` | 29 | 700 |
| 9 | **Invoices & Receipts** | register | `daycare_business_spreadsheet/sheets/invoices.py` | 31 | 700 |
| 10 | **Daily Reports** | register | `daycare_business_spreadsheet/sheets/daily_reports.py` | 41 | 600 |
| 11 | **Meal Tracker** | register | `daycare_business_spreadsheet/sheets/meals.py` | 22 | 600 |
| 12 | **Allergy & Medical** | register | `daycare_business_spreadsheet/sheets/medical.py` | 33 | 180 |
| 13 | **Incident Reports** | register | `daycare_business_spreadsheet/sheets/incidents.py` | 36 | 150 |
| 14 | **Staff Database** | register | `daycare_business_spreadsheet/sheets/staff.py` | 40 | 60 |
| 15 | **Staff Schedule** | register | `daycare_business_spreadsheet/sheets/schedule.py` | 32 | 750 |
| 16 | **Payroll Tracker** | register | `daycare_business_spreadsheet/sheets/payroll.py` | 29 | 350 |
| 17 | **Classroom Management** | register | `daycare_business_spreadsheet/sheets/classrooms.py` | 33 | 30 |
| 18 | **Expense Tracker** | register | `daycare_business_spreadsheet/sheets/expenses.py` | 23 | 500 |
| 19 | **Supply Inventory** | register | `daycare_business_spreadsheet/sheets/supplies.py` | 27 | 180 |
| 20 | **Parent Communication** | register | `daycare_business_spreadsheet/sheets/communication.py` | 24 | 350 |
| 21 | **Document Tracker** | register | `daycare_business_spreadsheet/sheets/documents.py` | 19 | 500 |
| 22 | **Monthly Reports** | page | `daycare_business_spreadsheet/sheets/monthly.py` | - | - |
| 23 | **Annual Reports** | page | `daycare_business_spreadsheet/sheets/annual.py` | - | - |

### Every file at a glance

| File | Lines | What it is |
|------|------:|------------|
| [`build_daycare_workbook.py`](#build_daycare_workbookpy) | 17 | Build the Daycare & Childcare Business Management Spreadsheet |
| [`daycare_business_spreadsheet/__init__.py`](#daycare_business_spreadsheet__init__py) | 19 | Daycare & Childcare Business Management Spreadsheet |
| [`daycare_business_spreadsheet/__main__.py`](#daycare_business_spreadsheet__main__py) | 6 | ``python -m daycare_business_spreadsheet`` |
| [`daycare_business_spreadsheet/cli.py`](#daycare_business_spreadsheetclipy) | 209 | Command line front end |
| [`daycare_business_spreadsheet/config.py`](#daycare_business_spreadsheetconfigpy) | 733 | Configuration for the Daycare & Childcare Business Management Spreadsheet |
| [`daycare_business_spreadsheet/layout.py`](#daycare_business_spreadsheetlayoutpy) | 253 | Layout engine |
| [`daycare_business_spreadsheet/formats.py`](#daycare_business_spreadsheetformatspy) | 202 | Cell format factory |
| [`daycare_business_spreadsheet/writer.py`](#daycare_business_spreadsheetwriterpy) | 458 | Sheet writing helpers shared by every tab of the workbook |
| [`daycare_business_spreadsheet/names.py`](#daycare_business_spreadsheetnamespy) | 63 | Workbook level defined names that point at the register sheets |
| [`daycare_business_spreadsheet/build.py`](#daycare_business_spreadsheetbuildpy) | 162 | Workbook builder: puts the twenty-three tabs together in order |
| [`daycare_business_spreadsheet/sample_data.py`](#daycare_business_spreadsheetsample_datapy) | 1,989 | Demo data: ten weeks at Little Acorns Childcare |
| [`daycare_business_spreadsheet/sheets/__init__.py`](#daycare_business_spreadsheetsheets__init__py) | 0 |  |
| [`daycare_business_spreadsheet/sheets/common.py`](#daycare_business_spreadsheetsheetscommonpy) | 509 | Small building blocks shared by the register sheets |
| [`daycare_business_spreadsheet/sheets/children.py`](#daycare_business_spreadsheetsheetschildrenpy) | 501 | `Child Enrollment` - the register every other tab hangs off |
| [`daycare_business_spreadsheet/sheets/parents.py`](#daycare_business_spreadsheetsheetsparentspy) | 316 | `Parent Database` - one row per family, however many children they have |
| [`daycare_business_spreadsheet/sheets/waitlist.py`](#daycare_business_spreadsheetsheetswaitlistpy) | 280 | `Waitlist` - the enquiries you have not turned into enrollments yet |
| [`daycare_business_spreadsheet/sheets/attendance.py`](#daycare_business_spreadsheetsheetsattendancepy) | 297 | `Attendance` - the daily sign in and sign out sheet |
| [`daycare_business_spreadsheet/sheets/tuition.py`](#daycare_business_spreadsheetsheetstuitionpy) | 221 | `Tuition & Fees` - every charge and every credit, one row at a time |
| [`daycare_business_spreadsheet/sheets/invoices.py`](#daycare_business_spreadsheetsheetsinvoicespy) | 311 | `Invoices & Receipts` - one ledger for everything that moves money |
| [`daycare_business_spreadsheet/sheets/expenses.py`](#daycare_business_spreadsheetsheetsexpensespy) | 210 | `Expense Tracker` - everything that goes out, categorised and budgeted |
| [`daycare_business_spreadsheet/sheets/daily_reports.py`](#daycare_business_spreadsheetsheetsdaily_reportspy) | 270 | `Daily Reports` - the sheet that goes home with the child |
| [`daycare_business_spreadsheet/sheets/meals.py`](#daycare_business_spreadsheetsheetsmealspy) | 172 | `Meal Tracker` - what was served, who ate it, and who must not have it |
| [`daycare_business_spreadsheet/sheets/medical.py`](#daycare_business_spreadsheetsheetsmedicalpy) | 232 | `Allergy & Medical` - conditions, plans, consents and the dose log |
| [`daycare_business_spreadsheet/sheets/incidents.py`](#daycare_business_spreadsheetsheetsincidentspy) | 248 | `Incident Reports` - accidents, injuries, near misses and who was told |
| [`daycare_business_spreadsheet/sheets/staff.py`](#daycare_business_spreadsheetsheetsstaffpy) | 287 | `Staff Database` - the team, their pay and their certificates |
| [`daycare_business_spreadsheet/sheets/schedule.py`](#daycare_business_spreadsheetsheetsschedulepy) | 256 | `Staff Schedule` - the rota, the clock and what the day actually cost |
| [`daycare_business_spreadsheet/sheets/payroll.py`](#daycare_business_spreadsheetsheetspayrollpy) | 242 | `Payroll Tracker` - turning rota hours into a wage bill |
| [`daycare_business_spreadsheet/sheets/classrooms.py`](#daycare_business_spreadsheetsheetsclassroomspy) | 320 | `Classroom Management` - how full each room is, and whether it is legal |
| [`daycare_business_spreadsheet/sheets/supplies.py`](#daycare_business_spreadsheetsheetssuppliespy) | 217 | `Supply Inventory` - nappies, wipes, paint, paper and lunch |
| [`daycare_business_spreadsheet/sheets/communication.py`](#daycare_business_spreadsheetsheetscommunicationpy) | 169 | `Parent Communication` - what was said, when, and who owes a reply |
| [`daycare_business_spreadsheet/sheets/documents.py`](#daycare_business_spreadsheetsheetsdocumentspy) | 186 | `Document Tracker` - the paperwork, and the day each piece expires |
| [`daycare_business_spreadsheet/sheets/dashboard.py`](#daycare_business_spreadsheetsheetsdashboardpy) | 566 | `Dashboard` - the one page that answers "how is the setting doing?" Twenty numbers across the top, the... |
| [`daycare_business_spreadsheet/sheets/getting_started.py`](#daycare_business_spreadsheetsheetsgetting_startedpy) | 337 | `Getting Started` - what this is, what to type, and what it will not do |
| [`daycare_business_spreadsheet/sheets/settings.py`](#daycare_business_spreadsheetsheetssettingspy) | 662 | `Settings & Lists` - every number and every list the workbook runs on |
| [`daycare_business_spreadsheet/sheets/monthly.py`](#daycare_business_spreadsheetsheetsmonthlypy) | 542 | `Monthly Reports` - one month, on one page you can print |
| [`daycare_business_spreadsheet/sheets/annual.py`](#daycare_business_spreadsheetsheetsannualpy) | 469 | `Annual Reports` - twelve months side by side, and what they add up to |
| [`tools/validate_daycare_workbook.py`](#toolsvalidate_daycare_workbookpy) | 892 | Static checks on a generated Daycare Business workbook |
| [`tools/check_parens.py`](#toolscheck_parenspy) | 121 | Parenthesis and quote balance check for every formula in a workbook |
| [`tools/check_daycare_sample.py`](#toolscheck_daycare_samplepy) | 438 | Does the demo nursery actually hang together? The validator checks the formulas; this checks the *story*... |
| [`tools/recalc_check.py`](#toolsrecalc_checkpy) | 216 | Really calculate every formula in a generated workbook and report any cell that ends up as an Excel error... |

---

## 1. Run it

Save the files below keeping these exact paths, then run the one liner. Nothing else is needed - no project file, no build step.

```bash
pip install XlsxWriter
python build_daycare_workbook.py
```

The workbook lands in `output/Daycare Business Management Spreadsheet.xlsx`.

### `build_daycare_workbook.py`

*17 lines*

> Build the Daycare & Childcare Business Management Spreadsheet.
>
>     python3 build_daycare_workbook.py
>     python3 build_daycare_workbook.py -o "output/My Nursery.xlsx" --blank
>     python3 build_daycare_workbook.py --currency "£" --tax 20 --rows 800
>
> Run `--help` for every option. The generator itself lives in the
> `daycare_business_spreadsheet` package next to this file.

````python
#!/usr/bin/env python3
"""Build the Daycare & Childcare Business Management Spreadsheet.

    python3 build_daycare_workbook.py
    python3 build_daycare_workbook.py -o "output/My Nursery.xlsx" --blank
    python3 build_daycare_workbook.py --currency "£" --tax 20 --rows 800

Run `--help` for every option. The generator itself lives in the
`daycare_business_spreadsheet` package next to this file.
"""

import sys

from daycare_business_spreadsheet.cli import main

if __name__ == "__main__":
    sys.exit(main())
````

---

## 2. The package

Configuration, the formula template engine, cell formats, the sheet writer, the cross sheet names and the demo nursery. These are the parts every tab leans on. `config.py` is also where the age bands and the staff-to-child ratios start out, so a setting in any country can be matched without touching a formula.

### `daycare_business_spreadsheet/__init__.py`

*19 lines*

> Daycare & Childcare Business Management Spreadsheet.
>
> Build the workbook with::
>
>     python -m daycare_business_spreadsheet -o "output/Daycare.xlsx"
>
> or from Python::
>
>     from daycare_business_spreadsheet import build_workbook
>     build_workbook("Daycare Business Management Spreadsheet.xlsx")

````python
"""Daycare & Childcare Business Management Spreadsheet.

Build the workbook with::

    python -m daycare_business_spreadsheet -o "output/Daycare.xlsx"

or from Python::

    from daycare_business_spreadsheet import build_workbook
    build_workbook("Daycare Business Management Spreadsheet.xlsx")
"""

from .build import build_workbook
from .config import (AUTHOR, PRODUCT_LISTING_TITLE, PRODUCT_NAME,
                     PRODUCT_SUBTITLE, PRODUCT_VERSION, Config)

__all__ = ["build_workbook", "Config", "AUTHOR", "PRODUCT_NAME",
           "PRODUCT_SUBTITLE", "PRODUCT_LISTING_TITLE", "PRODUCT_VERSION"]
__version__ = PRODUCT_VERSION
````

### `daycare_business_spreadsheet/__main__.py`

*6 lines*

> ``python -m daycare_business_spreadsheet``.

````python
"""``python -m daycare_business_spreadsheet``."""

from .cli import main

if __name__ == "__main__":
    raise SystemExit(main())
````

### `daycare_business_spreadsheet/cli.py`

*209 lines · 3 functions*

> Command line front end.
>
> python -m daycare_business_spreadsheet
> python -m daycare_business_spreadsheet -o "output/Daycare.xlsx" --blank
> python -m daycare_business_spreadsheet --currency "£" --tax 20 --rows 800

| Defines | What it does |
|---------|--------------|
| `build_parser()` |  |
| `config_from_args(args)` |  |
| `main(argv)` |  |

**Module constants:** `DEFAULT_OUTPUT` (os.path.join(...)).

````python
"""Command line front end.

    python -m daycare_business_spreadsheet
    python -m daycare_business_spreadsheet -o "output/Daycare.xlsx" --blank
    python -m daycare_business_spreadsheet --currency "£" --tax 20 --rows 800
"""

from __future__ import annotations

import argparse
import datetime as _dt
import os
import sys
from typing import List, Optional

from .build import SHEET_ORDER, build_workbook
from .config import (AGE_BANDS, AUTHOR, BILLING_BASES,
                     PRODUCT_LISTING_TITLE, PRODUCT_NAME, Config)

DEFAULT_OUTPUT = os.path.join(
    "output", "Daycare Business Management Spreadsheet.xlsx")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="daycare_business_spreadsheet",
        description=f"Build the {PRODUCT_NAME} workbook.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument("-o", "--output", default=DEFAULT_OUTPUT,
                        help="where to write the .xlsx file")
    parser.add_argument("--blank", action="store_true",
                        help="leave the registers empty (no demo "
                             "setting)")
    parser.add_argument("--no-protect", action="store_true",
                        help="do not protect the sheets (for "
                             "development)")
    parser.add_argument("--password", default=None,
                        help="sheet protection password")
    parser.add_argument("--author", default=AUTHOR,
                        help="document author / creator")
    parser.add_argument("--centre", default=None,
                        help="pre-fill the setting's name on Settings")
    parser.add_argument("--owner", default=None,
                        help="pre-fill the owner or director's name")
    parser.add_argument("--currency", default=None,
                        help="currency symbol, e.g. $ £ € A$")
    parser.add_argument("--currency-code", default=None,
                        help="three letter currency code, e.g. USD")
    parser.add_argument("--year", type=int, default=None,
                        help="the year the reports open on")
    parser.add_argument("--tax", type=float, default=None,
                        help="tax rate as a percentage, e.g. 20 for "
                             "20%% (childcare is exempt in many "
                             "places - the default is zero)")
    parser.add_argument("--tax-on", choices=["both", "tuition",
                                             "extras", "neither"],
                        default=None,
                        help="what the tax rate applies to")
    parser.add_argument("--capacity", type=int, default=None,
                        help="licensed places for the whole setting")
    parser.add_argument("--open-days", type=int, default=None,
                        help="days open a week")
    parser.add_argument("--open-hour", type=float, default=None,
                        help="opening time as a 24 hour number, "
                             "e.g. 7 for 7am")
    parser.add_argument("--close-hour", type=float, default=None,
                        help="closing time as a 24 hour number, "
                             "e.g. 18 for 6pm")
    parser.add_argument("--weeks-open", type=int, default=None,
                        help="trading weeks a year")
    parser.add_argument("--full-time-week", type=float, default=None,
                        help="the standard full-time weekly fee")
    parser.add_argument("--basis", choices=BILLING_BASES, default=None,
                        help="the billing basis new children start on")
    parser.add_argument("--sibling-discount", type=float, default=None,
                        help="sibling discount as a percentage, "
                             "e.g. 10")
    parser.add_argument("--hourly", type=float, default=None,
                        help="default staff hourly rate")
    parser.add_argument("--burden", type=float, default=None,
                        help="employer costs on wages as a percentage")
    parser.add_argument("--overtime-after", type=float, default=None,
                        help="hours a week before overtime starts")
    parser.add_argument("--target", type=float, default=None,
                        help="monthly revenue target")
    parser.add_argument("--rows", type=int, default=None,
                        help="scale every register (the default row "
                             "counts are scaled to fit this many "
                             "attendance rows)")
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
    if args.tax is not None:
        kwargs["tax_rate"] = round(args.tax / 100.0, 5)
    if args.tax_on:
        kwargs["tax_on_tuition"] = ("Yes" if args.tax_on in
                                    ("both", "tuition") else "No")
        kwargs["tax_on_extras"] = ("Yes" if args.tax_on in
                                   ("both", "extras") else "No")
    if args.capacity is not None:
        kwargs["licensed_capacity"] = args.capacity
    if args.open_days is not None:
        kwargs["open_days"] = args.open_days
    if args.open_hour is not None:
        kwargs["open_hour"] = args.open_hour
    if args.close_hour is not None:
        kwargs["close_hour"] = args.close_hour
    if args.weeks_open is not None:
        kwargs["weeks_open"] = args.weeks_open
    if args.full_time_week is not None:
        kwargs["full_time_week"] = args.full_time_week
    if args.basis:
        kwargs["default_billing_basis"] = args.basis
    if args.sibling_discount is not None:
        kwargs["sibling_discount"] = round(
            args.sibling_discount / 100.0, 4)
    if args.hourly is not None:
        kwargs["default_hourly"] = args.hourly
    if args.burden is not None:
        kwargs["payroll_burden"] = round(args.burden / 100.0, 4)
    if args.overtime_after is not None:
        kwargs["overtime_after_hours"] = args.overtime_after
    if args.target is not None:
        kwargs["monthly_revenue_target"] = args.target
        kwargs["sample_monthly_revenue_target"] = args.target
    if args.today:
        kwargs["today"] = _dt.date.fromisoformat(args.today)
    cfg = Config(**kwargs)
    if args.centre:
        cfg.centre.name = args.centre
    if args.owner:
        cfg.centre.owner = args.owner
    if args.rows:
        cfg.scale_rows(args.rows)
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

    tax_on = []
    if cfg.tax_on_tuition == "Yes":
        tax_on.append("tuition")
    if cfg.tax_on_extras == "Yes":
        tax_on.append("extras")

    bands = ", ".join(f"{label} 1:{per}"
                      for label, _f, _t, per, _m in AGE_BANDS[:3])

    print(f"{cfg.product_name}")
    print(f"  listing title : {PRODUCT_LISTING_TITLE}")
    print(f"  author        : {cfg.author}")
    print(f"  sheets        : {len(SHEET_ORDER)}")
    print(f"  year          : {cfg.year}")
    print(f"  setting       : {cfg.centre.name}")
    print(f"  capacity      : {cfg.licensed_capacity} places, "
          f"{cfg.open_days} days a week, "
          f"{cfg.open_hours_day:,.1f} hours a day")
    print(f"  ratios        : {bands} (editable on Settings)")
    print(f"  tax           : {cfg.tax_rate:.2%} on "
          f"{' and '.join(tax_on) if tax_on else 'nothing'}")
    print(f"  full-time fee : {cfg.currency_symbol}"
          f"{cfg.full_time_week:,.2f} a week, basis "
          f"'{cfg.default_billing_basis}'")
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

### `daycare_business_spreadsheet/config.py`

*733 lines · 3 classes · 1 function*

> Configuration for the Daycare & Childcare Business Management Spreadsheet.
>
> Everything a childcare provider might want to change - the currency, the
> age groups, the staff-to-child ratios, the fee basis, the late pickup
> charge, every drop-down list - lives in this one module, so the workbook
> can be regenerated for a different country, licence or price list
> without touching a single sheet.
>
> Nothing here assumes one country's licensing rules.  Ratios, age bands,
> tax, subsidy handling and document types are all settings, not
> hard-coded constants.
>
> Author of the generated workbooks: Novality Store.

| Defines | What it does |
|---------|--------------|
| `rid(kind, number)` | Build the ID the workbook itself will generate on row `number` |
| `class Palette` |  |
| `class CentreDefaults` | Sample values written into the business profile on Settings |
| `class Config` | Everything the generator needs to know to write one workbook |

**Module constants:** `AUTHOR` ('Novality Store'), `PRODUCT_NAME` ('Daycare & Childcare Business Management Spreadsheet'), `PRODUCT_SUBTITLE` ('Enrollment, Attendance, Tuition, Invoices, Daily Reports...), `PRODUCT_VERSION` ('1.0.0'), `PRODUCT_LISTING_TITLE` ('Daycare Business Spreadsheet, Childcare Management Syste...), `FORMULA_PASSWORD` ('premium'), `DISCLAIMER` ("This workbook is a business organising tool for a childc...), `PRIVACY_NOTE` ('Sensitive personal data. Keep this file in a password pr...), `PROGRAM_TYPES` (12 items), `ENROLLMENT_STATUSES` (8 items), `OCCUPYING_STATUSES` (3 items), `SCHEDULE_TYPES` (14 items), `ATTENDANCE_STATUSES` (11 items), `PRESENT_STATUSES` (3 items) and 53 more.

````python
"""
Configuration for the Daycare & Childcare Business Management Spreadsheet.

Everything a childcare provider might want to change - the currency, the
age groups, the staff-to-child ratios, the fee basis, the late pickup
charge, every drop-down list - lives in this one module, so the workbook
can be regenerated for a different country, licence or price list
without touching a single sheet.

Nothing here assumes one country's licensing rules.  Ratios, age bands,
tax, subsidy handling and document types are all settings, not
hard-coded constants.

Author of the generated workbooks: Novality Store.
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from datetime import date


# ---------------------------------------------------------------------------
# Product / branding constants
# ---------------------------------------------------------------------------

AUTHOR = "Novality Store"
PRODUCT_NAME = "Daycare & Childcare Business Management Spreadsheet"
PRODUCT_SUBTITLE = (
    "Enrollment, Attendance, Tuition, Invoices, Daily Reports, Staff, "
    "Classrooms and Profit in one workbook"
)
PRODUCT_VERSION = "1.0.0"

#: Title suggested for the shop listing.
PRODUCT_LISTING_TITLE = (
    "Daycare Business Spreadsheet, Childcare Management System, "
    "Enrollment Attendance Tuition Invoice Staff and Parent Tracker, "
    "Home Daycare and Preschool Excel Template"
)

#: Password that protects every sheet so formula cells cannot be edited
#: by accident.  Input cells stay unlocked, formula cells are locked.
FORMULA_PASSWORD = "premium"

#: The workbook organises a childcare business.  It is not a medical
#: record, not a licensing compliance system and not an accountant.
DISCLAIMER = (
    "This workbook is a business organising tool for a childcare "
    "setting. It is not accounting, payroll, legal, medical or "
    "licensing advice, it does not file or submit anything, and it "
    "does not make you compliant with any regulator. It does not send "
    "messages to parents and it does not connect to any childcare app "
    "or payment provider. Allergy, medication and incident pages are "
    "working notes for your own staff - they do not replace your "
    "setting's medical records, medication protocols, safeguarding "
    "procedures or emergency plans. Child and family records are "
    "sensitive personal data: keep this file somewhere access "
    "controlled and backed up, and set the ratios, age bands and "
    "document types to match the rules that actually apply where you "
    "operate."
)

#: Short line printed on the pages that carry the most sensitive data.
PRIVACY_NOTE = (
    "Sensitive personal data. Keep this file in a password protected, "
    "access controlled location, share it only with staff who need it, "
    "and delete records when your retention policy says to."
)


# ---------------------------------------------------------------------------
# Vocabularies - every one of these becomes a drop-down on Settings & Lists
# ---------------------------------------------------------------------------

#: The programmes a setting can offer.
PROGRAM_TYPES = [
    "Infant", "Young toddler", "Toddler", "Preschool", "Pre-K",
    "Kindergarten", "Before school", "After school", "Before & after",
    "Drop-in", "Holiday club", "Summer camp",
]

#: Where a child is up to with you.
ENROLLMENT_STATUSES = [
    "Enquiry", "Waitlisted", "Pending start", "Active", "On hold",
    "Notice given", "Withdrawn", "Graduated",
]

#: Statuses that mean the child currently occupies a seat.
OCCUPYING_STATUSES = ["Active", "On hold", "Notice given"]

#: How many days a week, and which part of the day.
SCHEDULE_TYPES = [
    "Full-time 5 days", "4 days", "3 days", "2 days", "1 day",
    "Half day 5 days", "Half day 3 days", "Mornings only",
    "Afternoons only", "Before school", "After school",
    "Before & after", "Drop-in", "Flexible hours",
]

#: What happened on the day.
ATTENDANCE_STATUSES = [
    "Present", "Late", "Early pickup", "Absent - sick",
    "Absent - family", "Absent - unexplained", "Holiday - child",
    "Public holiday", "Centre closed", "Snow day", "Not scheduled",
]

#: Attendance statuses that mean the child was in the building.
PRESENT_STATUSES = ["Present", "Late", "Early pickup"]

#: Statuses that mean a scheduled day was missed.
ABSENT_STATUSES = ["Absent - sick", "Absent - family",
                   "Absent - unexplained", "Holiday - child"]

#: Statuses where nobody was expected, so they do not count against the
#: attendance percentage.
NON_SESSION_STATUSES = ["Public holiday", "Centre closed", "Snow day",
                        "Not scheduled"]

#: How a family's tuition is worked out.  The owner picks per child -
#: nothing is assumed.
BILLING_BASES = [
    "Weekly rate", "Monthly rate", "Daily rate", "Hourly rate",
    "Per session", "Fortnightly rate", "Term rate", "No charge",
]

#: What a charge line is for.
FEE_TYPES = [
    "Tuition", "Registration fee", "Enrollment deposit",
    "Late pickup fee", "Late payment fee", "Meal fee", "Snack fee",
    "Transport fee", "Activity fee", "Field trip", "Supply fee",
    "Nappy & wipes fee", "Holiday club", "Extra day", "Extra hours",
    "Sibling discount", "Staff discount", "Promotional discount",
    "Subsidy credit", "Scholarship credit", "Goodwill credit",
    "Other charge",
]

#: Fee types that reduce the bill rather than add to it.  Entered as a
#: positive number and subtracted, so a parent never sees a minus sign.
CREDIT_FEE_TYPES = [
    "Sibling discount", "Staff discount", "Promotional discount",
    "Subsidy credit", "Scholarship credit", "Goodwill credit",
]

#: Where a charge came from - handy when a parent queries a bill.
CHARGE_SOURCES = [
    "Contracted schedule", "Extra day booked", "Extra hours",
    "Late pickup", "Meals", "Transport", "Activity or trip",
    "One-off request", "Registration", "Annual increase",
    "Manual adjustment",
]

#: Money movements on the Invoices & Receipts ledger.
LEDGER_TYPES = ["Invoice", "Payment", "Credit note", "Refund",
                "Write-off"]

#: Ledger types that increase what a family owes.
CHARGE_TYPES = ["Invoice", "Refund"]

#: How a family pays.
PAYMENT_METHODS = [
    "Cash", "Card", "Bank transfer", "Direct debit", "Standing order",
    "Cheque", "Mobile wallet", "Online payment", "Childcare voucher",
    "Government subsidy", "Employer paid", "Other",
]

#: Where an invoice is up to.
INVOICE_STATUSES = ["Draft", "Sent", "Part paid", "Paid", "Overdue",
                    "Written off", "Cancelled"]

#: Who a guardian is to the child.
RELATIONSHIPS = [
    "Mother", "Father", "Step-mother", "Step-father", "Grandmother",
    "Grandfather", "Legal guardian", "Foster carer", "Aunt", "Uncle",
    "Adult sibling", "Family friend", "Social worker", "Other",
]

#: How a family prefers to hear from you.
CONTACT_METHODS = [
    "Phone call", "Text message", "Email", "WhatsApp", "App message",
    "In person at pickup", "Paper note home", "Do not contact",
]

#: How a child seemed today.
MOODS = ["Happy", "Settled", "Playful", "Quiet", "Tired", "Clingy",
         "Tearful", "Unsettled", "Unwell", "Excited", "Frustrated"]

#: Sittings on the meal tracker.
MEAL_SITTINGS = ["Breakfast", "Morning snack", "Lunch",
                 "Afternoon snack", "Tea", "Dinner", "Bottle",
                 "Drink only"]

#: How much of it went in.
EATEN_AMOUNTS = ["All", "Most", "About half", "A little", "None",
                 "Refused", "Not offered", "Brought own"]

#: Dietary requirements and restrictions.
DIET_TYPES = [
    "No restrictions", "Vegetarian", "Vegan", "Halal", "Kosher",
    "Dairy free", "Gluten free", "Nut free", "Egg free", "Soy free",
    "Shellfish free", "Low sugar", "Texture modified", "Allergy plan",
    "Other - see notes",
]

#: How serious an allergy is - drives the colour coding.
ALLERGY_SEVERITY = ["None recorded", "Mild", "Moderate", "Severe",
                    "Anaphylaxis risk"]

#: What kind of medical record a row is.
MEDICAL_RECORD_TYPES = [
    "Allergy", "Intolerance", "Asthma", "Epilepsy", "Diabetes",
    "Eczema", "Medical condition", "Regular medication",
    "As-needed medication", "Dietary requirement", "Mobility need",
    "Additional need", "Immunisation", "Medication given",
    "Medication refused", "Other",
]

#: Medical rows that are an event rather than a standing profile.  The
#: medication administration log lives in the same register, flagged by
#: its record type, so a dose is always filed next to the plan it
#: belongs to and nothing has to be cross-referenced by hand.
MEDICAL_EVENT_TYPES = ["Medication given", "Medication refused"]

#: What happened.
INCIDENT_CATEGORIES = [
    "Minor injury", "Fall", "Bump to head", "Cut or graze", "Bruise",
    "Bite", "Scratch", "Allergic reaction", "Asthma episode",
    "Illness at the setting", "Medication error", "Behaviour incident",
    "Near miss", "Property damage", "Lost or found property",
    "Safeguarding concern", "Other",
]

#: Where it happened.
INCIDENT_LOCATIONS = [
    "Infant room", "Toddler room", "Preschool room", "Classroom",
    "Playground", "Garden", "Hallway", "Bathroom", "Nap room",
    "Kitchen", "Dining area", "Entrance", "Car park", "Outing",
    "Minibus", "Other",
]

#: How badly hurt, if at all.
INCIDENT_SEVERITY = ["No injury", "Minor", "Moderate", "Serious",
                     "Emergency services called"]

#: Jobs in a childcare setting.
STAFF_ROLES = [
    "Owner / Director", "Centre manager", "Deputy manager",
    "Room leader", "Lead teacher", "Teacher", "Assistant teacher",
    "Early years practitioner", "Infant room staff", "Floater",
    "Apprentice / trainee", "Substitute / supply", "Cook",
    "Kitchen assistant", "Cleaner", "Administrator", "Bus driver",
    "Volunteer", "Other",
]

#: Where a staff member is up to.
EMPLOYMENT_STATUSES = ["Active", "Probation", "On leave",
                       "Casual / bank", "Notice period", "Inactive",
                       "Left"]

#: How they are paid.
PAY_TYPES = ["Hourly", "Salary", "Daily rate", "Contractor",
             "Volunteer - unpaid"]

#: Shifts on the rota.
SHIFT_TYPES = ["Opening", "Early", "Morning", "Mid", "Afternoon",
               "Late", "Closing", "Full day", "Split shift",
               "Sub cover", "Training", "Admin", "Off", "Leave"]

#: Why somebody is not in.
LEAVE_TYPES = ["Annual leave", "Sick leave", "Unpaid leave",
               "Parental leave", "Compassionate leave", "Training",
               "Public holiday", "TOIL"]

#: How often a pay run happens.
PAY_PERIODS = ["Weekly", "Fortnightly", "Four weekly", "Monthly",
               "Twice monthly", "One off"]

#: Where the money goes, in the order a profit and loss wants it.
EXPENSE_CATEGORIES = [
    "Rent", "Mortgage & loan", "Utilities", "Food & groceries",
    "Nappies & hygiene", "Cleaning supplies", "Toys & learning materials",
    "Educational supplies", "Arts & crafts", "Furniture & equipment",
    "Staff wages", "Payroll taxes & on-costs", "Staff benefits",
    "Recruitment", "Training & qualifications", "Insurance",
    "Licensing & registration", "Inspection & compliance",
    "Repairs & maintenance", "Transport & vehicle", "Field trips",
    "Marketing & advertising", "Software & subscriptions",
    "Telephone & internet", "Office & admin", "Bank & card fees",
    "Professional fees", "Waste & recycling", "Security",
    "Other expense",
]

#: Categories whose cost is already on the Payroll Tracker.  The profit
#: and loss skips them unless the owner says otherwise on Settings.
PAYROLL_EXPENSE_CATEGORIES = ["Staff wages", "Payroll taxes & on-costs"]

#: What is in the store cupboard.
SUPPLY_CATEGORIES = [
    "Food & kitchen", "Drinks & formula", "Cleaning & sanitising",
    "Nappies & wipes", "Hygiene & bathroom", "Arts & crafts",
    "Learning materials", "Books", "Toys", "Outdoor equipment",
    "Office & stationery", "First aid", "Safety & PPE",
    "Linen & bedding", "Other supplies",
]

#: Paperwork a setting has to keep on top of.
DOCUMENT_TYPES = [
    "Enrollment agreement", "Emergency contact form",
    "Medical authorization", "Immunisation record",
    "Photo & video consent", "Pickup authorization",
    "Allergy action plan", "Medication consent",
    "Parent handbook acknowledgment", "Field trip consent",
    "Sunscreen consent", "Birth certificate sighted",
    "Subsidy paperwork", "Custody / court order",
    "Staff contract", "Staff ID check", "Background check",
    "First aid certificate", "CPR certificate",
    "Child protection training", "Food handling certificate",
    "Qualification certificate", "Reference check",
    "Business licence", "Insurance policy", "Fire inspection",
    "Health inspection", "Risk assessment", "Policy review", "Other",
]

#: Who a document belongs to.
DOCUMENT_HOLDERS = ["Child", "Parent / family", "Staff", "Business"]

#: Where a document is up to.
DOCUMENT_STATUSES = ["Not started", "Requested", "Chased", "Received",
                     "Signed", "On file", "Expired", "Not required"]

#: What a conversation with a parent was about.
COMM_CATEGORIES = [
    "General update", "Settling in", "Attendance", "Late pickup",
    "Behaviour", "Health & illness", "Medication", "Allergy",
    "Accident / incident", "Billing & fees", "Enrollment change",
    "Schedule change", "Meeting request", "Progress update",
    "Complaint", "Compliment", "Policy notice", "Closure notice",
    "Photo consent", "Other",
]

#: Where a conversation is up to.
COMM_STATUSES = ["Open", "Waiting on parent", "Waiting on us",
                 "Meeting booked", "Closed", "No reply needed"]

#: How a family found you.
INQUIRY_SOURCES = [
    "Google search", "Google Maps", "Facebook", "Instagram",
    "Word of mouth", "Parent referral", "Staff referral",
    "Sign outside", "Flyer or leaflet", "Local directory",
    "Childcare website", "Employer scheme", "School referral",
    "Walk in", "Returning family", "Sibling already here", "Other",
]

#: Where an enquiry is up to on the waitlist.
WAITLIST_STATUSES = [
    "New enquiry", "Information sent", "Tour booked", "Toured",
    "Place offered", "Offer accepted", "Offer declined", "Enrolled",
    "Waiting - no space", "Waiting - future date", "Lost", "Withdrawn",
]

#: Priority on the waitlist.
WAITLIST_PRIORITY = ["Sibling", "Staff child", "Returning family",
                     "Standard", "Low"]

#: Why a family left.
LEAVING_REASONS = [
    "Aged out / started school", "Moved away", "Parent left work",
    "Cost", "Changed provider", "Hours did not suit",
    "Parental leave", "Unhappy with service", "Not stated", "Other",
]

#: How often a recurring cost comes round.
FREQUENCIES = ["One off", "Weekly", "Fortnightly", "Four weekly",
               "Monthly", "Quarterly", "Twice a year", "Yearly"]

#: Months per occurrence, used to turn any frequency into a monthly cost.
FREQUENCY_MONTHS = {
    "Weekly": 52.0 / 12.0,
    "Fortnightly": 26.0 / 12.0,
    "Four weekly": 13.0 / 12.0,
    "Monthly": 1.0,
    "Quarterly": 1.0 / 3.0,
    "Twice a year": 1.0 / 6.0,
    "Yearly": 1.0 / 12.0,
}

#: Offered because some settings record it and some deliberately do not.
GENDERS = ["Girl", "Boy", "Prefer not to say", "Not recorded"]

#: Who started the conversation.
COMM_DIRECTIONS = ["Incoming", "Outgoing"]

#: Where a classroom is up to.
ROOM_STATUSES = ["Open", "Full", "Planned", "Closed", "Seasonal"]

#: Nappy / toileting entries on the daily report.
TOILETING_OPTIONS = ["Nappy - wet", "Nappy - soiled", "Nappy - dry",
                     "Potty - success", "Potty - tried", "Toilet",
                     "Accident", "Not applicable"]


# ---------------------------------------------------------------------------
# Age bands and the staff-to-child ratios that go with them.
#
# These are the single most country-specific numbers in the whole
# product, which is exactly why they are a table on the Settings tab
# rather than a constant in a formula.  Defaults are a middle-of-the-road
# starting point, nothing more.
# ---------------------------------------------------------------------------

#: (label, from months, to months, children per adult, max group size)
AGE_BANDS = [
    ("Infant", 0, 12, 4, 8),
    ("Young toddler", 12, 24, 5, 10),
    ("Toddler", 24, 36, 6, 12),
    ("Preschool", 36, 48, 10, 20),
    ("Pre-K", 48, 60, 12, 22),
    ("School age", 60, 156, 15, 30),
]


# ---------------------------------------------------------------------------
# ID prefixes.  Each one is an editable field on Settings with a defined
# name, so a setting can renumber its own records without touching a
# formula: (label, default, defined name, digits).
# ---------------------------------------------------------------------------

ID_PREFIXES = [
    ("Child IDs", "CHD-", "ChildPrefix", 4),
    ("Family IDs", "FAM-", "FamilyPrefix", 4),
    ("Waitlist IDs", "WL-", "WaitlistPrefix", 4),
    ("Attendance IDs", "ATT-", "AttendancePrefix", 5),
    ("Charge IDs", "FEE-", "FeePrefix", 5),
    ("Invoice IDs", "INV-", "InvoicePrefix", 5),
    ("Receipt IDs", "RCP-", "ReceiptPrefix", 5),
    ("Daily report IDs", "DR-", "ReportPrefix", 5),
    ("Meal log IDs", "MEA-", "MealPrefix", 5),
    ("Medical record IDs", "MED-", "MedicalPrefix", 4),
    ("Incident IDs", "INC-", "IncidentPrefix", 4),
    ("Staff IDs", "EMP-", "StaffPrefix", 3),
    ("Shift IDs", "SHF-", "ShiftPrefix", 5),
    ("Pay run IDs", "PAY-", "PayrollPrefix", 4),
    ("Classroom IDs", "ROOM-", "ClassroomPrefix", 2),
    ("Expense IDs", "EXP-", "ExpensePrefix", 4),
    ("Supply IDs", "SUP-", "SupplyPrefix", 3),
    ("Message IDs", "MSG-", "MessagePrefix", 4),
    ("Document IDs", "DOC-", "DocumentPrefix", 4),
]

#: defined name -> (prefix, digits), for the demo data generator.
PREFIX_BY_NAME = {name: (prefix, digits)
                  for _label, prefix, name, digits in ID_PREFIXES}


def rid(kind: str, number: int) -> str:
    """Build the ID the workbook itself will generate on row `number`.

    ``rid("Child", 7)`` -> ``"CHD-0007"``.  The demo data has to agree
    with the ID formulas exactly, because the IDs are what tie the
    twenty-three tabs together.
    """
    prefix, digits = PREFIX_BY_NAME[f"{kind}Prefix"]
    return f"{prefix}{number:0{digits}d}"


#: The twelve month names used by every monthly roll-up.
MONTH_NAMES = ["January", "February", "March", "April", "May", "June",
               "July", "August", "September", "October", "November",
               "December"]

#: Weekday names, Monday first.
WEEKDAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday",
                 "Saturday", "Sunday"]


# ---------------------------------------------------------------------------
# Colours - a childcare palette: teal, sky, warm amber, soft cream
# ---------------------------------------------------------------------------

class Palette:
    ink = "#1E333C"            # near black teal - titles
    deep = "#14505F"           # deep teal - money and staff tabs
    primary = "#1F6F8B"        # teal headers
    primary_mid = "#4D96AE"    # section bands
    primary_soft = "#E7F2F6"   # light band / alternating rows
    accent = "#E09B2D"         # warm amber accent
    accent_soft = "#FDF3E0"
    input_bg = "#FFFFFF"       # editable cell background
    input_border = "#D3E4EB"
    calc_bg = "#F5FAFC"        # locked formula cell background
    calc_border = "#E3EFF4"
    white = "#FFFFFF"
    grey_text = "#5B6B73"
    line = "#C9DDE5"
    good_bg = "#DBEDD7"
    good_text = "#2C5A28"
    warn_bg = "#FDEDD2"
    warn_text = "#8A5A00"
    bad_bg = "#F8DAD9"
    bad_text = "#9C2A24"
    info_bg = "#E2E7F3"
    info_text = "#36447A"
    note_bg = "#F7FBFC"


# ---------------------------------------------------------------------------
# Row capacity of every register
# ---------------------------------------------------------------------------

DEFAULT_ROWS = {
    "Dashboard": 0,
    "Getting Started": 0,
    "Settings & Lists": 0,
    "Child Enrollment": 150,
    "Parent Database": 200,
    "Waitlist": 120,
    "Attendance": 2000,
    "Tuition & Fees": 700,
    "Invoices & Receipts": 700,
    "Daily Reports": 600,
    "Meal Tracker": 600,
    "Allergy & Medical": 180,
    "Incident Reports": 150,
    "Staff Database": 60,
    "Staff Schedule": 750,
    "Payroll Tracker": 350,
    "Classroom Management": 30,
    "Expense Tracker": 500,
    "Supply Inventory": 180,
    "Parent Communication": 350,
    "Document Tracker": 500,
    "Monthly Reports": 0,
    "Annual Reports": 0,
}

#: The register the `--rows` switch scales everything else against.
SCALE_ANCHOR = "Attendance"


@dataclass
class CentreDefaults:
    """Sample values written into the business profile on Settings."""

    name: str = "Little Acorns Childcare"
    owner: str = "Marianne Okafor"
    brand: str = "Little Acorns"
    email: str = "hello@littleacorns.example"
    phone: str = "(614) 555 0142"
    address: str = "47 Willow Grove, Columbus, OH 43206"
    website: str = "littleacorns.example"
    licence: str = "OH-CC-2017-04418"
    tax_id: str = "EIN 84-3920117"
    opening_hours: str = "Mon-Fri 7:00am - 6:00pm"
    tagline: str = "A small, warm place to grow since 2017"


# ---------------------------------------------------------------------------
# The configuration object every sheet module reads
# ---------------------------------------------------------------------------

@dataclass
class Config:
    """Everything the generator needs to know to write one workbook."""

    author: str = AUTHOR
    product_name: str = PRODUCT_NAME
    product_subtitle: str = PRODUCT_SUBTITLE
    listing_title: str = PRODUCT_LISTING_TITLE
    version: str = PRODUCT_VERSION
    password: str = FORMULA_PASSWORD
    protect_sheets: bool = True
    include_sample_data: bool = True

    # Formatting ------------------------------------------------------------
    currency_symbol: str = "$"
    currency_code: str = "USD"
    date_format: str = "dd-mmm-yyyy"
    time_format: str = "h:mm AM/PM"
    decimals: int = 2
    week_start: str = "Monday"

    # The setting's own profile, pre-filled on Settings ---------------------
    centre: CentreDefaults = field(default_factory=CentreDefaults)
    plan_year: int = 0                   # 0 = the year of `today`

    # Opening hours and capacity -------------------------------------------
    open_hour: float = 7.0
    close_hour: float = 18.0
    open_days: int = 5                   # sessions a week
    licensed_capacity: int = 42          # the licence number on the wall
    classrooms: int = 4
    weeks_open: int = 50                 # trading weeks in a year
    term_weeks: int = 13                 # weeks in a term, for term fees

    # Fees ------------------------------------------------------------------
    default_billing_basis: str = "Weekly rate"
    full_time_week: float = 285.00
    part_time_day: float = 68.00
    hourly_rate_fee: float = 11.50
    registration_fee: float = 95.00
    enrollment_deposit: float = 200.00
    meal_fee_day: float = 6.50
    transport_fee_week: float = 25.00
    sibling_discount: float = 0.10       # second and later child
    staff_discount: float = 0.30
    late_pickup_grace: int = 10          # minutes before the clock starts
    late_pickup_fee: float = 15.00       # the first block
    late_pickup_per: float = 1.00        # per minute after the block
    late_pickup_block: int = 15          # minutes in the first block
    late_payment_fee: float = 25.00
    payment_terms_days: int = 7          # invoice due this many days out
    invoice_day: int = 1                 # day of the month you invoice
    tax_rate: float = 0.0                # childcare is often exempt
    tax_on_tuition: str = "No"
    tax_on_extras: str = "No"
    tax_registered: str = "No"
    card_fee_percent: float = 0.029
    card_fee_fixed: float = 0.30

    # Staff pay -------------------------------------------------------------
    default_pay_type: str = "Hourly"
    default_hourly: float = 16.50
    overtime_after_hours: float = 40.0   # hours a week before overtime
    overtime_multiplier: float = 1.5
    payroll_burden: float = 0.14         # employer tax, insurance, cover
    paid_break_minutes: int = 0          # unpaid break length, in minutes
    break_minutes: int = 30

    # Ratios and capacity ---------------------------------------------------
    ratio_warning_percent: float = 0.90  # flag a room at 90% full
    include_staff_children: str = "Yes"

    # Targets the dashboard measures against --------------------------------
    monthly_revenue_target: float = 42000.0
    annual_revenue_target: float = 504000.0
    monthly_profit_target: float = 6000.0
    target_occupancy: float = 0.90
    target_attendance: float = 0.92
    target_collection: float = 0.97      # share of billings collected
    enrollment_target: int = 40
    # the example setting is smaller than the defaults assume, so the
    # demo file carries targets it can actually be measured against
    sample_monthly_revenue_target: float = 30000.0
    sample_annual_revenue_target: float = 360000.0
    sample_monthly_profit_target: float = 3000.0
    sample_enrollment_target: int = 34

    # Supplies --------------------------------------------------------------
    reorder_weeks: float = 3.0
    stock_warning_weeks: float = 1.5

    # Reminders -------------------------------------------------------------
    document_warning_days: int = 45      # amber before an expiry
    document_urgent_days: int = 14
    birthday_window_days: int = 30
    absence_chase_days: int = 3          # unexplained absences in a row
    overdue_chase_days: int = 7

    # Behaviour -------------------------------------------------------------
    top_rows: int = 5                    # size of the "top / best" tables
    months_shown: int = 12
    rows: dict = field(default_factory=lambda: dict(DEFAULT_ROWS))
    today: date = field(default_factory=date.today)

    # ---------------------------------------------------------------- helpers
    def row_count(self, sheet: str) -> int:
        return int(self.rows.get(sheet, 100))

    @property
    def year(self) -> int:
        return self.plan_year or self.today.year

    @property
    def revenue_target_month(self) -> float:
        """Monthly revenue target, scaled to whoever is reading it."""
        return (self.sample_monthly_revenue_target
                if self.include_sample_data else self.monthly_revenue_target)

    @property
    def revenue_target_year(self) -> float:
        return (self.sample_annual_revenue_target
                if self.include_sample_data else self.annual_revenue_target)

    @property
    def profit_target_month(self) -> float:
        return (self.sample_monthly_profit_target
                if self.include_sample_data else self.monthly_profit_target)

    @property
    def enrollment_goal(self) -> int:
        return (self.sample_enrollment_target
                if self.include_sample_data else self.enrollment_target)

    @property
    def open_hours_day(self) -> float:
        return round(self.close_hour - self.open_hour, 2)

    @property
    def labor_cost_per_hour(self) -> float:
        """An hour of paid time once employer on-costs are added."""
        return round(self.default_hourly * (1 + self.payroll_burden), 2)

    @property
    def capacity_weeks_year(self) -> int:
        """Licensed child-weeks available in a year - the occupancy base."""
        return int(self.licensed_capacity * self.weeks_open)

    def scale_rows(self, main_rows: int) -> None:
        """Scale every register in place so Attendance holds N rows."""
        base = DEFAULT_ROWS[SCALE_ANCHOR]
        factor = max(0.02, float(main_rows) / float(base))
        self.rows = {k: (max(5, int(round(v * factor))) if v else v)
                     for k, v in self.rows.items()}

    def scaled(self, factor: float) -> "Config":
        """Return a copy with every register scaled."""
        rows = {k: max(5, int(round(v * factor))) if v else v
                for k, v in self.rows.items()}
        return replace(self, rows=rows)

    # Number format strings derived from the chosen currency ----------------
    @property
    def money_format(self) -> str:
        s = self.currency_symbol
        return f'{s}#,##0.00;[Red]-{s}#,##0.00'

    @property
    def money_format_0(self) -> str:
        s = self.currency_symbol
        return f'{s}#,##0;[Red]-{s}#,##0'

    @property
    def money_format_4(self) -> str:
        s = self.currency_symbol
        return f'{s}#,##0.0000;[Red]-{s}#,##0.0000'
````

### `daycare_business_spreadsheet/layout.py`

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

### `daycare_business_spreadsheet/formats.py`

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

### `daycare_business_spreadsheet/writer.py`

*458 lines · 1 class · 8 functions*

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

**Module constants:** `P` (Palette), `NAV_LINKS` (9 items), `PROTECT_OPTIONS` (15 entries).

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
    ("Children", "Child Enrollment"),
    ("Parents", "Parent Database"),
    ("Attendance", "Attendance"),
    ("Invoices", "Invoices & Receipts"),
    ("Staff", "Staff Database"),
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
            self.ws.write_formula(row, col, value, fmt)
        elif isinstance(value, (_dt.datetime, _dt.date, _dt.time)):
            self.ws.write_datetime(row, col, value, fmt)
        elif isinstance(value, bool):
            self.ws.write_string(row, col, "Yes" if value else "No", fmt)
        elif isinstance(value, (int, float)):
            self.ws.write_number(row, col, value, fmt)
        else:
            self.ws.write_string(row, col, str(value), fmt)

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
        self.ws.merge_range(row + 1, col, row + 1, col + span - 1, "", fmt)
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

### `daycare_business_spreadsheet/names.py`

*63 lines · 1 function*

> Workbook level defined names that point at the register sheets.
>
> The drop-down lists on the `Settings & Lists` tab define their own
> names.  These are the ones that have to follow live data - child IDs,
> family IDs, the team, the classrooms, the store cupboard - so they are
> built from the registry once every register exists.  Each one is an
> ``OFFSET`` range that grows and shrinks with the number of filled rows,
> which is why a child you enroll this morning turns up in the attendance
> drop-down this afternoon.

| Defines | What it does |
|---------|--------------|
| `define_cross_sheet_names(book, reg)` | Create the live lookup names |

**Module constants:** `CROSS_SHEET_NAMES` (29 entries).

````python
"""Workbook level defined names that point at the register sheets.

The drop-down lists on the `Settings & Lists` tab define their own
names.  These are the ones that have to follow live data - child IDs,
family IDs, the team, the classrooms, the store cupboard - so they are
built from the registry once every register exists.  Each one is an
``OFFSET`` range that grows and shrinks with the number of filled rows,
which is why a child you enroll this morning turns up in the attendance
drop-down this afternoon.
"""

from __future__ import annotations

from typing import Dict, Tuple

from .layout import Registry, dynamic_name

#: defined name -> (sheet, column key)
CROSS_SHEET_NAMES: Dict[str, Tuple[str, str]] = {
    # the spine of the workbook - every drop-down that must follow data
    "ChildList": ("Child Enrollment", "id"),
    "ChildNames": ("Child Enrollment", "name"),
    "FamilyList": ("Parent Database", "id"),
    "FamilyNames": ("Parent Database", "parent1"),
    "StaffList": ("Staff Database", "id"),
    "StaffNames": ("Staff Database", "name"),
    "ClassroomList": ("Classroom Management", "name"),
    "SupplyList": ("Supply Inventory", "id"),
    "SupplyNames": ("Supply Inventory", "name"),
    "InvoiceList": ("Invoices & Receipts", "id"),
    "WaitlistIDs": ("Waitlist", "id"),
    # columns the dashboard, the analytics and the reports scan
    "ChildStatusCol": ("Child Enrollment", "status"),
    "ChildRoomCol": ("Child Enrollment", "room"),
    "ChildProgramCol": ("Child Enrollment", "program"),
    "ChildDobCol": ("Child Enrollment", "dob"),
    "ChildFamilyCol": ("Child Enrollment", "family"),
    "AttendanceDates": ("Attendance", "date"),
    "AttendanceStatusCol": ("Attendance", "status"),
    "AttendanceChildCol": ("Attendance", "child"),
    "FeeMonths": ("Tuition & Fees", "month"),
    "FeeNet": ("Tuition & Fees", "net"),
    "LedgerMonths": ("Invoices & Receipts", "month"),
    "LedgerBilled": ("Invoices & Receipts", "billed"),
    "LedgerCollected": ("Invoices & Receipts", "collected"),
    "ExpenseMonths": ("Expense Tracker", "month"),
    "ExpenseCategoryCol": ("Expense Tracker", "category"),
    "ExpenseAmounts": ("Expense Tracker", "amount"),
    "PayrollMonths": ("Payroll Tracker", "month"),
    "PayrollCost": ("Payroll Tracker", "total_cost"),
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

### `daycare_business_spreadsheet/build.py`

*162 lines · 3 functions*

> Workbook builder: puts the twenty-three tabs together in order.
>
> Typical use::
>
>     from daycare_business_spreadsheet import build_workbook
>     build_workbook("Daycare Business Management Spreadsheet.xlsx")

| Defines | What it does |
|---------|--------------|
| `build_register(book, fmts, cfg, reg, module)` | Create one register sheet from its :class:`SheetSpec` |
| `build_workbook(path, cfg)` | Write the whole product to *path* and return the path |
| `_set_properties(book, cfg)` |  |

**Module constants:** `REGISTER_MODULES` (18 items), `SHEET_ORDER` (23 items), `PAGE_BUILDERS` (5 entries), `REGISTER_BY_NAME` ({module.NAME: module for module in REGISTER_MODULES}).

````python
"""Workbook builder: puts the twenty-three tabs together in order.

Typical use::

    from daycare_business_spreadsheet import build_workbook
    build_workbook("Daycare Business Management Spreadsheet.xlsx")
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
from .sheets import (annual, attendance, children, classrooms,
                     communication, daily_reports, dashboard,
                     documents, expenses, getting_started, incidents,
                     invoices, meals, medical, monthly, parents,
                     payroll, schedule, settings, staff, supplies,
                     tuition, waitlist)
from .writer import SheetWriter

#: register modules, in the order their specs are registered
REGISTER_MODULES = [
    classrooms, parents, children, waitlist, attendance, tuition,
    invoices, daily_reports, meals, medical, incidents, staff,
    schedule, payroll, expenses, supplies, communication, documents,
]

#: tab order of the finished workbook
SHEET_ORDER: List[str] = [
    "Dashboard",
    "Getting Started",
    "Settings & Lists",
    "Child Enrollment",
    "Parent Database",
    "Waitlist",
    "Attendance",
    "Tuition & Fees",
    "Invoices & Receipts",
    "Daily Reports",
    "Meal Tracker",
    "Allergy & Medical",
    "Incident Reports",
    "Staff Database",
    "Staff Schedule",
    "Payroll Tracker",
    "Classroom Management",
    "Expense Tracker",
    "Supply Inventory",
    "Parent Communication",
    "Document Tracker",
    "Monthly Reports",
    "Annual Reports",
]

#: sheets that are not registers - name -> builder(book, fmts, cfg, reg)
PAGE_BUILDERS: Dict[str, Callable] = {
    "Dashboard": dashboard.build,
    "Getting Started": getting_started.build,
    "Settings & Lists": settings.build,
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


def build_workbook(path: str = "Daycare Business Spreadsheet.xlsx",
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

    # 3. Live lookup names (child IDs, family IDs, rooms, the team).
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
        "subject": "Childcare enrollment, attendance, tuition, "
                   "invoicing, staff, classrooms and parent records",
        "author": cfg.author,
        "manager": cfg.author,
        "company": cfg.author,
        "category": "Small business spreadsheet template",
        "keywords": ("daycare spreadsheet, childcare management "
                     "system, child care business planner, enrollment "
                     "tracker, attendance tracker, tuition invoice "
                     "template, preschool spreadsheet, nursery "
                     "management, childminder tracker, staff "
                     "scheduling, classroom ratio, spreadsheet "
                     "template"),
        "comments": (f"{cfg.product_name}. Created by {cfg.author}. "
                     f"Version {cfg.version}. Formula cells are locked "
                     f"and every sheet is protected with the password "
                     f"'{cfg.password}'."),
        "status": "Final",
        "created": _dt.datetime(cfg.today.year, cfg.today.month,
                                cfg.today.day),
    })
````

### `daycare_business_spreadsheet/sample_data.py`

*1,989 lines · 1 class · 11 functions*

> Demo data: ten weeks at Little Acorns Childcare.
>
> A forty-two place setting in Columbus, Ohio.  Four rooms, twelve staff,
> twenty families and thirty children - some of them siblings, which is
> the whole point.  One family pays by the month, one by the hour, one
> is on a government subsidy, one is three weeks late and one left in
> March.  There is a child on an adrenaline pen, a staff member whose
> first aid certificate runs out in a fortnight, a bumped head that was
> reported to the parent in eleven minutes, and a room sitting at ninety
> per cent of its licensed places.
>
> Every figure hangs together.  The register feeds the billing, the
> billing feeds the ledger, the rota feeds the payroll and the payroll
> lands in the profit and loss - so when you change one thing you can
> watch the rest of the workbook move.
>
> Everything is dated relative to the build date, so the file never
> looks stale, and `--blank` strips the lot out for your own children.

| Defines | What it does |
|---------|--------------|
| `d(base, offset)` |  |
| `t(hour, minute)` |  |
| `money(value)` |  |
| `month_key(day)` |  |
| `month_name(day)` |  |
| `_month_start(day, back)` |  |
| `_month_end(first)` |  |
| `_weekdays(start, end)` |  |
| `_months_ago(day, months)` | `day` shifted back whole months, clamped to a valid day |
| `class Sample` | fields: classrooms, children, parents, waitlist, attendance, fees, ledger, daily_reports |
| `get(cfg)` | Build (and cache) the demo data for this configuration |
| `_build(cfg)` |  |

**Module constants:** `SEED` (20261031), `ROOMS` (4 items), `FAMILIES` (20 items), `CHILDREN` (34 items), `STAFF` (12 items), `SUPPLIES` (28 items).

````python
"""Demo data: ten weeks at Little Acorns Childcare.

A forty-two place setting in Columbus, Ohio.  Four rooms, twelve staff,
twenty families and thirty children - some of them siblings, which is
the whole point.  One family pays by the month, one by the hour, one
is on a government subsidy, one is three weeks late and one left in
March.  There is a child on an adrenaline pen, a staff member whose
first aid certificate runs out in a fortnight, a bumped head that was
reported to the parent in eleven minutes, and a room sitting at ninety
per cent of its licensed places.

Every figure hangs together.  The register feeds the billing, the
billing feeds the ledger, the rota feeds the payroll and the payroll
lands in the profit and loss - so when you change one thing you can
watch the rest of the workbook move.

Everything is dated relative to the build date, so the file never
looks stale, and `--blank` strips the lot out for your own children.
"""

from __future__ import annotations

import datetime as _dt
import random
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

from .config import Config, rid

SEED = 20261031


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def d(base: _dt.date, offset: int) -> _dt.date:
    return base + _dt.timedelta(days=offset)


def t(hour: int, minute: int = 0) -> _dt.time:
    return _dt.time(hour, minute)


def money(value: float) -> float:
    return round(value + 1e-9, 2)


def month_key(day: _dt.date) -> str:
    return f"{day.year:04d}-{day.month:02d}"


def month_name(day: _dt.date) -> str:
    return day.strftime("%B %Y")


def _month_start(day: _dt.date, back: int = 0) -> _dt.date:
    year, month = day.year, day.month - back
    while month <= 0:
        month += 12
        year -= 1
    return _dt.date(year, month, 1)


def _month_end(first: _dt.date) -> _dt.date:
    if first.month == 12:
        return _dt.date(first.year, 12, 31)
    return _dt.date(first.year, first.month + 1, 1) - _dt.timedelta(days=1)


def _weekdays(start: _dt.date, end: _dt.date) -> List[_dt.date]:
    out: List[_dt.date] = []
    day = start
    while day <= end:
        if day.weekday() < 5:
            out.append(day)
        day += _dt.timedelta(days=1)
    return out


def _months_ago(day: _dt.date, months: int) -> _dt.date:
    """`day` shifted back whole months, clamped to a valid day."""
    year, month = day.year, day.month - months
    while month <= 0:
        month += 12
        year -= 1
    last = _month_end(_dt.date(year, month, 1)).day
    return _dt.date(year, month, min(day.day, last))


@dataclass
class Sample:
    classrooms: List[Dict[str, Any]] = field(default_factory=list)
    children: List[Dict[str, Any]] = field(default_factory=list)
    parents: List[Dict[str, Any]] = field(default_factory=list)
    waitlist: List[Dict[str, Any]] = field(default_factory=list)
    attendance: List[Dict[str, Any]] = field(default_factory=list)
    fees: List[Dict[str, Any]] = field(default_factory=list)
    ledger: List[Dict[str, Any]] = field(default_factory=list)
    daily_reports: List[Dict[str, Any]] = field(default_factory=list)
    meals: List[Dict[str, Any]] = field(default_factory=list)
    medical: List[Dict[str, Any]] = field(default_factory=list)
    incidents: List[Dict[str, Any]] = field(default_factory=list)
    staff: List[Dict[str, Any]] = field(default_factory=list)
    schedule: List[Dict[str, Any]] = field(default_factory=list)
    payroll: List[Dict[str, Any]] = field(default_factory=list)
    expenses: List[Dict[str, Any]] = field(default_factory=list)
    supplies: List[Dict[str, Any]] = field(default_factory=list)
    messages: List[Dict[str, Any]] = field(default_factory=list)
    documents: List[Dict[str, Any]] = field(default_factory=list)


_CACHE: Dict[int, Sample] = {}


def get(cfg: Config) -> Sample:
    """Build (and cache) the demo data for this configuration."""
    key = id(cfg)
    if key not in _CACHE:
        _CACHE[key] = _build(cfg) if cfg.include_sample_data else Sample()
    return _CACHE[key]


# ---------------------------------------------------------------------------
# The rooms.  Capacities add up to the forty-two on the licence.
# ---------------------------------------------------------------------------

ROOMS: List[Dict[str, Any]] = [
    dict(name="Acorns", age_band="Infant", capacity=8, ratio_set=4,
         teacher=2, assistant=6, room_status="Open",
         notes="Babies from twelve weeks. Sleep room off the back, "
               "two cots short of full."),
    dict(name="Saplings", age_band="Young toddler", capacity=10,
         ratio_set=5, teacher=3, assistant=7, room_status="Open",
         notes="Walkers and early talkers. Busiest room in the "
               "building and the one the waitlist is for."),
    dict(name="Oaks", age_band="Toddler", capacity=12, ratio_set=6,
         teacher=4, assistant=8, room_status="Open",
         notes="Two to three year olds. Direct door to the garden."),
    dict(name="Willows", age_band="Preschool", capacity=12,
         ratio_set=10, teacher=5, assistant=9, room_status="Open",
         notes="School readiness, plus the before and after school "
               "children in term time."),
]


# ---------------------------------------------------------------------------
# The families.  Four of them have two children here, one has three -
# which is how the sibling discount and the shared invoice get a real
# workout rather than a demonstration.
# ---------------------------------------------------------------------------

FAMILIES: List[Tuple[Any, ...]] = [
    # surname, parent1, rel1, parent2, rel2, street, language, contact
    ("Okonkwo", "Adaeze Okonkwo", "Mother", "Chidi Okonkwo", "Father",
     "12 Bellflower Lane", "English", "App message"),
    ("Hartley", "Megan Hartley", "Mother", "Tom Hartley", "Father",
     "8 Kestrel Court", "English", "Email"),
    ("Nowak", "Kasia Nowak", "Mother", "", "",
     "214 Mulberry Street, Apt 3", "Polish", "Text message"),
    ("Delgado", "Rosa Delgado", "Mother", "Javier Delgado", "Father",
     "61 Sycamore Drive", "Spanish", "Phone call"),
    ("Brennan", "Claire Brennan", "Mother", "Niall Brennan", "Father",
     "19 Larkspur Way", "English", "App message"),
    ("Ahmed", "Yusra Ahmed", "Mother", "Imran Ahmed", "Father",
     "33 Chestnut Row", "Urdu", "WhatsApp"),
    ("Whitfield", "Jo Whitfield", "Mother", "Sam Whitfield",
     "Step-father", "7 Fieldfare Close", "English", "Email"),
    ("Petrova", "Elena Petrova", "Mother", "", "",
     "102 Hawthorn Avenue", "Russian", "Email"),
    ("Mbeki", "Thandi Mbeki", "Mother", "Sipho Mbeki", "Father",
     "45 Juniper Place", "English", "App message"),
    ("Clarke", "Hannah Clarke", "Mother", "Dan Clarke", "Father",
     "28 Rowan Gardens", "English", "Text message"),
    ("Nguyen", "Linh Nguyen", "Mother", "Minh Nguyen", "Father",
     "5 Alder Terrace", "Vietnamese", "App message"),
    ("O'Sullivan", "Aoife O'Sullivan", "Mother", "", "",
     "77 Primrose Hill", "English", "Phone call"),
    ("Ferreira", "Bianca Ferreira", "Mother", "Paulo Ferreira",
     "Father", "16 Magnolia Street", "Portuguese", "WhatsApp"),
    ("Lindqvist", "Anna Lindqvist", "Mother", "Erik Lindqvist",
     "Father", "90 Birchwood Lane", "English", "Email"),
    ("Grant", "Sara Grant", "Mother", "Michael Grant", "Father",
     "23 Foxglove Road", "English", "App message"),
    ("Bello", "Funmi Bello", "Mother", "Tunde Bello", "Father",
     "4 Bramble Court", "English", "Text message"),
    ("Kowalczyk", "Marta Kowalczyk", "Grandmother", "", "",
     "58 Cedar Walk", "Polish", "Phone call"),
    ("Reyes", "Camila Reyes", "Mother", "Diego Reyes", "Father",
     "131 Poplar Street", "Spanish", "App message"),
    ("Thornton", "Beth Thornton", "Mother", "Craig Thornton",
     "Father", "11 Heather Rise", "English", "Email"),
    ("Achebe", "Ngozi Achebe", "Mother", "", "",
     "66 Willow Grove", "English", "In person at pickup"),
]

_EMERGENCY = [
    ("Grace Mensah", "Grandmother"), ("Peter Hartley", "Grandfather"),
    ("Danuta Nowak", "Grandmother"), ("Luisa Delgado", "Aunt"),
    ("Mary Brennan", "Grandmother"), ("Sofia Ahmed", "Aunt"),
    ("Pauline Reed", "Family friend"), ("Irina Petrova", "Grandmother"),
    ("Nomsa Mbeki", "Aunt"), ("Jean Clarke", "Grandmother"),
    ("Hoa Nguyen", "Grandmother"), ("Dervla Cassidy", "Family friend"),
    ("Rita Ferreira", "Grandmother"), ("Lars Lindqvist", "Grandfather"),
    ("Pam Grant", "Grandmother"), ("Kehinde Bello", "Aunt"),
    ("Jerzy Kowalczyk", "Grandfather"), ("Ana Reyes", "Grandmother"),
    ("Lorna Thornton", "Grandmother"), ("Chika Achebe", "Aunt"),
]

_EMPLOYERS = [
    "Nationwide Children's Hospital", "Ohio State University",
    "Columbus City Schools", "Huntington Bank", "Self employed",
    "Kroger distribution", "COTA", "Grange Insurance",
    "Mount Carmel Health", "Honda of America", "Cardinal Health",
    "Freelance designer", "The Columbus Dispatch", "L Brands",
    "Franklin County", "Safelite", "Abbott Nutrition",
    "OhioHealth", "Bath & Body Works", "Donatos",
]

_DOCTORS = [
    ("Dr Imani Walker, Grove Family Practice", "(614) 555 0310"),
    ("Dr Peter Lang, Willow Road Surgery", "(614) 555 0344"),
    ("Dr Nina Castellano, Mulberry Clinic", "(614) 555 0377"),
    ("Dr Omar Haddad, Eastside Paediatrics", "(614) 555 0392"),
]


# ---------------------------------------------------------------------------
# The children.  `fam` is a 1 based index into FAMILIES, `months` is
# their age today, `days` the weekdays they attend (0 = Monday).
# ---------------------------------------------------------------------------

CHILDREN: List[Dict[str, Any]] = [
    # ---- Acorns, the infant room ------------------------------------
    dict(name="Obiageli Okonkwo", preferred="Obi", fam=1, months=9,
         gender="Girl", days=(0, 1, 2, 3, 4), schedule="Full-time 5 days",
         hours=10.0, basis="Weekly rate", status="Active", room="Acorns",
         program="Infant", start=-210, diet="No restrictions"),
    dict(name="Felix Hartley", preferred="Felix", fam=2, months=6,
         gender="Boy", days=(1, 2, 3), schedule="3 days", hours=9.5,
         basis="Weekly rate", status="Active", room="Acorns",
         program="Infant", start=-95, diet="Dairy free"),
    dict(name="Zofia Nowak", preferred="Zosia", fam=3, months=11,
         gender="Girl", days=(0, 1, 2, 3), schedule="4 days", hours=9.0,
         basis="Weekly rate", status="Active", room="Acorns",
         program="Infant", start=-150, diet="No restrictions"),
    dict(name="Mateo Delgado", preferred="Mateo", fam=4, months=13,
         gender="Boy", days=(0, 2, 4), schedule="3 days", hours=8.5,
         basis="Daily rate", status="Active", room="Acorns",
         program="Infant", start=-120, diet="No restrictions"),
    dict(name="Saoirse Brennan", preferred="Saoirse", fam=5, months=8,
         gender="Girl", days=(0, 1, 2, 3, 4), schedule="Full-time 5 days",
         hours=9.5, basis="Weekly rate", status="Active", room="Acorns",
         program="Infant", start=-70, diet="No restrictions"),
    dict(name="Bilal Ahmed", preferred="Bilal", fam=6, months=10,
         gender="Boy", days=(1, 3), schedule="2 days", hours=9.0,
         basis="Weekly rate", status="Active", room="Acorns",
         program="Infant", start=-55, diet="Halal"),
    # ---- Saplings, young toddlers (the busy room) --------------------
    dict(name="Ivy Whitfield", preferred="Ivy", fam=7, months=18,
         gender="Girl", days=(0, 1, 2, 3, 4), schedule="Full-time 5 days",
         hours=10.0, basis="Weekly rate", status="Active",
         room="Saplings", program="Young toddler", start=-320,
         diet="No restrictions"),
    dict(name="Nikolai Petrov", preferred="Niko", fam=8, months=20,
         gender="Boy", days=(0, 1, 2, 3), schedule="4 days", hours=9.0,
         basis="Weekly rate", status="Active", room="Saplings",
         program="Young toddler", start=-280, diet="No restrictions"),
    dict(name="Ayanda Mbeki", preferred="Aya", fam=9, months=15,
         gender="Girl", days=(1, 2, 3, 4), schedule="4 days", hours=9.5,
         basis="Weekly rate", status="Active", room="Saplings",
         program="Young toddler", start=-175, diet="Vegetarian"),
    dict(name="Rory Clarke", preferred="Rory", fam=10, months=22,
         gender="Boy", days=(0, 1, 2, 3, 4), schedule="Full-time 5 days",
         hours=10.0, basis="Monthly rate", status="Active",
         room="Saplings", program="Young toddler", start=-400,
         diet="Nut free"),
    dict(name="An Nguyen", preferred="An", fam=11, months=17,
         gender="Girl", days=(0, 2, 4), schedule="3 days", hours=9.0,
         basis="Weekly rate", status="Active", room="Saplings",
         program="Young toddler", start=-140, diet="No restrictions"),
    dict(name="Cillian O'Sullivan", preferred="Cillian", fam=12,
         months=23, gender="Boy", days=(0, 1, 2, 3, 4),
         schedule="Full-time 5 days", hours=9.5, basis="Weekly rate",
         status="Active", room="Saplings", program="Young toddler",
         start=-365, diet="Gluten free"),
    dict(name="Luca Ferreira", preferred="Luca", fam=13, months=19,
         gender="Boy", days=(1, 2, 3), schedule="3 days", hours=8.0,
         basis="Weekly rate", status="Active", room="Saplings",
         program="Young toddler", start=-110, diet="No restrictions"),
    dict(name="Elsa Lindqvist", preferred="Elsa", fam=14, months=16,
         gender="Girl", days=(0, 1, 3, 4), schedule="4 days", hours=9.0,
         basis="Weekly rate", status="Active", room="Saplings",
         program="Young toddler", start=-200, diet="No restrictions"),
    dict(name="Noah Grant", preferred="Noah", fam=15, months=21,
         gender="Boy", days=(0, 1, 2, 3, 4), schedule="Full-time 5 days",
         hours=10.0, basis="Weekly rate", status="Active",
         room="Saplings", program="Young toddler", start=-240,
         diet="No restrictions"),
    # ---- Oaks, toddlers ----------------------------------------------
    dict(name="Tobi Bello", preferred="Tobi", fam=16, months=28,
         gender="Boy", days=(0, 1, 2, 3, 4), schedule="Full-time 5 days",
         hours=9.5, basis="Weekly rate", status="Active", room="Oaks",
         program="Toddler", start=-430, diet="No restrictions"),
    dict(name="Amaka Okonkwo", preferred="Amaka", fam=1, months=31,
         gender="Girl", days=(0, 1, 2, 3, 4), schedule="Full-time 5 days",
         hours=10.0, basis="Weekly rate", status="Active", room="Oaks",
         program="Toddler", start=-560, diet="No restrictions"),
    dict(name="Jakub Kowalczyk", preferred="Kuba", fam=17, months=26,
         gender="Boy", days=(1, 2, 3, 4), schedule="4 days", hours=9.0,
         basis="Weekly rate", status="Active", room="Oaks",
         program="Toddler", start=-190, diet="No restrictions"),
    dict(name="Valentina Reyes", preferred="Valen", fam=18, months=33,
         gender="Girl", days=(0, 1, 2, 3, 4), schedule="Full-time 5 days",
         hours=9.0, basis="Weekly rate", status="Active", room="Oaks",
         program="Toddler", start=-300, diet="No restrictions"),
    dict(name="Harvey Thornton", preferred="Harvey", fam=19, months=29,
         gender="Boy", days=(0, 2, 4), schedule="3 days", hours=8.5,
         basis="Weekly rate", status="Active", room="Oaks",
         program="Toddler", start=-160, diet="Egg free"),
    dict(name="Chiamaka Achebe", preferred="Chi", fam=20, months=34,
         gender="Girl", days=(0, 1, 2, 3), schedule="4 days", hours=9.5,
         basis="Weekly rate", status="Active", room="Oaks",
         program="Toddler", start=-250, diet="No restrictions"),
    dict(name="Seren Hartley", preferred="Seren", fam=2, months=27,
         gender="Girl", days=(1, 2, 3), schedule="3 days", hours=9.5,
         basis="Weekly rate", status="Active", room="Oaks",
         program="Toddler", start=-320, diet="Dairy free"),
    dict(name="Oscar Whitfield", preferred="Oscar", fam=7, months=30,
         gender="Boy", days=(0, 1, 2, 3, 4), schedule="Full-time 5 days",
         hours=10.0, basis="Weekly rate", status="Active", room="Oaks",
         program="Toddler", start=-520, diet="No restrictions"),
    # ---- Willows, preschool and school age ----------------------------
    dict(name="Isla Brennan", preferred="Isla", fam=5, months=44,
         gender="Girl", days=(0, 1, 2, 3, 4), schedule="Full-time 5 days",
         hours=9.0, basis="Weekly rate", status="Active",
         room="Willows", program="Preschool", start=-700,
         diet="No restrictions"),
    dict(name="Dylan Grant", preferred="Dylan", fam=15, months=41,
         gender="Boy", days=(0, 1, 2, 3), schedule="4 days", hours=8.5,
         basis="Weekly rate", status="Active", room="Willows",
         program="Preschool", start=-480, diet="No restrictions"),
    dict(name="Priya Nguyen", preferred="Priya", fam=11, months=38,
         gender="Girl", days=(1, 2, 3, 4), schedule="4 days", hours=9.0,
         basis="Weekly rate", status="Active", room="Willows",
         program="Preschool", start=-330, diet="Vegetarian"),
    dict(name="Emilia Petrova", preferred="Emilia", fam=8, months=52,
         gender="Girl", days=(0, 1, 2, 3, 4), schedule="Half day 5 days",
         hours=5.0, basis="Hourly rate", status="Active",
         room="Willows", program="Pre-K", start=-600,
         diet="No restrictions"),
    dict(name="Theo Delgado", preferred="Theo", fam=4, months=47,
         gender="Boy", days=(0, 2, 4), schedule="3 days", hours=9.0,
         basis="Weekly rate", status="Active", room="Willows",
         program="Preschool", start=-410, diet="No restrictions"),
    dict(name="Jonah Okonkwo", preferred="Jonah", fam=1, months=68,
         gender="Boy", days=(0, 1, 2, 3, 4), schedule="Before & after",
         hours=3.5, basis="Per session", status="Active",
         room="Willows", program="Before & after", start=-430,
         diet="No restrictions"),
    dict(name="Maya Reyes", preferred="Maya", fam=18, months=57,
         gender="Girl", days=(0, 1, 2, 3, 4), schedule="Full-time 5 days",
         hours=9.5, basis="Weekly rate", status="Active",
         room="Willows", program="Pre-K", start=-640,
         diet="No restrictions"),
    # ---- not currently attending --------------------------------------
    dict(name="Arthur Lindqvist", preferred="Artie", fam=14, months=63,
         gender="Boy", days=(), schedule="Full-time 5 days", hours=9.0,
         basis="Weekly rate", status="Graduated", room="Willows",
         program="Pre-K", start=-900, end=-45,
         leaving="Aged out / started school", diet="No restrictions"),
    dict(name="Kofi Mbeki", preferred="Kofi", fam=9, months=36,
         gender="Boy", days=(), schedule="3 days", hours=9.0,
         basis="Weekly rate", status="Withdrawn", room="Oaks",
         program="Toddler", start=-380, end=-70, leaving="Moved away",
         diet="No restrictions"),
    dict(name="Esme Thornton", preferred="Esme", fam=19, months=7,
         gender="Girl", days=(), schedule="4 days", hours=9.0,
         basis="Weekly rate", status="Pending start", room="Acorns",
         program="Infant", start=21, diet="No restrictions"),
    dict(name="Rafael Ferreira", preferred="Rafa", fam=13, months=5,
         gender="Boy", days=(), schedule="3 days", hours=9.0,
         basis="Weekly rate", status="On hold", room="Acorns",
         program="Infant", start=-30, diet="No restrictions"),
]

#: Weekly full-time rate by room, before any pro-rata or discount.
_ROOM_RATE = {"Acorns": 325.0, "Saplings": 305.0, "Oaks": 285.0,
              "Willows": 265.0}


# ---------------------------------------------------------------------------
# The team.  Between them they use every pay type the workbook knows.
# ---------------------------------------------------------------------------

STAFF: List[Dict[str, Any]] = [
    dict(name="Marianne Okafor", role="Owner / Director", room="",
         pay_type="Salary", salary=64000, hourly=0, hours=40,
         status="Active", hire=-3100, qual="EdD Early Childhood",
         shift="Admin"),
    dict(name="Dawn Pritchard", role="Centre manager", room="",
         pay_type="Salary", salary=51000, hourly=0, hours=40,
         status="Active", hire=-1850, qual="BA Early Years, Level 5",
         shift="Full day"),
    dict(name="Grace Lim", role="Room leader", room="Acorns",
         pay_type="Hourly", salary=0, hourly=21.5, hours=38,
         status="Active", hire=-1420, qual="Level 3 Early Years",
         shift="Opening"),
    dict(name="Tasha Boyd", role="Room leader", room="Saplings",
         pay_type="Hourly", salary=0, hourly=20.75, hours=38,
         status="Active", hire=-1090, qual="Level 3 Early Years",
         shift="Early"),
    dict(name="Priyanka Rao", role="Room leader", room="Oaks",
         pay_type="Hourly", salary=0, hourly=20.5, hours=36,
         status="Active", hire=-880, qual="Level 3 Early Years",
         shift="Morning"),
    dict(name="Beth Carmody", role="Lead teacher", room="Willows",
         pay_type="Hourly", salary=0, hourly=22.0, hours=38,
         status="Active", hire=-1660, qual="BSc Child Development",
         shift="Full day"),
    dict(name="Hollie Sanderson", role="Assistant teacher",
         room="Acorns", pay_type="Hourly", salary=0, hourly=17.25,
         hours=34, status="Active", hire=-640,
         qual="Level 2 Early Years", shift="Mid"),
    dict(name="Danny Echeverria", role="Assistant teacher",
         room="Saplings", pay_type="Hourly", salary=0, hourly=17.0,
         hours=32, status="Active", hire=-420,
         qual="Level 2 Early Years", shift="Afternoon"),
    dict(name="Keisha Monroe", role="Assistant teacher", room="Oaks",
         pay_type="Hourly", salary=0, hourly=16.75, hours=30,
         status="Probation", hire=-70, qual="Level 2 in progress",
         shift="Late"),
    dict(name="Aiden Walsh", role="Apprentice / trainee",
         room="Willows", pay_type="Hourly", salary=0, hourly=13.5,
         hours=30, status="Active", hire=-240,
         qual="Apprenticeship, year 1", shift="Closing"),
    dict(name="Rosemary Hicks", role="Cook", room="",
         pay_type="Hourly", salary=0, hourly=18.0, hours=25,
         status="Active", hire=-1300, qual="Level 2 Food Safety",
         shift="Morning"),
    dict(name="Jess Turnbull", role="Substitute / supply", room="",
         pay_type="Daily rate", salary=0, hourly=0, day_rate=132,
         hours=0, status="Casual / bank", hire=-500,
         qual="Level 3 Early Years", shift="Sub cover"),
]

_STAFF_ROOM_ORDER = ["Acorns", "Saplings", "Oaks", "Willows"]


# ---------------------------------------------------------------------------
# Supplies
# ---------------------------------------------------------------------------

SUPPLIES: List[Tuple[Any, ...]] = [
    # name, category, unit, pack, supplier, lead, opening, recv, used,
    # waste, cost, min
    ("Nappies size 3", "Nappies & wipes", "Box of 150", 150,
     "Columbus Care Supply", 3, 9, 6, 11, 0, 38.90, 3),
    ("Nappies size 5", "Nappies & wipes", "Box of 120", 120,
     "Columbus Care Supply", 3, 6, 4, 7, 0, 41.50, 2),
    ("Baby wipes", "Nappies & wipes", "Case of 12", 12,
     "Columbus Care Supply", 3, 14, 8, 17, 1, 22.40, 5),
    ("Nappy sacks", "Nappies & wipes", "Roll pack", 200,
     "Columbus Care Supply", 3, 11, 0, 5, 0, 6.80, 4),
    ("Barrier cream", "Hygiene & bathroom", "Tub 500g", 1,
     "Ohio Nursery Direct", 5, 7, 2, 3, 0, 9.25, 3),
    ("Hand soap refill", "Hygiene & bathroom", "5 litre", 1,
     "Buckeye Janitorial", 4, 5, 3, 6, 0, 17.60, 2),
    ("Paper towels", "Hygiene & bathroom", "Case of 24", 24,
     "Buckeye Janitorial", 4, 8, 6, 12, 0, 29.95, 3),
    ("Toilet rolls", "Hygiene & bathroom", "Pack of 36", 36,
     "Buckeye Janitorial", 4, 6, 4, 8, 0, 24.50, 3),
    ("Disinfectant spray", "Cleaning & sanitising", "750ml", 1,
     "Buckeye Janitorial", 4, 16, 6, 14, 1, 4.95, 8),
    ("Surface sanitiser concentrate", "Cleaning & sanitising",
     "5 litre", 1, "Buckeye Janitorial", 4, 3, 2, 3, 0, 31.00, 2),
    ("Mop heads", "Cleaning & sanitising", "Each", 1,
     "Buckeye Janitorial", 7, 6, 0, 2, 0, 7.20, 3),
    ("Bin liners", "Cleaning & sanitising", "Roll of 50", 50,
     "Buckeye Janitorial", 4, 12, 4, 10, 0, 11.30, 5),
    ("Infant formula stage 1", "Drinks & formula", "800g tin", 1,
     "Ohio Nursery Direct", 5, 5, 4, 7, 0, 26.75, 3),
    ("Whole milk", "Drinks & formula", "Gallon", 1, "Kroger wholesale",
     1, 10, 32, 38, 2, 3.85, 8),
    ("Oat milk", "Drinks & formula", "Litre", 1, "Kroger wholesale",
     1, 6, 8, 11, 0, 2.95, 4),
    ("Porridge oats", "Food & kitchen", "3kg bag", 1,
     "Kroger wholesale", 2, 4, 3, 5, 0, 7.40, 2),
    ("Wholemeal bread", "Food & kitchen", "Loaf", 1,
     "Kroger wholesale", 1, 8, 26, 31, 3, 2.60, 6),
    ("Pasta", "Food & kitchen", "5kg", 1, "Kroger wholesale", 2, 3,
     2, 4, 0, 9.80, 2),
    ("Frozen mixed vegetables", "Food & kitchen", "2.5kg", 1,
     "Kroger wholesale", 2, 7, 6, 10, 0, 8.15, 4),
    ("Gluten free pasta", "Food & kitchen", "500g", 1,
     "Kroger wholesale", 2, 4, 2, 3, 0, 3.90, 3),
    ("Paint, poster, assorted", "Arts & crafts", "Pack of 6", 6,
     "Ohio Nursery Direct", 6, 5, 2, 4, 0, 18.40, 2),
    ("Sugar paper A2", "Arts & crafts", "Ream of 250", 250,
     "Ohio Nursery Direct", 6, 4, 1, 2, 0, 21.00, 2),
    ("Glue sticks", "Arts & crafts", "Pack of 12", 12,
     "Ohio Nursery Direct", 6, 7, 2, 5, 1, 9.60, 3),
    ("Playdough", "Arts & crafts", "Tub", 1, "Ohio Nursery Direct",
     6, 12, 0, 6, 2, 2.20, 6),
    ("First aid refill kit", "First aid", "Kit", 1,
     "Midwest Safety Co", 7, 3, 1, 1, 0, 34.00, 2),
    ("Disposable gloves", "Safety & PPE", "Box of 200", 200,
     "Midwest Safety Co", 5, 9, 6, 13, 0, 12.75, 4),
    ("Cot sheets", "Linen & bedding", "Each", 1, "Ohio Nursery Direct",
     10, 24, 0, 0, 2, 6.40, 12),
    ("Printer paper A4", "Office & stationery", "Ream", 1,
     "Columbus Office Hub", 3, 6, 2, 4, 0, 5.90, 3),
]


# ---------------------------------------------------------------------------
# The builder
# ---------------------------------------------------------------------------

def _build(cfg: Config) -> Sample:
    rng = random.Random(SEED)
    today = cfg.today
    s = Sample()

    # Monday of the current week - every weekly thing hangs off this.
    monday = today - _dt.timedelta(days=today.weekday())
    history_start = monday - _dt.timedelta(weeks=10)
    all_days = _weekdays(history_start, today)

    months = [_month_start(today, back) for back in (3, 2, 1, 0)]

    # ------------------------------------------------------- classrooms --
    staff_id = {i + 1: rid("Staff", i + 1) for i in range(len(STAFF))}
    for room in ROOMS:
        s.classrooms.append(dict(
            name=room["name"], age_band=room["age_band"],
            capacity=room["capacity"], ratio_set=room["ratio_set"],
            teacher=staff_id[room["teacher"]],
            assistant=staff_id[room["assistant"]],
            room_status=room["room_status"], notes=room["notes"]))

    # ---------------------------------------------------------- families --
    for index, fam in enumerate(FAMILIES, start=1):
        (surname, p1, rel1, p2, rel2, street, language, contact) = fam
        em_name, em_rel = _EMERGENCY[index - 1]
        doctor, doc_phone = _DOCTORS[index % len(_DOCTORS)]
        pickup = p1 + (f", {p2}" if p2 else "") + f", {em_name}"
        s.parents.append(dict(
            family_name=f"The {surname} family",
            parent1=p1, relation1=rel1,
            phone1=f"(614) 555 {1000 + index * 7:04d}",
            email1=f"{p1.split()[0].lower()}.{surname.lower()}"
                   f"@example.com".replace("'", ""),
            employer1=_EMPLOYERS[index - 1],
            work_phone1=f"(614) 555 {2000 + index * 3:04d}",
            parent2=p2, relation2=rel2,
            phone2=(f"(614) 555 {3000 + index * 5:04d}" if p2 else ""),
            email2=(f"{p2.split()[0].lower()}.{surname.lower()}"
                    f"@example.com".replace("'", "") if p2 else ""),
            address=street, city="Columbus",
            postcode=f"432{index % 10}{(index * 3) % 10}",
            contact_by=contact, language=language,
            em1_name=em_name,
            em1_phone=f"(614) 555 {4000 + index * 11:04d}",
            em1_relation=em_rel,
            em2_name=(_EMERGENCY[(index + 4) % len(_EMERGENCY)][0]
                      if index % 3 else ""),
            em2_phone=(f"(614) 555 {5000 + index * 13:04d}"
                       if index % 3 else ""),
            doctor=doctor, doctor_phone=doc_phone,
            pickup_list=pickup,
            pickup_ok="Yes" if index != 17 else "No",
            custody_notes=("Father has alternate weekends. Court order "
                           "on file - collection by mother only on "
                           "weekdays."
                           if index == 12 else ""),
            joined=_months_ago(today, 24 - (index % 20)),
            notes=("Grandmother is the main contact and does all "
                   "pickups." if index == 17 else
                   "Second child starting in the new year."
                   if index == 2 else "")))

    family_id = {i: rid("Family", i) for i in range(1, len(FAMILIES) + 1)}

    # ---------------------------------------------------------- children --
    seen_family: Dict[int, int] = {}
    kids: List[Dict[str, Any]] = []
    for index, kid in enumerate(CHILDREN, start=1):
        fam = kid["fam"]
        seen_family[fam] = seen_family.get(fam, 0) + 1
        seq = seen_family[fam]
        dob = _months_ago(today, kid["months"]) - _dt.timedelta(
            days=(index * 3) % 25)
        start = d(today, kid["start"])
        enrolled = start - _dt.timedelta(days=21 + (index % 30))
        days = kid["days"]
        per_week = len(days) if days else {
            "Full-time 5 days": 5, "4 days": 4, "3 days": 3,
            "2 days": 2, "Half day 5 days": 5,
            "Before & after": 5}.get(kid["schedule"], 3)
        base = _ROOM_RATE.get(kid["room"], 265.0)
        weekly = base * per_week / 5.0
        if kid["basis"] == "Weekly rate":
            rate = round(weekly / 5.0) * 5.0
        elif kid["basis"] == "Monthly rate":
            rate = round(weekly * 52 / 12 / 5) * 5.0
        elif kid["basis"] == "Daily rate":
            rate = round(base / 5.0 / 5) * 5.0
        elif kid["basis"] == "Hourly rate":
            rate = 9.50
        elif kid["basis"] == "Per session":
            rate = 14.00
        else:
            rate = round(weekly / 5.0) * 5.0
        usual_in = t(7, 30) if per_week >= 4 else t(8, 15)
        usual_out = t(17, 30) if kid["hours"] >= 9 else t(13, 0)
        if kid["schedule"] == "Before & after":
            usual_in, usual_out = t(7, 15), t(17, 45)
        parent = FAMILIES[fam - 1][1]
        partner = FAMILIES[fam - 1][3]
        pickup = parent + (f", {partner}" if partner else "") + \
            f", {_EMERGENCY[fam - 1][0]}"
        record = dict(
            name=kid["name"], preferred=kid["preferred"], dob=dob,
            gender=kid["gender"], family=family_id[fam],
            program=kid["program"], room=kid["room"],
            schedule=kid["schedule"], days_week=float(per_week),
            hours_day=kid["hours"], usual_in=usual_in,
            usual_out=usual_out, enrolled_date=enrolled,
            start_date=start,
            end_date=(d(today, kid["end"]) if kid.get("end") else None),
            leaving_reason=kid.get("leaving", ""),
            status=kid["status"], basis=kid["basis"], rate=money(rate),
            discount_pct=(0.1 if index == 20 else 0),
            subsidy_week=(95.0 if index in (3, 17) else
                          120.0 if index == 11 else 0),
            diet=kid["diet"],
            medical_notes="", pickup_list=pickup,
            pickup_password=f"{kid['preferred'].lower()}"
                            f"{(index * 7) % 90 + 10}",
            photo_ok="Yes" if index not in (12, 26) else "No",
            sunscreen_ok="Yes" if index != 19 else "No",
            outings_ok="Yes" if index not in (4, 26) else "No",
            contract="Signed" if index not in (31, 32) else "Requested",
            notes="")
        # a few care notes, so the register is not uniformly bland
        if index == 2:
            record["medical_notes"] = ("Dairy free - soya formula "
                                       "supplied by parent, labelled.")
        if index == 10:
            record["medical_notes"] = ("Severe nut allergy. Adrenaline "
                                       "pen in the red box by the "
                                       "Saplings door. Allergy plan on "
                                       "file.")
        if index == 12:
            record["medical_notes"] = "Coeliac - strictly gluten free."
        if index == 22:
            record["medical_notes"] = ("Mild asthma. Blue inhaler and "
                                       "spacer in room bag.")
        kids.append(record)
        record["_seq"] = seq
        record["_fam"] = fam
        record["_days"] = days
        record["_weekly"] = money(weekly)
    for item in kids:
        s.children.append({k: v for k, v in item.items()
                           if not k.startswith("_")})

    child_id = {i + 1: rid("Child", i + 1) for i in range(len(kids))}
    for index, item in enumerate(kids, start=1):
        item["_id"] = child_id[index]

    active = [k for k in kids if k["status"] == "Active"]
    by_id = {k["_id"]: k for k in kids}

    # ---------------------------------------------------------- waitlist --
    wl_names = [
        ("Nora Bianchi", "Giulia Bianchi", "Saplings", "Young toddler"),
        ("Caleb Osei", "Abena Osei", "Acorns", "Infant"),
        ("Robin Fraser", "Mhairi Fraser", "Saplings", "Young toddler"),
        ("Amira Haddad", "Layla Haddad", "Acorns", "Infant"),
        ("Jude Whitfield", "Jo Whitfield", "Acorns", "Infant"),
        ("Sofia Marchetti", "Elisa Marchetti", "Oaks", "Toddler"),
        ("Leo Kaminski", "Ola Kaminski", "Saplings", "Young toddler"),
        ("Zara Patel", "Nisha Patel", "Willows", "Preschool"),
        ("Henry Boateng", "Akosua Boateng", "Acorns", "Infant"),
        ("Mila Novak", "Petra Novak", "Saplings", "Young toddler"),
        ("Ethan Grant", "Sara Grant", "Acorns", "Infant"),
        ("Aarav Shah", "Devi Shah", "Oaks", "Toddler"),
        ("Freya Donnelly", "Roisin Donnelly", "Saplings",
         "Young toddler"),
        ("Malik Johnson", "Tanya Johnson", "Willows", "Preschool"),
        ("Clara Weiss", "Ingrid Weiss", "Oaks", "Toddler"),
        ("Rosa Delgado jr", "Rosa Delgado", "Acorns", "Infant"),
    ]
    wl_status = ["Enrolled", "Offer accepted", "Place offered", "Toured",
                 "Tour booked", "Information sent", "New enquiry",
                 "Waiting - no space", "Waiting - future date",
                 "Offer declined", "Lost"]
    sources = ["Parent referral", "Google search", "Sibling already here",
               "Word of mouth", "Facebook", "Sign outside",
               "Google Maps", "Walk in", "Employer scheme",
               "Returning family", "Local directory"]
    for index, (child_name, parent, room, program) in enumerate(
            wl_names, start=1):
        enquiry = d(today, -(130 - index * 7))
        status = wl_status[index % len(wl_status)]
        priority = ("Sibling" if "Whitfield" in child_name
                    or "Grant" in child_name or "Delgado" in child_name
                    else "Staff child" if index == 6
                    else "Returning family" if index == 10
                    else "Low" if index == 14 else "Standard")
        sib = ""
        if priority == "Sibling":
            for item in kids:
                if item["name"].split()[-1] == child_name.split()[-1]:
                    sib = item["_id"]
                    break
        toured = status in ("Toured", "Place offered", "Offer accepted",
                            "Enrolled", "Offer declined")
        offered = status in ("Place offered", "Offer accepted",
                             "Enrolled", "Offer declined")
        s.waitlist.append(dict(
            enquiry_date=enquiry, child_name=child_name,
            dob=_months_ago(today, 4 + (index * 5) % 40),
            parent_name=parent,
            phone=f"(614) 555 {6000 + index * 9:04d}",
            email=f"{parent.split()[0].lower()}@example.com",
            source=sources[index % len(sources)], program=program,
            room=room, days_wanted=float(3 + index % 3),
            start_wanted=d(today, -20 + index * 11),
            priority=priority, sibling=sib, status=status,
            info_sent=d(enquiry, 1),
            tour_date=(d(enquiry, 9) if toured else None),
            offered_date=(d(enquiry, 20) if offered else None),
            response_date=(d(enquiry, 24)
                           if status in ("Offer accepted", "Enrolled",
                                         "Offer declined") else None),
            enrolled_child=("" if status != "Enrolled"
                            else kids[index % len(active)]["_id"]),
            followed_up=d(enquiry, 30 + index),
            notes=("Wants the same days as her brother."
                   if priority == "Sibling" else
                   "Needs a September start, happy to wait."
                   if status == "Waiting - future date" else
                   "Chose a setting closer to work."
                   if status in ("Lost", "Offer declined") else "")))

    # -------------------------------------------------------- attendance --
    floor_staff = [staff_id[i] for i in (3, 4, 5, 6, 7, 8, 9, 10)]
    closed_day = None
    for day in reversed(all_days):
        if day.weekday() == 0 and (today - day).days > 20:
            closed_day = day
            break

    reasons_sick = ["Temperature overnight", "Chickenpox",
                    "Sickness bug - 48 hour rule", "Heavy cold",
                    "Ear infection", "Hand, foot and mouth"]
    reasons_family = ["Family holiday", "Grandparents visiting",
                      "Wedding", "Moving house", "Parent day off"]

    for day in all_days:
        if closed_day and day == closed_day:
            for kid in active:
                s.attendance.append(dict(
                    date=day, child=kid["_id"], status="Centre closed",
                    in_time=None, out_time=None, fee_raised="No",
                    dropped_by="", collected_by="",
                    reason="Staff training day - notified in August",
                    recorded_by=staff_id[2], temp_check="",
                    notes=""))
            continue
        for kid in active:
            if day.weekday() not in kid["_days"]:
                continue
            if day < kid["start_date"]:
                continue
            roll = rng.random()
            status = "Present"
            if roll > 0.975:
                status = "Absent - sick"
            elif roll > 0.955:
                status = "Absent - family"
            elif roll > 0.948:
                status = "Absent - unexplained"
            elif roll > 0.935:
                status = "Holiday - child"
            elif roll > 0.885:
                status = "Late"
            elif roll > 0.862:
                status = "Early pickup"
            present = status in ("Present", "Late", "Early pickup")
            in_t = out_t = None
            dropped = collected = ""
            parent = FAMILIES[kid["_fam"] - 1][1]
            partner = FAMILIES[kid["_fam"] - 1][3]
            if present:
                base_in = kid["usual_in"]
                jitter = rng.choice([-10, -5, 0, 0, 5, 8])
                if status == "Late":
                    jitter = rng.choice([35, 42, 55, 70])
                minutes = base_in.hour * 60 + base_in.minute + jitter
                in_t = t(minutes // 60, minutes % 60)
                base_out = kid["usual_out"]
                jout = rng.choice([-8, -3, 0, 0, 4, 9, 22])
                if status == "Early pickup":
                    jout = -rng.choice([75, 95, 120])
                mo = base_out.hour * 60 + base_out.minute + jout
                out_t = t(mo // 60, mo % 60)
                dropped = parent if rng.random() < 0.7 else (
                    partner or _EMERGENCY[kid["_fam"] - 1][0])
                collected = partner or parent
                if rng.random() < 0.12:
                    collected = _EMERGENCY[kid["_fam"] - 1][0]
            reason = ""
            if status == "Absent - sick":
                reason = rng.choice(reasons_sick)
            elif status == "Absent - family":
                reason = rng.choice(reasons_family)
            elif status == "Holiday - child":
                reason = "Booked holiday - two weeks' notice given"
            s.attendance.append(dict(
                date=day, child=kid["_id"], status=status,
                in_time=in_t, out_time=out_t,
                fee_raised="No" if status == "Holiday - child" else "Yes",
                dropped_by=dropped, collected_by=collected,
                reason=reason, recorded_by=rng.choice(floor_staff),
                temp_check=("36.6" if present and rng.random() < 0.4
                            else ""),
                notes=""))

    # a handful of deliberate loose ends on the most recent days
    recent = [r for r in s.attendance
              if r["date"] == all_days[-1] and r["status"] == "Present"]
    for row in recent[:2]:
        row["out_time"] = None
        row["notes"] = "Still here at close - check the book."
    if len(recent) > 4:
        recent[4]["collected_by"] = "Neighbour - Mrs Dunn"
        recent[4]["notes"] = ("Mum phoned at 2pm to authorise. Added to "
                              "the list the next morning.")

    # -------------------------------------------------- charges & ledger --
    # Work out every charge first, group it by family and month, then
    # raise one invoice per family per month and pay most of them.
    charges: List[Dict[str, Any]] = []

    def weeks_in(first: _dt.date, last: _dt.date) -> float:
        return round(len(_weekdays(first, last)) / 5.0, 2)

    for first in months:
        last = min(_month_end(first), today)
        label = month_name(first)
        span = weeks_in(first, last)
        for kid in active:
            if kid["start_date"] > last:
                continue
            basis = kid["basis"]
            if basis == "Monthly rate":
                qty, rate = 1.0, kid["rate"]
            elif basis == "Daily rate":
                qty = float(len([x for x in _weekdays(first, last)
                                 if x.weekday() in kid["_days"]]))
                rate = kid["rate"]
            elif basis == "Hourly rate":
                qty = round(len([x for x in _weekdays(first, last)
                                 if x.weekday() in kid["_days"]])
                            * kid["hours_day"], 1)
                rate = kid["rate"]
            elif basis == "Per session":
                qty = float(len([x for x in _weekdays(first, last)
                                 if x.weekday() in kid["_days"]]) * 2)
                rate = kid["rate"]
            else:
                qty, rate = span, kid["rate"]
            charges.append(dict(
                date=first, child=kid["_id"], family_set="",
                fee_type="Tuition",
                description=f"Childcare - {label}",
                period_from=first, period_to=last, qty=qty,
                rate=money(rate), discount_pct=0,
                auto_source="Contracted schedule", notes="",
                _fam=kid["_fam"], _month=first))
            # sibling discount on the second and later child
            if kid["_seq"] > 1:
                charges.append(dict(
                    date=first, child=kid["_id"], family_set="",
                    fee_type="Sibling discount",
                    description=f"Sibling discount - {label}",
                    period_from=first, period_to=last, qty=1.0,
                    rate=money(qty * rate * 0.1), discount_pct=0,
                    auto_source="Contracted schedule", notes="",
                    _fam=kid["_fam"], _month=first))
            if kid["subsidy_week"]:
                charges.append(dict(
                    date=first, child=kid["_id"], family_set="",
                    fee_type="Subsidy credit",
                    description=f"Publicly funded hours - {label}",
                    period_from=first, period_to=last, qty=span,
                    rate=money(kid["subsidy_week"]), discount_pct=0,
                    auto_source="Manual adjustment",
                    notes="Paid direct by the county scheme.",
                    _fam=kid["_fam"], _month=first))

    # extras: late pickups, meals, trips, registration
    late_rows = [r for r in s.attendance
                 if r["out_time"] and r["status"] in
                 ("Present", "Late", "Early pickup")]
    rng.shuffle(late_rows)
    for row in late_rows[:14]:
        kid = by_id[row["child"]]
        first = _month_start(row["date"])
        charges.append(dict(
            date=row["date"], child=kid["_id"], family_set="",
            fee_type="Late pickup fee",
            description="Collected after the contracted time",
            period_from=row["date"], period_to=row["date"], qty=1.0,
            rate=15.0, discount_pct=0, auto_source="Late pickup",
            notes="", _fam=kid["_fam"], _month=first))
    for index, kid in enumerate(active):
        if index % 3:
            continue
        first = months[-2]
        charges.append(dict(
            date=d(first, 9), child=kid["_id"], family_set="",
            fee_type="Field trip",
            description="Farm visit - coach and entry",
            period_from=d(first, 9), period_to=d(first, 9), qty=1.0,
            rate=12.0, discount_pct=0, auto_source="Activity or trip",
            notes="", _fam=kid["_fam"], _month=first))
    for kid in active[:6]:
        first = months[-1]
        charges.append(dict(
            date=d(first, 2), child=kid["_id"], family_set="",
            fee_type="Meal fee",
            description="Hot lunches - extra days",
            period_from=first, period_to=_month_end(first), qty=4.0,
            rate=4.5, discount_pct=0, auto_source="Meals", notes="",
            _fam=kid["_fam"], _month=first))
    for kid in kids:
        if kid["status"] != "Pending start":
            continue
        first = months[-1]
        charges.append(dict(
            date=d(today, -9), child=kid["_id"], family_set="",
            fee_type="Registration fee",
            description="One-off registration",
            period_from=d(today, -9), period_to=d(today, -9), qty=1.0,
            rate=75.0, discount_pct=0, auto_source="Registration",
            notes="Non refundable.", _fam=kid["_fam"], _month=first))

    # group by (family, month) and raise the invoices
    groups: Dict[Tuple[int, _dt.date], List[Dict[str, Any]]] = {}
    for charge in charges:
        groups.setdefault((charge["_fam"], charge["_month"]),
                          []).append(charge)

    def net_of(charge: Dict[str, Any]) -> float:
        gross = charge["qty"] * charge["rate"]
        credit = charge["fee_type"] in (
            "Sibling discount", "Staff discount", "Promotional discount",
            "Subsidy credit", "Scholarship credit", "Goodwill credit")
        return money(-gross if credit else gross)

    ledger: List[Dict[str, Any]] = []
    invoice_for: Dict[Tuple[int, _dt.date], str] = {}

    def add_ledger(row: Dict[str, Any]) -> str:
        ledger.append(row)
        kind = ("Receipt" if row["type"] in ("Payment", "Refund")
                else "Invoice")
        return rid(kind, len(ledger))

    for first in months:
        label = month_name(first)
        for fam in range(1, len(FAMILIES) + 1):
            lines = groups.get((fam, first))
            if not lines:
                continue
            issue = d(first, 0)
            if first == months[-1]:
                issue = min(d(first, 0), today)
            ref = add_ledger(dict(
                date=issue, type="Invoice", family=family_id[fam],
                child="", period=label,
                description=f"Childcare fees - {label}", amount=None,
                against="", method="", reference="",
                terms=cfg.payment_terms_days,
                sent="Yes", notes=""))
            invoice_for[(fam, first)] = ref
        # payments, once the invoices for the month exist
        for fam in range(1, len(FAMILIES) + 1):
            ref = invoice_for.get((fam, first))
            if not ref:
                continue
            total = money(sum(net_of(c) for c in groups[(fam, first)]))
            if total <= 0:
                continue
            is_current = first == months[-1]
            chase = fam in (3, 11, 17)
            if is_current and chase:
                continue                      # the three that are late
            if is_current and fam % 4 == 0:
                part = money(total * 0.5)
                add_ledger(dict(
                    date=d(first, 6), type="Payment",
                    family=family_id[fam], child="", period=label,
                    description="Part payment", amount=part,
                    against=ref, method="Card",
                    reference=f"CRD{fam:03d}{first.month:02d}", terms=0,
                    sent="Yes", notes="Second half agreed for the "
                                      "28th."))
                continue
            method = ("Direct debit" if fam % 3 == 0 else
                      "Bank transfer" if fam % 3 == 1 else "Card")
            if fam == 5:
                method = "Childcare voucher"
            if fam == 9:
                method = "Government subsidy"
            add_ledger(dict(
                date=d(first, 4 + (fam % 9)), type="Payment",
                family=family_id[fam], child="", period=label,
                description=f"Payment received - {label}",
                amount=total, against=ref, method=method,
                reference=f"{method[:3].upper()}{fam:03d}"
                          f"{first.month:02d}",
                terms=0, sent="Yes", notes=""))

    # a deposit, a credit note, a refund and a write-off
    pending = [k for k in kids if k["status"] == "Pending start"]
    if pending:
        add_ledger(dict(
            date=d(today, -8), type="Payment",
            family=pending[0]["family"], child=pending[0]["_id"],
            period="Deposit", description="Enrollment deposit",
            amount=150.0, against="", method="Bank transfer",
            reference="DEP0041", terms=0, sent="Yes",
            notes="Held against the first invoice in January."))
    first_prev = months[-2]
    ref_prev = invoice_for.get((2, first_prev))
    if ref_prev:
        add_ledger(dict(
            date=d(first_prev, 19), type="Credit note",
            family=family_id[2], child="", period=month_name(first_prev),
            description="Two days charged in error",
            amount=118.0, against=ref_prev, method="", reference="",
            terms=0, sent="Yes",
            notes="Seren was away on the agreed closure days."))
    ref_old = invoice_for.get((14, months[0]))
    if ref_old:
        add_ledger(dict(
            date=d(months[1], 3), type="Refund",
            family=family_id[14], child="", period=month_name(months[0]),
            description="Overpayment returned",
            amount=96.0, against=ref_old, method="Bank transfer",
            reference="RFD0007", terms=0, sent="Yes", notes=""))
    ref_gone = invoice_for.get((9, months[0]))
    if ref_gone:
        add_ledger(dict(
            date=d(months[1], 14), type="Write-off",
            family=family_id[9], child="", period=month_name(months[0]),
            description="Family moved away - unrecoverable",
            amount=142.5, against=ref_gone, method="", reference="",
            terms=0, sent="Yes",
            notes="Three letters, no forwarding address."))

    s.ledger = [{k: v for k, v in row.items() if not k.startswith("_")}
                for row in ledger]

    for charge in charges:
        charge["invoice"] = invoice_for.get(
            (charge["_fam"], charge["_month"]), "")
    s.fees = [{k: v for k, v in c.items() if not k.startswith("_")}
              for c in charges]

    # ------------------------------------------------------ daily reports --
    report_days = all_days[-10:]
    report_rooms = ("Acorns", "Saplings", "Oaks")
    present_key = {(r["child"], r["date"]) for r in s.attendance
                   if r["status"] in ("Present", "Late", "Early pickup")}
    moods = ["Happy", "Settled", "Playful", "Quiet", "Tired", "Excited",
             "Clingy"]
    doings = [
        "Water play in the garden, then story corner",
        "Painting with sponges and a long nap",
        "Music and movement, bubbles outside",
        "Sand tray, dough and a walk to the post box",
        "Construction blocks and a picnic lunch outdoors",
        "Messy play with cooked spaghetti - well worth the clean up",
        "Autumn leaves collage and puddle jumping",
        "Heuristic play baskets and singing",
    ]
    learning = ["Taking turns", "Early mark making", "Self feeding",
                "Listening and attention", "Gross motor - climbing",
                "Naming colours", "Sharing with a friend",
                "Pouring and scooping"]
    for day in report_days:
        for kid in active:
            if kid["room"] not in report_rooms:
                continue
            if (kid["_id"], day) not in present_key:
                continue
            infant = kid["room"] == "Acorns"
            nap_from = t(12, 15) if not infant else t(9, 30)
            nap_to = t(14, 0) if not infant else t(11, 0)
            s.daily_reports.append(dict(
                date=day, child=kid["_id"], mood=rng.choice(moods),
                behaviour=rng.choice([
                    "Settled straight away at drop off.",
                    "A bit tearful at first, fine within ten minutes.",
                    "Full of energy all morning.",
                    "Quiet after lunch, wanted a cuddle.",
                    "Very independent today.",
                ]),
                activities=rng.choice(doings),
                learning=rng.choice(learning),
                outdoor_min=rng.choice([45, 60, 75, 90, 30]),
                nap_from=nap_from, nap_to=nap_to,
                nap2_min=(rng.choice([0, 30, 40]) if infant else 0),
                nappies_wet=(rng.choice([3, 4, 5]) if infant
                             else rng.choice([0, 1, 2])),
                nappies_soiled=(rng.choice([1, 1, 2]) if infant
                                else rng.choice([0, 0, 1])),
                toileting=("Nappy - wet" if infant
                           else rng.choice(["Potty - success",
                                            "Toilet", "Potty - tried",
                                            "Nappy - dry"])),
                potty_notes=("" if infant else
                             rng.choice(["Asked to go twice today.",
                                         "Dry all afternoon.", "", ""])),
                water=rng.choice([3, 4, 5, 6]),
                bottles=(rng.choice([180, 210, 240]) if infant else 0),
                teacher=rng.choice(floor_staff),
                observations=rng.choice([
                    "Stacked four blocks without help - first time.",
                    "Used two words together at snack.",
                    "Shared the trike without being asked.",
                    "Walked the length of the garden unaided.",
                    "Named three colours correctly.",
                ]),
                parent_notes=rng.choice([
                    "Could you send in a spare jumper tomorrow?",
                    "Nappies are running low.",
                    "", "", "A lovely day - see you in the morning.",
                ]),
                shared="Yes"))

    # ----------------------------------------------------------- meals ----
    meal_days = all_days[-4:]
    menu = {
        "Breakfast": ["Porridge with banana", "Toast and fruit",
                      "Weetabix and milk", "Scrambled egg on toast"],
        "Lunch": ["Chicken casserole, rice, green beans",
                  "Pasta bolognese and salad",
                  "Fish pie, peas and sweetcorn",
                  "Lentil curry, rice and naan",
                  "Shepherd's pie and broccoli"],
        "Afternoon snack": ["Apple slices and crackers",
                            "Yoghurt and berries",
                            "Breadsticks and hummus",
                            "Banana and rice cakes"],
    }
    portions = ["All", "Most", "About half", "A little", "None",
                "Refused"]
    weights = [38, 30, 16, 9, 4, 3]
    for day in meal_days:
        for sitting in ("Breakfast", "Lunch", "Afternoon snack"):
            food = rng.choice(menu[sitting])
            for kid in active:
                if (kid["_id"], day) not in present_key:
                    continue
                if sitting == "Breakfast" and rng.random() < 0.45:
                    continue
                infant = kid["room"] == "Acorns"
                alt = ""
                if kid["diet"] == "Gluten free":
                    alt = "Gluten free alternative"
                elif kid["diet"] == "Dairy free" and "milk" in food.lower():
                    alt = "Oat milk"
                elif kid["diet"] == "Vegetarian" and sitting == "Lunch":
                    alt = "Vegetarian portion"
                elif kid["diet"] == "Halal" and sitting == "Lunch":
                    alt = "Halal portion"
                s.meals.append(dict(
                    date=day, sitting=sitting, child=kid["_id"],
                    food=food,
                    portion=rng.choices(portions, weights=weights)[0],
                    alternative=alt,
                    drink="Milk" if infant else rng.choice(
                        ["Water", "Milk", "Water"]),
                    water_ml=rng.choice([120, 150, 180, 200]),
                    own_food="No", logged_by=rng.choice(floor_staff),
                    notes=""))
    # one bottle round for the babies
    for kid in active:
        if kid["room"] != "Acorns":
            continue
        if (kid["_id"], meal_days[-1]) not in present_key:
            continue
        s.meals.append(dict(
            date=meal_days[-1], sitting="Bottle", child=kid["_id"],
            food="Formula, parent supplied", portion="All",
            alternative="", drink="Formula", water_ml=210,
            own_food="Yes", logged_by=staff_id[3],
            notes="Made up fresh at 11:15."))

    # --------------------------------------------------------- medical ----
    med: List[Dict[str, Any]] = []

    def add_med(**row):
        base = dict(child="", record_type="", detail="", severity="",
                    reaction="", action="", medication="", dose="",
                    stored="", trained="", consent="", consent_date=None,
                    prescriber="", doctor="", doctor_phone="",
                    reviewed=None, expires=None, given_date=None,
                    given_time=None, given_by="", witness="", notes="")
        base.update(row)
        med.append(base)

    def kid_by_name(fragment: str) -> Optional[Dict[str, Any]]:
        for item in kids:
            if fragment in item["name"]:
                return item
        return None

    nut = kid_by_name("Rory Clarke")
    asthma = kid_by_name("Seren Hartley")
    coeliac = kid_by_name("Cillian")
    dairy = kid_by_name("Felix Hartley")
    eczema = kid_by_name("Noah Grant")
    egg = kid_by_name("Harvey Thornton")

    if nut:
        add_med(child=nut["_id"], record_type="Allergy",
                detail="Peanuts and all tree nuts",
                severity="Anaphylaxis risk",
                reaction="Hives within minutes, lip swelling, "
                         "difficulty breathing.",
                action="Give adrenaline pen immediately, call 911, "
                       "then phone the parents. Do not wait to see if "
                       "it settles.",
                medication="Adrenaline auto-injector 0.15mg",
                dose="0.15mg, outer thigh, through clothing if needed",
                stored="Red allergy box, Saplings door, unlocked",
                trained="Grace Lim, Tasha Boyd, Dawn Pritchard, "
                        "Beth Carmody",
                consent="Yes", consent_date=d(today, -300),
                prescriber="Dr Imani Walker",
                doctor="Grove Family Practice",
                doctor_phone="(614) 555 0310",
                reviewed=d(today, -120), expires=d(today, 95),
                notes="Allergy action plan with photograph on the "
                      "kitchen wall and in the room file.")
        add_med(child=nut["_id"], record_type="Immunisation",
                detail="Routine schedule up to date",
                severity="None recorded", consent="Yes",
                consent_date=d(today, -300),
                doctor="Grove Family Practice",
                doctor_phone="(614) 555 0310",
                reviewed=d(today, -300), notes="Copy on file.")
    if asthma:
        add_med(child=asthma["_id"], record_type="Asthma",
                detail="Mild, triggered by colds and cold air",
                severity="Moderate",
                reaction="Wheeze and a tight cough, usually after "
                         "running outside.",
                action="Two puffs through the spacer, sit her up, wait "
                       "five minutes. Phone home if it does not settle.",
                medication="Salbutamol inhaler 100mcg",
                dose="2 puffs via spacer, repeat once after 5 minutes",
                stored="Oaks room medical bag",
                trained="Priyanka Rao, Keisha Monroe, Dawn Pritchard",
                consent="Yes", consent_date=d(today, -210),
                prescriber="Dr Peter Lang",
                doctor="Willow Road Surgery",
                doctor_phone="(614) 555 0344",
                reviewed=d(today, -60), expires=d(today, 18),
                notes="Inhaler expires soon - new one requested.")
        for offset, hour in ((-12, 11), (-5, 14), (-2, 10)):
            add_med(child=asthma["_id"], record_type="Medication given",
                    detail="Wheezy after outdoor play",
                    severity="Moderate",
                    medication="Salbutamol inhaler 100mcg",
                    dose="2 puffs via spacer", consent="Yes",
                    consent_date=d(today, -210),
                    given_date=d(today, offset), given_time=t(hour, 20),
                    given_by=staff_id[5], witness="Keisha Monroe",
                    notes="Settled within ten minutes. Parent told at "
                          "pickup.")
    if coeliac:
        add_med(child=coeliac["_id"], record_type="Intolerance",
                detail="Coeliac disease - strictly gluten free",
                severity="Severe",
                reaction="Stomach pain and diarrhoea for days.",
                action="Separate chopping board and toaster. Check "
                       "every label, including playdough.",
                consent="Yes", consent_date=d(today, -340),
                doctor="Eastside Paediatrics",
                doctor_phone="(614) 555 0392",
                reviewed=d(today, -95),
                notes="Kitchen has the full list. Rosemary checks "
                      "deliveries.")
    if dairy:
        add_med(child=dairy["_id"], record_type="Dietary requirement",
                detail="Cow's milk protein allergy",
                severity="Moderate",
                reaction="Rash and reflux.",
                action="Soya formula supplied by the parent, labelled "
                       "and dated in the fridge.",
                consent="Yes", consent_date=d(today, -80),
                doctor="Mulberry Clinic",
                doctor_phone="(614) 555 0377",
                reviewed=d(today, -80), notes="")
    if eczema:
        add_med(child=eczema["_id"], record_type="Eczema",
                detail="Dry patches behind the knees and elbows",
                severity="Mild",
                action="Emollient after water play and before nap.",
                medication="Emollient cream", dose="Thin layer",
                stored="Saplings changing shelf",
                trained="Tasha Boyd, Danny Echeverria",
                consent="Yes", consent_date=d(today, -150),
                reviewed=d(today, -150), expires=d(today, -5),
                notes="Consent needs renewing - parent asked at "
                      "pickup.")
        add_med(child=eczema["_id"], record_type="Medication given",
                detail="Emollient after water play",
                medication="Emollient cream", dose="Thin layer",
                consent="Yes", consent_date=d(today, -150),
                given_date=d(today, -3), given_time=t(15, 10),
                given_by=staff_id[8], witness="Tasha Boyd", notes="")
    if egg:
        add_med(child=egg["_id"], record_type="Allergy",
                detail="Egg - baked egg tolerated",
                severity="Mild",
                reaction="Mouth itching, occasional hives.",
                action="Avoid loose egg. Baked goods are fine.",
                consent="Yes", consent_date=d(today, -130),
                doctor="Grove Family Practice",
                doctor_phone="(614) 555 0310",
                reviewed=d(today, -130), notes="")
    # a scattering of ordinary records so the tab is not all drama
    for index, kid in enumerate(active[:10]):
        add_med(child=kid["_id"], record_type="Immunisation",
                detail="Routine schedule up to date",
                severity="None recorded", consent="Yes",
                consent_date=d(today, -200 - index * 9),
                doctor=_DOCTORS[index % len(_DOCTORS)][0],
                doctor_phone=_DOCTORS[index % len(_DOCTORS)][1],
                reviewed=d(today, -200 - index * 9),
                expires=(d(today, 40) if index == 3 else None),
                notes="")
    s.medical = med

    # -------------------------------------------------------- incidents ---
    incident_rows = [
        ("Bump to head", "Playground", "Minor",
         "Tripped on the edge of the soft surface while running to the "
         "trikes and went down onto her hands and forehead.",
         "Small red mark to the left forehead, no swelling.",
         "Cold compress for five minutes, comforted, observed for the "
         "rest of the session.", 11, True, -2, 10, 45),
        ("Bite", "Toddler room", "Minor",
         "Two children reached for the same truck. One bit the other "
         "on the forearm.",
         "Oval mark to the right forearm, skin not broken.",
         "Washed with soap and water, cold compress, both children "
         "supported.", 14, True, -5, 15, 30),
        ("Cut or graze", "Garden", "Minor",
         "Slipped on wet decking coming down the slide.",
         "Graze to the right knee, approximately 2cm.",
         "Cleaned with saline, plaster applied with parental consent "
         "on file.", 9, True, -8, 11, 20),
        ("Fall", "Hallway", "Minor",
         "Running from the cloakroom and lost footing on the mat.",
         "No visible injury, shaken.",
         "Comforted, checked over, back to play within five minutes.",
         7, True, -12, 13, 55),
        ("Allergic reaction", "Dining area", "Serious",
         "Developed hives on the neck and chest within two minutes of "
         "sitting down for lunch. Food had been checked but a visiting "
         "grandparent had brought in a cake.",
         "Hives to neck and upper chest, no breathing difficulty.",
         "Removed from the table, antihistamine not given (no consent "
         "on file), parents phoned immediately, monitored one to one "
         "until collected.", 4, True, -17, 8, 160),
        ("Illness at the setting", "Preschool room", "No injury",
         "Temperature of 38.4 at the 11am check, flushed and "
         "lethargic.",
         "None.",
         "Moved to the quiet corner, fluids offered, parents called.",
         6, True, -21, 12, 35),
        ("Asthma episode", "Playground", "Moderate",
         "Wheezy and coughing after ten minutes of running outside on "
         "a cold morning.",
         "None - respiratory.",
         "Two puffs of her own inhaler via spacer, sat quietly, "
         "settled within ten minutes.", 22, True, -12, 10, 25),
        ("Near miss", "Entrance", "No injury",
         "The outer door was found on the latch rather than locked "
         "when the cleaner left. No child was in the lobby.",
         "None.",
         "Door secured immediately, lock checked, cleaning contractor "
         "briefed, entry procedure re-read at the staff meeting.",
         0, False, -26, 17, 0),
        ("Bruise", "Toddler room", "Minor",
         "Walked into the corner of the book trolley.",
         "Small bruise to the left shin.",
         "Cold compress and a cuddle.", 17, True, -30, 14, 50),
        ("Minor injury", "Nap room", "Minor",
         "Rolled off the low mat while waking from a nap.",
         "No mark, startled.",
         "Comforted and checked over.", 2, True, -34, 13, 70),
        ("Cut or graze", "Outing", "Minor",
         "Caught a finger on the gate latch at the farm.",
         "Small graze to the index finger.",
         "Cleaned with a wipe from the outing bag, plaster applied.",
         24, True, -38, 11, 240),
        ("Behaviour incident", "Toddler room", "No injury",
         "Repeated throwing of resources at tidy up time, two "
         "occasions in one morning.",
         "None.",
         "Supported with a quiet activity, strategy agreed with the "
         "parent at pickup.", 10, True, -41, 16, 90),
        ("Bump to head", "Classroom", "Moderate",
         "Climbed onto a chair to reach the window and fell backwards "
         "onto the carpet.",
         "Swelling to the back of the head, no loss of consciousness.",
         "Cold compress, head injury observation sheet started, "
         "parents phoned at once and advised to collect.",
         19, True, -46, 9, 8),
        ("Medication error", "Preschool room", "No injury",
         "Inhaler given at 13:40 rather than the 13:00 noted on the "
         "plan - clock in the room was slow.",
         "None - no adverse effect.",
         "Parent and manager informed the same afternoon, room clock "
         "replaced, medication times moved to the kitchen board.",
         22, True, -52, 15, 20),
        ("Lost or found property", "Entrance", "No injury",
         "A coat went home with the wrong family.",
         "None.",
         "Both families phoned, coat swapped the next morning, name "
         "labels reminder sent out.", 13, False, -58, 17, 0),
        ("Fall", "Garden", "Minor",
         "Jumped from the second step of the climbing frame.",
         "Grazed palms, both hands.",
         "Washed and dried, no plaster needed.", 15, True, -63, 10, 60),
        ("Scratch", "Infant room", "Minor",
         "Another baby reached out while they were both on the mat.",
         "Two short scratches to the cheek.",
         "Cleaned, parents shown at pickup, nails checked.",
         1, True, -69, 14, 110),
        ("Illness at the setting", "Toddler room", "No injury",
         "Vomited once after lunch, no temperature.",
         "None.",
         "Parents called, 48 hour exclusion explained, area cleaned "
         "and sanitised.", 20, True, -75, 13, 40),
    ]
    for index, row in enumerate(incident_rows, start=1):
        (category, location, severity, description, injury, first_aid,
         kid_index, notify, day_offset, hour, notify_mins) = row
        when = d(today, day_offset)
        while when.weekday() > 4:
            when -= _dt.timedelta(days=1)
        kid = kids[kid_index] if kid_index else None
        notify_at = (hour * 60 + 15 + notify_mins)
        first_aider = rng.choice(floor_staff)
        s.incidents.append(dict(
            date=when, time=t(hour, 15),
            child=(kid["_id"] if kid else ""), location=location,
            category=category, severity=severity,
            description=description, injury=injury,
            first_aid=first_aid, first_aider=first_aider,
            staff=rng.choice(floor_staff),
            witness=rng.choice(["Dawn Pritchard", "Grace Lim",
                                "Beth Carmody", "Hollie Sanderson"]),
            notified="Yes" if notify else "No",
            notified_date=when if notify else None,
            notified_time=(t(min(18, notify_at // 60), notify_at % 60)
                           if notify else None),
            notified_how=("Phone call" if notify_mins < 60
                          else "In person at pickup") if notify else "",
            notified_by=first_aider if notify else "",
            signed="Yes" if notify and index > 2 else "No",
            signed_date=(d(when, 1) if notify and index > 2 else None),
            follow_up=("Head injury sheet completed and sent home. "
                       "Parent to call if drowsy overnight."
                       if "head" in category.lower() else
                       "Review the allergy procedure for visitor food "
                       "at the next staff meeting."
                       if category == "Allergic reaction" else
                       "Lock checked daily for a fortnight."
                       if category == "Near miss" else ""),
            follow_by=(d(when, 7) if index % 3 == 0 else None),
            follow_done=("Yes" if index % 3 == 0 and index > 6
                         else "No" if index % 3 == 0 else ""),
            reported_out=("County licensing - notified within 24 hours"
                          if severity == "Serious" else ""),
            medical_sought=("Parent took her to the GP the same "
                            "afternoon as a precaution"
                            if severity == "Serious" else ""),
            notes=""))

    # ------------------------------------------------------------ staff ---
    for index, person in enumerate(STAFF, start=1):
        hire = d(today, person["hire"])
        first_aid_exp = d(today, 14 if index == 7 else
                          -20 if index == 10 else
                          120 + index * 23)
        s.staff.append(dict(
            name=person["name"], role=person["role"],
            room=person["room"],
            phone=f"(614) 555 {7000 + index * 13:04d}",
            email=f"{person['name'].split()[0].lower()}"
                  f"@littleacorns.example",
            address=f"{index * 7 + 3} "
                    f"{['Oak', 'Elm', 'Vine', 'Maple', 'Cedar'][index % 5]}"
                    f" Street, Columbus",
            dob=_months_ago(today, 260 + index * 17),
            hire_date=hire, status=person["status"],
            end_date=None, pay_type=person["pay_type"],
            hourly=money(person["hourly"]) if person["hourly"] else None,
            salary=money(person["salary"]) if person["salary"] else None,
            contract_hours=float(person["hours"]) if person["hours"]
            else None,
            day_rate=(money(person["day_rate"])
                      if person.get("day_rate") else None),
            first_aid_exp=first_aid_exp,
            cpr_exp=d(today, 90 + index * 19),
            safeguard_exp=d(today, 45 + index * 31),
            check_exp=d(today, 200 + index * 27),
            check_done=d(hire, -14),
            qualification=person["qual"],
            training=rng.choice([
                "Paediatric first aid, food hygiene level 2",
                "Safeguarding level 2, allergy awareness",
                "SEND awareness, behaviour support",
                "Food hygiene level 2, fire marshal",
            ]),
            em_name=rng.choice(["Partner - see file", "Mother",
                                "Sister", "Brother"]),
            em_phone=f"(614) 555 {8000 + index * 7:04d}",
            notes=("First aid certificate expires in two weeks - "
                   "course booked." if index == 7 else
                   "First aid certificate has lapsed. Off the ratio "
                   "count until renewed." if index == 10 else
                   "Probation review due at three months."
                   if person["status"] == "Probation" else
                   "Books through the agency, usually a day's notice."
                   if person["pay_type"] == "Daily rate" else "")))

    # --------------------------------------------------------- schedule ---
    sched_days = [day for day in all_days if (today - day).days <= 35]
    shift_times = {
        "Opening": (t(6, 45), t(15, 15)),
        "Early": (t(7, 0), t(15, 30)),
        "Morning": (t(8, 0), t(16, 0)),
        "Mid": (t(8, 30), t(16, 30)),
        "Afternoon": (t(9, 30), t(18, 0)),
        "Late": (t(9, 45), t(18, 15)),
        "Closing": (t(10, 0), t(18, 15)),
        "Full day": (t(7, 30), t(17, 30)),
        "Admin": (t(9, 0), t(15, 0)),
        "Sub cover": (t(8, 0), t(17, 0)),
    }
    leave_block = {
        4: (-16, -12, "Annual leave"),
        8: (-6, -5, "Sick leave"),
        11: (-23, -22, "Training"),
    }
    for day in sched_days:
        for index, person in enumerate(STAFF, start=1):
            if person["status"] == "Casual / bank":
                # the sub only works when somebody is away
                continue
            shift = person["shift"]
            leave = leave_block.get(index)
            on_leave = False
            if leave:
                low, high, kind = leave
                if d(today, low) <= day <= d(today, high):
                    on_leave = True
            if on_leave:
                s.schedule.append(dict(
                    date=day, staff=staff_id[index], room="",
                    shift="Leave", leave_type=leave[2],
                    start=None, finish=None, break_min=0,
                    clock_in=None, clock_out=None, substitute="",
                    notes=""))
                continue
            if index == 11 and day.weekday() == 4:
                shift = "Off"
            if index == 9 and day.weekday() in (3, 4):
                shift = "Off"
            if shift == "Off":
                s.schedule.append(dict(
                    date=day, staff=staff_id[index], room="",
                    shift="Off", leave_type="", start=None,
                    finish=None, break_min=0, clock_in=None,
                    clock_out=None, substitute="", notes=""))
                continue
            start, finish = shift_times.get(shift, (t(8, 0), t(16, 0)))
            jitter_in = rng.choice([-4, -2, 0, 0, 1, 3, 6])
            jitter_out = rng.choice([-3, 0, 0, 2, 5, 11])
            mi = start.hour * 60 + start.minute + jitter_in
            mo = finish.hour * 60 + finish.minute + jitter_out
            s.schedule.append(dict(
                date=day, staff=staff_id[index],
                room=person["room"], shift=shift, leave_type="",
                start=start, finish=finish,
                break_min=30 if person["hours"] >= 30 else 15,
                clock_in=t(mi // 60, mi % 60),
                clock_out=t(mo // 60, mo % 60),
                substitute="", notes=""))
    # the substitute covers the week of annual leave
    low, high, _kind = leave_block[4]
    for day in sched_days:
        if d(today, low) <= day <= d(today, high):
            s.schedule.append(dict(
                date=day, staff=staff_id[12], room="Saplings",
                shift="Sub cover", leave_type="", start=t(8, 0),
                finish=t(17, 0), break_min=30, clock_in=t(7, 58),
                clock_out=t(17, 4), substitute=staff_id[4],
                notes="Covering Tasha's annual leave."))

    # ---------------------------------------------------------- payroll ---
    period_end = monday - _dt.timedelta(days=3)      # last Friday
    for run in range(3):
        end = period_end - _dt.timedelta(days=14 * run)
        start = end - _dt.timedelta(days=13)
        for index, person in enumerate(STAFF, start=1):
            if person["status"] == "Casual / bank" and run != 1:
                continue
            monthly = person["pay_type"] == "Salary"
            s.payroll.append(dict(
                period_end=end,
                period_start=start,
                pay_period="Monthly" if monthly else "Fortnightly",
                staff=staff_id[index],
                bonus=(150.0 if index == 3 and run == 1 else 0),
                adjust=(-24.0 if index == 10 and run == 0 else 0),
                deductions=money(
                    (person["salary"] / 12 * 0.21) if monthly else
                    (person["hourly"] * person["hours"] * 2 * 0.19)
                    if person["hourly"] else 0),
                paid="No" if run == 0 else "Yes",
                paid_date=None if run == 0 else d(end, 5),
                method="Bank transfer",
                notes=("Long service bonus." if index == 3 and run == 1
                       else "Correction for a short shift in the "
                            "previous run."
                       if index == 10 and run == 0 else
                       "Agency cover - invoiced separately."
                       if person["pay_type"] == "Daily rate" else "")))

    # --------------------------------------------------------- expenses ---
    recurring = [
        ("Rent", "Monthly rent - 47 Willow Grove", "Grove Property LLC",
         3200.0, 0.0, "Bank transfer", "Monthly", 1),
        ("Utilities", "Gas and electricity", "AEP Ohio", 612.40, 0.0,
         "Direct debit", "Monthly", 8),
        ("Telephone & internet", "Broadband and phone line",
         "Spectrum Business", 109.00, 0.0, "Direct debit", "Monthly",
         8),
        ("Insurance", "Public liability and employers liability",
         "Grange Insurance", 308.75, 0.0, "Direct debit", "Monthly", 2),
        ("Software & subscriptions", "Parent app and payroll software",
         "Brightwheel", 94.00, 0.0, "Card", "Monthly", 5),
        ("Waste & recycling", "Commercial waste collection",
         "Rumpke", 68.50, 0.0, "Direct debit", "Monthly", 12),
        ("Cleaning supplies", "Monthly janitorial order",
         "Buckeye Janitorial", 243.80, 19.50, "Card", "Monthly", 10),
        ("Nappies & hygiene", "Nappies, wipes and creams",
         "Columbus Care Supply", 418.60, 33.50, "Card", "Monthly", 6),
    ]
    for first in months:
        last = min(_month_end(first), today)
        for (category, description, vendor, amount, tax, method,
             freq, day_of) in recurring:
            when = d(first, day_of - 1)
            if when > last:
                continue
            s.expenses.append(dict(
                date=when, category=category, description=description,
                vendor=vendor, amount=money(amount), tax=money(tax),
                method=method, receipt="Yes",
                receipt_ref=f"{vendor[:3].upper()}-{when:%y%m}",
                frequency=freq, notes=""))
        # weekly food shop
        for week_start in _weekdays(first, last)[::5]:
            s.expenses.append(dict(
                date=week_start, category="Food & groceries",
                description="Weekly food order",
                vendor="Kroger wholesale",
                amount=money(rng.uniform(380, 520)), tax=0.0,
                method="Card", receipt="Yes",
                receipt_ref=f"KRG-{week_start:%y%m%d}",
                frequency="Weekly", notes=""))

    one_offs = [
        ("Toys & learning materials", "Replacement home corner set",
         "Ohio Nursery Direct", 486.00, 38.90, -82),
        ("Repairs & maintenance", "Boiler service and two radiators",
         "Hewitt Heating", 612.00, 0.0, -70),
        ("Training & qualifications", "Paediatric first aid, four staff",
         "Midwest Safety Co", 560.00, 0.0, -61),
        ("Marketing & advertising", "Local parenting magazine, autumn",
         "Columbus Family Press", 240.00, 19.20, -55),
        ("Licensing & registration", "Annual licence renewal",
         "Ohio Department of Job and Family Services", 375.00, 0.0,
         -49),
        ("Furniture & equipment", "Six stacking cots",
         "Ohio Nursery Direct", 918.00, 73.44, -44),
        ("Arts & crafts", "Autumn term craft order",
         "Ohio Nursery Direct", 186.40, 14.90, -38),
        ("Field trips", "Farm visit coach hire", "Buckeye Coaches",
         320.00, 0.0, -34),
        ("Professional fees", "Quarterly bookkeeping",
         "Lane & Co Accountants", 420.00, 0.0, -30),
        ("Inspection & compliance", "Fire extinguisher service",
         "Midwest Safety Co", 148.00, 11.84, -27),
        ("Recruitment", "Job advert - assistant teacher",
         "Indeed", 149.00, 0.0, -24),
        ("Office & admin", "Printer toner and stationery",
         "Columbus Office Hub", 96.30, 7.70, -20),
        ("Repairs & maintenance", "Garden fence panel replaced",
         "Hewitt Property Services", 275.00, 0.0, -16),
        ("Bank & card fees", "Card processing - last month",
         "Stripe", 138.45, 0.0, -12),
        ("Educational supplies", "Phonics and story books",
         "Ohio Nursery Direct", 212.80, 17.02, -9),
        ("Staff benefits", "Team lunch - end of term",
         "Donatos", 145.00, 11.60, -6),
        ("Security", "Door entry system call-out", "SafeLock Ohio",
         185.00, 0.0, -4),
        ("Transport & vehicle", "Minibus fuel and check",
         "Shell / Hall's Garage", 164.20, 0.0, -3),
        # the two wage rows that prove the P&L does not double count
        ("Staff wages", "October payroll - posted from the bank",
         "Payroll", 18940.00, 0.0, -2),
        ("Payroll taxes & on-costs", "Employer taxes - October",
         "IRS / Ohio", 2651.60, 0.0, -2),
    ]
    for (category, description, vendor, amount, tax, offset) in one_offs:
        s.expenses.append(dict(
            date=d(today, offset), category=category,
            description=description, vendor=vendor,
            amount=money(amount), tax=money(tax),
            method=rng.choice(["Card", "Bank transfer", "Card"]),
            receipt="Yes" if offset % 7 else "No",
            receipt_ref=(f"{vendor[:3].upper()}-{abs(offset):04d}"
                         if offset % 7 else ""),
            frequency="One off",
            notes=("Counted from the Payroll Tracker, not from here - "
                   "see the Settings switch."
                   if category in ("Staff wages",
                                   "Payroll taxes & on-costs") else "")))

    # --------------------------------------------------------- supplies ---
    for index, item in enumerate(SUPPLIES, start=1):
        (name, category, unit, pack, supplier, lead, opening, recv,
         used, waste, cost, minimum) = item
        s.supplies.append(dict(
            name=name, category=category, unit=unit, pack_size=pack,
            supplier=supplier,
            supplier_phone=f"(614) 555 {9000 + index * 3:04d}",
            lead_days=lead, opening=float(opening),
            received=float(recv), used=float(used), waste=float(waste),
            adjust=(-1.0 if index == 9 else 0.0), cost=money(cost),
            min_level=float(minimum),
            last_ordered=d(today, -(7 + index * 2)),
            last_counted=d(today, -(index % 7)),
            notes=("Count was two out at the last stock take."
                   if index == 9 else
                   "Parent supplies her own - do not reorder."
                   if index == 13 else "")))

    # ---------------------------------------------------------- messages --
    message_rows = [
        (1, "Billing & fees", "Outgoing", "Email",
         "October invoice", "Sent the October invoice and the "
         "breakdown of the two extra days.", "Closed", -27, 0, None),
        (3, "Billing & fees", "Outgoing", "Phone call",
         "Overdue balance", "Left a voicemail about the September "
         "balance. Promised a call back.", "Waiting on parent", -9, 0,
         None),
        (3, "Billing & fees", "Outgoing", "Email",
         "Overdue balance - second note", "Emailed a statement and "
         "offered a payment plan.", "Waiting on parent", -3, 7, None),
        (11, "Billing & fees", "Outgoing", "Text message",
         "October balance", "Texted a reminder. Parent says the "
         "subsidy payment is late at their end.", "Open", -5, 4, None),
        (17, "Billing & fees", "Incoming", "Phone call",
         "Payment arrangement", "Grandmother called to agree a "
         "fortnightly arrangement from November.", "Waiting on us", -2,
         1, None),
        (2, "Settling in", "Outgoing", "App message",
         "Felix's first fortnight", "Sent a short note and two photos "
         "about how the settling in is going.", "Closed", -34, 0, None),
        (5, "Health & illness", "Incoming", "Phone call",
         "Temperature overnight", "Mum phoned to say Saoirse had a "
         "temperature and would not be in.", "Closed", -6, 0, None),
        (10, "Allergy", "Incoming", "In person at pickup",
         "Nut allergy plan review", "Dad asked to go through the "
         "allergy plan again before the Christmas party.",
         "Meeting booked", -8, 5, 6),
        (10, "Accident / incident", "Outgoing", "Phone call",
         "Hives at lunch", "Phoned straight away about the reaction. "
         "Full report sent the same afternoon.", "Closed", -17, 0,
         None),
        (19, "Accident / incident", "Outgoing", "Phone call",
         "Bump to the head", "Phoned within eight minutes. Parent "
         "collected at 10:30.", "Closed", -46, 0, None),
        (4, "Schedule change", "Incoming", "Email",
         "Extra day from January", "Asked whether Mateo could move "
         "from three days to four.", "Waiting on us", -4, 2, None),
        (7, "Enrollment change", "Incoming", "App message",
         "Second child", "Asked about a place for Jude in the Acorns "
         "room next spring.", "Waiting on us", -1, 3, None),
        (13, "Late pickup", "Outgoing", "Phone call",
         "Late collection", "Phoned at 18:05. Dad arrived 18:22. "
         "Late fee explained.", "Closed", -13, 0, None),
        (6, "Policy notice", "Outgoing", "Email",
         "Fee increase from January", "Sent the annual letter - 4% "
         "from 1 January with eight weeks' notice.", "Closed", -40, 0,
         None),
        (8, "Progress update", "Outgoing", "App message",
         "Emilia's half-term summary", "Shared the learning journal "
         "and three observations.", "Closed", -21, 0, None),
        (9, "Complaint", "Incoming", "Email",
         "Lost coat", "Unhappy that a coat went home with another "
         "family. Apologised, coat returned next morning, name labels "
         "reminder sent to all.", "Closed", -57, 0, None),
        (15, "Compliment", "Incoming", "Email",
         "Thank you", "Lovely note about how Noah has settled with "
         "Tasha.", "No reply needed", -11, 0, None),
        (20, "Compliment", "Incoming", "In person at pickup",
         "Thank you for the trip", "Said the farm trip photos made "
         "her week.", "No reply needed", -31, 0, None),
        (12, "Attendance", "Outgoing", "Phone call",
         "Unexplained absence", "Called after two days with no "
         "message. All fine - family were away.", "Closed", -19, 0,
         None),
        (18, "Meeting request", "Incoming", "Email",
         "Parent evening slot", "Asked for a later slot on the "
         "Thursday.", "Meeting booked", -7, 2, 4),
        (14, "Health & illness", "Outgoing", "Phone call",
         "Sickness exclusion", "Explained the 48 hour rule after a "
         "vomiting episode.", "Closed", -74, 0, None),
        (16, "General update", "Outgoing", "Email",
         "Christmas closure dates", "Sent the closure dates and the "
         "January return date.", "Closed", -15, 0, None),
        (1, "Photo consent", "Outgoing", "App message",
         "Photo consent renewal", "Asked for the annual photo consent "
         "to be re-signed.", "Waiting on parent", -10, 5, None),
        (2, "Medication", "Outgoing", "Phone call",
         "Inhaler expiry", "Asked for a replacement inhaler before "
         "the current one expires.", "Waiting on parent", -4, 3, None),
        (20, "Behaviour", "Outgoing", "In person at pickup",
         "Throwing at tidy up", "Agreed a consistent approach at home "
         "and here. Review in a fortnight.", "Open", -41, 14, None),
        (5, "Closure notice", "Outgoing", "Email",
         "Staff training day", "Reminder of the training day closure.",
         "Closed", -26, 0, None),
    ]
    staff_logger = [staff_id[2], staff_id[1], staff_id[4], staff_id[3],
                    staff_id[6]]
    for index, row in enumerate(message_rows, start=1):
        (fam, category, direction, method, subject, text, status,
         offset, follow, meeting) = row
        when = d(today, offset)
        kid_ref = ""
        for item in kids:
            if item["_fam"] == fam:
                kid_ref = item["_id"]
                break
        closed = status in ("Closed", "No reply needed")
        s.messages.append(dict(
            date=when, time=t(9 + index % 8, (index * 7) % 60),
            family=family_id[fam],
            child=kid_ref if index % 3 else "",
            direction=direction, method=method, category=category,
            subject=subject, notes=text,
            staff=staff_logger[index % len(staff_logger)],
            follow_date=(d(when, follow) if follow else None),
            follow_action=("Call back with an answer."
                           if status == "Waiting on us" else
                           "Chase if nothing by the date."
                           if status == "Waiting on parent" else ""),
            status=status,
            closed_date=(d(when, 1) if closed else None),
            meeting=(d(today, meeting) if meeting else None)))

    # --------------------------------------------------------- documents --
    child_docs = [
        ("Enrollment agreement", "Yes", None),
        ("Emergency contact form", "Yes", 365),
        ("Immunisation record", "Yes", None),
        ("Photo & video consent", "No", 365),
        ("Pickup authorization", "Yes", 365),
    ]
    for index, kid in enumerate(kids, start=1):
        for offset, (doc, required, life) in enumerate(child_docs):
            if offset == 4 and index % 3:
                continue
            if offset == 3 and index % 2:
                continue
            received = kid["enrolled_date"]
            status = "On file"
            expires = None
            if life:
                expires = d(received, life)
                if expires < today:
                    status = "Expired"
            if index in (31, 32) and offset in (0, 2):
                status = "Requested"
                received = None
                expires = None
            if index == 24 and offset == 2:
                status = "Chased"
                received = None
            s.documents.append(dict(
                holder="Child", who=kid["_id"], doc_type=doc,
                required=required, status=status,
                requested=d(received or today, -10),
                received=received, expires=expires, renewed=received,
                renew_every=(12 if life else None),
                owner=staff_id[2],
                location="Child files, cabinet 2",
                reference="", notes=""))
    staff_docs = [
        ("Staff contract", "Yes", None, "hire"),
        ("Background check", "Yes", None, "check_exp"),
        ("First aid certificate", "Yes", None, "first_aid_exp"),
        ("Child protection training", "Yes", None, "safeguard_exp"),
        ("Qualification certificate", "No", None, None),
    ]
    for index, person in enumerate(STAFF, start=1):
        record = s.staff[index - 1]
        for doc, required, _life, source in staff_docs:
            expires = record.get(source) if source and source != "hire" \
                else None
            status = "On file"
            if expires and expires < today:
                status = "Expired"
            if index == 9 and doc == "Child protection training":
                status = "Requested"
                expires = None
            s.documents.append(dict(
                holder="Staff", who=staff_id[index], doc_type=doc,
                required=required, status=status,
                requested=d(record["hire_date"], -14),
                received=record["hire_date"] if status != "Requested"
                else None,
                expires=expires, renewed=record["hire_date"],
                renew_every=(36 if doc == "First aid certificate"
                             else None),
                owner=staff_id[2], location="Staff files, locked drawer",
                reference="", notes=("Course booked for the 14th."
                                     if status == "Requested" else "")))
    business_docs = [
        ("Business licence", d(today, 205), "OH-CC-2017-04418",
         "Office, licence folder"),
        ("Insurance policy", d(today, 58), "GRG-884102-11",
         "Office, insurance folder"),
        ("Fire inspection", d(today, 26), "FI-2026-0881",
         "Office, compliance folder"),
        ("Health inspection", d(today, 141), "HI-2026-0442",
         "Office, compliance folder"),
        ("Risk assessment", d(today, -11), "RA-ROOMS-2026",
         "Shared drive, Policies"),
        ("Policy review", d(today, 74), "POL-REVIEW-2026",
         "Shared drive, Policies"),
        ("Food handling certificate", d(today, 312), "FH-RH-2026",
         "Kitchen folder"),
    ]
    for doc, expires, reference, where in business_docs:
        status = "Expired" if expires < today else "On file"
        s.documents.append(dict(
            holder="Business", who="", doc_type=doc, required="Yes",
            status=status, requested=None,
            received=_months_ago(expires, 12), expires=expires,
            renewed=_months_ago(expires, 12), renew_every=12,
            owner=staff_id[1], location=where, reference=reference,
            notes=("Overdue - the annual room review was not signed "
                   "off." if status == "Expired" else "")))

    return s
````

---

## 3. The children and the families

The enrolment record that works out an age in months, drops a child into an age band, counts the siblings in a family and turns any billing basis into a weekly figure; the family record that carries one balance for several children; and the waiting list with its queue position. The shared helpers in `common.py` are where the alert colours, the lookups and the sequence keys live.

### `daycare_business_spreadsheet/sheets/__init__.py`

*0 lines*

````python

````

### `daycare_business_spreadsheet/sheets/common.py`

*509 lines · 49 functions*

> Small building blocks shared by the register sheets.

| Defines | What it does |
|---------|--------------|
| `id_col(trigger, prefix, digits, width, header, note)` | Auto numbered ID column (blank until the trigger column is filled) |
| `month_col(date_key, key, header, width, hidden)` | `yyyy-mm` text column - the key every monthly roll-up groups on |
| `year_col(date_key, key, header, width)` |  |
| `weekday_col(date_key, key, header, width, hidden)` |  |
| `week_col(date_key, key, header, width)` | Week number of the year - the rota and the payroll group on it |
| `week_start_col(date_key, key, header, width)` | The Monday of that week - a far safer grouping key than WEEKNUM |
| `alert_cf()` | Traffic light colouring for the 'Action' / 'Check' columns |
| `status_cf()` | Colouring for the enrollment, attendance, invoice and staff statuses |
| `severity_cf()` | Allergy and incident severity - the one place red really matters |
| `bar_cf(color)` |  |
| `scale_cf()` |  |
| `money_cf()` |  |
| `loss_cf()` | Red when a profit column goes negative |
| `owing_cf()` | Amber on anything still owed, blue on a credit balance |
| `dup_cf()` |  |
| `flag_yes_no_cf()` |  |
| `flag_no_bad_cf()` | For the ones where 'No' is a real problem, not a preference |
| `target_cf(target_name)` | Green at or above a named target, amber below it |
| `yesno(note)` |  |
| `score_validation(maximum)` |  |
| `blank(trigger, body)` | `=IF(trigger="","",body)` with the standard row placeholders |
| `lookup(sheet, target, key_col, key, missing)` | Generic `INDEX/MATCH` against any register, blank row safe |
| `lookup_num(sheet, target, key_col, key)` | Numeric `INDEX/MATCH` - 0 when the ID is missing or the cell empty |
| `lookup_child(target, key, missing)` |  |
| `lookup_child_num(target, key)` |  |
| `lookup_family(target, key, missing)` |  |
| `lookup_staff(target, key, missing)` |  |
| `lookup_staff_num(target, key)` |  |
| `lookup_room(target, key, missing)` |  |
| `lookup_supply(target, key, missing)` |  |
| `in_list(cell, values)` | `OR(cell="a",cell="b",...)` written as a COUNTIF so it stays short |
| `is_present(cell)` | True when an attendance status means the child was in the building |
| `is_absent(cell)` | True when a scheduled session was missed |
| `is_session(cell)` | True when a session was expected at all (the attendance base) |
| `is_enrolled(cell)` | True when an enrollment status means the child holds a place |
| `is_credit(cell)` | True when a fee type reduces the bill instead of adding to it |
| `count_present(child_cell, extra)` |  |
| `count_absent(child_cell, extra)` |  |
| `age_months(dob_cell, on_cell)` | Whole months between a date of birth and a date |
| `age_years_text(dob_cell, on_cell)` | `3y 7m` - how a nursery actually says a child's age |
| `next_birthday(dob_cell)` | The next anniversary of `dob_cell` on or after today |
| `hours_between(start, finish)` | Decimal hours between two time cells, tolerant of an overnight end |
| `rank_desc(key, guard, header, width, note)` | Hidden dense rank, 1 = the biggest value in `key` |
| `rank_asc(key, guard, header, width, note)` | Hidden dense rank, 1 = the smallest value in `key` |
| `min_number(value_range, cond, cap)` | Smallest positive number in `value_range` where `cond` holds |
| `seq_key(parent, trigger)` | Hidden `1, 2, 3 ...` within a parent, plus a `parent|n` key |
| `tax_on(amount, is_tuition)` | `amount` x the tax rate, if that kind of line is taxable |
| `card_fee(amount, method_cell)` | What the card processor keeps out of a payment |
| `late_pickup_charge(minutes_cell)` | The late collection charge for N minutes past closing |

**Module constants:** `ROW_OFFSET` (HEADER_ROW).

````python
"""Small building blocks shared by the register sheets."""

from __future__ import annotations

from typing import Any, Dict, List

from ..config import (ABSENT_STATUSES, CREDIT_FEE_TYPES,
                      NON_SESSION_STATUSES, OCCUPYING_STATUSES,
                      PRESENT_STATUSES)
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
    ``ChildPrefix`` or a quoted literal such as ``'"INC-"'``.
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


def week_col(date_key: str, key: str = "week", header: str = "Week",
             width: float = 7) -> Col:
    """Week number of the year - the rota and the payroll group on it."""
    return Col(key, header, width, "int",
               formula=(f'=IF({{{date_key}}}{{r}}="","",'
                        f'WEEKNUM({{{date_key}}}{{r}},2))'),
               hidden=True)


def week_start_col(date_key: str, key: str = "week_start",
                   header: str = "Week beginning", width: float = 12) -> Col:
    """The Monday of that week - a far safer grouping key than WEEKNUM."""
    return Col(key, header, width, "date",
               formula=(f'=IF({{{date_key}}}{{r}}="","",'
                        f'{{{date_key}}}{{r}}-WEEKDAY({{{date_key}}}{{r}},3))'),
               hidden=True,
               note="The Monday of that week. Weekly totals group on it.")


# ---------------------------------------------------------------------------
# Conditional formatting sets
# ---------------------------------------------------------------------------

def alert_cf() -> List[Dict[str, Any]]:
    """Traffic light colouring for the 'Action' / 'Check' columns."""
    return [
        cf_text("OK", "good", "begins with"),
        cf_text("Overdue", "bad"),
        cf_text("Expired", "bad"),
        cf_text("Missing", "bad"),
        cf_text("Over capacity", "bad"),
        cf_text("Over ratio", "bad"),
        cf_text("Under ratio", "bad"),
        cf_text("Unsafe", "bad"),
        cf_text("Allergy", "bad"),
        cf_text("Not notified", "bad"),
        cf_text("No check-out", "bad"),
        cf_text("Unexplained", "bad"),
        cf_text("Duplicate", "bad"),
        cf_text("Loss", "bad"),
        cf_text("Unpaid", "bad"),
        cf_text("Not authorised", "bad"),
        cf_text("Chase", "warn"),
        cf_text("Due", "warn"),
        cf_text("Check", "warn"),
        cf_text("Action", "warn"),
        cf_text("Expiring", "warn"),
        cf_text("Nearly full", "warn"),
        cf_text("Order", "warn"),
        cf_text("Low stock", "warn"),
        cf_text("Confirm", "warn"),
        cf_text("Add", "warn"),
        cf_text("Invoice", "warn"),
        cf_text("Follow up", "warn"),
        cf_text("Today", "info"),
        cf_text("Tomorrow", "info"),
        cf_text("Upcoming", "info"),
        cf_text("Birthday", "info"),
        cf_text("Starting", "info"),
        cf_text("Space", "info"),
    ]


def status_cf() -> List[Dict[str, Any]]:
    """Colouring for the enrollment, attendance, invoice and staff statuses."""
    return [
        cf_text("Active", "good"),
        cf_text("Present", "good"),
        cf_text("Paid", "good"),
        cf_text("On file", "good"),
        cf_text("Signed", "good"),
        cf_text("Received", "good"),
        cf_text("Enrolled", "good"),
        cf_text("Offer accepted", "good"),
        cf_text("In stock", "good"),
        cf_text("Complete", "good"),
        cf_text("Closed", "good"),
        cf_text("Yes", "good"),
        cf_text("Pending start", "info"),
        cf_text("Sent", "info"),
        cf_text("Tour booked", "info"),
        cf_text("Toured", "info"),
        cf_text("Place offered", "info"),
        cf_text("New enquiry", "info"),
        cf_text("Probation", "info"),
        cf_text("Scheduled", "info"),
        cf_text("Meeting booked", "info"),
        cf_text("Late", "warn"),
        cf_text("Part paid", "warn"),
        cf_text("Draft", "warn"),
        cf_text("Requested", "warn"),
        cf_text("Chased", "warn"),
        cf_text("Waitlisted", "warn"),
        cf_text("On hold", "warn"),
        cf_text("Notice period", "warn"),
        cf_text("Notice given", "warn"),
        cf_text("On leave", "warn"),
        cf_text("Open", "warn"),
        cf_text("Low stock", "warn"),
        cf_text("Early pickup", "warn"),
        cf_text("Absent", "bad"),
        cf_text("Overdue", "bad"),
        cf_text("Expired", "bad"),
        cf_text("Missing", "bad"),
        cf_text("Written off", "bad"),
        cf_text("Cancelled", "bad"),
        cf_text("Out of stock", "bad"),
        cf_text("Offer declined", "bad"),
        cf_text("Lost", "bad"),
        cf_text("Withdrawn", "muted"),
        cf_text("Graduated", "muted"),
        cf_text("Inactive", "muted"),
        cf_text("Left", "muted"),
        cf_text("Not required", "muted"),
        cf_text("Not scheduled", "muted"),
        cf_text("Centre closed", "muted"),
    ]


def severity_cf() -> List[Dict[str, Any]]:
    """Allergy and incident severity - the one place red really matters."""
    return [
        cf_text("Anaphylaxis", "bold_bad"),
        cf_text("Emergency services", "bold_bad"),
        cf_text("Severe", "bad"),
        cf_text("Serious", "bad"),
        cf_text("Moderate", "warn"),
        cf_text("Mild", "info"),
        cf_text("Minor", "info"),
        cf_text("No injury", "good"),
        cf_text("None recorded", "muted"),
    ]


def bar_cf(color: str = "#4D96AE") -> List[Dict[str, Any]]:
    return [{"type": "data_bar", "bar_color": color}]


def scale_cf() -> List[Dict[str, Any]]:
    return [{"type": "3_color_scale", "min_color": "#F8DAD9",
             "mid_color": "#FDEDD2", "max_color": "#DBEDD7"}]


def money_cf() -> List[Dict[str, Any]]:
    return [{"type": "cell", "criteria": ">", "value": 0, "fmt": "bold_good"}]


def loss_cf() -> List[Dict[str, Any]]:
    """Red when a profit column goes negative."""
    return [{"type": "cell", "criteria": "<", "value": 0, "fmt": "bold_bad"},
            {"type": "cell", "criteria": ">", "value": 0, "fmt": "good"}]


def owing_cf() -> List[Dict[str, Any]]:
    """Amber on anything still owed, blue on a credit balance."""
    return [{"type": "cell", "criteria": ">", "value": 0.004, "fmt": "warn"},
            {"type": "cell", "criteria": "<", "value": -0.004, "fmt": "info"}]


def dup_cf() -> List[Dict[str, Any]]:
    return [cf_text("Duplicate", "bad"), cf_text("OK", "good")]


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


def yesno(note: str = "") -> Dict[str, Any]:
    return yes_no(note)


def score_validation(maximum: int = 5) -> Dict[str, Any]:
    return {
        "validate": "integer",
        "criteria": "between",
        "minimum": 1,
        "maximum": maximum,
        "error_title": f"Score 1-{maximum}",
        "error_message": f"Enter a whole number from 1 (poor) to {maximum} "
                         f"(excellent).",
    }


def blank(trigger: str, body: str) -> str:
    """`=IF(trigger="","",body)` with the standard row placeholders."""
    return f'=IF({{{trigger}}}{{r}}="","",{body})'


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


def lookup_child(target: str, key: str = "child",
                 missing: str = "Not in Child Enrollment") -> str:
    return lookup("Child Enrollment", target, "id", key, missing)


def lookup_child_num(target: str, key: str = "child") -> str:
    return lookup_num("Child Enrollment", target, "id", key)


def lookup_family(target: str, key: str = "family",
                  missing: str = "Not in Parent Database") -> str:
    return lookup("Parent Database", target, "id", key, missing)


def lookup_staff(target: str, key: str = "staff",
                 missing: str = "Not in Staff Database") -> str:
    return lookup("Staff Database", target, "id", key, missing)


def lookup_staff_num(target: str, key: str = "staff") -> str:
    return lookup_num("Staff Database", target, "id", key)


def lookup_room(target: str, key: str = "room",
                missing: str = "Not in Classroom Management") -> str:
    return lookup("Classroom Management", target, "name", key, missing)


def lookup_supply(target: str, key: str = "supply",
                  missing: str = "Not in Supply Inventory") -> str:
    return lookup("Supply Inventory", target, "id", key, missing)


# ---------------------------------------------------------------------------
# Shared formula fragments
# ---------------------------------------------------------------------------

def in_list(cell: str, values: List[str]) -> str:
    """`OR(cell="a",cell="b",...)` written as a COUNTIF so it stays short.

    `OR()` evaluates every argument, which makes long chains both slow
    and unsafe next to a lookup, so membership tests are done with an
    arithmetic sum of equality tests instead.
    """
    tests = "+".join(f'({cell}="{v}")' for v in values)
    return f"({tests})>0"


def is_present(cell: str) -> str:
    """True when an attendance status means the child was in the building."""
    return in_list(cell, PRESENT_STATUSES)


def is_absent(cell: str) -> str:
    """True when a scheduled session was missed."""
    return in_list(cell, ABSENT_STATUSES)


def is_session(cell: str) -> str:
    """True when a session was expected at all (the attendance base)."""
    tests = "+".join(f'({cell}="{v}")' for v in NON_SESSION_STATUSES)
    return f'IF({cell}="",FALSE,({tests})=0)'


def is_enrolled(cell: str) -> str:
    """True when an enrollment status means the child holds a place."""
    return in_list(cell, OCCUPYING_STATUSES)


def is_credit(cell: str) -> str:
    """True when a fee type reduces the bill instead of adding to it."""
    return in_list(cell, CREDIT_FEE_TYPES)


#: `SUMIFS` across the attendance register for one child, counting only
#: the statuses that mean they were actually here.
def count_present(child_cell: str, extra: str = "") -> str:
    parts = "+".join(
        f'COUNTIFS([[Attendance.child]],{child_cell},'
        f'[[Attendance.status]],"{status}"{extra})'
        for status in PRESENT_STATUSES)
    return parts


def count_absent(child_cell: str, extra: str = "") -> str:
    parts = "+".join(
        f'COUNTIFS([[Attendance.child]],{child_cell},'
        f'[[Attendance.status]],"{status}"{extra})'
        for status in ABSENT_STATUSES)
    return parts


def age_months(dob_cell: str, on_cell: str = "TodayDate") -> str:
    """Whole months between a date of birth and a date.

    `DATEDIF` is the one Excel function that gets this right across a
    leap year, and it has existed since Excel 5, so it is safe here.
    """
    return f'IF({dob_cell}="","",DATEDIF({dob_cell},MAX({dob_cell},{on_cell}),"m"))'


def age_years_text(dob_cell: str, on_cell: str = "TodayDate") -> str:
    """`3y 7m` - how a nursery actually says a child's age."""
    return (f'IF({dob_cell}="","",'
            f'DATEDIF({dob_cell},MAX({dob_cell},{on_cell}),"y")&"y "&'
            f'DATEDIF({dob_cell},MAX({dob_cell},{on_cell}),"ym")&"m")')


def next_birthday(dob_cell: str) -> str:
    """The next anniversary of `dob_cell` on or after today."""
    return (f'IF({dob_cell}="","",'
            f'IF(DATE(YEAR(TodayDate),MONTH({dob_cell}),DAY({dob_cell}))'
            f'<TodayDate,'
            f'DATE(YEAR(TodayDate)+1,MONTH({dob_cell}),DAY({dob_cell})),'
            f'DATE(YEAR(TodayDate),MONTH({dob_cell}),DAY({dob_cell}))))')


def hours_between(start: str, finish: str) -> str:
    """Decimal hours between two time cells, tolerant of an overnight end."""
    return (f'IF({start}="","",IF({finish}="","",'
            f'ROUND((MOD({finish}-{start},1))*24,2)))')


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


def min_number(value_range: str, cond: str, cap: int = 999999) -> str:
    """Smallest positive number in `value_range` where `cond` holds.

    Excel 2010 has no MINIFS, and a plain SUMPRODUCT(MIN(...)) trips over
    the empty strings that formula columns leave behind, so the values are
    coerced with IFERROR(...*1,0) and the non matching rows are swapped for
    a big number.  Returns `cap` when nothing matches - test for that.
    """
    safe = f"IFERROR({value_range}*1,0)"
    hit = f"(({cond})*({safe}>0))"
    return f"SUMPRODUCT(MIN({hit}*{safe}+(1-{hit})*{cap}))"


def seq_key(parent: str, trigger: str = "") -> List[Col]:
    """Hidden `1, 2, 3 ...` within a parent, plus a `parent|n` key.

    Printables and per-parent summaries use `MATCH(parent & "|" & n, key, 0)`
    to pull the nth line of one invoice without an array formula.
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

def tax_on(amount: str, is_tuition: str) -> str:
    """`amount` x the tax rate, if that kind of line is taxable.

    Childcare is tax exempt in a great many places, which is why the
    default is zero and both switches are set on the Settings tab rather
    than assumed here.
    """
    return (f'ROUND({amount}*IF({is_tuition},'
            f'IF(TaxOnTuition="Yes",TaxRate,0),'
            f'IF(TaxOnExtras="Yes",TaxRate,0)),2)')


def card_fee(amount: str, method_cell: str) -> str:
    """What the card processor keeps out of a payment."""
    return (f'IF({method_cell}="Card",ROUND({amount}*CardFeePercent+'
            f'CardFeeFixed,2),0)')


def late_pickup_charge(minutes_cell: str) -> str:
    """The late collection charge for N minutes past closing.

    Free for the grace period, then a flat first block, then a per minute
    rate - which is how most settings actually word it in the handbook.
    """
    return (f'IF({minutes_cell}<=LatePickupGrace,0,'
            f'ROUND(LatePickupFee+MAX(0,{minutes_cell}-LatePickupGrace-'
            f'LatePickupBlock)*LatePickupPerMin,2))')
````

### `daycare_business_spreadsheet/sheets/children.py`

*501 lines · 2 functions*

> `Child Enrollment` - the register every other tab hangs off.
>
> One row per child: who they are, how old they are today, which room
> they are in, which family pays, what that family is contracted to pay,
> how often the child actually turns up, and what is still outstanding on
> their paperwork and their account.
>
> Age is calculated, never typed.  So is the age band, the weekly fee,
> the attendance percentage, the balance and the list of missing
> documents.  The only things you type are the things only you know.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Enrollment summary strip |

**Module constants:** `NAME` ('Child Enrollment').

````python
"""`Child Enrollment` - the register every other tab hangs off.

One row per child: who they are, how old they are today, which room
they are in, which family pays, what that family is contracted to pay,
how often the child actually turns up, and what is still outstanding on
their paperwork and their account.

Age is calculated, never typed.  So is the age band, the weekly fee,
the attendance percentage, the balance and the list of missing
documents.  The only things you type are the things only you know.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec, max_if
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check, time_check
from .common import (age_months, age_years_text, alert_cf, bar_cf, blank,
                     count_absent, count_present, id_col, lookup_family,
                     next_birthday, owing_cf, rank_desc, status_cf,
                     yesno)

NAME = "Child Enrollment"


def spec(cfg: Config) -> SheetSpec:
    med_child = "[[Allergy & Medical.child]]"
    med_sev = "[[Allergy & Medical.severity]]"
    doc_who = "[[Document Tracker.who]]"
    doc_req = "[[Document Tracker.required]]"
    doc_status = "[[Document Tracker.status]]"

    def severity_count(level: str) -> str:
        return f'COUNTIFS({med_child},{{id}}{{r}},{med_sev},"{level}")'

    columns = [
        id_col("name", "ChildPrefix", digits=4, width=11,
               header="Child ID",
               note="Automatic. Attendance, fees, meals, daily reports "
                    "and incidents all point at this ID."),
        Col("name", "Child's full name", 24, "text",
            note="As it appears on the birth certificate and the "
                 "enrollment agreement. One row per child, always."),
        Col("preferred", "Known as", 14, "text",
            note="What you actually call them. Printed on the daily "
                 "report."),
        Col("dob", "Date of birth", 12, "date", validate=date_check(),
            note="Everything age related in this workbook is worked "
                 "out from this one cell - age, age band, room "
                 "suitability, birthdays, moving-up dates."),
        Col("age_months", "Age (months)", 11, "int",
            formula="=" + age_months("{dob}{r}"), hidden=True),
        Col("age", "Age", 10, "text",
            formula="=" + age_years_text("{dob}{r}"),
            note="Calculated. Recalculates every time you open the "
                 "file."),
        Col("age_band", "Age band", 14, "text",
            formula=blank(
                "dob",
                'IFERROR(INDEX(AgeBandList,MATCH({age_months}{r},'
                'AgeBandFrom,1))&"","")'),
            note="From the age band table on Settings. Tells you when "
                 "a child has grown out of the room they are in."),
        Col("gender", "Gender", 11, "text",
            validate=list_from("GenderList"),
            note="Optional. Some settings record it, some deliberately "
                 "do not - leave it blank if you do not."),
        Col("family", "Family ID", 11, "id",
            validate=list_from("FamilyList"),
            note="The family that pays. Two siblings share one family "
                 "ID - that is what makes one invoice per family and "
                 "the sibling discount possible."),
        Col("family_name", "Family", 18, "text",
            formula=lookup_family("family_name")),
        Col("parent_name", "Main contact", 20, "text",
            formula=lookup_family("parent1")),
        Col("parent_phone", "Phone", 15, "text",
            formula=lookup_family("phone1")),
        Col("parent_email", "Email", 24, "text",
            formula=lookup_family("email1"), hidden=True),
        Col("family_seq", "Sibling no.", 9, "int",
            formula=blank(
                "family",
                'COUNTIFS({family_first}:{family}{r},{family}{r})'),
            hidden=True,
            note="1 for the first child of a family, 2 for the next. "
                 "The sibling discount reads this."),
        Col("family_key", "Family key", 12, "text",
            formula=blank("family",
                          '{family}{r}&"|"&{family_seq}{r}'),
            hidden=True),
        # -------------------------------------------------- the place ---
        Col("program", "Programme", 14, "text",
            validate=list_from("ProgramList"), cf=status_cf()),
        Col("room", "Classroom", 16, "text",
            validate=list_from("ClassroomList"),
            note="Pick from the rooms on Classroom Management. The "
                 "capacity and ratio figures count this column."),
        Col("teacher", "Key teacher", 18, "text",
            formula=blank(
                "room",
                'IFERROR(INDEX([[Classroom Management.teacher_name]],'
                'MATCH({room}{r},[[Classroom Management.name]],0))&"",'
                '"")'),
            note="The lead teacher of whichever room the child is in."),
        Col("schedule", "Schedule", 16, "text",
            validate=list_from("ScheduleTypeList")),
        Col("days_week", "Days a week", 11, "number1",
            validate=number_check(0, 7),
            note="Contracted sessions a week. Daily and per-session "
                 "fees multiply by this."),
        Col("hours_day", "Hours a day", 11, "number1",
            validate=number_check(0, 24),
            note="Contracted hours. Hourly fees multiply by this and "
                 "by the days a week."),
        Col("usual_in", "Usual drop-off", 12, "time",
            validate=time_check(),
            note="The contracted arrival time. Attendance measures "
                 "late arrivals against it."),
        Col("usual_out", "Usual pickup", 12, "time",
            validate=time_check(),
            note="The contracted collection time. The late pickup fee "
                 "is calculated from this, not from closing time."),
        Col("enrolled_date", "Enrolled on", 12, "date",
            validate=date_check(),
            note="The day the paperwork was signed."),
        Col("start_date", "Starts", 11, "date", validate=date_check(),
            note="First session. A start date in the future with "
                 "status 'Pending start' is a place held, not a child "
                 "in the room."),
        Col("end_date", "Last day", 11, "date", validate=date_check(),
            note="Leave blank while the child is still with you."),
        Col("leaving_reason", "Leaving because", 18, "text",
            validate=list_from("LeavingReasonList"), hidden=True,
            note="Worth filling in. Six of these in a row tells you "
                 "something an individual conversation never will."),
        Col("status", "Status", 13, "status",
            validate=list_from("EnrollmentStatusList"), cf=status_cf(),
            note="Active, On hold and Notice given all occupy a place. "
                 "Withdrawn and Graduated free it up."),
        Col("holds_place", "Holds a place?", 11, "flag",
            formula=blank(
                "name",
                'IF(({status}{r}="Active")+({status}{r}="On hold")+'
                '({status}{r}="Notice given")>0,"Yes","No")'),
            hidden=True),
        Col("months_with_us", "Months with us", 12, "int",
            formula=blank(
                "start_date",
                'DATEDIF(MIN({start_date}{r},TodayDate),'
                'IF({end_date}{r}="",TodayDate,{end_date}{r}),"m")'),
            hidden=True),
        # ------------------------------------------------- the money ----
        Col("basis", "Billing basis", 13, "text",
            validate=list_from("BillingBasisList"),
            note="How this family's tuition is worked out. Different "
                 "children can be on different bases - a full-timer on "
                 "a monthly rate next to a drop-in on an hourly one."),
        Col("rate", "Agreed rate", 12, "money",
            validate=number_check(0, 100000),
            note="The rate for the basis you picked: the weekly rate "
                 "if weekly, the hourly rate if hourly, and so on."),
        Col("auto_discount", "Sibling %", 9, "percent",
            formula=blank(
                "name",
                'IF({family_seq}{r}>1,SiblingDiscount,0)'),
            hidden=True),
        Col("discount_pct", "Discount", 10, "percent",
            validate=number_check(0, 1),
            note="Leave blank and the sibling discount applies "
                 "automatically to the second and later child of a "
                 "family. Type a figure here to override it."),
        Col("discount_used", "Discount applied", 12, "percent",
            formula=blank(
                "name",
                'IF({discount_pct}{r}<>"",{discount_pct}{r},'
                '{auto_discount}{r})'),
            note="What is actually coming off. Typed beats automatic."),
        Col("subsidy_week", "Subsidy a week", 12, "money",
            validate=number_check(0, 100000),
            note="Government or employer help paid on this child's "
                 "behalf. It comes off what the family owes, and it is "
                 "still revenue to you."),
        Col("gross_week", "Full rate a week", 13, "money",
            formula=blank(
                "name",
                'ROUND(IF({basis}{r}="Weekly rate",{rate}{r},'
                'IF({basis}{r}="Fortnightly rate",{rate}{r}/2,'
                'IF({basis}{r}="Monthly rate",{rate}{r}*12/52,'
                'IF({basis}{r}="Term rate",{rate}{r}/MAX(1,TermWeeks),'
                'IF({basis}{r}="Daily rate",{rate}{r}*{days_week}{r},'
                'IF({basis}{r}="Per session",{rate}{r}*{days_week}{r},'
                'IF({basis}{r}="Hourly rate",'
                '{rate}{r}*{hours_day}{r}*{days_week}{r},0))))))),2)'),
            note="Every basis converted to one comparable weekly "
                 "number, before any discount."),
        Col("weekly_net", "Weekly fee", 12, "money",
            formula=blank(
                "name",
                'ROUND(MAX(0,{gross_week}{r}*(1-{discount_used}{r})-'
                '{subsidy_week}{r}),2)'),
            cf=bar_cf(Palette.accent),
            note="After the discount and after the subsidy - what the "
                 "family actually pays in a normal week."),
        Col("monthly_net", "Monthly fee", 12, "money",
            formula=blank(
                "name",
                'ROUND({weekly_net}{r}*52/12,2)'),
            note="The weekly fee annualised and divided by twelve, so "
                 "that every month is the same amount."),
        Col("annual_net", "A year", 12, "money",
            formula=blank("name",
                          'ROUND({weekly_net}{r}*WeeksOpen,2)'),
            hidden=True,
            note="Weekly fee x the weeks you open. What this child is "
                 "worth over a year."),
        Col("billed", "Billed to date", 12, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Tuition & Fees.net]],'
                '[[Tuition & Fees.child]],{id}{r}),2)'),
            note="Every charge raised against this child on the "
                 "Tuition & Fees tab, discounts and credits included."),
        Col("balance", "Family balance", 12, "money",
            formula=blank(
                "family",
                'IFERROR(INDEX([[Parent Database.balance]],'
                'MATCH({family}{r},[[Parent Database.id]],0))*1,0)'),
            cf=owing_cf(),
            note="The whole family's balance, not just this child's - "
                 "because you invoice the family, not the child."),
        # ---------------------------------------------- how often in ----
        Col("sessions", "Sessions booked", 13, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Attendance.child]],{id}{r})-'
                '(' + "+".join(
                    f'COUNTIFS([[Attendance.child]],{{id}}{{r}},'
                    f'[[Attendance.status]],"{s}")'
                    for s in ("Public holiday", "Centre closed",
                              "Snow day", "Not scheduled")) + ')'),
            hidden=True,
            note="Days a session was expected. Closures and days the "
                 "child was never booked for do not count."),
        Col("days_in", "Days attended", 12, "int",
            formula=blank("name", count_present("{id}{r}")),
            cf=bar_cf()),
        Col("days_out", "Days absent", 11, "int",
            formula=blank("name", count_absent("{id}{r}")),
            cf=[{"type": "cell", "criteria": ">=", "value": 10,
                 "fmt": "warn"}]),
        Col("days_sick", "Sick days", 10, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Attendance.child]],{id}{r},'
                '[[Attendance.status]],"Absent - sick")'),
            hidden=True),
        Col("attendance_pct", "Attendance", 11, "percent",
            formula=blank(
                "name",
                'IF({sessions}{r}<=0,"",{days_in}{r}/{sessions}{r})'),
            cf=[{"type": "cell", "criteria": ">=",
                 "value": "=TargetAttendance", "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0.7,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<",
                 "value": "=TargetAttendance", "fmt": "warn"}],
            note="Days attended divided by sessions expected. "
                 "Closures and unbooked days are excluded, so this is "
                 "a fair figure."),
        Col("last_in", "Last attended", 12, "date",
            formula=blank(
                "name",
                'IF({days_in}{r}=0,"",' +
                max_if("[[Attendance.date]]",
                       "[[Attendance.child]]={id}{r}",
                       '[[Attendance.present]]=1') + ")")),
        Col("days_since", "Days since", 10, "int",
            formula=blank(
                "name",
                'IF({last_in}{r}="","",TodayDate-{last_in}{r})'),
            hidden=True),
        Col("late_count", "Late arrivals", 11, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Attendance.child]],{id}{r},'
                '[[Attendance.status]],"Late")'),
            hidden=True),
        Col("late_pickups", "Late pickups", 11, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Attendance.child]],{id}{r},'
                '[[Attendance.late_out]],">0")'),
            cf=[{"type": "cell", "criteria": ">=", "value": 3,
                 "fmt": "warn"}],
            note="Collections after the contracted pickup time. Three "
                 "of them is a conversation, not a fee."),
        # -------------------------------------------------- care needs --
        Col("diet", "Diet", 14, "text",
            validate=list_from("DietList"),
            note="Feeds the meal tracker, where it sits next to every "
                 "plate you serve."),
        Col("allergy_level", "Allergy risk", 13, "status",
            formula=blank(
                "name",
                f'IF({severity_count("Anaphylaxis risk")}>0,'
                f'"Anaphylaxis risk",'
                f'IF({severity_count("Severe")}>0,"Severe",'
                f'IF({severity_count("Moderate")}>0,"Moderate",'
                f'IF({severity_count("Mild")}>0,"Mild",'
                f'"None recorded"))))'),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Anaphylaxis", "fmt": "bold_bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Severe", "fmt": "bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Moderate", "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "None", "fmt": "muted"}],
            note="The worst severity on the Allergy & Medical tab for "
                 "this child. Nothing to type - record the allergy "
                 "there and it appears here, on the meal tracker and "
                 "on the daily report."),
        Col("allergy_count", "Allergies", 9, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Allergy & Medical.child]],{id}{r},'
                '[[Allergy & Medical.record_type]],"Allergy")'),
            hidden=True),
        Col("first_allergy", "Main allergy", 18, "text",
            formula=blank(
                "name",
                'IFERROR(INDEX([[Allergy & Medical.detail]],'
                'MATCH({id}{r}&"|1",[[Allergy & Medical.key]],0))&"",'
                '"")'),
            hidden=True),
        Col("medical_notes", "Care notes", 28, "wrap",
            note="Asthma inhaler in the red box, needs a nap at 12, "
                 "frightened of the hand dryer. The things a cover "
                 "member of staff has to know."),
        Col("pickup_list", "Authorised to collect", 30, "wrap",
            note="Names of everybody allowed to collect this child. "
                 "Attendance checks what you type at pickup against "
                 "this list and flags anything it cannot find."),
        Col("pickup_password", "Pickup password", 14, "text",
            note="A word only the family knows, for the day somebody "
                 "unexpected arrives."),
        Col("photo_ok", "Photos OK?", 10, "flag", validate=yesno(),
            cf=[{"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "warn"}],
            note="Check this before anything goes on social media."),
        Col("sunscreen_ok", "Sunscreen OK?", 11, "flag",
            validate=yesno(), hidden=True),
        Col("outings_ok", "Outings OK?", 10, "flag", validate=yesno(),
            hidden=True),
        # ------------------------------------------------- paperwork ----
        Col("docs_required", "Docs required", 11, "int",
            formula=blank(
                "name",
                f'COUNTIFS({doc_who},{{id}}{{r}},{doc_req},"Yes")'),
            hidden=True),
        Col("docs_ok", "Docs on file", 11, "int",
            formula=blank(
                "name",
                f'COUNTIFS({doc_who},{{id}}{{r}},{doc_req},"Yes",'
                f'{doc_status},"On file")+'
                f'COUNTIFS({doc_who},{{id}}{{r}},{doc_req},"Yes",'
                f'{doc_status},"Signed")+'
                f'COUNTIFS({doc_who},{{id}}{{r}},{doc_req},"Yes",'
                f'{doc_status},"Received")'),
            hidden=True),
        Col("docs_missing", "Docs missing", 11, "int",
            formula=blank(
                "name",
                '{docs_required}{r}-{docs_ok}{r}'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "=", "value": 0,
                 "fmt": "good"}],
            note="Required documents for this child that are not yet "
                 "received, signed or on file. Set them up on the "
                 "Document Tracker."),
        Col("contract", "Agreement", 12, "status",
            validate=list_from("DocumentStatusList"), cf=status_cf(),
            note="The enrollment agreement and consent forms. "
                 "Shorthand for the full picture on the Document "
                 "Tracker."),
        # -------------------------------------------------- birthdays ---
        Col("next_bday", "Next birthday", 12, "date",
            formula="=" + next_birthday("{dob}{r}")),
        Col("bday_in", "Days to birthday", 12, "int",
            formula=blank("dob", '{next_bday}{r}-TodayDate'),
            hidden=True),
        Col("turning", "Turning", 8, "int",
            formula=blank(
                "dob", 'YEAR({next_bday}{r})-YEAR({dob}{r})'),
            hidden=True),
        Col("anniversary", "Years with us", 11, "number1",
            formula=blank(
                "start_date",
                'ROUND((IF({end_date}{r}="",TodayDate,{end_date}{r})-'
                '{start_date}{r})/365,1)'),
            hidden=True),
        Col("moves_up", "Moves up on", 12, "date",
            formula=blank(
                "dob",
                'IF({age_band}{r}="","",'
                'IFERROR(EDATE({dob}{r},INDEX(AgeBandTo,'
                'MATCH({age_months}{r},AgeBandFrom,1))),""))'),
            note="The day the child ages out of their band. Plan the "
                 "room move before the parent asks about it."),
        Col("notes", "Notes", 28, "wrap"),
        rank_desc("annual_net", header="Value rank"),
        Col("alert", "Action", 34, "status",
            formula=blank(
                "name",
                'IF({dob}{r}="","Add - no date of birth",'
                'IF({family}{r}="","Add - no family linked",'
                'IF(({status}{r}="Withdrawn")+({status}{r}="Graduated")'
                '>0,"",'
                'IF({docs_missing}{r}>0,"Missing - "&{docs_missing}{r}&'
                '" document(s) outstanding",'
                'IF({balance}{r}>0.004,"Chase - family owes money",'
                'IF({room}{r}="","Add - no classroom",'
                'IF({basis}{r}="","Add - no billing basis",'
                'IF(IF({weekly_net}{r}=0,{basis}{r}<>"No charge",FALSE),'
                '"Check - fee works out at zero",'
                'IF(IF({moves_up}{r}="",FALSE,'
                '{moves_up}{r}<=TodayDate+30),'
                '"Due - ages out of the band soon",'
                'IF(IF({bday_in}{r}="",FALSE,'
                '{bday_in}{r}<=BirthdayWindow),'
                '"Birthday - turning "&{turning}{r}&" in "&'
                '{bday_in}{r}&" days",'
                'IF(IF({attendance_pct}{r}="",FALSE,'
                '{attendance_pct}{r}<0.7),'
                '"Check - attendance below 70%","OK")))))))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Child Enrollment & Registration",
        subtitle="Every child, their family, their room, their fee and "
                 "the paperwork still outstanding.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).children,
        help_text="Type the name, the date of birth and the family ID. "
                  "Age, age band, the weekly fee, attendance, the "
                  "balance and the missing paperwork all calculate "
                  "themselves.",
        extra_notes=[
            "The sibling discount is automatic: the first child of a "
            "family pays full price, the second and later ones get "
            "the percentage set on Settings. Type a figure in the "
            "Discount column to override it for one child.",
            "'Holds a place' is what the capacity figures count. A "
            "child on notice still occupies a seat until their last "
            "day - which is exactly when you need to know about it.",
            "Filter the Action column every Monday morning. Missing "
            "documents, outstanding balances and children about to "
            "age out of their room are the three things that cost "
            "money if you find them late.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Enrollment summary strip."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Enrollment at a glance", last)
    sheet.table_head(row, 0, [
        "On the register", "Active", "Pending start", "On notice",
        "Withdrawn", "Average age (months)", "Weekly contracted fees",
        "Average weekly fee", "Average attendance",
        "Children with documents missing", "Allergy plans on file"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Child Enrollment.name]],"<>")',
        '=COUNTIFS([[Child Enrollment.status]],"Active")',
        '=COUNTIFS([[Child Enrollment.status]],"Pending start")',
        '=COUNTIFS([[Child Enrollment.status]],"Notice given")',
        '=COUNTIFS([[Child Enrollment.status]],"Withdrawn")',
        '=IFERROR(ROUND(AVERAGEIF([[Child Enrollment.age_months]],'
        '">0"),0),0)',
        '=ROUND(SUMIFS([[Child Enrollment.weekly_net]],'
        '[[Child Enrollment.holds_place]],"Yes"),2)',
        '=IFERROR(ROUND(SUMIFS([[Child Enrollment.weekly_net]],'
        '[[Child Enrollment.holds_place]],"Yes")/'
        'MAX(1,COUNTIFS([[Child Enrollment.holds_place]],"Yes")),2),0)',
        '=IFERROR(AVERAGEIF([[Child Enrollment.attendance_pct]],">0"),'
        '"")',
        '=COUNTIFS([[Child Enrollment.docs_missing]],">0")',
        '=COUNTIFS([[Child Enrollment.allergy_count]],">0")',
    ], ["int", "int", "int", "int", "int", "int", "money", "money",
        "percent", "int", "int"])
````

### `daycare_business_spreadsheet/sheets/parents.py`

*316 lines · 2 functions*

> `Parent Database` - one row per family, however many children they have.
>
> The mistake most childcare templates make is one row per parent, which
> means a family with two children appears twice, gets two invoices and
> shows two half-balances.  Here a family is a family: one row, one
> balance, one invoice, as many children as they like linked to it.
>
> Emergency contacts, authorised collectors and the preferred way of
> being contacted all live here too, because those are family facts
> rather than child facts.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Family book summary |

**Module constants:** `NAME` ('Parent Database').

````python
"""`Parent Database` - one row per family, however many children they have.

The mistake most childcare templates make is one row per parent, which
means a family with two children appears twice, gets two invoices and
shows two half-balances.  Here a family is a family: one row, one
balance, one invoice, as many children as they like linked to it.

Emergency contacts, authorised collectors and the preferred way of
being contacted all live here too, because those are family facts
rather than child facts.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec, max_if, min_if
from ..sample_data import get as sample
from ..writer import list_from, date_check
from .common import (alert_cf, bar_cf, blank, id_col, owing_cf,
                     rank_desc, status_cf)

NAME = "Parent Database"


def spec(cfg: Config) -> SheetSpec:
    child_family = "[[Child Enrollment.family]]"
    child_status = "[[Child Enrollment.status]]"
    led_family = "[[Invoices & Receipts.family]]"

    # the earliest due date among this family's overdue invoices, and
    # the sentinel min_if() hands back when there are none
    _oldest_due = min_if("[[Invoices & Receipts.due_num]]",
                         f"{led_family}={{id}}{{r}}",
                         '[[Invoices & Receipts.status]]="Overdue"')
    led_type = "[[Invoices & Receipts.type]]"

    def nth_child(n: int) -> str:
        return (f'IFERROR(INDEX([[Child Enrollment.name]],'
                f'MATCH({{id}}{{r}}&"|{n}",'
                f'[[Child Enrollment.family_key]],0))&"","")')

    columns = [
        id_col("parent1", "FamilyPrefix", digits=4, width=11,
               header="Family ID",
               note="Automatic. Children, invoices, payments and "
                    "messages all point at this ID. One family, one "
                    "ID, however many children."),
        Col("family_name", "Family name", 16, "text",
            note="The surname you would write on an envelope. Used on "
                 "invoices and in every list that groups by family."),
        # ------------------------------------------------ first parent --
        Col("parent1", "Parent / guardian 1", 22, "text",
            note="The main contact. Filling this in is what creates "
                 "the family ID."),
        Col("relation1", "Relationship", 13, "text",
            validate=list_from("RelationshipList")),
        Col("phone1", "Phone", 15, "text"),
        Col("email1", "Email", 26, "text"),
        Col("employer1", "Employer", 18, "text", hidden=True,
            note="Optional. Useful when an employer pays part of the "
                 "fees or runs a childcare voucher scheme."),
        Col("work_phone1", "Work phone", 14, "text", hidden=True),
        # ----------------------------------------------- second parent --
        Col("parent2", "Parent / guardian 2", 22, "text"),
        Col("relation2", "Relationship", 13, "text",
            validate=list_from("RelationshipList")),
        Col("phone2", "Phone", 15, "text"),
        Col("email2", "Email", 26, "text", hidden=True),
        # --------------------------------------------------- where ------
        Col("address", "Address", 30, "text"),
        Col("city", "Town / city", 16, "text"),
        Col("postcode", "Postcode / ZIP", 12, "text"),
        Col("contact_by", "Contact by", 15, "text",
            validate=list_from("ContactMethodList"),
            note="How this family actually wants to hear from you. "
                 "Respect it and you will get replies."),
        Col("language", "Main language", 14, "text", hidden=True,
            note="Worth recording if you ever send a notice home."),
        # --------------------------------------------- emergency ------
        Col("em1_name", "Emergency contact 1", 20, "text",
            note="Somebody who is not parent 1 or 2, and who can get "
                 "here quickly."),
        Col("em1_phone", "Phone", 15, "text"),
        Col("em1_relation", "Relationship", 13, "text",
            validate=list_from("RelationshipList"), hidden=True),
        Col("em2_name", "Emergency contact 2", 20, "text",
            hidden=True),
        Col("em2_phone", "Phone", 15, "text", hidden=True),
        Col("doctor", "Family doctor", 20, "text", hidden=True),
        Col("doctor_phone", "Doctor's phone", 15, "text", hidden=True),
        Col("pickup_list", "Authorised to collect", 32, "wrap",
            note="Everybody this family has authorised, separated by "
                 "commas. Individual children can have their own list "
                 "on the enrollment tab - the attendance check looks "
                 "at both."),
        Col("pickup_ok", "Pickup list signed?", 13, "flag",
            validate=list_from("YesNoList"),
            cf=[{"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "bad"}],
            note="Has the family signed the pickup authorisation form? "
                 "Until they have, only the parents collect."),
        Col("custody_notes", "Custody / legal notes", 28, "wrap",
            hidden=True,
            note="Court orders and collection restrictions. Sensitive - "
                 "restrict who can see this file."),
        # ---------------------------------------------- their children --
        Col("children", "Children", 9, "int",
            formula=blank(
                "parent1",
                f'COUNTIFS({child_family},{{id}}{{r}})'),
            cf=bar_cf()),
        Col("children_active", "In now", 8, "int",
            formula=blank(
                "parent1",
                f'COUNTIFS({child_family},{{id}}{{r}},'
                f'{child_status},"Active")'),
            note="Children currently attending. A family with none is "
                 "an old family, not a lost one - they may be back."),
        Col("child1", "First child", 18, "text",
            formula=blank("parent1", nth_child(1))),
        Col("child2", "Second child", 18, "text",
            formula=blank("parent1", nth_child(2))),
        Col("child3", "Third child", 18, "text",
            formula=blank("parent1", nth_child(3)), hidden=True),
        Col("more_children", "And others", 11, "text",
            formula=blank(
                "parent1",
                'IF({children}{r}>3,"+"&({children}{r}-3)&" more","")'),
            hidden=True),
        Col("weekly_fees", "Fees a week", 12, "money",
            formula=blank(
                "parent1",
                'ROUND(SUMIFS([[Child Enrollment.weekly_net]],'
                f'{child_family},{{id}}{{r}},'
                '[[Child Enrollment.holds_place]],"Yes"),2)'),
            cf=bar_cf(Palette.accent),
            note="Everything this family is contracted to pay in a "
                 "normal week, across all their children."),
        Col("monthly_fees", "Fees a month", 12, "money",
            formula=blank("parent1",
                          'ROUND({weekly_fees}{r}*52/12,2)'),
            hidden=True),
        # -------------------------------------------------- the money ---
        Col("billed", "Billed", 12, "money",
            formula=blank(
                "parent1",
                f'ROUND(SUMIFS([[Invoices & Receipts.owed]],'
                f'{led_family},{{id}}{{r}}),2)'),
            hidden=True,
            note="Net effect of every invoice, credit note, write-off "
                 "and refund on this family's account."),
        Col("invoiced", "Invoiced", 12, "money",
            formula=blank(
                "parent1",
                f'ROUND(SUMIFS([[Invoices & Receipts.total]],'
                f'{led_family},{{id}}{{r}},{led_type},"Invoice"),2)')),
        Col("paid", "Paid", 12, "money",
            formula=blank(
                "parent1",
                f'ROUND(SUMIFS([[Invoices & Receipts.collected]],'
                f'{led_family},{{id}}{{r}}),2)'),
            note="Payments received less any refunds given back."),
        Col("balance", "Balance", 12, "money",
            formula=blank(
                "parent1",
                'ROUND({billed}{r}-{paid}{r},2)'),
            cf=owing_cf(),
            note="Positive means the family owes you. Negative means "
                 "they are in credit - usually a deposit waiting to be "
                 "used."),
        Col("overdue", "Overdue", 12, "money",
            formula=blank(
                "parent1",
                'ROUND(SUMIFS([[Invoices & Receipts.balance]],'
                f'{led_family},{{id}}{{r}},'
                '[[Invoices & Receipts.status]],"Overdue"),2)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0.004,
                 "fmt": "bold_bad"}],
            note="Only the invoices that are past their due date. "
                 "This is the chase list."),
        Col("oldest_due", "Oldest unpaid", 12, "date",
            # the earliest due date still unpaid - the one that has
            # been outstanding longest. Reads the numeric mirror of
            # the due date, never the dated column, which is blank on
            # payments and on empty rows.
            formula=blank(
                "parent1",
                'IF({overdue}{r}<=0.004,"",IF(' +
                _oldest_due +
                '>=2958465,"",' + _oldest_due + '))'),
            hidden=True),
        Col("last_paid", "Last payment", 12, "date",
            formula=blank(
                "parent1",
                'IF({paid}{r}=0,"",' +
                max_if("[[Invoices & Receipts.date]]",
                       f"{led_family}={{id}}{{r}}",
                       f'{led_type}="Payment"') + ")")),
        Col("days_since_paid", "Days since", 9, "int",
            formula=blank(
                "parent1",
                'IF({last_paid}{r}="","",TodayDate-{last_paid}{r})'),
            hidden=True),
        Col("subsidy", "On subsidy?", 11, "flag",
            formula=blank(
                "parent1",
                'IF(SUMIFS([[Child Enrollment.subsidy_week]],'
                f'{child_family},{{id}}{{r}})>0,"Yes","No")'),
            hidden=True),
        # ------------------------------------------------ relationship --
        Col("joined", "With us since", 12, "date",
            validate=date_check(),
            note="Leave blank and nothing breaks - it is here for the "
                 "families who have been with you for years."),
        Col("messages", "Messages logged", 12, "int",
            formula=blank(
                "parent1",
                'COUNTIFS([[Parent Communication.family]],{id}{r})'),
            hidden=True),
        Col("open_items", "Open items", 10, "int",
            formula=blank(
                "parent1",
                'COUNTIFS([[Parent Communication.family]],{id}{r},'
                '[[Parent Communication.status]],"Open")+'
                'COUNTIFS([[Parent Communication.family]],{id}{r},'
                '[[Parent Communication.status]],"Waiting on us")'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}],
            note="Conversations waiting on you. Not on the parent - "
                 "on you."),
        Col("status", "Family status", 13, "status",
            formula=blank(
                "parent1",
                'IF({children}{r}=0,"No children linked",'
                'IF({overdue}{r}>0.004,"Overdue",'
                'IF({children_active}{r}=0,"Inactive",'
                'IF({balance}{r}>0.004,"Owing","Active"))))'),
            cf=status_cf()),
        Col("notes", "Notes", 28, "wrap"),
        rank_desc("weekly_fees", guard='{parent1}{r}=""',
                  header="Value rank"),
        Col("alert", "Action", 32, "status",
            formula=blank(
                "parent1",
                'IF(COUNTIF({parent1_first}:{parent1_last},'
                '{parent1}{r})>1,"Duplicate - same contact twice",'
                'IF({children}{r}=0,"Add - no children linked yet",'
                'IF({overdue}{r}>0.004,"Overdue - '
                + str(cfg.currency_symbol) +
                '"&TEXT({overdue}{r},"#,##0.00")&" past due",'
                'IF({phone1}{r}="","Missing - no phone number",'
                'IF({em1_name}{r}="","Missing - no emergency contact",'
                'IF({pickup_ok}{r}="No","Not authorised - pickup list '
                'unsigned",'
                'IF({open_items}{r}>0,"Follow up - "&{open_items}{r}&'
                '" message(s) waiting on you",'
                'IF({balance}{r}>0.004,"Chase - balance outstanding",'
                '"OK"))))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Parent & Guardian Database",
        subtitle="One row per family. Siblings share it, invoices "
                 "follow it, and the balance is the family's, not the "
                 "child's.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=3,
        samples=sample(cfg).parents,
        help_text="Type the family's details once. Link each child to "
                  "the family ID on the Child Enrollment tab and the "
                  "children, the weekly fees and the balance appear "
                  "here on their own.",
        extra_notes=[
            "Children, fees, invoices, payments, balance and overdue "
            "amounts are all calculated. The only money column you "
            "ever type on this tab is nothing at all - it comes from "
            "Invoices & Receipts.",
            "Emergency contacts and the authorised collection list "
            "are the two columns worth auditing termly. The "
            "attendance tab checks collections against them.",
            "Custody notes and anything else sensitive should be a "
            "pointer to a secure file, not the detail itself.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Family book summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "The family book", last)
    sheet.table_head(row, 0, [
        "Families", "With a child attending", "Siblings families",
        "Weekly contracted fees", "Total invoiced", "Total paid",
        "Outstanding", "Overdue", "Families owing", "On subsidy",
        "Missing emergency contact"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Parent Database.parent1]],"<>")',
        '=COUNTIFS([[Parent Database.children_active]],">0")',
        '=COUNTIFS([[Parent Database.children]],">1")',
        '=ROUND(SUM([[Parent Database.weekly_fees]]),2)',
        '=ROUND(SUM([[Parent Database.invoiced]]),2)',
        '=ROUND(SUM([[Parent Database.paid]]),2)',
        '=ROUND(SUM([[Parent Database.balance]]),2)',
        '=ROUND(SUM([[Parent Database.overdue]]),2)',
        '=COUNTIFS([[Parent Database.balance]],">0.004")',
        '=COUNTIFS([[Parent Database.subsidy]],"Yes")',
        '=COUNTIFS([[Parent Database.parent1]],"<>")-'
        'COUNTIFS([[Parent Database.em1_name]],"<>")',
    ], ["int", "int", "int", "money", "money", "money", "money",
        "money", "int", "int", "int"])
````

### `daycare_business_spreadsheet/sheets/waitlist.py`

*280 lines · 2 functions*

> `Waitlist` - the enquiries you have not turned into enrollments yet.
>
> A childcare waitlist is a sales pipeline with nappies.  Somebody calls,
> you send information, they tour, you offer, they accept or they do not.
> Every stage has a date, every stage has a drop-out rate, and the
> difference between a setting that fills its rooms and one that does not
> is usually nothing more than following the list up.
>
> The position number ranks the people still waiting by priority first
> and then by how long they have waited, which is how most settings say
> they allocate places - and now they can prove it.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Pipeline summary |

**Module constants:** `NAME` ('Waitlist'), `CLOSED` (4 items).

````python
"""`Waitlist` - the enquiries you have not turned into enrollments yet.

A childcare waitlist is a sales pipeline with nappies.  Somebody calls,
you send information, they tour, you offer, they accept or they do not.
Every stage has a date, every stage has a drop-out rate, and the
difference between a setting that fills its rooms and one that does not
is usually nothing more than following the list up.

The position number ranks the people still waiting by priority first
and then by how long they have waited, which is how most settings say
they allocate places - and now they can prove it.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (age_months, age_years_text, alert_cf, blank,
                     id_col, status_cf)

NAME = "Waitlist"

#: Statuses that mean the enquiry is finished with, one way or another.
CLOSED = ["Enrolled", "Offer declined", "Lost", "Withdrawn"]


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("child_name", "WaitlistPrefix", digits=4, width=10,
               header="Enquiry ID"),
        Col("enquiry_date", "Enquiry date", 12, "date",
            validate=date_check(),
            note="The day they first got in touch. The waiting time "
                 "and the position in the queue both measure from "
                 "here."),
        Col("child_name", "Child's name", 20, "text"),
        Col("dob", "Date of birth", 12, "date", validate=date_check(),
            note="Needed to work out which room they would go into, "
                 "and when."),
        Col("age_months", "Age (months)", 10, "int",
            formula="=" + age_months("{dob}{r}"), hidden=True),
        Col("age", "Age now", 9, "text",
            formula="=" + age_years_text("{dob}{r}")),
        Col("age_at_start", "Age at start", 11, "text",
            formula=blank(
                "start_wanted",
                age_years_text("{dob}{r}", "{start_wanted}{r}")),
            hidden=True),
        Col("band_wanted", "Age band then", 13, "text",
            formula=blank(
                "dob",
                'IFERROR(INDEX(AgeBandList,MATCH('
                'DATEDIF({dob}{r},MAX({dob}{r},'
                'IF({start_wanted}{r}="",TodayDate,'
                '{start_wanted}{r})),"m"),AgeBandFrom,1))&"","")'),
            note="The band the child will be in on the day they want "
                 "to start - which is not always the band they are in "
                 "today."),
        Col("parent_name", "Parent / guardian", 20, "text"),
        Col("phone", "Phone", 15, "text"),
        Col("email", "Email", 24, "text"),
        Col("source", "Heard about you via", 18, "text",
            validate=list_from("InquirySourceList"),
            note="The single most useful marketing column in the "
                 "workbook. Fill it in every time."),
        # -------------------------------------------------- what they want
        Col("program", "Programme wanted", 15, "text",
            validate=list_from("ProgramList")),
        Col("room", "Room wanted", 16, "text",
            validate=list_from("ClassroomList"),
            note="Pick a room and the workbook tells you whether there "
                 "is a seat in it right now."),
        Col("days_wanted", "Days a week", 10, "number1",
            validate=number_check(0, 7)),
        Col("start_wanted", "Wants to start", 12, "date",
            validate=date_check()),
        Col("priority", "Priority", 13, "text",
            validate=list_from("WaitlistPriorityList"),
            note="Siblings and staff children usually come first. "
                 "Whatever your policy is, record it here and the "
                 "queue position follows it."),
        Col("priority_rank", "Priority order", 10, "int",
            formula=blank(
                "child_name",
                'IF({priority}{r}="Sibling",1,'
                'IF({priority}{r}="Staff child",2,'
                'IF({priority}{r}="Returning family",3,'
                'IF({priority}{r}="Low",5,4))))'),
            hidden=True),
        Col("sibling", "Sibling already here", 11, "id",
            validate=list_from("ChildList"),
            note="The child ID of a brother or sister already with "
                 "you. Links the enquiry to the family you know."),
        Col("sibling_name", "Sibling", 16, "text",
            formula=blank(
                "sibling",
                'IFERROR(INDEX([[Child Enrollment.name]],'
                'MATCH({sibling}{r},[[Child Enrollment.id]],0))&"","")'),
            hidden=True),
        # ------------------------------------------------------ pipeline
        Col("status", "Status", 15, "status",
            validate=list_from("WaitlistStatusList"), cf=status_cf()),
        Col("info_sent", "Info sent", 11, "date", validate=date_check(),
            hidden=True),
        Col("tour_date", "Tour", 11, "date", validate=date_check(),
            note="Settings that get families through the door convert "
                 "two to three times better. Book the tour."),
        Col("offered_date", "Offered", 11, "date",
            validate=date_check()),
        Col("response_date", "They replied", 11, "date",
            validate=date_check(), hidden=True),
        Col("enrolled_child", "Became child", 11, "id",
            validate=list_from("ChildList"),
            note="Once they enroll, put the new child ID here. It "
                 "closes the loop between marketing and revenue."),
        Col("closed", "Closed?", 9, "flag",
            formula=blank(
                "child_name",
                'IF(' + "+".join(f'({{status}}{{r}}="{s}")'
                                 for s in CLOSED) + '>0,"Yes","No")'),
            hidden=True),
        Col("converted", "Enrolled?", 10, "flag",
            formula=blank(
                "child_name",
                'IF({status}{r}="Enrolled","Yes",'
                'IF({enrolled_child}{r}<>"","Yes","No"))'),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "good"}]),
        Col("days_waiting", "Days waiting", 11, "int",
            formula=blank(
                "enquiry_date",
                'IF({closed}{r}="Yes",'
                'IF({response_date}{r}<>"",'
                '{response_date}{r}-{enquiry_date}{r},'
                'IF({offered_date}{r}<>"",'
                '{offered_date}{r}-{enquiry_date}{r},'
                'TodayDate-{enquiry_date}{r})),'
                'TodayDate-{enquiry_date}{r})'),
            cf=[{"type": "cell", "criteria": ">", "value": 90,
                 "fmt": "warn"}]),
        Col("days_quiet", "Days since contact", 12, "int",
            formula=blank(
                "enquiry_date",
                'IF({closed}{r}="Yes","",'
                'TodayDate-MAX({enquiry_date}{r},{info_sent}{r},'
                '{tour_date}{r},{offered_date}{r},'
                '{response_date}{r}))'),
            cf=[{"type": "cell", "criteria": ">=", "value": 14,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">=", "value": 7,
                 "fmt": "warn"}],
            note="How long since anything happened on this enquiry. "
                 "Over a week and it is going cold."),
        Col("position", "Queue position", 12, "int",
            formula=blank(
                "child_name",
                'IF({closed}{r}="Yes","",'
                'COUNTIFS({closed_first}:{closed_last},"No",'
                '{room_first}:{room_last},{room}{r},'
                '{priority_rank_first}:{priority_rank_last},'
                '"<"&{priority_rank}{r})+'
                'COUNTIFS({closed_first}:{closed_last},"No",'
                '{room_first}:{room_last},{room}{r},'
                '{priority_rank_first}:{priority_rank_last},'
                '{priority_rank}{r},'
                '{enquiry_date_first}:{enquiry_date_last},'
                '"<="&{enquiry_date}{r}))'),
            note="Where they stand in the queue for the room they "
                 "want: priority first, then how long they have "
                 "waited. Nothing to type."),
        Col("seats_now", "Seats free now", 11, "int",
            formula=blank(
                "room",
                'IFERROR(INDEX([[Classroom Management.seats]],'
                'MATCH({room}{r},[[Classroom Management.name]],0))*1,'
                '0)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "good"}],
            note="Live, from the classroom tab. A free seat and "
                 "somebody waiting for it is money sitting on the "
                 "table."),
        Col("value_week", "Worth a week", 12, "money",
            formula=blank(
                "days_wanted",
                'ROUND(IF({days_wanted}{r}>=5,FullTimeWeek,'
                '{days_wanted}{r}*PartTimeDay),2)'),
            note="Rough value at your standard rates, so you can see "
                 "what the waiting list is worth."),
        Col("followed_up", "Last follow-up", 12, "date",
            validate=date_check(), hidden=True),
        Col("notes", "Notes", 30, "wrap"),
        Col("alert", "Action", 32, "status",
            formula=blank(
                "child_name",
                'IF({closed}{r}="Yes","",'
                'IF({room}{r}="","Add - which room do they want?",'
                'IF(IF({seats_now}{r}<=0,FALSE,'
                '({status}{r}="Waiting - no space")+'
                '({status}{r}="Toured")+'
                '({status}{r}="New enquiry")+'
                '({status}{r}="Information sent")>0),'
                '"Space - a seat just opened, offer it",'
                'IF({days_quiet}{r}>=14,"Chase - two weeks with no '
                'contact",'
                'IF(IF({offered_date}{r}="",FALSE,'
                'IF({response_date}{r}<>"",FALSE,'
                'TodayDate-{offered_date}{r}>7)),'
                '"Chase - offer made, no reply",'
                'IF({status}{r}="Toured","Follow up - toured, now '
                'offer them a place",'
                'IF({tour_date}{r}="","Confirm - book them a tour",'
                'IF({days_quiet}{r}>=7,"Follow up - going quiet",'
                '"OK"))))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Waitlist & Enquiries",
        subtitle="Every enquiry, where it got to, how long they have "
                 "waited and which of them you can say yes to today.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=3,
        samples=sample(cfg).waitlist,
        help_text="Log every enquiry the day it arrives. The queue "
                  "position, the waiting time, the free seats and the "
                  "follow-up list all calculate themselves.",
        extra_notes=[
            "The queue position only counts enquiries that are still "
            "open and still want the same room, so it is a real "
            "position rather than a row number.",
            "'Seats free now' reads the classroom tab live. When it "
            "goes above zero the Action column tells you to offer the "
            "place - that is the one alert on this tab that makes "
            "money.",
            "Fill in 'Heard about you via' every single time. After "
            "thirty enquiries it tells you which half of your "
            "marketing budget to stop spending.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Pipeline summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "The pipeline", last)
    sheet.table_head(row, 0, [
        "Enquiries logged", "Still open", "Tours booked", "Toured",
        "Offers out", "Enrolled", "Lost or declined",
        "Enquiry to enrollment", "Tour to enrollment",
        "Average days waiting", "Waiting list worth a week"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Waitlist.child_name]],"<>")',
        '=COUNTIFS([[Waitlist.closed]],"No")',
        '=COUNTIFS([[Waitlist.tour_date]],"<>")',
        '=COUNTIFS([[Waitlist.status]],"Toured")',
        '=COUNTIFS([[Waitlist.offered_date]],"<>")',
        '=COUNTIFS([[Waitlist.converted]],"Yes")',
        '=COUNTIFS([[Waitlist.status]],"Offer declined")+'
        'COUNTIFS([[Waitlist.status]],"Lost")+'
        'COUNTIFS([[Waitlist.status]],"Withdrawn")',
        '=IF(COUNTIF([[Waitlist.child_name]],"<>")<=0,"",'
        'COUNTIFS([[Waitlist.converted]],"Yes")/'
        'COUNTIF([[Waitlist.child_name]],"<>"))',
        '=IF(COUNTIFS([[Waitlist.tour_date]],"<>")<=0,"",'
        'COUNTIFS([[Waitlist.tour_date]],"<>",'
        '[[Waitlist.converted]],"Yes")/'
        'COUNTIFS([[Waitlist.tour_date]],"<>"))',
        '=IFERROR(ROUND(AVERAGEIF([[Waitlist.days_waiting]],">=0"),0),'
        '0)',
        '=ROUND(SUMIFS([[Waitlist.value_week]],'
        '[[Waitlist.closed]],"No"),2)',
    ], ["int", "int", "int", "int", "int", "int", "int", "percent",
        "percent", "int", "money"])
````

---

## 4. The register

Check in, check out, who collected them and whether that person was on the authorised list. Late arrivals, early pickups, missing check-outs and the late collection charge all fall out of the two time columns.

### `daycare_business_spreadsheet/sheets/attendance.py`

*297 lines · 2 functions*

> `Attendance` - the daily sign in and sign out sheet.
>
> One row per child per day.  Everything else follows from it: the
> attendance percentage on the enrollment tab, the ratio figures on the
> classroom tab, the late pickup charges on the fee tab, the "who is in
> today" tile on the dashboard and every monthly report in the file.
>
> Three things this tab does that a paper sign-in sheet cannot.  It
> measures late arrivals and late collections against each child's own
> contracted times rather than against the centre's opening hours.  It
> checks the name of whoever collected the child against the authorised
> list and says so when it cannot find them.  And it refuses to let a
> missing check-out time quietly disappear.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Attendance summary |

**Module constants:** `NAME` ('Attendance').

````python
"""`Attendance` - the daily sign in and sign out sheet.

One row per child per day.  Everything else follows from it: the
attendance percentage on the enrollment tab, the ratio figures on the
classroom tab, the late pickup charges on the fee tab, the "who is in
today" tile on the dashboard and every monthly report in the file.

Three things this tab does that a paper sign-in sheet cannot.  It
measures late arrivals and late collections against each child's own
contracted times rather than against the centre's opening hours.  It
checks the name of whoever collected the child against the authorised
list and says so when it cannot find them.  And it refuses to let a
missing check-out time quietly disappear.
"""

from __future__ import annotations

from ..config import (ABSENT_STATUSES, NON_SESSION_STATUSES,
                      PRESENT_STATUSES, Config, Palette)
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, time_check
from .common import (alert_cf, blank, id_col, late_pickup_charge,
                     lookup_child, lookup_staff, month_col, status_cf,
                     week_start_col, weekday_col)

NAME = "Attendance"


def spec(cfg: Config) -> SheetSpec:
    present_test = "+".join(f'({{status}}{{r}}="{s}")'
                            for s in PRESENT_STATUSES)
    absent_test = "+".join(f'({{status}}{{r}}="{s}")'
                           for s in ABSENT_STATUSES)
    nonsession_test = "+".join(f'({{status}}{{r}}="{s}")'
                               for s in NON_SESSION_STATUSES)

    columns = [
        id_col("date", "AttendancePrefix", digits=5, width=11,
               header="Entry ID"),
        Col("date", "Date", 11, "date", validate=date_check(),
            note="One row per child per day. Fill a whole day down the "
                 "sheet in the morning and finish the times as people "
                 "arrive."),
        weekday_col("date"),
        month_col("date"),
        week_start_col("date"),
        Col("child", "Child", 11, "id", validate=list_from("ChildList"),
            note="Pick the child ID. Everything else about them fills "
                 "in automatically."),
        Col("child_name", "Child's name", 20, "text",
            formula=lookup_child("name")),
        Col("room", "Room", 15, "text",
            formula=lookup_child("room", missing=""),
            note="From the enrollment tab. The ratio figures on the "
                 "classroom tab count this column."),
        Col("program", "Programme", 13, "text",
            formula=lookup_child("program", missing=""), hidden=True),
        Col("daykey", "Day key", 14, "text",
            formula=blank("child", '{child}{r}&"|"&{date}{r}'),
            hidden=True,
            note="Child and date joined together. The daily report "
                 "and the meal tracker find this row with it."),
        Col("status", "Status", 16, "status",
            validate=list_from("AttendanceStatusList"), cf=status_cf(),
            note="Present, Late and Early pickup all mean the child "
                 "was here. Public holiday, Centre closed, Snow day "
                 "and Not scheduled mean nobody was expected, so they "
                 "do not count against the attendance percentage."),
        Col("session", "Session?", 8, "int",
            formula=blank(
                "status",
                f'IF(({nonsession_test})>0,0,1)'),
            hidden=True,
            note="1 when a session was expected. The denominator of "
                 "every attendance percentage in the workbook."),
        Col("present", "Present?", 8, "int",
            formula=blank("status", f'IF(({present_test})>0,1,0)'),
            hidden=True),
        Col("absent", "Absent?", 8, "int",
            formula=blank("status", f'IF(({absent_test})>0,1,0)'),
            hidden=True),
        # ----------------------------------------------------- times ----
        Col("in_time", "Checked in", 11, "time", validate=time_check(),
            note="The time they actually arrived. Type 8:15 AM or "
                 "08:15 - both work."),
        Col("out_time", "Checked out", 11, "time",
            validate=time_check(),
            note="Leave blank and the Action column will not let you "
                 "forget. A sign-in sheet with no sign-out is the one "
                 "thing an inspector always finds."),
        Col("hours", "Hours", 8, "number1",
            formula=blank(
                "in_time",
                'IF({out_time}{r}="","",'
                'ROUND((MOD({out_time}{r}-{in_time}{r},1))*24,2))'),
            note="Time in the building. Hourly fees and the "
                 "'contracted versus actual' comparison use it."),
        Col("due_in", "Due in", 10, "time",
            formula=blank(
                "child",
                'IFERROR(INDEX([[Child Enrollment.usual_in]],'
                'MATCH({child}{r},[[Child Enrollment.id]],0))*1,"")'),
            hidden=True),
        Col("due_out", "Due out", 10, "time",
            formula=blank(
                "child",
                'IFERROR(INDEX([[Child Enrollment.usual_out]],'
                'MATCH({child}{r},[[Child Enrollment.id]],0))*1,"")'),
            hidden=True),
        Col("late_in", "Late in (min)", 11, "int",
            formula=blank(
                "in_time",
                'IF({due_in}{r}="","",'
                'MAX(0,ROUND(({in_time}{r}-{due_in}{r})*1440,0)))'),
            cf=[{"type": "cell", "criteria": ">=", "value": 30,
                 "fmt": "warn"}],
            note="Minutes after this child's own contracted drop-off "
                 "time, not after opening time."),
        Col("early_out", "Left early (min)", 12, "int",
            formula=blank(
                "out_time",
                'IF({due_out}{r}="","",'
                'MAX(0,ROUND(({due_out}{r}-{out_time}{r})*1440,0)))'),
            hidden=True),
        Col("late_out", "Late pickup (min)", 13, "int",
            formula=blank(
                "out_time",
                'IF({due_out}{r}="","",'
                'MAX(0,ROUND(({out_time}{r}-{due_out}{r})*1440,0)))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}],
            note="Minutes after the contracted collection time. The "
                 "charge below is worked out from it."),
        Col("late_fee", "Late fee due", 11, "money",
            formula=blank(
                "late_out",
                'IF({late_out}{r}<=0,0,' +
                late_pickup_charge("{late_out}{r}") + ")"),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}],
            note="Free for the grace minutes, then the first block, "
                 "then per minute - all three set on Settings. It is "
                 "not charged until you raise it on the Tuition & "
                 "Fees tab."),
        Col("fee_raised", "Fee charged?", 11, "flag",
            validate=list_from("YesNoList"),
            note="Say Yes once you have added the charge to Tuition & "
                 "Fees, so the Action column stops reminding you."),
        # ------------------------------------------------ who and why ---
        Col("dropped_by", "Dropped off by", 18, "text"),
        Col("collected_by", "Collected by", 18, "text",
            note="Type the name of whoever actually took the child. "
                 "It is checked against the authorised list on the "
                 "child's and the family's record."),
        Col("auth_list", "Authorised list", 22, "text",
            formula=blank(
                "child",
                'IFERROR(INDEX([[Child Enrollment.pickup_list]],'
                'MATCH({child}{r},[[Child Enrollment.id]],0))&"","")&'
                '" "&IFERROR(INDEX([[Parent Database.pickup_list]],'
                'MATCH(IFERROR(INDEX([[Child Enrollment.family]],'
                'MATCH({child}{r},[[Child Enrollment.id]],0))&"",""),'
                '[[Parent Database.id]],0))&"","")&" "&'
                'IFERROR(INDEX([[Parent Database.parent1]],'
                'MATCH(IFERROR(INDEX([[Child Enrollment.family]],'
                'MATCH({child}{r},[[Child Enrollment.id]],0))&"",""),'
                '[[Parent Database.id]],0))&"","")&" "&'
                'IFERROR(INDEX([[Parent Database.parent2]],'
                'MATCH(IFERROR(INDEX([[Child Enrollment.family]],'
                'MATCH({child}{r},[[Child Enrollment.id]],0))&"",""),'
                '[[Parent Database.id]],0))&"","")'),
            hidden=True,
            note="The child's own list, the family's list and both "
                 "parents' names, joined together so the check below "
                 "can search it."),
        Col("pickup_ok", "Authorised?", 11, "flag",
            formula=blank(
                "collected_by",
                'IF({auth_list}{r}="","Not checked",'
                'IF(ISNUMBER(SEARCH({collected_by}{r},'
                '{auth_list}{r})),"Yes","No"))'),
            cf=[{"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "bold_bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "good"}],
            note="Searches the authorised names for whoever you typed. "
                 "'No' does not mean stop - it means check, and that "
                 "is exactly what it is for."),
        Col("reason", "Absence reason", 22, "text",
            note="Why they are not in. 'Unexplained' three days "
                 "running is a safeguarding prompt, not an admin "
                 "chore."),
        Col("recorded_by", "Recorded by", 11, "id",
            validate=list_from("StaffList")),
        Col("recorded_name", "Staff", 16, "text",
            formula=lookup_staff("name", "recorded_by"), hidden=True),
        Col("temp_check", "Temp / health check", 12, "text",
            hidden=True,
            note="Only if your setting does them. Left blank it costs "
                 "you nothing."),
        Col("notes", "Notes", 26, "wrap"),
        Col("run_absent", "Absent in a row", 11, "int",
            formula=blank(
                "child",
                'IF({absent}{r}=0,0,'
                'COUNTIFS({child_first}:{child}{r},{child}{r},'
                '{absent_first}:{absent}{r},1,'
                '{date_first}:{date}{r},">"&({date}{r}-'
                'AbsenceChaseDays)))'),
            hidden=True,
            note="Absences by this child inside the chase window set "
                 "on Settings."),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "date",
                'IF({child}{r}="","Add - no child picked",'
                'IF({child_name}{r}="Not in Child Enrollment",'
                '"Check - child ID not found",'
                'IF({status}{r}="","Add - no status yet",'
                'IF(IF({present}{r}=0,FALSE,{in_time}{r}=""),'
                '"Missing - no check-in time",'
                'IF(IF({present}{r}=0,FALSE,'
                'IF({out_time}{r}<>"",FALSE,{date}{r}<TodayDate)),'
                '"No check-out - find out who took them",'
                'IF({pickup_ok}{r}="No","Not authorised - check who '
                'collected",'
                'IF(IF({late_fee}{r}<=0,FALSE,{fee_raised}{r}<>"Yes"),'
                '"Invoice - late pickup fee not charged",'
                'IF(IF({absent}{r}=0,FALSE,{reason}{r}=""),'
                '"Unexplained - call the parent",'
                'IF({run_absent}{r}>=AbsenceChaseDays,'
                '"Chase - absent several days running",'
                'IF({status}{r}="Late","Late - "&{late_in}{r}&'
                '" minutes","OK"))))))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Daily Attendance, Check-In & Check-Out",
        subtitle="Who was in, when they arrived, when they left, who "
                 "collected them - and what that late pickup is worth.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.deep,
        freeze_cols=2,
        samples=sample(cfg).attendance,
        help_text="One row per child per day. Pick the date, the "
                  "child and the status, then fill the times in as "
                  "they happen. Everything to the right is worked out "
                  "for you.",
        extra_notes=[
            "Late arrivals and late collections are measured against "
            "each child's own contracted times on the enrollment tab, "
            "so a child contracted until 3pm is late at 3:10 even "
            "though you are open until 6.",
            "The authorised check searches the child's pickup list, "
            "the family's pickup list and both parents' names. It is "
            "a prompt to look, never a decision - the adult in the "
            "room makes that.",
            "Statuses that mean nobody was expected (Public holiday, "
            "Centre closed, Snow day, Not scheduled) are left out of "
            "the attendance percentage, which is the only way to get "
            "a number that is actually fair to the family.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Attendance summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Attendance summary", last)
    sheet.table_head(row, 0, [
        "Entries logged", "Sessions expected", "Attended", "Absences",
        "Attendance rate", "In today", "Absent today", "Late arrivals",
        "Late pickups", "Late fees earned", "Late fees not charged",
        "Missing check-outs"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Attendance.date]],"<>")',
        '=SUM([[Attendance.session]])',
        '=SUM([[Attendance.present]])',
        '=SUM([[Attendance.absent]])',
        '=IF(SUM([[Attendance.session]])<=0,"",'
        'SUM([[Attendance.present]])/SUM([[Attendance.session]]))',
        '=SUMIFS([[Attendance.present]],[[Attendance.date]],TodayDate)',
        '=SUMIFS([[Attendance.absent]],[[Attendance.date]],TodayDate)',
        '=COUNTIFS([[Attendance.status]],"Late")',
        '=COUNTIFS([[Attendance.late_out]],">0")',
        '=ROUND(SUM([[Attendance.late_fee]]),2)',
        '=ROUND(SUMIFS([[Attendance.late_fee]],'
        '[[Attendance.fee_raised]],"<>Yes"),2)',
        '=COUNTIFS([[Attendance.alert]],"No check-out*")',
    ], ["int", "int", "int", "int", "percent", "int", "int", "int",
        "int", "money", "money", "int"])
````

---

## 5. The money

Charges first, invoices second. Every fee is a line on Tuition & Fees - tuition, registration, meals, transport, late pickup, discounts and subsidies - and an invoice is simply the sum of the lines pointed at it. Payments, credit notes, refunds and write-offs sit in the same ledger, so billed and collected can never drift apart.

### `daycare_business_spreadsheet/sheets/tuition.py`

*221 lines · 2 functions*

> `Tuition & Fees` - every charge and every credit, one row at a time.
>
> This is the charge register, not the invoice.  Tuition for a week, a
> registration fee, a late pickup, a field trip, the sibling discount, a
> subsidy credit - each one is a line here, and each one carries the
> invoice number it was billed on.  Invoices & Receipts then adds the
> lines up.
>
> Keeping charges and invoices apart is what makes partial billing,
> mid-month changes and credits possible without any of the arithmetic
> going wrong.  Discounts and subsidies are entered as positive numbers
> and subtracted, so a parent never has to read a minus sign.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Charge register summary |

**Module constants:** `NAME` ('Tuition & Fees').

````python
"""`Tuition & Fees` - every charge and every credit, one row at a time.

This is the charge register, not the invoice.  Tuition for a week, a
registration fee, a late pickup, a field trip, the sibling discount, a
subsidy credit - each one is a line here, and each one carries the
invoice number it was billed on.  Invoices & Receipts then adds the
lines up.

Keeping charges and invoices apart is what makes partial billing,
mid-month changes and credits possible without any of the arithmetic
going wrong.  Discounts and subsidies are entered as positive numbers
and subtracted, so a parent never has to read a minus sign.
"""

from __future__ import annotations

from ..config import CREDIT_FEE_TYPES, Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, id_col, lookup_child,
                     month_col, seq_key, status_cf, tax_on)

NAME = "Tuition & Fees"


def spec(cfg: Config) -> SheetSpec:
    credit_test = "+".join(f'({{fee_type}}{{r}}="{t}")'
                           for t in CREDIT_FEE_TYPES)

    columns = [
        id_col("date", "FeePrefix", digits=5, width=11,
               header="Charge ID"),
        Col("date", "Date", 11, "date", validate=date_check(),
            note="The date the charge belongs to. Monthly reports "
                 "group on it."),
        month_col("date"),
        Col("child", "Child", 11, "id", validate=list_from("ChildList"),
            note="Which child the charge is for. Leave blank only for "
                 "a family-level charge such as a registration fee "
                 "covering siblings."),
        Col("child_name", "Child's name", 18, "text",
            formula=lookup_child("name")),
        Col("family", "Family", 11, "id",
            formula=blank(
                "child",
                'IFERROR(INDEX([[Child Enrollment.family]],'
                'MATCH({child}{r},[[Child Enrollment.id]],0))&"","")'),
            note="From the child's record. The invoice goes to this "
                 "family."),
        Col("family_set", "Family override", 11, "id",
            validate=list_from("FamilyList"), hidden=True,
            note="Only needed for a charge that has no child against "
                 "it at all."),
        Col("family_used", "Billed to", 11, "id",
            formula=blank(
                "date",
                'IF({family}{r}<>"",{family}{r},{family_set}{r})'),
            hidden=True),
        Col("fee_type", "Charge type", 17, "text",
            validate=list_from("FeeTypeList"), cf=status_cf(),
            note="Discounts, subsidies and credits are in this list "
                 "too. Enter them as a positive amount - the workbook "
                 "knows to take them off."),
        Col("is_credit", "Credit?", 8, "flag",
            formula=blank("fee_type",
                          f'IF(({credit_test})>0,"Yes","No")'),
            hidden=True),
        Col("description", "Description", 30, "text",
            note="What a parent would need to read to recognise the "
                 "charge. 'Tuition, week beginning 5 Oct' beats "
                 "'fees'."),
        Col("period_from", "Covers from", 11, "date",
            validate=date_check(), hidden=True),
        Col("period_to", "Covers to", 11, "date",
            validate=date_check(), hidden=True),
        Col("qty", "Quantity", 9, "number1",
            validate=number_check(0, 10000),
            note="Weeks, days, hours, sessions or trips - whatever the "
                 "rate is per. Leave blank and it counts as 1."),
        Col("rate", "Rate", 11, "money",
            validate=number_check(-100000, 100000),
            note="The unit price. For a flat charge, put the whole "
                 "amount here and leave the quantity blank."),
        Col("gross", "Gross", 11, "money",
            formula=blank(
                "date",
                'ROUND(IF({qty}{r}="",1,{qty}{r})*{rate}{r},2)')),
        Col("discount_pct", "Discount %", 10, "percent",
            validate=number_check(0, 1), hidden=True),
        Col("discount", "Discount", 11, "money",
            formula=blank(
                "date",
                'ROUND({gross}{r}*{discount_pct}{r},2)'),
            hidden=True),
        Col("subtotal", "Subtotal", 11, "money",
            formula=blank("date",
                          'ROUND({gross}{r}-{discount}{r},2)')),
        Col("tax", "Tax", 10, "money",
            formula=blank(
                "date",
                tax_on("{subtotal}{r}", '{fee_type}{r}="Tuition"')),
            note="Zero unless you switch tax on for tuition or extras "
                 "on the Settings tab. Childcare is exempt in a great "
                 "many places, so nothing is assumed."),
        Col("net", "Amount", 12, "money",
            formula=blank(
                "date",
                'ROUND(IF({is_credit}{r}="Yes",-1,1)*'
                '({subtotal}{r}+{tax}{r}),2)'),
            cf=bar_cf(Palette.accent),
            note="What this line adds to the bill. Discounts, "
                 "subsidies and credits come out negative even though "
                 "you typed a positive number."),
        # -------------------------------------------------- invoicing --
        Col("invoice", "Invoice", 12, "text",
            validate=list_from("InvoiceList"),
            note="The invoice number this line was billed on. Leave it "
                 "blank until you raise the invoice and the Action "
                 "column will remind you."),
        Col("billed", "Billed?", 9, "flag",
            formula=blank("date",
                          'IF({invoice}{r}="","No","Yes")'),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "good"},
                {"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "warn"}]),
        Col("invoice_ok", "Invoice found?", 11, "flag",
            formula=blank(
                "invoice",
                'IF(COUNTIF([[Invoices & Receipts.id]],'
                '{invoice}{r})>0,"Yes","No")'),
            hidden=True),
        # Line 1, 2, 3 ... within one invoice, and the `INV-00012|3`
        # key the printable invoice uses to pull a line without an
        # array formula.
        *seq_key("invoice"),
        Col("auto_source", "Raised by", 13, "text",
            validate=list_from("ChargeSourceList"), hidden=True,
            note="Where the charge came from - the schedule, a late "
                 "pickup, an extra day. Useful when you are checking "
                 "a bill a parent has queried."),
        Col("notes", "Notes", 24, "wrap"),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "date",
                'IF(IF({child}{r}="",{family_set}{r}="",FALSE),'
                '"Add - no child or family on this charge",'
                'IF({fee_type}{r}="","Add - pick a charge type",'
                'IF({rate}{r}="","Add - no rate entered",'
                'IF(IF({invoice}{r}="",FALSE,{invoice_ok}{r}="No"),'
                '"Check - that invoice number does not exist",'
                'IF(IF({invoice}{r}<>"",FALSE,'
                'TodayDate-{date}{r}>PaymentTermsDays),'
                '"Invoice - charge is '
                '"&(TodayDate-{date}{r})&" days old and unbilled",'
                'IF({invoice}{r}="","Invoice - not billed yet",'
                '"OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Tuition & Fee Charges",
        subtitle="Every charge and every credit, ready to be pulled "
                 "onto an invoice. Discounts and subsidies are typed "
                 "positive and come off automatically.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).fees,
        help_text="Add a line for anything you want to bill. Put the "
                  "invoice number against it once it is billed, and "
                  "Invoices & Receipts adds the lines up for you.",
        extra_notes=[
            "The rate card lives on the Child Enrollment tab: basis, "
            "agreed rate, discount and subsidy per child. This tab is "
            "the actual charges that came out of it.",
            "Late pickup fees are calculated on the Attendance tab "
            "from each child's own contracted collection time. Copy "
            "the amount here to charge it, then mark it as charged "
            "over there.",
            "A charge with no invoice number against it is a charge "
            "nobody has been asked to pay. That is what the Action "
            "column is watching for.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Charge register summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Charges raised", last)
    sheet.table_head(row, 0, [
        "Charge lines", "Gross charges", "Discounts & credits",
        "Tax", "Net charged", "Tuition", "Extras & fees",
        "Billed", "Not yet billed", "This month"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Tuition & Fees.date]],"<>")',
        '=ROUND(SUMIFS([[Tuition & Fees.subtotal]],'
        '[[Tuition & Fees.is_credit]],"No"),2)',
        '=ROUND(SUMIFS([[Tuition & Fees.subtotal]],'
        '[[Tuition & Fees.is_credit]],"Yes"),2)',
        '=ROUND(SUM([[Tuition & Fees.tax]]),2)',
        '=ROUND(SUM([[Tuition & Fees.net]]),2)',
        '=ROUND(SUMIFS([[Tuition & Fees.net]],'
        '[[Tuition & Fees.fee_type]],"Tuition"),2)',
        '=ROUND(SUM([[Tuition & Fees.net]])-'
        'SUMIFS([[Tuition & Fees.net]],'
        '[[Tuition & Fees.fee_type]],"Tuition"),2)',
        '=ROUND(SUMIFS([[Tuition & Fees.net]],'
        '[[Tuition & Fees.billed]],"Yes"),2)',
        '=ROUND(SUMIFS([[Tuition & Fees.net]],'
        '[[Tuition & Fees.billed]],"No"),2)',
        '=ROUND(SUMIFS([[Tuition & Fees.net]],'
        '[[Tuition & Fees.month]],ThisMonthKey),2)',
    ], ["int", "money", "money", "money", "money", "money", "money",
        "money", "money", "money"])
````

### `daycare_business_spreadsheet/sheets/invoices.py`

*311 lines · 2 functions*

> `Invoices & Receipts` - one ledger for everything that moves money.
>
> Every row is one movement on a family's account:
>
>     Invoice      you asked them for money        account goes up
>     Payment      they gave you some              account goes down
>     Credit note  you cancelled part of a bill    account goes down
>     Write-off    you gave up on a bill           account goes down
>     Refund       you gave money back             account goes up
>
> That is the whole model, and it is why partial payments, instalments,
> deposits, credits and refunds all work without a single extra column.
> Pay an invoice in four goes and you get four Payment rows pointing at
> it; the invoice knows how much is left.
>
> Two figures are kept deliberately separate.  **Billed** is revenue -
> what you charged, less credits and write-offs.  **Collected** is cash -
> what actually arrived, less refunds.  They are different numbers, they
> answer different questions, and mixing them up is the single most
> common way a childcare business talks itself into trouble.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Ledger summary |

**Module constants:** `NAME` ('Invoices & Receipts').

````python
"""`Invoices & Receipts` - one ledger for everything that moves money.

Every row is one movement on a family's account:

    Invoice      you asked them for money        account goes up
    Payment      they gave you some              account goes down
    Credit note  you cancelled part of a bill    account goes down
    Write-off    you gave up on a bill           account goes down
    Refund       you gave money back             account goes up

That is the whole model, and it is why partial payments, instalments,
deposits, credits and refunds all work without a single extra column.
Pay an invoice in four goes and you get four Payment rows pointing at
it; the invoice knows how much is left.

Two figures are kept deliberately separate.  **Billed** is revenue -
what you charged, less credits and write-offs.  **Collected** is cash -
what actually arrived, less refunds.  They are different numbers, they
answer different questions, and mixing them up is the single most
common way a childcare business talks itself into trouble.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, card_fee, lookup_family,
                     month_col, owing_cf, status_cf)
from .common import ROW_OFFSET

NAME = "Invoices & Receipts"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        # An invoice gets an invoice number and a receipt gets a
        # receipt number, from the same row - which is exactly what a
        # parent expects to see on the piece of paper.
        Col("id", "Reference", 12, "id",
            formula=('=IF({date}{r}="","",'
                     'IF({type}{r}="","",'
                     'IF(({type}{r}="Payment")+({type}{r}="Refund")>0,'
                     'ReceiptPrefix,InvoicePrefix)&'
                     f'TEXT(ROW()-{ROW_OFFSET},"00000")))'),
            note="Automatic. Invoice rows get an invoice number, "
                 "payment and refund rows get a receipt number. Both "
                 "come from the row, so they never collide."),
        Col("date", "Date", 11, "date", validate=date_check()),
        month_col("date"),
        Col("type", "Type", 12, "status",
            validate=list_from("LedgerTypeList"), cf=status_cf(),
            note="Invoice, Payment, Credit note, Refund or Write-off. "
                 "This one cell decides what the row does to the "
                 "family's balance, to revenue and to cash."),
        Col("family", "Family", 11, "id",
            validate=list_from("FamilyList"),
            note="Whose account this touches. Invoices go to the "
                 "family, not the child, so siblings share one bill."),
        Col("family_name", "Family name", 18, "text",
            formula=lookup_family("family_name")),
        Col("child", "Child (optional)", 11, "id",
            validate=list_from("ChildList"), hidden=True,
            note="Only when you want one child's name on the "
                 "document. The money still belongs to the family."),
        Col("period", "Period", 12, "text",
            note="What the invoice covers in plain words - 'October "
                 "2026' or 'w/c 12 Oct'. Printed on the invoice."),
        Col("description", "Description", 30, "text"),
        # --------------------------------------------------- amounts ----
        Col("charges", "From charge lines", 14, "money",
            formula=blank(
                "date",
                'IF({type}{r}<>"Invoice",0,'
                'ROUND(SUMIFS([[Tuition & Fees.net]],'
                '[[Tuition & Fees.invoice]],{id}{r}),2))'),
            note="Everything on the Tuition & Fees tab tagged with "
                 "this invoice number. Add the lines there and this "
                 "fills itself in."),
        Col("lines", "Lines", 7, "int",
            formula=blank(
                "date",
                'IF({type}{r}<>"Invoice","",'
                'COUNTIFS([[Tuition & Fees.invoice]],{id}{r}))'),
            hidden=True),
        Col("amount", "Amount typed", 12, "money",
            validate=number_check(0, 1000000),
            note="For payments, refunds, credit notes and write-offs, "
                 "this is the amount. For an invoice, leave it blank "
                 "if you are using charge lines - type a figure only "
                 "for a one-off bill with no lines behind it."),
        Col("total", "Total", 12, "money",
            formula=blank(
                "date",
                'ROUND(IF({type}{r}="Invoice",'
                'IF({lines}{r}>0,{charges}{r},{amount}{r}),'
                '{amount}{r}),2)'),
            cf=bar_cf(Palette.accent),
            note="The figure on the document. An invoice uses its "
                 "charge lines when it has any, and the typed amount "
                 "when it does not."),
        # ------------------------------------------------- settlement ---
        Col("against", "Pays invoice", 12, "text",
            validate=list_from("InvoiceList"),
            note="On a payment, refund, credit note or write-off, the "
                 "invoice number it applies to. A payment with nothing "
                 "here still reduces the family's balance - it just "
                 "is not matched to a bill."),
        Col("method", "Method", 15, "text",
            validate=list_from("PaymentMethodList")),
        Col("reference", "Reference", 14, "text", hidden=True,
            note="Cheque number, transaction ID, the reference on the "
                 "bank statement."),
        Col("terms", "Terms (days)", 10, "int",
            validate=number_check(0, 365, decimal=False), hidden=True,
            note="Leave blank to use the standard payment terms from "
                 "Settings."),
        Col("due_date", "Due", 11, "date",
            formula=blank(
                "date",
                'IF({type}{r}<>"Invoice","",{date}{r}+'
                'IF({terms}{r}="",PaymentTermsDays,{terms}{r}))'),
            note="Invoice date plus your payment terms."),
        Col("due_num", "Due (number)", 10, "number",
            # the due date as a plain number in every row, including
            # the empty ones. Other sheets search this column instead
            # of the dated one: multiplying a blank by a number is an
            # error in Excel, and a column of zeros never is.
            formula='=IF({date}{r}="",0,IF({due_date}{r}="",0,'
                    '{due_date}{r}))',
            hidden=True),
        Col("paid", "Paid so far", 12, "money",
            formula=blank(
                "date",
                'IF({type}{r}<>"Invoice","",'
                'ROUND(SUMIFS({total_first}:{total_last},'
                '{type_first}:{type_last},"Payment",'
                '{against_first}:{against_last},{id}{r})-'
                'SUMIFS({total_first}:{total_last},'
                '{type_first}:{type_last},"Refund",'
                '{against_first}:{against_last},{id}{r}),2))'),
            note="Every payment pointing at this invoice, less any "
                 "refund. Partial payments just work."),
        Col("credited", "Credited", 11, "money",
            formula=blank(
                "date",
                'IF({type}{r}<>"Invoice","",'
                'ROUND(SUMIFS({total_first}:{total_last},'
                '{type_first}:{type_last},"Credit note",'
                '{against_first}:{against_last},{id}{r})+'
                'SUMIFS({total_first}:{total_last},'
                '{type_first}:{type_last},"Write-off",'
                '{against_first}:{against_last},{id}{r}),2))'),
            hidden=True),
        Col("balance", "Still owed", 12, "money",
            formula=blank(
                "date",
                'IF({type}{r}<>"Invoice","",'
                'ROUND({total}{r}-{paid}{r}-{credited}{r},2))'),
            cf=owing_cf(),
            note="What is left on this invoice."),
        Col("status", "Status", 12, "status",
            formula=blank(
                "date",
                'IF({type}{r}<>"Invoice",{type}{r},'
                'IF({sent}{r}<>"Yes","Draft",'
                'IF({balance}{r}<=0.004,"Paid",'
                'IF(IF({due_date}{r}="",FALSE,'
                '{due_date}{r}<TodayDate),"Overdue",'
                'IF({paid}{r}>0.004,"Part paid","Sent")))))'),
            cf=status_cf(),
            note="Draft until you mark it sent, then Sent, Part paid, "
                 "Paid or Overdue - worked out, never typed."),
        Col("sent", "Sent to parent?", 11, "flag",
            validate=list_from("YesNoList"),
            note="An invoice you have not sent is not overdue, so the "
                 "workbook will not call it overdue."),
        Col("days_overdue", "Days overdue", 11, "int",
            formula=blank(
                "date",
                'IF({status}{r}<>"Overdue","",'
                'TodayDate-{due_date}{r})'),
            cf=[{"type": "cell", "criteria": ">=", "value": 30,
                 "fmt": "bold_bad"},
                {"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "bad"}]),
        # ------------------------------------- what it does to the books
        Col("billed", "Revenue effect", 12, "money",
            formula=blank(
                "date",
                'ROUND(IF({type}{r}="Invoice",{total}{r},'
                'IF({type}{r}="Credit note",-{total}{r},'
                'IF({type}{r}="Write-off",-{total}{r},0))),2)'),
            note="Revenue. Invoices add, credit notes and write-offs "
                 "take away. Payments are not revenue - they are the "
                 "same money arriving later."),
        Col("collected", "Cash effect", 12, "money",
            formula=blank(
                "date",
                'ROUND(IF({type}{r}="Payment",{total}{r},'
                'IF({type}{r}="Refund",-{total}{r},0)),2)'),
            note="Cash. Payments in, refunds out. Never counted as "
                 "revenue a second time."),
        Col("owed", "Account effect", 12, "money",
            formula=blank(
                "date",
                'ROUND(IF(({type}{r}="Invoice")+({type}{r}="Refund")'
                '>0,{total}{r},-{total}{r}),2)'),
            hidden=True,
            note="What this row does to the family's balance. The "
                 "parent tab adds the column up."),
        Col("fee", "Processing fee", 11, "money",
            formula=blank(
                "date",
                'IF({type}{r}<>"Payment",0,' +
                card_fee("{total}{r}", "{method}{r}") + ")"),
            hidden=True,
            note="What a card payment costs you. Set the rates on "
                 "Settings."),
        Col("notes", "Notes", 24, "wrap"),
        Col("alert", "Action", 32, "status",
            formula=blank(
                "date",
                'IF({type}{r}="","Add - pick a row type",'
                'IF({family}{r}="","Add - no family on this row",'
                'IF({family_name}{r}="Not in Parent Database",'
                '"Check - family ID not found",'
                'IF(IF({type}{r}<>"Invoice",FALSE,'
                'IF({lines}{r}>0,FALSE,{amount}{r}="")),'
                '"Add - no charge lines and no amount",'
                'IF(IF({type}{r}="Invoice",FALSE,{amount}{r}=""),'
                '"Add - no amount entered",'
                'IF(IF({type}{r}="Invoice",FALSE,'
                'IF({type}{r}="Payment",FALSE,{against}{r}="")),'
                '"Check - not matched to an invoice",'
                'IF({status}{r}="Overdue","Overdue - "&'
                '{days_overdue}{r}&" days, chase it",'
                'IF(IF({type}{r}<>"Invoice",FALSE,{sent}{r}<>"Yes"),'
                '"Action - invoice still a draft",'
                'IF(IF({type}{r}<>"Payment",FALSE,{method}{r}=""),'
                '"Add - how did they pay?","OK")))))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Invoices, Payments & Receipts",
        subtitle="One ledger for invoices, payments, credits, refunds "
                 "and write-offs. Revenue and cash kept firmly apart.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.deep,
        freeze_cols=2,
        samples=sample(cfg).ledger,
        help_text="Add an Invoice row, tag the charge lines on "
                  "Tuition & Fees with its number, then add a Payment "
                  "row each time money arrives. Everything else is "
                  "worked out.",
        extra_notes=[
            "To take a payment in instalments, add one Payment row "
            "per instalment, all pointing at the same invoice number. "
            "The invoice shows what is left after each one.",
            "A deposit is just a Payment with no invoice behind it "
            "yet. It sits on the family's account as a credit until "
            "you raise the bill it belongs to - which is why it can "
            "never be counted as income twice.",
            "'Revenue effect' and 'Cash effect' are two different "
            "columns on purpose. The profit figures use revenue; the "
            "collection rate uses cash.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Ledger summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "The ledger", last)
    sheet.table_head(row, 0, [
        "Rows", "Invoices raised", "Invoiced", "Credits & write-offs",
        "Revenue", "Collected", "Outstanding", "Overdue",
        "Overdue invoices", "Collection rate", "Processing fees",
        "This month revenue", "This month collected"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Invoices & Receipts.date]],"<>")',
        '=COUNTIFS([[Invoices & Receipts.type]],"Invoice")',
        '=ROUND(SUMIFS([[Invoices & Receipts.total]],'
        '[[Invoices & Receipts.type]],"Invoice"),2)',
        '=ROUND(SUMIFS([[Invoices & Receipts.total]],'
        '[[Invoices & Receipts.type]],"Credit note")+'
        'SUMIFS([[Invoices & Receipts.total]],'
        '[[Invoices & Receipts.type]],"Write-off"),2)',
        '=ROUND(SUM([[Invoices & Receipts.billed]]),2)',
        '=ROUND(SUM([[Invoices & Receipts.collected]]),2)',
        '=ROUND(SUM([[Invoices & Receipts.billed]])-'
        'SUM([[Invoices & Receipts.collected]]),2)',
        '=ROUND(SUMIFS([[Invoices & Receipts.balance]],'
        '[[Invoices & Receipts.status]],"Overdue"),2)',
        '=COUNTIFS([[Invoices & Receipts.status]],"Overdue")',
        '=IF(SUM([[Invoices & Receipts.billed]])<=0,"",'
        'SUM([[Invoices & Receipts.collected]])/'
        'SUM([[Invoices & Receipts.billed]]))',
        '=ROUND(SUM([[Invoices & Receipts.fee]]),2)',
        '=ROUND(SUMIFS([[Invoices & Receipts.billed]],'
        '[[Invoices & Receipts.month]],ThisMonthKey),2)',
        '=ROUND(SUMIFS([[Invoices & Receipts.collected]],'
        '[[Invoices & Receipts.month]],ThisMonthKey),2)',
    ], ["int", "int", "money", "money", "money", "money", "money",
        "money", "int", "percent", "money", "money", "money"])
````

### `daycare_business_spreadsheet/sheets/expenses.py`

*210 lines · 2 functions*

> `Expense Tracker` - everything that goes out, categorised and budgeted.
>
> One row per payment.  Categories come from a list so that a month's
> spending can actually be added up, recurring costs are converted to a
> monthly equivalent so that an annual insurance premium does not look
> like a disaster in March, and each row carries the budget for its own
> category so the variance is visible where the money was spent.
>
> Wages are the one trap.  Staff costs already come from the Payroll
> Tracker, so rows in the wage categories are excluded from the profit
> and loss by default.  If you would rather keep everything here, switch
> 'Payroll sits in expenses' to Yes on Settings - but do one or the
> other, never both.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Spending summary |

**Module constants:** `NAME` ('Expense Tracker').

````python
"""`Expense Tracker` - everything that goes out, categorised and budgeted.

One row per payment.  Categories come from a list so that a month's
spending can actually be added up, recurring costs are converted to a
monthly equivalent so that an annual insurance premium does not look
like a disaster in March, and each row carries the budget for its own
category so the variance is visible where the money was spent.

Wages are the one trap.  Staff costs already come from the Payroll
Tracker, so rows in the wage categories are excluded from the profit
and loss by default.  If you would rather keep everything here, switch
'Payroll sits in expenses' to Yes on Settings - but do one or the
other, never both.
"""

from __future__ import annotations

from ..config import FREQUENCY_MONTHS, PAYROLL_EXPENSE_CATEGORIES, \
    Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, id_col, month_col,
                     status_cf)

NAME = "Expense Tracker"


def spec(cfg: Config) -> SheetSpec:
    wage_test = "+".join(f'({{category}}{{r}}="{c}")'
                         for c in PAYROLL_EXPENSE_CATEGORIES)

    # Each recurring frequency turned into "times per month".
    freq = "0"
    for label, months in FREQUENCY_MONTHS.items():
        freq = f'IF({{frequency}}{{r}}="{label}",{months:.6f},{freq})'

    columns = [
        id_col("date", "ExpensePrefix", digits=4, width=10,
               header="Expense ID"),
        Col("date", "Date", 11, "date", validate=date_check()),
        month_col("date"),
        Col("category", "Category", 22, "text",
            validate=list_from("ExpenseCategoryList"), cf=status_cf(),
            note="Pick from the list. Typing your own text means the "
                 "monthly category reports will not find it."),
        Col("description", "What it was", 30, "text"),
        Col("vendor", "Paid to", 20, "text"),
        Col("amount", "Amount", 12, "money",
            validate=number_check(-100000, 1000000),
            note="Net of tax if you reclaim it, gross if you do not. "
                 "Be consistent - the workbook does not mind which."),
        Col("tax", "Tax", 10, "money",
            validate=number_check(0, 100000), hidden=True),
        Col("total", "Total paid", 12, "money",
            formula=blank(
                "date",
                'ROUND({amount}{r}+IF({tax}{r}="",0,{tax}{r}),2)'),
            cf=bar_cf(Palette.accent)),
        Col("method", "Paid by", 14, "text",
            validate=list_from("PaymentMethodList")),
        Col("receipt", "Receipt?", 9, "flag",
            validate=list_from("YesNoList"),
            cf=[{"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "good"}],
            note="The column your accountant will thank you for."),
        Col("receipt_ref", "Receipt ref", 13, "text", hidden=True,
            note="Where the paper or the photo lives. A folder name "
                 "is enough."),
        Col("frequency", "Recurring", 13, "text",
            validate=list_from("FrequencyList"),
            note="How often this cost comes round. One off is fine - "
                 "it just means nothing is projected forward."),
        Col("monthly", "Monthly equivalent", 13, "money",
            formula=blank(
                "date",
                f'ROUND({{total}}{{r}}*({freq}),2)'),
            note="A yearly insurance premium divided by twelve, a "
                 "weekly food bill multiplied by 4.33. This is the "
                 "column that tells you your real monthly overhead."),
        Col("is_wages", "Payroll category?", 11, "flag",
            formula=blank("category",
                          f'IF(({wage_test})>0,"Yes","No")'),
            hidden=True),
        Col("in_pl", "Counts in profit?", 11, "flag",
            formula=blank(
                "date",
                'IF({is_wages}{r}<>"Yes","Yes",'
                'IF(PayrollInExpenses="Yes","Yes","No"))'),
            cf=[{"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "muted"}],
            note="Wage rows are left out of the profit figures "
                 "because the Payroll Tracker already counts them. "
                 "Change the switch on Settings if you would rather "
                 "do it the other way round."),
        Col("pl_amount", "In the P&L", 12, "money",
            formula=blank(
                "date",
                'IF({in_pl}{r}="Yes",{total}{r},0)'),
            hidden=True),
        # ------------------------------------------------- the budget --
        Col("budget", "Monthly budget", 12, "money",
            formula=blank(
                "category",
                'IFERROR(INDEX(BudgetAmount,MATCH({category}{r},'
                'BudgetCategory,0))*1,0)'),
            hidden=True,
            note="From the budget table on Settings & Lists."),
        Col("cat_month", "Category this month", 13, "money",
            formula=blank(
                "date",
                'ROUND(SUMIFS({total_first}:{total_last},'
                '{category_first}:{category_last},{category}{r},'
                '{month_first}:{month_last},{month}{r}),2)'),
            hidden=True),
        Col("variance", "Budget variance", 13, "money",
            formula=blank(
                "category",
                'IF({budget}{r}<=0,"",'
                'ROUND({cat_month}{r}-{budget}{r},2))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<=", "value": 0,
                 "fmt": "good"}],
            note="Actual minus budget for this category in this "
                 "month. Positive means over budget."),
        Col("child_cost", "Cost per child", 11, "money",
            formula=blank(
                "date",
                'IF(ChildrenEnrolled<=0,"",'
                'ROUND({total}{r}/ChildrenEnrolled,2))'),
            hidden=True),
        Col("notes", "Notes", 24, "wrap"),
        Col("alert", "Action", 28, "status",
            formula=blank(
                "date",
                'IF({category}{r}="","Add - pick a category",'
                'IF({amount}{r}="","Add - no amount",'
                'IF({receipt}{r}<>"Yes","Missing - no receipt kept",'
                'IF(IF({budget}{r}<=0,FALSE,'
                '{cat_month}{r}>{budget}{r}*1.1),'
                '"Check - category over budget this month",'
                'IF({vendor}{r}="","Add - who was it paid to?",'
                'IF({frequency}{r}="","Add - one off or recurring?",'
                '"OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Expense & Supply Spending",
        subtitle="Everything that goes out, categorised, converted to "
                 "a monthly equivalent and measured against budget.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).expenses,
        help_text="One row per payment. Pick a category from the "
                  "list, say how often it recurs, and the monthly "
                  "equivalent and the budget variance follow.",
        extra_notes=[
            "'Monthly equivalent' is the number to use when you are "
            "working out what the setting costs to run. A yearly "
            "premium and a weekly food bill are not comparable until "
            "both are expressed per month.",
            "Wage rows are deliberately excluded from the profit "
            "figures, because the Payroll Tracker already counts "
            "them. The switch is on Settings if you want it the "
            "other way round.",
            "Set a monthly budget per category on Settings & Lists "
            "and the variance column tells you which categories are "
            "running away from you, in the month they do it.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Spending summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Spending summary", last)
    sheet.table_head(row, 0, [
        "Payments logged", "Total paid", "Counted in profit",
        "Excluded (payroll)", "Monthly overhead", "This month",
        "Last month", "Budget a month", "Over / under budget",
        "Missing receipts"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Expense Tracker.date]],"<>")',
        '=ROUND(SUM([[Expense Tracker.total]]),2)',
        '=ROUND(SUM([[Expense Tracker.pl_amount]]),2)',
        '=ROUND(SUM([[Expense Tracker.total]])-'
        'SUM([[Expense Tracker.pl_amount]]),2)',
        '=ROUND(SUMIFS([[Expense Tracker.monthly]],'
        '[[Expense Tracker.frequency]],"<>One off"),2)',
        '=ROUND(SUMIFS([[Expense Tracker.total]],'
        '[[Expense Tracker.month]],ThisMonthKey),2)',
        '=ROUND(SUMIFS([[Expense Tracker.total]],'
        '[[Expense Tracker.month]],'
        'TEXT(EDATE(TodayDate,-1),"yyyy-mm")),2)',
        '=ROUND(BudgetMonthlyTotal,2)',
        '=ROUND(SUMIFS([[Expense Tracker.total]],'
        '[[Expense Tracker.month]],ThisMonthKey)-BudgetMonthlyTotal,2)',
        '=COUNTIFS([[Expense Tracker.receipt]],"<>Yes",'
        '[[Expense Tracker.date]],"<>")',
    ], ["int", "money", "money", "money", "money", "money", "money",
        "money", "money", "int"])
````

---

## 6. The day a parent reads about

The daily report, the meal log that knows the child's diet and allergy level, the medical and medication record, and the incident form with the clock running from the accident to the moment a parent was told.

### `daycare_business_spreadsheet/sheets/daily_reports.py`

*270 lines · 2 functions*

> `Daily Reports` - the sheet that goes home with the child.
>
> One row per child per day: how they were, what they did, what they ate,
> how they slept, how many nappies, and the sentence the parent actually
> reads at the kitchen table.
>
> Arrival and departure times are pulled from the attendance register
> rather than typed again, meal counts come from the meal tracker, and
> the end-of-day summary writes itself out of the columns beside it.  The
> staff member filling this in at 5:45pm types about eight things.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Daily report summary |

**Module constants:** `NAME` ('Daily Reports').

````python
"""`Daily Reports` - the sheet that goes home with the child.

One row per child per day: how they were, what they did, what they ate,
how they slept, how many nappies, and the sentence the parent actually
reads at the kitchen table.

Arrival and departure times are pulled from the attendance register
rather than typed again, meal counts come from the meal tracker, and
the end-of-day summary writes itself out of the columns beside it.  The
staff member filling this in at 5:45pm types about eight things.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check, time_check
from .common import (alert_cf, blank, id_col, lookup_child,
                     lookup_staff, month_col, status_cf, weekday_col)

NAME = "Daily Reports"


def spec(cfg: Config) -> SheetSpec:
    def from_attendance(target: str) -> str:
        return blank(
            "daykey",
            f'IFERROR(INDEX([[Attendance.{target}]],'
            f'MATCH({{daykey}}{{r}},[[Attendance.daykey]],0))*1,"")')

    columns = [
        id_col("date", "ReportPrefix", digits=5, width=11,
               header="Report ID"),
        Col("date", "Date", 11, "date", validate=date_check()),
        weekday_col("date", hidden=True),
        month_col("date"),
        Col("child", "Child", 11, "id", validate=list_from("ChildList")),
        Col("child_name", "Child's name", 18, "text",
            formula=lookup_child("name")),
        Col("preferred", "Known as", 13, "text",
            formula=lookup_child("preferred", missing=""), hidden=True),
        Col("room", "Room", 14, "text",
            formula=lookup_child("room", missing="")),
        Col("daykey", "Day key", 13, "text",
            formula=blank("child", '{child}{r}&"|"&{date}{r}'),
            hidden=True),
        Col("arrived", "Arrived", 10, "time",
            formula=from_attendance("in_time"),
            note="From the attendance register, matched on child and "
                 "date. Nothing to type twice."),
        Col("departed", "Left", 10, "time",
            formula=from_attendance("out_time")),
        Col("hours", "Hours", 8, "number1",
            formula=from_attendance("hours"), hidden=True),
        # ------------------------------------------------- how they were
        Col("mood", "Mood", 12, "text", validate=list_from("MoodList"),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Tearful", "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "Unwell", "fmt": "bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Unsettled", "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "Happy", "fmt": "good"},
                {"type": "text", "criteria": "containing",
                 "value": "Settled", "fmt": "good"}]),
        Col("behaviour", "Behaviour note", 24, "wrap", hidden=True),
        Col("activities", "What we did", 34, "wrap",
            note="Two or three things. Parents want specifics - "
                 "'painted a rainbow and told everyone about it' "
                 "beats 'arts and crafts'."),
        Col("learning", "Learning focus", 26, "wrap",
            note="The skill behind the activity. This is the column "
                 "that makes a daily report feel like a nursery "
                 "rather than a babysitter."),
        Col("outdoor_min", "Outdoors (min)", 11, "int",
            validate=number_check(0, 600, decimal=False),
            cf=[{"type": "cell", "criteria": ">=", "value": 60,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 30,
                 "fmt": "warn"}]),
        # ------------------------------------------------------- naps ---
        Col("nap_from", "Nap from", 10, "time", validate=time_check()),
        Col("nap_to", "Nap to", 10, "time", validate=time_check()),
        Col("nap_hours", "Slept", 8, "number1",
            formula=blank(
                "nap_from",
                'IF({nap_to}{r}="","",'
                'ROUND((MOD({nap_to}{r}-{nap_from}{r},1))*24,2))')),
        Col("nap2_min", "Second nap (min)", 12, "int",
            validate=number_check(0, 400, decimal=False), hidden=True),
        Col("sleep_total", "Total sleep", 10, "number1",
            formula=blank(
                "date",
                'ROUND(IF({nap_hours}{r}="",0,{nap_hours}{r})+'
                'IF({nap2_min}{r}="",0,{nap2_min}{r}/60),2)'),
            hidden=True),
        # --------------------------------------------- nappies & potty --
        Col("nappies_wet", "Nappies wet", 10, "int",
            validate=number_check(0, 30, decimal=False)),
        Col("nappies_soiled", "Nappies soiled", 11, "int",
            validate=number_check(0, 30, decimal=False)),
        Col("nappies", "Nappy changes", 11, "int",
            formula=blank(
                "date",
                'IF({nappies_wet}{r}="",0,{nappies_wet}{r})+'
                'IF({nappies_soiled}{r}="",0,{nappies_soiled}{r})'),
            hidden=True),
        Col("toileting", "Toileting", 14, "text",
            validate=list_from("ToiletingList"),
            note="For children who are out of nappies or on their way."),
        Col("potty_notes", "Potty training", 22, "wrap", hidden=True,
            note="Progress, accidents, what worked. Parents want to "
                 "keep it consistent at home."),
        # -------------------------------------------------- food & drink
        Col("meals_logged", "Meals logged", 11, "int",
            formula=blank(
                "daykey",
                'COUNTIFS([[Meal Tracker.daykey]],{daykey}{r})'),
            note="From the meal tracker. Log the plates there and "
                 "they are counted here."),
        Col("ate_well", "Ate well", 9, "int",
            formula=blank(
                "daykey",
                'COUNTIFS([[Meal Tracker.daykey]],{daykey}{r},'
                '[[Meal Tracker.portion]],"All")+'
                'COUNTIFS([[Meal Tracker.daykey]],{daykey}{r},'
                '[[Meal Tracker.portion]],"Most")'),
            hidden=True),
        Col("refused", "Refused", 9, "int",
            formula=blank(
                "daykey",
                'COUNTIFS([[Meal Tracker.daykey]],{daykey}{r},'
                '[[Meal Tracker.portion]],"Refused")+'
                'COUNTIFS([[Meal Tracker.daykey]],{daykey}{r},'
                '[[Meal Tracker.portion]],"None")'),
            cf=[{"type": "cell", "criteria": ">=", "value": 2,
                 "fmt": "warn"}],
            hidden=True),
        Col("eating", "Eating", 13, "text",
            formula=blank(
                "date",
                'IF({meals_logged}{r}=0,"Not logged",'
                'IF({refused}{r}>={meals_logged}{r},"Refused most",'
                'IF({ate_well}{r}>={meals_logged}{r},"Ate everything",'
                '"Ate some")))'),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Refused", "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "everything", "fmt": "good"},
                {"type": "text", "criteria": "containing",
                 "value": "Not logged", "fmt": "muted"}]),
        Col("water", "Drinks", 8, "int",
            validate=number_check(0, 30, decimal=False),
            note="Cups or bottles offered and taken."),
        Col("bottles", "Bottles (ml)", 10, "int",
            validate=number_check(0, 3000, decimal=False), hidden=True),
        # --------------------------------------------------- the people -
        Col("teacher", "Written by", 11, "id",
            validate=list_from("StaffList")),
        Col("teacher_name", "Staff", 16, "text",
            formula=lookup_staff("name", "teacher"), hidden=True),
        Col("observations", "Teacher's observations", 34, "wrap",
            note="The one thing you noticed that the parent did not "
                 "see. This is the whole value of a daily report."),
        Col("parent_notes", "Message to parents", 30, "wrap",
            note="Bring wellies tomorrow. Nearly out of nappies. We "
                 "need the trip form back."),
        Col("incidents", "Incidents today", 10, "int",
            formula=blank(
                "date",
                'COUNTIFS([[Incident Reports.child]],{child}{r},'
                '[[Incident Reports.date]],{date}{r})'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "bad"}],
            note="Pulled from the incident tab so nothing is sent "
                 "home without you knowing it happened."),
        Col("summary", "End of day summary", 44, "wrap",
            formula=blank(
                "date",
                'IF({child_name}{r}="","",'
                'IF({preferred}{r}="",{child_name}{r},'
                '{preferred}{r})&" was "&'
                'LOWER(IF({mood}{r}="","settled",{mood}{r}))&" today. "&'
                'IF({sleep_total}{r}>0,"Slept "&'
                'TEXT({sleep_total}{r},"0.0")&"h. ","")&'
                'IF({meals_logged}{r}>0,{eating}{r}&" at "&'
                '{meals_logged}{r}&" meal(s). ","")&'
                'IF({nappies}{r}>0,{nappies}{r}&" nappy changes. ","")&'
                'IF({outdoor_min}{r}>0,{outdoor_min}{r}&'
                '" minutes outside. ","")&'
                'IF({incidents}{r}>0,"Please read the incident form '
                'we have sent home. ","")&'
                'IF({parent_notes}{r}="","",{parent_notes}{r}))'),
            note="Written for you out of the columns to the left. "
                 "Copy it into your app, print it, or hand the row "
                 "over as it is."),
        Col("shared", "Sent to parent?", 11, "flag",
            validate=list_from("YesNoList"), cf=status_cf()),
        Col("alert", "Action", 28, "status",
            formula=blank(
                "date",
                'IF({child}{r}="","Add - no child picked",'
                'IF({mood}{r}="","Add - how were they today?",'
                'IF(IF({activities}{r}="",{observations}{r}="",FALSE),'
                '"Add - nothing recorded for the parent",'
                'IF({incidents}{r}>0,"Action - incident form goes '
                'home too",'
                'IF(IF({shared}{r}="Yes",FALSE,{date}{r}<TodayDate),'
                '"Chase - not sent to the parent yet",'
                'IF({refused}{r}>=2,"Check - refused more than one '
                'meal","OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Daily Child Reports",
        subtitle="Mood, activities, naps, nappies, meals and the "
                 "end-of-day summary - written for you from the "
                 "columns beside it.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).daily_reports,
        help_text="Pick the date and the child. Arrival times, meals "
                  "and incidents come from the other tabs; you fill "
                  "in the mood, the activities and the observation.",
        extra_notes=[
            "The summary column writes a full sentence out of "
            "everything else on the row. Filter to one day, widen the "
            "column and print - that is your take-home report.",
            "Meals come from the Meal Tracker and arrival times from "
            "Attendance, both matched on the child and the date. If a "
            "time is blank here, it is blank on the attendance sheet.",
            "If a child had an incident today, the summary says so. "
            "No daily report should be the first a parent hears of it.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Daily report summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Reports written", last)
    sheet.table_head(row, 0, [
        "Reports logged", "Written today", "Sent to parents",
        "Not sent yet", "Average sleep (hours)",
        "Average outdoor minutes", "Average nappy changes",
        "Children unsettled", "Days with a refused meal"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Daily Reports.date]],"<>")',
        '=COUNTIFS([[Daily Reports.date]],TodayDate)',
        '=COUNTIFS([[Daily Reports.shared]],"Yes")',
        '=COUNTIF([[Daily Reports.date]],"<>")-'
        'COUNTIFS([[Daily Reports.shared]],"Yes")',
        '=IFERROR(ROUND(AVERAGEIF([[Daily Reports.sleep_total]],">0"),'
        '1),0)',
        '=IFERROR(ROUND(AVERAGEIF([[Daily Reports.outdoor_min]],">0"),'
        '0),0)',
        '=IFERROR(ROUND(AVERAGEIF([[Daily Reports.nappies]],">0"),1),0)',
        '=COUNTIFS([[Daily Reports.mood]],"Unsettled")+'
        'COUNTIFS([[Daily Reports.mood]],"Tearful")',
        '=COUNTIFS([[Daily Reports.refused]],">0")',
    ], ["int", "int", "int", "int", "number1", "int", "number1", "int",
        "int"])
````

### `daycare_business_spreadsheet/sheets/meals.py`

*172 lines · 2 functions*

> `Meal Tracker` - what was served, who ate it, and who must not have it.
>
> One row per child per sitting.  The reason this tab exists and a paper
> menu does not is the two columns in the middle: every plate carries the
> child's recorded diet and their allergy level, pulled live from the
> medical tab, so the person serving lunch sees the warning on the same
> line as the food.
>
> It is a prompt, not a safety system.  The allergy plan on the wall is
> the safety system.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Meal log summary |

**Module constants:** `NAME` ('Meal Tracker').

````python
"""`Meal Tracker` - what was served, who ate it, and who must not have it.

One row per child per sitting.  The reason this tab exists and a paper
menu does not is the two columns in the middle: every plate carries the
child's recorded diet and their allergy level, pulled live from the
medical tab, so the person serving lunch sees the warning on the same
line as the food.

It is a prompt, not a safety system.  The allergy plan on the wall is
the safety system.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, lookup_child,
                     lookup_staff, month_col, severity_cf, status_cf)

NAME = "Meal Tracker"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "MealPrefix", digits=5, width=11,
               header="Meal ID"),
        Col("date", "Date", 11, "date", validate=date_check()),
        month_col("date"),
        Col("sitting", "Sitting", 14, "text",
            validate=list_from("MealSittingList"), cf=status_cf(),
            note="Breakfast, lunch, snack, bottle. One row per child "
                 "per sitting."),
        Col("child", "Child", 11, "id", validate=list_from("ChildList")),
        Col("child_name", "Child's name", 18, "text",
            formula=lookup_child("name")),
        Col("room", "Room", 14, "text",
            formula=lookup_child("room", missing=""), hidden=True),
        Col("daykey", "Day key", 13, "text",
            formula=blank("child", '{child}{r}&"|"&{date}{r}'),
            hidden=True),
        # ------------------------------------------------ the warnings --
        Col("diet", "Diet on file", 15, "text",
            formula=lookup_child("diet", missing=""),
            cf=[{"type": "text", "criteria": "not containing",
                 "value": "No restrictions", "fmt": "info"}],
            note="From the child's enrollment record. Nothing to type "
                 "- change it there and every future plate follows."),
        Col("allergy_level", "Allergy risk", 13, "status",
            formula=lookup_child("allergy_level", missing=""),
            cf=severity_cf(),
            note="The worst allergy severity recorded for this child "
                 "on the Allergy & Medical tab. Red means read the "
                 "plan before you serve anything."),
        Col("allergy_main", "Allergic to", 18, "text",
            formula=lookup_child("first_allergy", missing=""),
            cf=[{"type": "text", "criteria": "not containing",
                 "value": "zzzz", "fmt": "warn"}],
            note="The first allergy on the child's medical record. "
                 "Where there is more than one, the medical tab has "
                 "the full list."),
        # --------------------------------------------------- the plate --
        Col("food", "Food served", 30, "text",
            note="What actually went on the plate, not what the menu "
                 "said. They differ more often than anyone admits."),
        Col("portion", "Amount eaten", 13, "text",
            validate=list_from("EatenAmountList"),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "All", "fmt": "good"},
                {"type": "text", "criteria": "containing",
                 "value": "Most", "fmt": "good"},
                {"type": "text", "criteria": "containing",
                 "value": "None", "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "Refused", "fmt": "bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Not offered", "fmt": "muted"}]),
        Col("refused", "Refused?", 9, "flag",
            formula=blank(
                "portion",
                'IF(({portion}{r}="Refused")+({portion}{r}="None")>0,'
                '"Yes","No")'),
            hidden=True),
        Col("alternative", "Alternative given", 22, "text",
            hidden=True,
            note="What you offered instead. Worth recording for a "
                 "fussy eater's parent."),
        Col("drink", "Drink", 16, "text"),
        Col("water_ml", "Fluid (ml)", 10, "int",
            validate=number_check(0, 3000, decimal=False),
            note="Only if you track it. Infant rooms usually do, "
                 "preschool rooms usually do not."),
        Col("own_food", "Brought own?", 10, "flag",
            validate=list_from("YesNoList"), hidden=True),
        Col("logged_by", "Logged by", 11, "id",
            validate=list_from("StaffList")),
        Col("logged_name", "Staff", 15, "text",
            formula=lookup_staff("name", "logged_by"), hidden=True),
        Col("notes", "Notes", 26, "wrap"),
        Col("alert", "Action", 32, "status",
            formula=blank(
                "date",
                'IF({child}{r}="","Add - no child picked",'
                'IF(({allergy_level}{r}="Anaphylaxis risk")+'
                '({allergy_level}{r}="Severe")>0,'
                '"Allergy - "&{allergy_main}{r}&", check the plan",'
                'IF({sitting}{r}="","Add - which sitting?",'
                'IF({food}{r}="","Add - what was served?",'
                'IF({portion}{r}="","Add - how much did they eat?",'
                'IF({refused}{r}="Yes","Check - refused, tell the '
                'parent",'
                'IF({allergy_level}{r}="Moderate","Check - moderate '
                'allergy on file","OK")))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Meal & Nutrition Tracker",
        subtitle="Every plate, with the child's diet and allergy "
                 "level on the same line as the food.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).meals,
        help_text="Pick the date, the sitting and the child. The diet "
                  "and the allergy warning appear on their own - you "
                  "type what was served and how much went in.",
        extra_notes=[
            "The allergy columns are pulled live from the Allergy & "
            "Medical tab. Record an allergy once, and it appears on "
            "every future meal for that child without anybody having "
            "to remember.",
            "This is a prompt, not a safeguard. The written allergy "
            "plan, the kitchen's own checks and the adult standing at "
            "the table are the safeguard.",
            "'Refused' twice in a day shows up on the child's daily "
            "report automatically, which saves the awkward "
            "conversation at the door.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Meal log summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "What got eaten", last)
    sheet.table_head(row, 0, [
        "Plates logged", "Ate it all", "Ate most", "Ate some",
        "Refused", "Refusal rate", "Allergy plates served",
        "Special diets served", "Fluids logged (ml)"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Meal Tracker.date]],"<>")',
        '=COUNTIFS([[Meal Tracker.portion]],"All")',
        '=COUNTIFS([[Meal Tracker.portion]],"Most")',
        '=COUNTIFS([[Meal Tracker.portion]],"About half")+'
        'COUNTIFS([[Meal Tracker.portion]],"A little")',
        '=COUNTIFS([[Meal Tracker.refused]],"Yes")',
        '=IF(COUNTIF([[Meal Tracker.portion]],"<>")<=0,"",'
        'COUNTIFS([[Meal Tracker.refused]],"Yes")/'
        'COUNTIF([[Meal Tracker.portion]],"<>"))',
        '=COUNTIFS([[Meal Tracker.allergy_level]],"Severe")+'
        'COUNTIFS([[Meal Tracker.allergy_level]],"Anaphylaxis risk")',
        '=COUNTIF([[Meal Tracker.diet]],"<>")-'
        'COUNTIFS([[Meal Tracker.diet]],"No restrictions")',
        '=SUM([[Meal Tracker.water_ml]])',
    ], ["int", "int", "int", "int", "int", "percent", "int", "int",
        "int"])
````

### `daycare_business_spreadsheet/sheets/medical.py`

*232 lines · 2 functions*

> `Allergy & Medical` - conditions, plans, consents and the dose log.
>
> Two kinds of row live here, and the record type tells them apart.
>
> A **profile** row is a standing fact: this child has a peanut allergy,
> this child has an inhaler, this child must not have dairy.  It carries
> the severity, the reaction, what to do, who prescribed it and when the
> plan or the medicine runs out.
>
> An **administration** row is an event: at 2:15pm on Tuesday, this
> member of staff gave this dose.  Keeping both in one register means a
> dose is always filed next to the plan it came from, and the expiry
> dates are never in a different file from the medicine.
>
> This is a working record for your own staff.  It is not a medical
> record, it does not replace a medication protocol, and it should be
> stored somewhere only the people who need it can open.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Medical register summary |

**Module constants:** `NAME` ('Allergy & Medical').

````python
"""`Allergy & Medical` - conditions, plans, consents and the dose log.

Two kinds of row live here, and the record type tells them apart.

A **profile** row is a standing fact: this child has a peanut allergy,
this child has an inhaler, this child must not have dairy.  It carries
the severity, the reaction, what to do, who prescribed it and when the
plan or the medicine runs out.

An **administration** row is an event: at 2:15pm on Tuesday, this
member of staff gave this dose.  Keeping both in one register means a
dose is always filed next to the plan it came from, and the expiry
dates are never in a different file from the medicine.

This is a working record for your own staff.  It is not a medical
record, it does not replace a medication protocol, and it should be
stored somewhere only the people who need it can open.
"""

from __future__ import annotations

from ..config import MEDICAL_EVENT_TYPES, Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, time_check
from .common import (alert_cf, blank, id_col, lookup_child,
                     lookup_staff, seq_key, severity_cf, status_cf)

NAME = "Allergy & Medical"


def spec(cfg: Config) -> SheetSpec:
    event_test = "+".join(f'({{record_type}}{{r}}="{t}")'
                          for t in MEDICAL_EVENT_TYPES)

    columns = [
        id_col("child", "MedicalPrefix", digits=4, width=11,
               header="Record ID"),
        Col("child", "Child", 11, "id", validate=list_from("ChildList"),
            note="Which child this record belongs to. Everything else "
                 "on the row describes them."),
        Col("child_name", "Child's name", 18, "text",
            formula=lookup_child("name")),
        Col("room", "Room", 13, "text",
            formula=lookup_child("room", missing=""), hidden=True),
        *seq_key("child"),
        Col("record_type", "Record type", 17, "text",
            validate=list_from("MedicalRecordList"), cf=status_cf(),
            note="Allergy, condition, medication, immunisation - or "
                 "'Medication given' to log a single dose."),
        Col("is_event", "Dose log?", 8, "flag",
            formula=blank("record_type",
                          f'IF(({event_test})>0,"Yes","No")'),
            hidden=True,
            note="Yes on the rows that record one dose rather than a "
                 "standing instruction."),
        Col("detail", "Detail", 24, "text",
            note="Peanuts. Asthma. Lactose. Amoxicillin. Whatever the "
                 "record is actually about - one thing per row."),
        Col("severity", "Severity", 15, "status",
            validate=list_from("AllergySeverityList"), cf=severity_cf(),
            note="Drives the warning on the child's record, on every "
                 "meal and on the dashboard. 'Anaphylaxis risk' turns "
                 "everything red, which is the point."),
        Col("reaction", "What happens", 28, "wrap",
            note="What a reaction looks like, in the words a new "
                 "member of staff would recognise it by."),
        Col("action", "What to do", 34, "wrap",
            note="The steps, in order, including when to call an "
                 "ambulance. Short sentences. This is read in a hurry."),
        Col("medication", "Medication", 20, "text"),
        Col("dose", "Dose & route", 20, "text",
            note="Exactly as the prescription says it. Never as "
                 "anybody remembers it."),
        Col("stored", "Stored where", 18, "text",
            note="Red box in the office. Fridge, top shelf. The "
                 "sentence that saves two minutes in an emergency."),
        Col("trained", "Staff trained", 18, "text", hidden=True,
            note="Who is trained to give this. Names, not job titles."),
        # ---------------------------------------------- authorisation ---
        Col("consent", "Parent consent", 12, "flag",
            validate=list_from("YesNoList"),
            cf=[{"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "bold_bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "good"}],
            note="Written consent on file. Without it, nothing is "
                 "given - and the Action column says so."),
        Col("consent_date", "Consent dated", 12, "date",
            validate=date_check(), hidden=True),
        Col("prescriber", "Prescribed by", 18, "text", hidden=True),
        Col("doctor", "Doctor / clinic", 20, "text"),
        Col("doctor_phone", "Doctor's phone", 14, "text"),
        Col("reviewed", "Plan dated", 11, "date",
            validate=date_check(),
            note="When the plan was last written or reviewed with the "
                 "parent."),
        Col("expires", "Expires", 11, "date", validate=date_check(),
            note="The medicine's expiry date, or the date the plan "
                 "needs reviewing. Whichever comes first."),
        Col("days_left", "Days left", 10, "int",
            formula=blank(
                "expires", '{expires}{r}-TodayDate'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bold_bad"},
                {"type": "cell", "criteria": "<=",
                 "value": "=DocUrgentDays", "fmt": "bad"},
                {"type": "cell", "criteria": "<=",
                 "value": "=DocWarningDays", "fmt": "warn"}]),
        # ------------------------------------------------- the dose log -
        Col("given_date", "Given on", 11, "date",
            validate=date_check(),
            note="Only on a 'Medication given' row. Leave blank on a "
                 "standing record."),
        Col("given_time", "Given at", 10, "time",
            validate=time_check()),
        Col("given_by", "Given by", 11, "id",
            validate=list_from("StaffList"),
            note="The member of staff who administered it. Required "
                 "on every dose row - that is the whole point of the "
                 "log."),
        Col("given_name", "Staff", 16, "text",
            formula=lookup_staff("name", "given_by")),
        Col("witness", "Witnessed by", 14, "text", hidden=True,
            note="Many settings require a second pair of eyes. If "
                 "yours does, this is where it goes."),
        Col("doses_logged", "Doses logged", 11, "int",
            formula=blank(
                "child",
                'IF({is_event}{r}="Yes","",'
                'COUNTIFS({child_first}:{child_last},{child}{r},'
                '{detail_first}:{detail_last},{detail}{r},'
                '{is_event_first}:{is_event_last},"Yes"))'),
            hidden=True,
            note="On a standing record, how many dose rows exist for "
                 "the same child and the same medicine."),
        Col("last_dose", "Last dose", 11, "date",
            formula=blank(
                "child",
                'IF({is_event}{r}="Yes","",'
                'IF({doses_logged}{r}=0,"",'
                'SUMPRODUCT(MAX(([[Allergy & Medical.child]]='
                '{child}{r})*([[Allergy & Medical.detail]]='
                '{detail}{r})*([[Allergy & Medical.is_event]]="Yes")*'
                '[[Allergy & Medical.given_date]]))))'),
            hidden=True),
        Col("notes", "Notes", 28, "wrap"),
        Col("alert", "Action", 34, "status",
            formula=blank(
                "child",
                'IF({child_name}{r}="Not in Child Enrollment",'
                '"Check - child ID not found",'
                'IF({record_type}{r}="","Add - what kind of record is '
                'this?",'
                'IF(IF({is_event}{r}<>"Yes",FALSE,{given_by}{r}=""),'
                '"Missing - no staff name against this dose",'
                'IF(IF({is_event}{r}<>"Yes",FALSE,{given_date}{r}=""),'
                '"Missing - no date against this dose",'
                'IF(IF({medication}{r}="",FALSE,{consent}{r}<>"Yes"),'
                '"Not authorised - no written parent consent",'
                'IF(IF({expires}{r}="",FALSE,{days_left}{r}<0),'
                '"Expired - renew it before it is needed",'
                'IF(IF({expires}{r}="",FALSE,'
                '{days_left}{r}<=DocUrgentDays),'
                '"Expiring - "&{days_left}{r}&" days left",'
                'IF(IF(({severity}{r}="Severe")+'
                '({severity}{r}="Anaphylaxis risk")>0,'
                '{action}{r}="",FALSE),'
                '"Missing - severe allergy with no written plan",'
                'IF(IF({expires}{r}="",FALSE,'
                '{days_left}{r}<=DocWarningDays),'
                '"Due - review the plan with the parent","OK")))))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Allergy, Medical & Medication Records",
        subtitle="Conditions, allergy plans, consents, expiry dates "
                 "and the dose log - one child's record in one place.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.bad_text,
        freeze_cols=3,
        samples=sample(cfg).medical,
        help_text="One row per condition, allergy or medication. To "
                  "log a dose, add a row with record type 'Medication "
                  "given' - it files itself next to the plan.",
        extra_notes=[
            "This is a working note for your staff. It is not a "
            "medical record and it does not replace your medication "
            "protocol, your allergy plans or your emergency "
            "procedures. Keep the signed originals where your policy "
            "says they live.",
            "Severity is the column that travels. 'Severe' or "
            "'Anaphylaxis risk' here turns the child's enrollment "
            "record red and puts a warning on every meal they are "
            "served.",
            "This tab holds sensitive personal data about children. "
            "Restrict who can open this file, keep it encrypted or "
            "access controlled, and delete records in line with your "
            "retention policy.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Medical register summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Medical register at a glance", last)
    sheet.table_head(row, 0, [
        "Records", "Children with a record", "Allergies",
        "Anaphylaxis risk", "Severe", "On medication",
        "Missing consent", "Plans expired", "Expiring soon",
        "Doses logged"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Allergy & Medical.child]],"<>")',
        '=COUNTIFS([[Child Enrollment.allergy_count]],">0")',
        '=COUNTIFS([[Allergy & Medical.record_type]],"Allergy")',
        '=COUNTIFS([[Allergy & Medical.severity]],"Anaphylaxis risk")',
        '=COUNTIFS([[Allergy & Medical.severity]],"Severe")',
        '=COUNTIF([[Allergy & Medical.medication]],"<>")',
        '=COUNTIFS([[Allergy & Medical.medication]],"<>",'
        '[[Allergy & Medical.consent]],"<>Yes")',
        '=COUNTIFS([[Allergy & Medical.days_left]],"<0")',
        '=COUNTIFS([[Allergy & Medical.days_left]],">=0",'
        '[[Allergy & Medical.days_left]],"<="&DocWarningDays)',
        '=COUNTIFS([[Allergy & Medical.is_event]],"Yes")',
    ], ["int", "int", "int", "int", "int", "int", "int", "int", "int",
        "int"])
````

### `daycare_business_spreadsheet/sheets/incidents.py`

*248 lines · 2 functions*

> `Incident Reports` - accidents, injuries, near misses and who was told.
>
> An incident form has one job beyond the obvious: to prove, months
> later, that it was written up the same day, that a parent was told
> within the hour and that somebody signed it.  Those three facts are
> what gets checked, and all three are columns here.
>
> The write-up column assembles a readable paragraph from the fields
> beside it, so the form you print for the parent reads like a form and
> not like a spreadsheet row.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Incident summary |

**Module constants:** `NAME` ('Incident Reports').

````python
"""`Incident Reports` - accidents, injuries, near misses and who was told.

An incident form has one job beyond the obvious: to prove, months
later, that it was written up the same day, that a parent was told
within the hour and that somebody signed it.  Those three facts are
what gets checked, and all three are columns here.

The write-up column assembles a readable paragraph from the fields
beside it, so the form you print for the parent reads like a form and
not like a spreadsheet row.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, time_check
from .common import (alert_cf, blank, id_col, lookup_child,
                     lookup_staff, month_col, severity_cf, status_cf)

NAME = "Incident Reports"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "IncidentPrefix", digits=4, width=11,
               header="Incident ID",
               note="Automatic. Quote it to the parent and in your "
                    "own records."),
        Col("date", "Date", 11, "date", validate=date_check()),
        month_col("date"),
        Col("time", "Time", 10, "time", validate=time_check(),
            note="As close as you can. 'Around 2pm' helps nobody six "
                 "months later."),
        Col("child", "Child", 11, "id", validate=list_from("ChildList")),
        Col("child_name", "Child's name", 18, "text",
            formula=lookup_child("name")),
        Col("age_at", "Age then", 9, "text",
            formula=blank(
                "child",
                'IFERROR(DATEDIF(INDEX([[Child Enrollment.dob]],'
                'MATCH({child}{r},[[Child Enrollment.id]],0)),'
                '{date}{r},"y")&"y","")'),
            hidden=True),
        Col("room", "Room", 14, "text",
            formula=lookup_child("room", missing=""), hidden=True),
        Col("location", "Where", 16, "text",
            validate=list_from("IncidentLocationList")),
        Col("category", "What happened", 18, "text",
            validate=list_from("IncidentCategoryList"), cf=status_cf()),
        Col("severity", "Severity", 16, "status",
            validate=list_from("IncidentSeverityList"),
            cf=severity_cf()),
        Col("description", "Description", 40, "wrap",
            note="What happened, in order, in plain words. Facts "
                 "only - what you saw, not what you concluded."),
        Col("injury", "Injury & where", 28, "wrap",
            note="Body part and what it looked like. 'Graze, left "
                 "knee, 2cm' is a record; 'hurt leg' is not."),
        Col("first_aid", "First aid given", 30, "wrap"),
        Col("first_aider", "First aid by", 11, "id",
            validate=list_from("StaffList")),
        Col("first_aider_name", "First aider", 16, "text",
            formula=lookup_staff("name", "first_aider")),
        Col("staff", "Staff present", 11, "id",
            validate=list_from("StaffList"),
            note="Whoever was supervising at the time."),
        Col("staff_name", "Supervising", 16, "text",
            formula=lookup_staff("name", "staff"), hidden=True),
        Col("witness", "Witness", 18, "text"),
        # ------------------------------------------- telling the parent -
        Col("notified", "Parent told?", 11, "flag",
            validate=list_from("YesNoList"),
            cf=[{"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "bold_bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "good"}]),
        Col("notified_date", "Told on", 11, "date",
            validate=date_check()),
        Col("notified_time", "Told at", 10, "time",
            validate=time_check()),
        Col("notified_how", "Told by", 14, "text",
            validate=list_from("ContactMethodList")),
        Col("notified_by", "Told by whom", 11, "id",
            validate=list_from("StaffList"), hidden=True),
        Col("notify_mins", "Minutes to tell", 11, "int",
            formula=blank(
                "notified_date",
                'IF({time}{r}="","",IF({notified_time}{r}="","",'
                'ROUND((({notified_date}{r}+{notified_time}{r})-'
                '({date}{r}+{time}{r}))*1440,0)))'),
            cf=[{"type": "cell", "criteria": ">", "value": 240,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<=", "value": 60,
                 "fmt": "good"}],
            note="How long between the incident and telling the "
                 "parent. The number everybody asks about afterwards."),
        Col("signed", "Parent signed?", 11, "flag",
            validate=list_from("YesNoList"),
            cf=[{"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "warn"},
                {"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "good"}]),
        Col("signed_date", "Signed on", 11, "date",
            validate=date_check(), hidden=True),
        # ------------------------------------------------- afterwards ---
        Col("follow_up", "Follow-up action", 30, "wrap",
            note="What changed because of this. A gate catch, a rule, "
                 "a risk assessment. An incident with no follow-up "
                 "happens twice."),
        Col("follow_by", "Follow up by", 11, "date",
            validate=date_check()),
        Col("follow_done", "Done?", 8, "flag",
            validate=list_from("YesNoList"), cf=status_cf()),
        Col("reported_out", "Reported to", 18, "text",
            note="Only where your local rules require it - a "
                 "regulator, an insurer, a safeguarding lead. Blank "
                 "is fine and common."),
        Col("medical_sought", "Medical attention", 14, "text",
            hidden=True,
            note="None, first aid only, doctor, hospital. Record it "
                 "even when the answer is none."),
        Col("complete", "Form complete?", 11, "flag",
            formula=blank(
                "date",
                'IF({description}{r}="","No",'
                'IF({severity}{r}="","No",'
                'IF({notified}{r}<>"Yes","No",'
                'IF({first_aid}{r}="","No","Yes"))))'),
            cf=[{"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "good"}]),
        Col("writeup", "Printable write-up", 46, "wrap",
            formula=blank(
                "date",
                'IF({child_name}{r}="","",'
                '"On "&TEXT({date}{r},"dddd d mmmm yyyy")&'
                'IF({time}{r}="",""," at "&TEXT({time}{r},"h:mm AM/PM"))&'
                '", "&{child_name}{r}&" was involved in a "&'
                'LOWER(IF({category}{r}="","n incident",'
                '{category}{r}))&" in the "&'
                'LOWER(IF({location}{r}="","setting",{location}{r}))&'
                '". "&{description}{r}&'
                'IF({injury}{r}=""," "," Injury: "&{injury}{r}&".")&'
                'IF({first_aid}{r}=""," "," First aid given: "&'
                '{first_aid}{r}&".")&'
                'IF({first_aider_name}{r}=""," "," Treated by "&'
                '{first_aider_name}{r}&".")&'
                'IF({notified}{r}<>"Yes"," The parent has not yet '
                'been informed."," The parent was informed on "&'
                'TEXT({notified_date}{r},"d mmm yyyy")&'
                'IF({notified_time}{r}=""," "," at "&'
                'TEXT({notified_time}{r},"h:mm AM/PM"))&'
                'IF({notified_how}{r}=""," by "&LOWER('
                '{notified_how}{r})&".","."))&'
                'IF({follow_up}{r}=""," "," Follow-up: "&'
                '{follow_up}{r}&"."))'),
            note="A readable paragraph assembled from the columns on "
                 "this row. Widen the column, print the row, get it "
                 "signed."),
        Col("notes", "Internal notes", 26, "wrap",
            note="For you, not for the parent's copy."),
        Col("alert", "Action", 34, "status",
            formula=blank(
                "date",
                'IF({child}{r}="","Add - no child on this report",'
                'IF({description}{r}="","Missing - no description '
                'written",'
                'IF({severity}{r}="","Add - how serious was it?",'
                'IF({notified}{r}<>"Yes","Not notified - tell the '
                'parent today",'
                'IF({signed}{r}<>"Yes","Chase - parent has not signed '
                'the form",'
                'IF(IF({follow_by}{r}="",FALSE,'
                'IF({follow_done}{r}="Yes",FALSE,'
                '{follow_by}{r}<TodayDate)),'
                '"Overdue - follow-up action not done",'
                'IF(IF({follow_up}{r}="",'
                '({severity}{r}="Serious")+'
                '({severity}{r}="Emergency services called")>0,FALSE),'
                '"Action - serious incident with no follow-up",'
                'IF({complete}{r}="No","Action - form not complete",'
                '"OK"))))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Incident & Accident Reports",
        subtitle="What happened, who dealt with it, when the parent "
                 "was told and what changed afterwards.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.bad_text,
        freeze_cols=2,
        samples=sample(cfg).incidents,
        help_text="Fill the form in the same day. The write-up "
                  "column turns the row into a printable paragraph "
                  "for the parent to read and sign.",
        extra_notes=[
            "'Minutes to tell' is the column that matters if anybody "
            "ever questions how an incident was handled. Under an "
            "hour is green.",
            "A serious incident with nothing in the follow-up column "
            "is flagged until you write something there. That is "
            "deliberate - the follow-up is the part that stops it "
            "happening again.",
            "This workbook does not report anything to anybody. Where "
            "your local rules require an incident to be notified to a "
            "regulator or insurer, that is a separate job - record "
            "that you did it in 'Reported to'.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Incident summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Incident record", last)
    sheet.table_head(row, 0, [
        "Incidents logged", "This month", "No injury", "Minor",
        "Moderate or worse", "Parents told", "Not yet told",
        "Signed by parent", "Average minutes to tell",
        "Follow-ups outstanding", "Forms incomplete"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Incident Reports.date]],"<>")',
        '=COUNTIFS([[Incident Reports.month]],ThisMonthKey)',
        '=COUNTIFS([[Incident Reports.severity]],"No injury")',
        '=COUNTIFS([[Incident Reports.severity]],"Minor")',
        '=COUNTIFS([[Incident Reports.severity]],"Moderate")+'
        'COUNTIFS([[Incident Reports.severity]],"Serious")+'
        'COUNTIFS([[Incident Reports.severity]],'
        '"Emergency services called")',
        '=COUNTIFS([[Incident Reports.notified]],"Yes")',
        '=COUNTIF([[Incident Reports.date]],"<>")-'
        'COUNTIFS([[Incident Reports.notified]],"Yes")',
        '=COUNTIFS([[Incident Reports.signed]],"Yes")',
        '=IFERROR(ROUND(AVERAGEIF([[Incident Reports.notify_mins]],'
        '">=0"),0),0)',
        '=COUNTIFS([[Incident Reports.follow_by]],"<>",'
        '[[Incident Reports.follow_done]],"<>Yes")',
        '=COUNTIFS([[Incident Reports.complete]],"No")',
    ], ["int", "int", "int", "int", "int", "int", "int", "int", "int",
        "int", "int"])
````

---

## 7. The team

Staff records with four certificate expiry dates rolled into one countdown, a rota that catches a double shift and splits basic from overtime hour by hour, and a payroll tab that knows the difference between what somebody is paid and what they cost.

### `daycare_business_spreadsheet/sheets/staff.py`

*287 lines · 2 functions*

> `Staff Database` - the team, their pay and their certificates.
>
> Childcare is the one small business where a lapsed certificate is not
> an administrative annoyance but a reason you cannot open the room.
> First aid, CPR, safeguarding training and background checks all have
> dates, and all four are watched here: the soonest expiry across the
> lot is calculated into one column, and the dashboard counts whatever
> is about to run out.
>
> Pay lives here too, as a rate rather than an amount, so the Staff
> Schedule can turn hours into money and the Payroll Tracker can turn
> money into a cost.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Team summary |

**Module constants:** `NAME` ('Staff Database').

````python
"""`Staff Database` - the team, their pay and their certificates.

Childcare is the one small business where a lapsed certificate is not
an administrative annoyance but a reason you cannot open the room.
First aid, CPR, safeguarding training and background checks all have
dates, and all four are watched here: the soonest expiry across the
lot is calculated into one column, and the dashboard counts whatever
is about to run out.

Pay lives here too, as a rate rather than an amount, so the Staff
Schedule can turn hours into money and the Payroll Tracker can turn
money into a cost.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, id_col, status_cf)

NAME = "Staff Database"


def spec(cfg: Config) -> SheetSpec:
    doc_who = "[[Document Tracker.who]]"
    doc_req = "[[Document Tracker.required]]"
    doc_status = "[[Document Tracker.status]]"

    columns = [
        id_col("name", "StaffPrefix", digits=3, width=10,
               header="Staff ID",
               note="Automatic. The rota, the payroll, the attendance "
                    "register and every incident form point at it."),
        Col("name", "Name", 22, "text"),
        Col("role", "Role", 20, "text", validate=list_from("RoleList"),
            cf=status_cf()),
        Col("room", "Main room", 16, "text",
            validate=list_from("ClassroomList"),
            note="Where they usually are. The rota can send them "
                 "anywhere - this is just their base."),
        Col("phone", "Phone", 15, "text"),
        Col("email", "Email", 24, "text"),
        Col("address", "Address", 28, "text", hidden=True),
        Col("dob", "Date of birth", 11, "date", validate=date_check(),
            hidden=True),
        Col("hire_date", "Started", 11, "date", validate=date_check()),
        Col("status", "Status", 14, "status",
            validate=list_from("EmploymentStatusList"), cf=status_cf(),
            note="Only Active and Probation staff are counted as "
                 "available on the dashboard."),
        Col("end_date", "Left on", 11, "date", validate=date_check(),
            hidden=True),
        Col("years", "Years here", 9, "number1",
            formula=blank(
                "hire_date",
                'ROUND((IF({end_date}{r}="",TodayDate,{end_date}{r})-'
                '{hire_date}{r})/365,1)'),
            hidden=True),
        # ---------------------------------------------------- the pay ---
        Col("pay_type", "Paid as", 13, "text",
            validate=list_from("PayTypeList"),
            note="Hourly, salaried, daily or contractor. Payroll "
                 "handles all four differently, and volunteers cost "
                 "nothing at all."),
        Col("hourly", "Hourly rate", 11, "money",
            validate=number_check(0, 1000),
            note="For hourly staff. Leave blank for salaried ones - "
                 "the effective rate below works theirs out."),
        Col("salary", "Salary a year", 13, "money",
            validate=number_check(0, 1000000)),
        Col("contract_hours", "Hours a week", 11, "number1",
            validate=number_check(0, 80),
            note="Contracted hours. Used to turn a salary into an "
                 "hourly rate and to spot anyone being over-rota'd."),
        Col("day_rate", "Day rate", 11, "money",
            validate=number_check(0, 10000), hidden=True),
        Col("rate", "Effective hourly", 12, "money",
            formula=blank(
                "name",
                'ROUND(IF({pay_type}{r}="Volunteer - unpaid",0,'
                'IF({pay_type}{r}="Hourly",{hourly}{r},'
                'IF({pay_type}{r}="Salary",'
                'IF({contract_hours}{r}>0,'
                '{salary}{r}/(WeeksPaid*{contract_hours}{r}),0),'
                'IF({pay_type}{r}="Daily rate",'
                'IF({contract_hours}{r}>0,'
                '{day_rate}{r}/({contract_hours}{r}/5),0),'
                '{hourly}{r})))),2)'),
            cf=bar_cf(Palette.accent),
            note="One comparable hourly figure however somebody is "
                 "paid. The rota costs its hours at this rate."),
        Col("cost_hour", "Cost an hour", 11, "money",
            formula=blank(
                "name",
                'ROUND({rate}{r}*IF(({pay_type}{r}="Contractor")+'
                '({pay_type}{r}="Volunteer - unpaid")>0,1,'
                '1+PayrollBurden),2)'),
            hidden=True,
            note="The rate plus employer costs. Contractors and "
                 "volunteers carry none."),
        # ------------------------------------------------ certificates --
        Col("first_aid_exp", "First aid expires", 12, "date",
            validate=date_check()),
        Col("cpr_exp", "CPR expires", 12, "date",
            validate=date_check()),
        Col("safeguard_exp", "Safeguarding expires", 13, "date",
            validate=date_check()),
        Col("check_exp", "Background check expires", 14, "date",
            validate=date_check(),
            note="Whatever the check is called where you are - DBS, "
                 "Garda vetting, state background check. Record when "
                 "it needs renewing."),
        Col("check_done", "Check completed", 12, "date",
            validate=date_check(), hidden=True),
        Col("qualification", "Qualification", 22, "text",
            note="Level and subject. The thing an inspector asks to "
                 "see the certificate for."),
        Col("training", "Other training", 26, "wrap", hidden=True),
        Col("next_expiry", "Next expiry", 12, "date",
            formula=blank(
                "name",
                'IF(({first_aid_exp}{r}="")*({cpr_exp}{r}="")*'
                '({safeguard_exp}{r}="")*({check_exp}{r}="")=1,"",'
                'MIN(IF({first_aid_exp}{r}="",99999,'
                '{first_aid_exp}{r}),'
                'IF({cpr_exp}{r}="",99999,{cpr_exp}{r}),'
                'IF({safeguard_exp}{r}="",99999,{safeguard_exp}{r}),'
                'IF({check_exp}{r}="",99999,{check_exp}{r})))'),
            note="The soonest of the four. One column to filter on "
                 "instead of four to remember."),
        Col("days_to_expiry", "Days left", 9, "int",
            formula=blank(
                "next_expiry", '{next_expiry}{r}-TodayDate'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bold_bad"},
                {"type": "cell", "criteria": "<=",
                 "value": "=DocUrgentDays", "fmt": "bad"},
                {"type": "cell", "criteria": "<=",
                 "value": "=DocWarningDays", "fmt": "warn"}]),
        Col("certs_ok", "Certificates", 13, "status",
            formula=blank(
                "name",
                'IF({next_expiry}{r}="","Nothing recorded",'
                'IF({days_to_expiry}{r}<0,"Expired",'
                'IF({days_to_expiry}{r}<=DocUrgentDays,"Expiring",'
                'IF({days_to_expiry}{r}<=DocWarningDays,"Due",'
                '"On file"))))'),
            cf=status_cf()),
        Col("docs_missing", "Docs missing", 10, "int",
            formula=blank(
                "name",
                f'COUNTIFS({doc_who},{{id}}{{r}},{doc_req},"Yes")-'
                f'COUNTIFS({doc_who},{{id}}{{r}},{doc_req},"Yes",'
                f'{doc_status},"On file")-'
                f'COUNTIFS({doc_who},{{id}}{{r}},{doc_req},"Yes",'
                f'{doc_status},"Signed")-'
                f'COUNTIFS({doc_who},{{id}}{{r}},{doc_req},"Yes",'
                f'{doc_status},"Received")'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "bad"}]),
        # ------------------------------------------------- their week ---
        Col("hours_week", "Rota'd this week", 12, "number1",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Staff Schedule.hours]],'
                '[[Staff Schedule.staff]],{id}{r},'
                '[[Staff Schedule.week_start]],'
                'TodayDate-WEEKDAY(TodayDate,3)),2)'),
            note="Hours on the rota for the week that contains today."),
        Col("hours_month", "Hours this month", 12, "number1",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Staff Schedule.actual_hours]],'
                '[[Staff Schedule.staff]],{id}{r},'
                '[[Staff Schedule.month]],ThisMonthKey),2)'),
            hidden=True),
        Col("over_contract", "Over contract", 10, "number1",
            formula=blank(
                "name",
                'IF({contract_hours}{r}="","",'
                'ROUND(MAX(0,{hours_week}{r}-{contract_hours}{r}),2))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}],
            hidden=True),
        Col("pay_month", "Paid this month", 12, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Payroll Tracker.gross]],'
                '[[Payroll Tracker.staff]],{id}{r},'
                '[[Payroll Tracker.month]],ThisMonthKey),2)'),
            hidden=True),
        Col("children_room", "Children in room", 11, "int",
            formula=blank(
                "room",
                'COUNTIFS([[Child Enrollment.room]],{room}{r},'
                '[[Child Enrollment.holds_place]],"Yes")'),
            hidden=True),
        Col("leave_days", "Leave taken", 10, "number1",
            formula=blank(
                "name",
                'COUNTIFS([[Staff Schedule.staff]],{id}{r},'
                '[[Staff Schedule.shift]],"Leave")'),
            hidden=True),
        Col("em_name", "Emergency contact", 20, "text"),
        Col("em_phone", "Emergency phone", 15, "text"),
        Col("notes", "Notes", 26, "wrap"),
        Col("alert", "Action", 34, "status",
            formula=blank(
                "name",
                'IF(({status}{r}="Left")+({status}{r}="Inactive")>0,"",'
                'IF({next_expiry}{r}="","Missing - no certificate '
                'dates recorded",'
                'IF({days_to_expiry}{r}<0,"Expired - "&'
                'TEXT({next_expiry}{r},"d mmm")&", they cannot be '
                'counted in ratio",'
                'IF({days_to_expiry}{r}<=DocUrgentDays,'
                '"Expiring - renew within "&{days_to_expiry}{r}&" days",'
                'IF({docs_missing}{r}>0,"Missing - "&'
                '{docs_missing}{r}&" document(s) outstanding",'
                'IF({pay_type}{r}="","Add - how are they paid?",'
                'IF(IF({rate}{r}=0,{pay_type}{r}<>"Volunteer - '
                'unpaid",FALSE),"Add - no pay rate set",'
                'IF({days_to_expiry}{r}<=DocWarningDays,'
                '"Due - book the renewal",'
                'IF({em_phone}{r}="","Missing - no emergency contact",'
                '"OK")))))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Staff & Employee Database",
        subtitle="The team, what they are paid, and exactly how many "
                 "days until a certificate runs out.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.deep,
        freeze_cols=2,
        samples=sample(cfg).staff,
        help_text="One row per person. Fill in the pay rate and the "
                  "four certificate dates - the rota, the payroll and "
                  "the renewal warnings all run off them.",
        extra_notes=[
            "'Effective hourly' turns a salary, a day rate or an "
            "hourly wage into one comparable number, so the rota can "
            "cost a shift however that person happens to be paid.",
            "'Next expiry' is the soonest of first aid, CPR, "
            "safeguarding and the background check. Sort by it once "
            "a month and you will never be caught out.",
            "Employer on-costs are added on top of the rate for "
            "everyone except contractors and volunteers. Set the "
            "percentage on the Settings tab.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Team summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "The team", last)
    sheet.table_head(row, 0, [
        "On the books", "Active", "On leave", "Average hourly rate",
        "Average cost an hour", "Contracted hours a week",
        "Certificates expired", "Expiring soon", "Documents missing",
        "Average years here"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Staff Database.name]],"<>")',
        '=COUNTIFS([[Staff Database.status]],"Active")+'
        'COUNTIFS([[Staff Database.status]],"Probation")',
        '=COUNTIFS([[Staff Database.status]],"On leave")',
        '=IFERROR(ROUND(AVERAGEIF([[Staff Database.rate]],">0"),2),0)',
        '=IFERROR(ROUND(AVERAGEIF([[Staff Database.cost_hour]],">0"),'
        '2),0)',
        '=ROUND(SUMIFS([[Staff Database.contract_hours]],'
        '[[Staff Database.status]],"Active"),1)',
        '=COUNTIFS([[Staff Database.certs_ok]],"Expired")',
        '=COUNTIFS([[Staff Database.certs_ok]],"Expiring")+'
        'COUNTIFS([[Staff Database.certs_ok]],"Due")',
        '=SUMIFS([[Staff Database.docs_missing]],'
        '[[Staff Database.status]],"Active")',
        '=IFERROR(ROUND(AVERAGEIF([[Staff Database.years]],">0"),1),0)',
    ], ["int", "int", "int", "money", "money", "number1", "int", "int",
        "int", "number1"])
````

### `daycare_business_spreadsheet/sheets/schedule.py`

*256 lines · 2 functions*

> `Staff Schedule` - the rota, the clock and what the day actually cost.
>
> One row per person per shift.  Scheduled hours and clocked hours are
> kept in separate columns so the variance between the two is visible,
> because in childcare the variance is where the money goes.
>
> Overtime is allocated properly rather than fudged.  Each row knows how
> many hours that person has already been rota'd in the same week up to
> and including this shift, so the hours that tip them over the weekly
> threshold are the ones charged at the overtime multiplier - not the
> whole shift, and not nothing.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Rota summary |

**Module constants:** `NAME` ('Staff Schedule'), `NON_FLOOR` (4 items).

````python
"""`Staff Schedule` - the rota, the clock and what the day actually cost.

One row per person per shift.  Scheduled hours and clocked hours are
kept in separate columns so the variance between the two is visible,
because in childcare the variance is where the money goes.

Overtime is allocated properly rather than fudged.  Each row knows how
many hours that person has already been rota'd in the same week up to
and including this shift, so the hours that tip them over the weekly
threshold are the ones charged at the overtime multiplier - not the
whole shift, and not nothing.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check, time_check
from .common import (alert_cf, bar_cf, blank, id_col, lookup_staff,
                     lookup_staff_num, month_col, status_cf,
                     week_start_col, weekday_col)

NAME = "Staff Schedule"

#: Shifts that do not put an adult on the floor.
NON_FLOOR = ["Off", "Leave", "Training", "Admin"]


def spec(cfg: Config) -> SheetSpec:
    floor_test = "+".join(f'({{shift}}{{r}}="{s}")' for s in NON_FLOOR)

    columns = [
        id_col("date", "ShiftPrefix", digits=5, width=10,
               header="Shift ID"),
        Col("date", "Date", 11, "date", validate=date_check()),
        weekday_col("date"),
        week_start_col("date"),
        month_col("date"),
        Col("staff", "Staff", 10, "id", validate=list_from("StaffList")),
        Col("staff_name", "Name", 18, "text",
            formula=lookup_staff("name")),
        Col("role", "Role", 16, "text",
            formula=lookup_staff("role"), hidden=True),
        Col("pay_type", "Paid as", 12, "text",
            formula=lookup_staff("pay_type"), hidden=True),
        Col("room", "Room", 15, "text",
            validate=list_from("ClassroomList"),
            note="Where they are working that shift. The ratio "
                 "figures on the classroom tab count this column, "
                 "not the person's usual room."),
        Col("shift", "Shift", 13, "text",
            validate=list_from("ShiftTypeList"), cf=status_cf(),
            note="Off, Leave, Training and Admin do not put an adult "
                 "on the floor, so they are excluded from the ratio "
                 "counts."),
        Col("counts", "Counts in ratio?", 11, "flag",
            formula=blank("shift",
                          f'IF(({floor_test})>0,"No","Yes")'),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "good"},
                {"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "muted"}]),
        Col("leave_type", "Leave type", 15, "text",
            validate=list_from("LeaveTypeList"), hidden=True),
        # ------------------------------------------------- rota'd hours
        Col("start", "Starts", 10, "time", validate=time_check()),
        Col("finish", "Ends", 10, "time", validate=time_check()),
        Col("break_min", "Break (min)", 10, "int",
            validate=number_check(0, 240, decimal=False),
            note="Unpaid break. Leave blank to use the default on "
                 "Settings."),
        Col("hours", "Rota'd hours", 11, "number1",
            # anchored on the date, not the start time: a leave day
            # has no shift but still has to read zero rather than
            # blank, or every hours column downstream turns to
            # #VALUE!
            formula=blank(
                "date",
                'IF({start}{r}="",0,IF({finish}{r}="",0,'
                'ROUND(MAX(0,(MOD({finish}{r}-{start}{r},1))*24-'
                '(IF({break_min}{r}="",BreakMinutes,{break_min}{r})/60))'
                ',2)))'),
            cf=bar_cf()),
        # ------------------------------------------------- actual hours
        Col("clock_in", "Clocked in", 10, "time",
            validate=time_check(),
            note="Leave blank and the rota'd hours are used. Fill it "
                 "in and the variance column shows the difference."),
        Col("clock_out", "Clocked out", 10, "time",
            validate=time_check()),
        Col("actual_hours", "Actual hours", 11, "number1",
            formula=blank(
                "date",
                'IF({clock_in}{r}="",{hours}{r},'
                'IF({clock_out}{r}="",{hours}{r},'
                'ROUND(MAX(0,(MOD({clock_out}{r}-{clock_in}{r},1))*24-'
                '(IF({break_min}{r}="",BreakMinutes,{break_min}{r})/60))'
                ',2)))'),
            note="What they actually worked. Falls back to the rota "
                 "when nobody clocked in or out, so the payroll is "
                 "never short."),
        Col("variance", "Over / under", 10, "number1",
            formula=blank(
                "date",
                'ROUND({actual_hours}{r}-{hours}{r},2)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0.25,
                 "fmt": "warn"},
                {"type": "cell", "criteria": "<", "value": -0.25,
                 "fmt": "info"}],
            hidden=True),
        # ----------------------------------------------------- overtime
        Col("week_to_date", "Week to date", 11, "number1",
            formula=blank(
                "date",
                'ROUND(SUMIFS({actual_hours_first}:{actual_hours}{r},'
                '{staff_first}:{staff}{r},{staff}{r},'
                '{week_start_first}:{week_start}{r},'
                '{week_start}{r}),2)'),
            hidden=True,
            note="Hours this person has worked in the same week up to "
                 "and including this shift. Overtime is worked out "
                 "from it."),
        Col("ot_hours", "Overtime", 9, "number1",
            formula=blank(
                "date",
                'ROUND(MAX(0,MIN({actual_hours}{r},'
                '{week_to_date}{r}-OvertimeAfter)),2)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}],
            note="Only the hours on this shift that tip the person "
                 "over the weekly threshold - not the whole shift."),
        Col("basic_hours", "Basic hours", 10, "number1",
            formula=blank(
                "date",
                'ROUND({actual_hours}{r}-{ot_hours}{r},2)'),
            hidden=True),
        Col("rate", "Rate", 10, "money",
            formula=lookup_staff_num("rate"), hidden=True),
        Col("cost", "Shift cost", 11, "money",
            formula=blank(
                "date",
                'ROUND(({basic_hours}{r}+{ot_hours}{r}*'
                'OvertimeMultiplier)*{rate}{r},2)'),
            cf=bar_cf(Palette.accent),
            note="Basic hours plus overtime at the multiplier, at "
                 "that person's effective rate. Before employer "
                 "on-costs."),
        Col("cost_full", "With on-costs", 11, "money",
            formula=blank(
                "date",
                'ROUND({cost}{r}*(1+IF(({pay_type}{r}="Contractor")+'
                '({pay_type}{r}="Volunteer - unpaid")>0,0,'
                'PayrollBurden)),2)'),
            hidden=True),
        # ----------------------------------------------------- cover ---
        Col("substitute", "Covered by", 10, "id",
            validate=list_from("StaffList"),
            note="When somebody else picks the shift up. The original "
                 "name stays so you can see what happened."),
        Col("sub_name", "Cover", 16, "text",
            formula=lookup_staff("name", "substitute"), hidden=True),
        Col("clash", "Double booked", 9, "int",
            formula=blank(
                "date",
                'IF({start}{r}="",0,'
                'COUNTIFS({staff_first}:{staff_last},{staff}{r},'
                '{date_first}:{date_last},{date}{r},'
                '{start_first}:{start_last},"<"&{finish}{r},'
                '{finish_first}:{finish_last},">"&{start}{r},'
                '{counts_first}:{counts_last},"Yes"))'),
            hidden=True,
            note="Overlapping shifts for the same person on the same "
                 "day. Anything above 1 is a clash."),
        Col("notes", "Notes", 24, "wrap"),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "date",
                'IF({staff}{r}="","Add - nobody on this shift",'
                'IF({staff_name}{r}="Not in Staff Database",'
                '"Check - staff ID not found",'
                'IF({shift}{r}="","Add - what kind of shift?",'
                'IF({clash}{r}>1,"Duplicate - double booked at the '
                'same time",'
                'IF(IF({counts}{r}<>"Yes",FALSE,{room}{r}=""),'
                '"Add - no room, so it does not count in ratio",'
                'IF(IF({counts}{r}<>"Yes",FALSE,{start}{r}=""),'
                '"Add - no start time",'
                'IF(IF({clock_in}{r}="",FALSE,'
                'IF({clock_out}{r}<>"",FALSE,{date}{r}<TodayDate)),'
                '"Missing - clocked in and never out",'
                'IF({ot_hours}{r}>0,"Check - "&'
                'TEXT({ot_hours}{r},"0.0")&" overtime hours",'
                '"OK"))))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Staff Schedule & Timesheet",
        subtitle="The rota, the clock, the overtime and what each "
                 "shift costs - with the ratio counts feeding the "
                 "classroom tab.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).schedule,
        help_text="One row per person per shift. Rota it in advance, "
                  "fill in the clock times afterwards, and the hours, "
                  "the overtime and the cost all follow.",
        extra_notes=[
            "Actual hours fall back to the rota'd hours when nobody "
            "clocked in or out, so a forgotten clock-out never "
            "silently removes somebody's pay.",
            "Overtime is allocated shift by shift: the workbook adds "
            "up what each person has already worked that week and "
            "only charges the hours above the threshold at the "
            "higher rate.",
            "'Counts in ratio' is what the classroom tab adds up. "
            "Training, admin, leave and days off are excluded, which "
            "is why a room can look staffed on the rota and still be "
            "short on the floor.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Rota summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Rota summary", last)
    sheet.table_head(row, 0, [
        "Shifts", "Rota'd hours", "Actual hours", "Variance",
        "Overtime hours", "Wage cost", "Cost with on-costs",
        "Average cost an hour", "Leave days", "Double bookings",
        "This month cost"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Staff Schedule.date]],"<>")',
        '=ROUND(SUM([[Staff Schedule.hours]]),1)',
        '=ROUND(SUM([[Staff Schedule.actual_hours]]),1)',
        '=ROUND(SUM([[Staff Schedule.actual_hours]])-'
        'SUM([[Staff Schedule.hours]]),1)',
        '=ROUND(SUM([[Staff Schedule.ot_hours]]),1)',
        '=ROUND(SUM([[Staff Schedule.cost]]),2)',
        '=ROUND(SUM([[Staff Schedule.cost_full]]),2)',
        '=IF(SUM([[Staff Schedule.actual_hours]])<=0,"",'
        'ROUND(SUM([[Staff Schedule.cost]])/'
        'SUM([[Staff Schedule.actual_hours]]),2))',
        '=COUNTIFS([[Staff Schedule.shift]],"Leave")',
        '=COUNTIFS([[Staff Schedule.clash]],">1")',
        '=ROUND(SUMIFS([[Staff Schedule.cost]],'
        '[[Staff Schedule.month]],ThisMonthKey),2)',
    ], ["int", "number1", "number1", "number1", "number1", "money",
        "money", "money", "int", "int", "money"])
````

### `daycare_business_spreadsheet/sheets/payroll.py`

*242 lines · 2 functions*

> `Payroll Tracker` - turning rota hours into a wage bill.
>
> One row per person per pay run.  Hours come straight off the Staff
> Schedule for the dates the run covers, split into basic and overtime,
> and costed at that person's effective rate.  Salaried staff are paid
> their slice of the year regardless of hours; contractors and
> volunteers carry no employer on-costs at all.
>
> `Total cost` is the only wage figure the profit and loss reads, which
> is why the expense tracker deliberately skips its own wage categories
> unless you tell it otherwise on Settings.  Paying twice for the same
> staff is the easiest mistake to make in a small business spreadsheet
> and the hardest to spot afterwards.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Payroll summary |

**Module constants:** `NAME` ('Payroll Tracker').

````python
"""`Payroll Tracker` - turning rota hours into a wage bill.

One row per person per pay run.  Hours come straight off the Staff
Schedule for the dates the run covers, split into basic and overtime,
and costed at that person's effective rate.  Salaried staff are paid
their slice of the year regardless of hours; contractors and
volunteers carry no employer on-costs at all.

`Total cost` is the only wage figure the profit and loss reads, which
is why the expense tracker deliberately skips its own wage categories
unless you tell it otherwise on Settings.  Paying twice for the same
staff is the easiest mistake to make in a small business spreadsheet
and the hardest to spot afterwards.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, id_col, lookup_staff,
                     lookup_staff_num, month_col, status_cf)

NAME = "Payroll Tracker"


def spec(cfg: Config) -> SheetSpec:
    sch_staff = "[[Staff Schedule.staff]]"
    sch_date = "[[Staff Schedule.date]]"

    def hours_in_period(column: str) -> str:
        return (f'ROUND(SUMIFS([[Staff Schedule.{column}]],'
                f'{sch_staff},{{staff}}{{r}},'
                f'{sch_date},">="&{{period_start}}{{r}},'
                f'{sch_date},"<="&{{period_end}}{{r}}),2)')

    columns = [
        id_col("period_end", "PayrollPrefix", digits=4, width=11,
               header="Pay run ID"),
        Col("period_end", "Period ends", 12, "date",
            validate=date_check(),
            note="The last day the run covers. Monthly totals group "
                 "on this date."),
        month_col("period_end"),
        Col("period_start", "Period starts", 12, "date",
            validate=date_check(),
            note="The first day. The hours below are added up between "
                 "these two dates, inclusive."),
        Col("pay_period", "Run type", 13, "text",
            validate=list_from("PayPeriodList"), cf=status_cf()),
        Col("staff", "Staff", 10, "id", validate=list_from("StaffList")),
        Col("staff_name", "Name", 18, "text",
            formula=lookup_staff("name")),
        Col("pay_type", "Paid as", 13, "text",
            formula=lookup_staff("pay_type")),
        Col("rate", "Rate", 10, "money",
            formula=lookup_staff_num("rate"),
            note="The effective hourly rate from the staff tab."),
        # --------------------------------------------------- the hours -
        Col("basic_hours", "Basic hours", 11, "number1",
            formula=blank("staff", hours_in_period("basic_hours")),
            cf=bar_cf(),
            note="Straight off the rota for the dates this run "
                 "covers. Nothing to add up by hand."),
        Col("ot_hours", "Overtime hours", 11, "number1",
            formula=blank("staff", hours_in_period("ot_hours")),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}]),
        Col("total_hours", "Total hours", 11, "number1",
            formula=blank(
                "staff",
                'ROUND({basic_hours}{r}+{ot_hours}{r},2)'),
            hidden=True),
        Col("leave_days", "Leave days", 9, "int",
            formula=blank(
                "staff",
                f'COUNTIFS({sch_staff},{{staff}}{{r}},'
                f'{sch_date},">="&{{period_start}}{{r}},'
                f'{sch_date},"<="&{{period_end}}{{r}},'
                f'[[Staff Schedule.shift]],"Leave")'),
            hidden=True),
        # ---------------------------------------------------- the pay --
        Col("periods_year", "Runs a year", 9, "int",
            formula=blank(
                "pay_period",
                'IF({pay_period}{r}="Weekly",52,'
                'IF({pay_period}{r}="Fortnightly",26,'
                'IF({pay_period}{r}="Four weekly",13,'
                'IF({pay_period}{r}="Twice monthly",24,'
                'IF({pay_period}{r}="Monthly",12,1)))))'),
            hidden=True),
        Col("basic_pay", "Basic pay", 11, "money",
            formula=blank(
                "staff",
                'ROUND(IF({pay_type}{r}="Salary",'
                'IFERROR(INDEX([[Staff Database.salary]],'
                'MATCH({staff}{r},[[Staff Database.id]],0))*1,0)/'
                'MAX(1,{periods_year}{r}),'
                '{basic_hours}{r}*{rate}{r}),2)'),
            note="Hourly staff are paid for the hours they worked. "
                 "Salaried staff get their slice of the year however "
                 "many hours the rota shows."),
        Col("ot_pay", "Overtime pay", 11, "money",
            formula=blank(
                "staff",
                'ROUND(IF({pay_type}{r}="Salary",0,'
                '{ot_hours}{r}*{rate}{r}*OvertimeMultiplier),2)')),
        Col("bonus", "Bonus", 10, "money",
            validate=number_check(0, 100000)),
        Col("adjust", "Adjustment", 11, "money",
            validate=number_check(-100000, 100000),
            note="Anything one-off: a backdated rise, a correction "
                 "from last run, an agreed deduction (type it "
                 "negative)."),
        Col("gross", "Gross pay", 12, "money",
            formula=blank(
                "staff",
                'ROUND({basic_pay}{r}+{ot_pay}{r}+'
                'IF({bonus}{r}="",0,{bonus}{r})+'
                'IF({adjust}{r}="",0,{adjust}{r}),2)'),
            cf=bar_cf(Palette.accent)),
        Col("employer_cost", "Employer costs", 12, "money",
            formula=blank(
                "staff",
                'ROUND(IF(({pay_type}{r}="Contractor")+'
                '({pay_type}{r}="Volunteer - unpaid")>0,0,'
                '{gross}{r}*PayrollBurden),2)'),
            note="Employer tax, insurance and holiday cover as a "
                 "percentage of gross - set it on Settings. "
                 "Contractors and volunteers carry none."),
        Col("total_cost", "Total cost", 12, "money",
            formula=blank(
                "staff",
                'ROUND({gross}{r}+{employer_cost}{r},2)'),
            cf=bar_cf(Palette.deep),
            note="What this person cost the business this run. The "
                 "only wage figure the profit and loss reads."),
        Col("deductions", "Deductions", 11, "money",
            validate=number_check(0, 100000), hidden=True,
            note="Tax and anything else withheld. For your own "
                 "records - this workbook does not calculate anybody's "
                 "tax."),
        Col("net_pay", "Net to pay", 12, "money",
            formula=blank(
                "staff",
                'ROUND({gross}{r}-IF({deductions}{r}="",0,'
                '{deductions}{r}),2)'),
            note="Gross less whatever you entered as deductions. It "
                 "is arithmetic, not a tax calculation."),
        Col("cost_hour", "Cost an hour", 10, "money",
            formula=blank(
                "staff",
                'IF({total_hours}{r}<=0,"",'
                'ROUND({total_cost}{r}/{total_hours}{r},2))'),
            hidden=True),
        # ------------------------------------------------------- paid --
        Col("paid", "Paid?", 8, "flag",
            validate=list_from("YesNoList"), cf=status_cf()),
        Col("paid_date", "Paid on", 11, "date", validate=date_check()),
        Col("method", "Method", 14, "text",
            validate=list_from("PaymentMethodList"), hidden=True),
        Col("notes", "Notes", 24, "wrap"),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "period_end",
                'IF({staff}{r}="","Add - nobody on this pay run",'
                'IF({staff_name}{r}="Not in Staff Database",'
                '"Check - staff ID not found",'
                'IF({period_start}{r}="","Add - no period start date",'
                'IF({period_start}{r}>{period_end}{r},'
                '"Check - the period runs backwards",'
                'IF(IF({pay_type}{r}="Salary",FALSE,'
                '{total_hours}{r}<=0),'
                '"Check - no rota hours in this period",'
                'IF(IF({pay_type}{r}="Volunteer - unpaid",FALSE,'
                '{gross}{r}<=0),"Check - gross pay is zero",'
                'IF(IF({paid}{r}="Yes",FALSE,'
                '{period_end}{r}<TodayDate-7),'
                '"Overdue - pay run not marked as paid",'
                'IF({paid}{r}<>"Yes","Due - not paid yet","OK")))))))'
                ')'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Payroll Tracker",
        subtitle="Hours from the rota, pay from the staff tab, "
                 "employer costs on top - one row per person per run.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.deep,
        freeze_cols=2,
        samples=sample(cfg).payroll,
        help_text="Set the two period dates and pick the person. The "
                  "hours, the overtime, the gross and the employer "
                  "costs all come from the rota and the staff record.",
        extra_notes=[
            "This is not a payroll bureau. It works out what staff "
            "cost you; it does not calculate anybody's income tax, "
            "social security or statutory deductions, and it files "
            "nothing with anybody.",
            "'Total cost' is gross plus employer on-costs, and it is "
            "the figure the profit and loss uses. Leave the wage "
            "categories out of the Expense Tracker, or switch "
            "'Payroll sits in expenses' to Yes on Settings - one or "
            "the other, never both.",
            "Salaried staff are paid their slice of the year whatever "
            "the rota says, so a quiet week does not quietly cut "
            "their pay.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Payroll summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Payroll summary", last)
    sheet.table_head(row, 0, [
        "Pay runs", "People paid", "Basic hours", "Overtime hours",
        "Gross pay", "Employer costs", "Total cost",
        "Average cost an hour", "This month cost", "Not yet paid"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Payroll Tracker.period_end]],"<>")',
        '=COUNTIFS([[Payroll Tracker.paid]],"Yes")',
        '=ROUND(SUM([[Payroll Tracker.basic_hours]]),1)',
        '=ROUND(SUM([[Payroll Tracker.ot_hours]]),1)',
        '=ROUND(SUM([[Payroll Tracker.gross]]),2)',
        '=ROUND(SUM([[Payroll Tracker.employer_cost]]),2)',
        '=ROUND(SUM([[Payroll Tracker.total_cost]]),2)',
        '=IF(SUM([[Payroll Tracker.total_hours]])<=0,"",'
        'ROUND(SUM([[Payroll Tracker.total_cost]])/'
        'SUM([[Payroll Tracker.total_hours]]),2))',
        '=ROUND(SUMIFS([[Payroll Tracker.total_cost]],'
        '[[Payroll Tracker.month]],ThisMonthKey),2)',
        '=ROUND(SUMIFS([[Payroll Tracker.net_pay]],'
        '[[Payroll Tracker.paid]],"<>Yes"),2)',
    ], ["int", "int", "number1", "number1", "money", "money", "money",
        "money", "money", "money"])
````

---

## 8. Rooms, stock and paperwork

Occupancy and empty seats per room, the live staff-to-child ratio against the band that room is set to, the stock cupboard counted the honest way, the parent conversation log and the document tracker that watches every renewal date in the building.

### `daycare_business_spreadsheet/sheets/classrooms.py`

*320 lines · 2 functions*

> `Classroom Management` - how full each room is, and whether it is legal.
>
> Two numbers run a childcare setting: how many children are in a room,
> and how many adults are standing in it.  This tab holds both, works out
> the seats you have left to sell and the staff the room needs right now,
> and goes red before the inspector does.
>
> Every ratio comes from the age band table on Settings, and every one of
> them can be overridden per room, because the number of two-year-olds one
> adult may supervise is a local licensing decision and not something a
> spreadsheet should have an opinion about.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Whole-setting capacity summary |

**Module constants:** `NAME` ('Classroom Management').

````python
"""`Classroom Management` - how full each room is, and whether it is legal.

Two numbers run a childcare setting: how many children are in a room,
and how many adults are standing in it.  This tab holds both, works out
the seats you have left to sell and the staff the room needs right now,
and goes red before the inspector does.

Every ratio comes from the age band table on Settings, and every one of
them can be overridden per room, because the number of two-year-olds one
adult may supervise is a local licensing decision and not something a
spreadsheet should have an opinion about.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check
from .common import (alert_cf, blank, bar_cf, id_col, lookup_staff,
                     status_cf)

NAME = "Classroom Management"


def spec(cfg: Config) -> SheetSpec:
    child_room = "[[Child Enrollment.room]]"
    child_status = "[[Child Enrollment.status]]"
    att_room = "[[Attendance.room]]"
    att_date = "[[Attendance.date]]"
    att_status = "[[Attendance.status]]"
    sch_room = "[[Staff Schedule.room]]"
    sch_date = "[[Staff Schedule.date]]"
    sch_counts = "[[Staff Schedule.counts]]"

    # Children holding a place in this room: Active, On hold or Notice
    # given.  Withdrawn and Graduated free the seat up.
    def occupying(extra: str = "") -> str:
        return ("+".join(
            f'COUNTIFS({child_room},{{name}}{{r}},{child_status},"{s}"'
            f'{extra})'
            for s in ("Active", "On hold", "Notice given")))

    columns = [
        id_col("name", "ClassroomPrefix", digits=2, width=10,
               header="Room ID",
               note="Automatic. Children and shifts point at the room "
                    "name, not this ID, so renaming a room is a one "
                    "cell job."),
        Col("name", "Room name", 20, "text",
            note="What you actually call it - Acorn Room, Baby Unit, "
                 "Pre-K 2. Every other tab picks the room from this "
                 "list, so type it once and type it carefully."),
        Col("age_band", "Age band", 15, "text",
            validate=list_from("AgeBandList"),
            note="Sets the default staff-to-child ratio for this room. "
                 "The bands and their ratios are yours to edit on "
                 "Settings."),
        Col("age_from", "From (months)", 11, "int",
            formula=blank(
                "age_band",
                'IFERROR(INDEX(AgeBandFrom,MATCH({age_band}{r},'
                'AgeBandList,0))*1,0)')),
        Col("age_to", "To (months)", 11, "int",
            formula=blank(
                "age_band",
                'IFERROR(INDEX(AgeBandTo,MATCH({age_band}{r},'
                'AgeBandList,0))*1,0)')),
        Col("capacity", "Licensed places", 13, "int",
            validate=number_check(0, 500, decimal=False),
            note="The most children allowed in this room at once, from "
                 "your licence - not the number of chairs."),
        Col("ratio_default", "Band ratio", 10, "int",
            formula=blank(
                "age_band",
                'IFERROR(INDEX(AgeBandRatio,MATCH({age_band}{r},'
                'AgeBandList,0))*1,0)'),
            hidden=True),
        Col("ratio_set", "Ratio override", 12, "int",
            validate=number_check(0, 40, decimal=False),
            note="Children per adult, if this room differs from the "
                 "band default. Leave blank to use the band."),
        Col("ratio", "Children per adult", 14, "int",
            formula=blank(
                "name",
                'IF({ratio_set}{r}<>"",{ratio_set}{r},'
                'IF({ratio_default}{r}>0,{ratio_default}{r},""))'),
            note="The ratio this room is actually judged against."),
        Col("max_group", "Max group size", 12, "int",
            formula=blank(
                "age_band",
                'IFERROR(INDEX(AgeBandMax,MATCH({age_band}{r},'
                'AgeBandList,0))*1,0)'),
            hidden=True),
        Col("teacher", "Lead teacher", 11, "id",
            validate=list_from("StaffList"),
            note="Staff ID of whoever leads the room."),
        Col("teacher_name", "Teacher", 18, "text",
            formula=lookup_staff("name", "teacher")),
        Col("assistant", "Assistant", 11, "id",
            validate=list_from("StaffList")),
        Col("assistant_name", "Assistant name", 18, "text",
            formula=lookup_staff("name", "assistant"), hidden=True),
        Col("room_status", "Room status", 12, "status",
            validate=list_from("RoomStatusList"), cf=status_cf(),
            note="Closed and Planned rooms are left out of the "
                 "occupancy figures on the dashboard."),
        # ------------------------------------------------- who is in it --
        Col("enrolled", "Children enrolled", 14, "int",
            formula=blank("name", occupying()),
            cf=bar_cf(),
            note="Active, On hold and Notice given - everyone still "
                 "holding a place. Withdrawn children free the seat."),
        Col("active", "Active only", 11, "int",
            formula=blank(
                "name",
                f'COUNTIFS({child_room},{{name}}{{r}},'
                f'{child_status},"Active")'),
            hidden=True),
        Col("starting", "Starting soon", 12, "int",
            formula=blank(
                "name",
                f'COUNTIFS({child_room},{{name}}{{r}},'
                f'{child_status},"Pending start")'),
            note="Offered a place and not in the room yet. They will "
                 "take a seat - plan for them."),
        Col("leaving", "Leaving soon", 11, "int",
            formula=blank(
                "name",
                f'COUNTIFS({child_room},{{name}}{{r}},'
                f'{child_status},"Notice given")'),
            hidden=True),
        Col("seats", "Seats free", 10, "int",
            formula=blank(
                "name",
                'IF({capacity}{r}="","",{capacity}{r}-{enrolled}{r})'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bold_bad"},
                {"type": "cell", "criteria": "=", "value": 0,
                 "fmt": "warn"},
                {"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "good"}],
            note="Places you could still sell today. A negative number "
                 "means you are over your licence."),
        Col("occupancy", "Occupancy", 10, "percent",
            formula=blank(
                "name",
                'IF({capacity}{r}="","",IF({capacity}{r}<=0,"",'
                '{enrolled}{r}/{capacity}{r}))'),
            cf=[{"type": "cell", "criteria": ">", "value": 1,
                 "fmt": "bold_bad"},
                {"type": "cell", "criteria": ">=",
                 "value": "=RatioWarning", "fmt": "warn"},
                {"type": "cell", "criteria": "<",
                 "value": "=RatioWarning", "fmt": "good"}],
            note="Enrolled divided by licensed places."),
        Col("waiting", "On the waitlist", 13, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Waitlist.room]],{name}{r},'
                '[[Waitlist.converted]],"No")'),
            note="Families still waiting for this room. A full room "
                 "with a queue behind it is a price rise, not a "
                 "problem."),
        # ------------------------------------------------------- today --
        Col("present_today", "In today", 10, "int",
            formula=blank(
                "name",
                "+".join(
                    f'COUNTIFS({att_room},{{name}}{{r}},'
                    f'{att_date},TodayDate,{att_status},"{s}")'
                    for s in ("Present", "Late", "Early pickup"))),
            note="From the attendance register, for today's date."),
        Col("staff_today", "Adults in today", 13, "number1",
            formula=blank(
                "name",
                f'SUMIFS([[Staff Schedule.hours]],{sch_room},'
                f'{{name}}{{r}},{sch_date},TodayDate,'
                f'{sch_counts},"Yes")/MAX(1,OpenHoursDay)'),
            note="Rota hours in this room today divided by the length "
                 "of the day - so a half-day cover counts as half an "
                 "adult, which is what the ratio really means."),
        Col("staff_needed", "Adults needed", 12, "number1",
            formula=blank(
                "name",
                'IF({ratio}{r}="","",IF({ratio}{r}<=0,"",'
                'ROUNDUP({present_today}{r}/{ratio}{r},0)))'),
            note="For the children actually in the building today."),
        Col("staff_needed_full", "Adults at full", 12, "number1",
            formula=blank(
                "name",
                'IF({ratio}{r}="","",IF({ratio}{r}<=0,"",'
                'ROUNDUP({enrolled}{r}/{ratio}{r},0)))'),
            hidden=True,
            note="What the room needs on a day when everybody turns "
                 "up. Staff to this, not to today."),
        Col("ratio_now", "Ratio today", 11, "text",
            formula=blank(
                "name",
                'IF({staff_today}{r}<=0,IF({present_today}{r}=0,'
                '"No session","No cover"),'
                '"1:"&ROUND({present_today}{r}/{staff_today}{r},1))'),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "No cover", "fmt": "bold_bad"}],
            note="One adult to this many children, right now."),
        Col("ratio_ok", "Ratio met?", 10, "flag",
            formula=blank(
                "name",
                'IF({present_today}{r}=0,"No session",'
                'IF({ratio}{r}="","Set ratio",'
                'IF({staff_today}{r}>={staff_needed}{r},"Yes","No")))'),
            cf=[{"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "bold_bad"},
                {"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "good"}]),
        # -------------------------------------------------- the money ----
        Col("revenue", "Room revenue a week", 15, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Child Enrollment.weekly_net]],'
                '[[Child Enrollment.room]],{name}{r}),2)'),
            cf=bar_cf(Palette.accent),
            note="What the children in this room are contracted to pay "
                 "in a week, after discounts and subsidy."),
        Col("revenue_seat", "Per place a week", 13, "money",
            formula=blank(
                "name",
                'IF({capacity}{r}="","",IF({capacity}{r}<=0,"",'
                'ROUND({revenue}{r}/{capacity}{r},2)))'),
            hidden=True,
            note="Weekly revenue divided by licensed places - the "
                 "number to compare rooms on."),
        Col("empty_cost", "Empty places cost", 14, "money",
            formula=blank(
                "name",
                'IF({seats}{r}="","",IF({enrolled}{r}<=0,"",'
                'ROUND(MAX(0,{seats}{r})*({revenue}{r}/'
                'MAX(1,{enrolled}{r})),2)))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}],
            note="What those empty seats would bring in a week at the "
                 "rate the room already charges. This is the number to "
                 "put next to an advertising budget."),
        Col("notes", "Notes", 30, "wrap"),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "name",
                'IF({capacity}{r}="","Add - set the licensed places",'
                'IF({ratio}{r}="","Add - pick an age band or ratio",'
                'IF({enrolled}{r}>{capacity}{r},'
                '"Over capacity - stop enrolling",'
                'IF(IF({present_today}{r}=0,FALSE,'
                '{staff_today}{r}<{staff_needed}{r}),'
                '"Over ratio - get cover into the room",'
                'IF({teacher}{r}="","Add - no lead teacher",'
                'IF({enrolled}{r}>={capacity}{r},'
                '"Nearly full - open the waitlist",'
                'IF({occupancy}{r}="","OK",'
                'IF({occupancy}{r}>=RatioWarning,'
                '"Nearly full - last places",'
                'IF({seats}{r}>=3,"Space - '
                + str(cfg.currency_symbol) +
                '"&TEXT({empty_cost}{r},"#,##0")&" a week empty",'
                '"OK")))))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Classroom Management",
        subtitle="Capacity, occupancy, staff-to-child ratios and what "
                 "every empty place is costing you a week.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).classrooms,
        help_text="Name each room, give it an age band and the number "
                  "of places on your licence. Enrolment, attendance, "
                  "ratios and revenue all fill themselves in from the "
                  "other tabs.",
        extra_notes=[
            "Ratios come from the age band table on Settings & Lists. "
            "They are a starting point, not a legal opinion - set them "
            "to whatever your own licence says and the whole workbook "
            "follows.",
            "'Adults in today' is rota hours divided by the length of "
            "your day, so somebody covering a four hour shift counts "
            "as part of an adult rather than a whole one.",
            "'Empty places cost' is the fastest argument for a "
            "marketing budget you will find in this file.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Whole-setting capacity summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "The whole setting", last)
    sheet.table_head(row, 0, [
        "Rooms open", "Licensed places", "Children enrolled",
        "Seats free", "Occupancy", "In today", "Adults in today",
        "Rooms over ratio", "Weekly contracted revenue"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIFS([[Classroom Management.room_status]],"Open")',
        '=SUM([[Classroom Management.capacity]])',
        '=SUM([[Classroom Management.enrolled]])',
        '=SUM([[Classroom Management.capacity]])-'
        'SUM([[Classroom Management.enrolled]])',
        '=IF(SUM([[Classroom Management.capacity]])<=0,"",'
        'SUM([[Classroom Management.enrolled]])/'
        'SUM([[Classroom Management.capacity]]))',
        '=SUM([[Classroom Management.present_today]])',
        '=ROUND(SUM([[Classroom Management.staff_today]]),1)',
        '=COUNTIFS([[Classroom Management.ratio_ok]],"No")',
        '=ROUND(SUM([[Classroom Management.revenue]]),2)',
    ], ["int", "int", "int", "int", "percent", "int", "number1", "int",
        "money"])
````

### `daycare_business_spreadsheet/sheets/supplies.py`

*217 lines · 2 functions*

> `Supply Inventory` - nappies, wipes, paint, paper and lunch.
>
> A childcare setting does not run out of toys.  It runs out of nappies
> in size 4 on a Tuesday afternoon, and somebody drives to the shop.
> This tab exists to stop that: opening stock plus everything in, less
> everything used, wasted or written off, equals what is on the shelf -
> and how many weeks that will last at the rate you have actually been
> using it.
>
>     stock = opening + received - used - waste + adjustments

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Store cupboard summary |

**Module constants:** `NAME` ('Supply Inventory').

````python
"""`Supply Inventory` - nappies, wipes, paint, paper and lunch.

A childcare setting does not run out of toys.  It runs out of nappies
in size 4 on a Tuesday afternoon, and somebody drives to the shop.
This tab exists to stop that: opening stock plus everything in, less
everything used, wasted or written off, equals what is on the shelf -
and how many weeks that will last at the rate you have actually been
using it.

    stock = opening + received - used - waste + adjustments
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, bar_cf, blank, id_col, status_cf)

NAME = "Supply Inventory"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "SupplyPrefix", digits=3, width=10,
               header="Item ID"),
        Col("name", "Item", 26, "text",
            note="Be specific. 'Nappies size 4' and 'Nappies size 5' "
                 "run out on different days."),
        Col("category", "Category", 20, "text",
            validate=list_from("SupplyCategoryList"), cf=status_cf()),
        Col("unit", "Unit", 11, "text",
            note="Pack, box, litre, each. Whatever you count in."),
        Col("pack_size", "Items per unit", 11, "int",
            validate=number_check(0, 10000, decimal=False),
            hidden=True),
        Col("supplier", "Supplier", 20, "text"),
        Col("supplier_phone", "Supplier phone", 15, "text",
            hidden=True),
        Col("lead_days", "Lead time (days)", 11, "int",
            validate=number_check(0, 180, decimal=False), hidden=True,
            note="How long an order takes to arrive. The reorder "
                 "warning allows for it."),
        # -------------------------------------------------- movements --
        Col("opening", "Opening stock", 11, "number1",
            validate=number_check(0, 100000),
            note="What was on the shelf when you started tracking. "
                 "The date is on Settings."),
        Col("received", "Received", 10, "number1",
            validate=number_check(0, 100000),
            note="Everything that has come in since. Add to this "
                 "figure each time a delivery arrives."),
        Col("used", "Used", 10, "number1",
            validate=number_check(0, 100000)),
        Col("waste", "Waste / damaged", 11, "number1",
            validate=number_check(0, 100000)),
        Col("adjust", "Adjustment", 10, "number1",
            validate=number_check(-100000, 100000),
            note="For stock counts that do not agree with the "
                 "arithmetic. Positive adds, negative takes away."),
        Col("stock", "On the shelf", 11, "number1",
            formula=blank(
                "name",
                'ROUND(IF({opening}{r}="",0,{opening}{r})+'
                'IF({received}{r}="",0,{received}{r})-'
                'IF({used}{r}="",0,{used}{r})-'
                'IF({waste}{r}="",0,{waste}{r})+'
                'IF({adjust}{r}="",0,{adjust}{r}),2)'),
            cf=[{"type": "cell", "criteria": "<=", "value": 0,
                 "fmt": "bold_bad"},
                {"type": "data_bar", "bar_color": Palette.primary_mid}],
            note="Opening plus received, less used, less waste, plus "
                 "adjustments. If it goes negative, something was "
                 "used that was never booked in."),
        # ----------------------------------------------------- money ---
        Col("cost", "Cost a unit", 11, "money",
            validate=number_check(0, 100000)),
        Col("value", "Stock value", 11, "money",
            formula=blank(
                "name",
                'ROUND(MAX(0,{stock}{r})*IF({cost}{r}="",0,'
                '{cost}{r}),2)'),
            cf=bar_cf(Palette.accent)),
        Col("spend", "Spent so far", 11, "money",
            formula=blank(
                "name",
                'ROUND((IF({received}{r}="",0,{received}{r}))*'
                'IF({cost}{r}="",0,{cost}{r}),2)'),
            hidden=True),
        # ------------------------------------------------- how long ----
        Col("weekly_use", "Used a week", 11, "number1",
            formula=blank(
                "name",
                'ROUND(IF({used}{r}="",0,{used}{r})/'
                'MAX(1,TrackingWeeks),2)'),
            note="Everything used, divided by the weeks you have been "
                 "tracking. Your real rate, not an estimate."),
        Col("weeks_cover", "Weeks left", 10, "number1",
            formula=blank(
                "name",
                'IF({weekly_use}{r}<=0,"",'
                'ROUND(MAX(0,{stock}{r})/{weekly_use}{r},1))'),
            cf=[{"type": "cell", "criteria": "<=", "value": 1,
                 "fmt": "bold_bad"},
                {"type": "cell", "criteria": "<=",
                 "value": "=StockWarningWeeks", "fmt": "warn"},
                {"type": "cell", "criteria": ">",
                 "value": "=ReorderWeeks", "fmt": "good"}],
            note="How long what is on the shelf will last at the rate "
                 "you have been getting through it."),
        Col("min_level", "Minimum level", 11, "number1",
            validate=number_check(0, 100000),
            note="Never go below this. Leave blank and the warning "
                 "uses the weeks-of-cover figure instead."),
        Col("order_qty", "Order this much", 12, "number1",
            formula=blank(
                "name",
                'IF({weekly_use}{r}<=0,'
                'IF({min_level}{r}="","",'
                'MAX(0,ROUND({min_level}{r}-{stock}{r},0))),'
                'MAX(0,ROUND({weekly_use}{r}*ReorderWeeks-'
                '{stock}{r},0)))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}],
            note="Enough to bring you back up to the cover you asked "
                 "for on Settings. Zero means you are fine."),
        Col("order_cost", "Order value", 11, "money",
            formula=blank(
                "name",
                'IF({order_qty}{r}="","",'
                'ROUND({order_qty}{r}*IF({cost}{r}="",0,{cost}{r}),'
                '2))'),
            hidden=True),
        Col("last_ordered", "Last ordered", 11, "date",
            validate=date_check()),
        Col("last_counted", "Last counted", 11, "date",
            validate=date_check(), hidden=True),
        Col("status", "Stock status", 12, "status",
            formula=blank(
                "name",
                'IF({stock}{r}<=0,"Out of stock",'
                'IF(IF({min_level}{r}="",FALSE,'
                '{stock}{r}<={min_level}{r}),"Low stock",'
                'IF({weeks_cover}{r}="","In stock",'
                'IF({weeks_cover}{r}<=StockWarningWeeks,"Low stock",'
                '"In stock"))))'),
            cf=status_cf()),
        Col("notes", "Notes", 24, "wrap"),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "name",
                'IF({stock}{r}<0,"Check - stock has gone negative",'
                'IF({stock}{r}=0,"Out of stock - order today",'
                'IF({status}{r}="Low stock","Order - "&'
                'IF({order_qty}{r}="","now",{order_qty}{r}&" "&'
                'IF({unit}{r}="","units",LOWER({unit}{r}))),'
                'IF({cost}{r}="","Add - no unit cost, no stock value",'
                'IF({supplier}{r}="","Add - no supplier recorded",'
                'IF(IF({weeks_cover}{r}="",FALSE,'
                '{weeks_cover}{r}>26),'
                '"Check - over six months of stock on the shelf",'
                '"OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Supply Inventory",
        subtitle="What is on the shelf, how fast it goes, and what to "
                 "order before somebody has to drive to the shop.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).supplies,
        help_text="List what you keep in stock. Update the received "
                  "and used figures as you go, and the shelf count, "
                  "the weeks of cover and the order quantity follow.",
        extra_notes=[
            "Stock is opening plus received, less used, less waste, "
            "plus adjustments. A negative figure is not a bug - it "
            "means something left the cupboard that was never booked "
            "in.",
            "'Weeks left' uses your actual consumption since the "
            "tracking start date on Settings, so it gets more "
            "accurate the longer you use it.",
            "The order quantity brings you back up to the weeks of "
            "cover you asked for on Settings - raise it before a "
            "holiday, lower it if you are short of storage.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Store cupboard summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "The store cupboard", last)
    sheet.table_head(row, 0, [
        "Items tracked", "In stock", "Low stock", "Out of stock",
        "Stock value", "Spent on supplies", "To order now",
        "Order value", "Items with no supplier"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Supply Inventory.name]],"<>")',
        '=COUNTIFS([[Supply Inventory.status]],"In stock")',
        '=COUNTIFS([[Supply Inventory.status]],"Low stock")',
        '=COUNTIFS([[Supply Inventory.status]],"Out of stock")',
        '=ROUND(SUM([[Supply Inventory.value]]),2)',
        '=ROUND(SUM([[Supply Inventory.spend]]),2)',
        '=COUNTIFS([[Supply Inventory.order_qty]],">0")',
        '=ROUND(SUM([[Supply Inventory.order_cost]]),2)',
        '=COUNTIF([[Supply Inventory.name]],"<>")-'
        'COUNTIF([[Supply Inventory.supplier]],"<>")',
    ], ["int", "int", "int", "int", "money", "money", "int", "money",
        "int"])
````

### `daycare_business_spreadsheet/sheets/communication.py`

*169 lines · 2 functions*

> `Parent Communication` - what was said, when, and who owes a reply.
>
> Every setting has the conversation that went wrong because nobody
> wrote it down.  This is one row per exchange: who, when, how, what
> about, and - the only column that really matters - whether the ball is
> still in your court.
>
> The workbook does not send anything.  It records that you did.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Communication summary |

**Module constants:** `NAME` ('Parent Communication').

````python
"""`Parent Communication` - what was said, when, and who owes a reply.

Every setting has the conversation that went wrong because nobody
wrote it down.  This is one row per exchange: who, when, how, what
about, and - the only column that really matters - whether the ball is
still in your court.

The workbook does not send anything.  It records that you did.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, time_check
from .common import (alert_cf, blank, id_col, lookup_child,
                     lookup_family, lookup_staff, month_col, status_cf)

NAME = "Parent Communication"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "MessagePrefix", digits=4, width=10,
               header="Message ID"),
        Col("date", "Date", 11, "date", validate=date_check()),
        month_col("date"),
        Col("time", "Time", 9, "time", validate=time_check(),
            hidden=True),
        Col("family", "Family", 11, "id",
            validate=list_from("FamilyList")),
        Col("family_name", "Family name", 18, "text",
            formula=lookup_family("family_name")),
        Col("parent", "Spoke to", 20, "text",
            formula=lookup_family("parent1"), hidden=True),
        Col("child", "About child", 11, "id",
            validate=list_from("ChildList"),
            note="Optional - some conversations are about the family, "
                 "not one child."),
        Col("child_name", "Child", 18, "text",
            formula=lookup_child("name", missing="")),
        Col("direction", "In or out", 11, "text",
            validate=list_from("DirectionList"),
            note="Did they contact you, or you them? The ratio tells "
                 "you something about how a term is going."),
        Col("method", "How", 15, "text",
            validate=list_from("ContactMethodList")),
        Col("category", "About", 18, "text",
            validate=list_from("CommCategoryList"), cf=status_cf(),
            note="Pick from the list so that a term's worth of "
                 "conversations can be counted by subject."),
        Col("subject", "Subject", 28, "text"),
        Col("notes", "What was said", 40, "wrap",
            note="Facts and what you agreed. Write it as though "
                 "somebody else will read it back to you in six "
                 "months, because one day somebody will."),
        Col("staff", "Logged by", 10, "id",
            validate=list_from("StaffList")),
        Col("staff_name", "Staff", 16, "text",
            formula=lookup_staff("name"), hidden=True),
        # ------------------------------------------------- follow up ---
        Col("follow_date", "Follow up by", 12, "date",
            validate=date_check()),
        Col("follow_action", "Action agreed", 28, "wrap"),
        Col("status", "Status", 15, "status",
            validate=list_from("CommStatusList"), cf=status_cf(),
            note="'Waiting on us' is the one to filter every "
                 "morning."),
        Col("closed_date", "Closed on", 11, "date",
            validate=date_check(), hidden=True),
        Col("days_open", "Days open", 9, "int",
            formula=blank(
                "date",
                'IF(({status}{r}="Closed")+'
                '({status}{r}="No reply needed")>0,'
                'IF({closed_date}{r}="","",'
                '{closed_date}{r}-{date}{r}),'
                'TodayDate-{date}{r})'),
            cf=[{"type": "cell", "criteria": ">=", "value": 14,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">=", "value": 7,
                 "fmt": "warn"}]),
        Col("overdue", "Follow-up late?", 10, "flag",
            formula=blank(
                "date",
                'IF({follow_date}{r}="","No",'
                'IF(({status}{r}="Closed")+'
                '({status}{r}="No reply needed")>0,"No",'
                'IF({follow_date}{r}<TodayDate,"Yes","No")))'),
            hidden=True),
        Col("meeting", "Meeting booked", 11, "date",
            validate=date_check(), hidden=True,
            note="Parent evenings, settling-in reviews, progress "
                 "chats."),
        Col("alert", "Action", 32, "status",
            formula=blank(
                "date",
                'IF({family}{r}="","Add - which family?",'
                'IF({family_name}{r}="Not in Parent Database",'
                '"Check - family ID not found",'
                'IF({category}{r}="","Add - what was it about?",'
                'IF({overdue}{r}="Yes","Overdue - follow-up date has '
                'passed",'
                'IF({status}{r}="Waiting on us","Follow up - the ball '
                'is with you",'
                'IF(IF({status}{r}="Waiting on parent",'
                '{days_open}{r}>=7,FALSE),'
                '"Chase - a week with no reply",'
                'IF(IF({status}{r}="Open",{days_open}{r}>=14,FALSE),'
                '"Chase - open two weeks",'
                'IF({notes}{r}="","Add - nothing written down",'
                '"OK"))))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Parent Communication Log",
        subtitle="Every conversation, what was agreed and whether "
                 "anybody is still waiting on you.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).messages,
        help_text="Log the conversation while you remember it. Pick "
                  "the family, say what it was about, and set a "
                  "follow-up date if anything was agreed.",
        extra_notes=[
            "This workbook does not send emails, texts or app "
            "notifications. It is a record that you spoke to "
            "somebody, not a way of speaking to them.",
            "Filter the status to 'Waiting on us' at the start of "
            "each day. That is the whole tab's value in one click.",
            "The open items count on each family's record comes from "
            "here, so a parent with three unanswered questions shows "
            "up on the Parent Database too.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Communication summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Conversations", last)
    sheet.table_head(row, 0, [
        "Logged", "This month", "Incoming", "Outgoing",
        "Open", "Waiting on us", "Waiting on parent",
        "Follow-ups overdue", "Average days open", "Complaints",
        "Compliments"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Parent Communication.date]],"<>")',
        '=COUNTIFS([[Parent Communication.month]],ThisMonthKey)',
        '=COUNTIFS([[Parent Communication.direction]],"Incoming")',
        '=COUNTIFS([[Parent Communication.direction]],"Outgoing")',
        '=COUNTIFS([[Parent Communication.status]],"Open")',
        '=COUNTIFS([[Parent Communication.status]],"Waiting on us")',
        '=COUNTIFS([[Parent Communication.status]],'
        '"Waiting on parent")',
        '=COUNTIFS([[Parent Communication.overdue]],"Yes")',
        '=IFERROR(ROUND(AVERAGEIF([[Parent Communication.days_open]],'
        '">=0"),1),0)',
        '=COUNTIFS([[Parent Communication.category]],"Complaint")',
        '=COUNTIFS([[Parent Communication.category]],"Compliment")',
    ], ["int", "int", "int", "int", "int", "int", "int", "int",
        "number1", "int", "int"])
````

### `daycare_business_spreadsheet/sheets/documents.py`

*186 lines · 2 functions*

> `Document Tracker` - the paperwork, and the day each piece expires.
>
> Children have enrollment agreements and immunisation records.  Staff
> have background checks and first aid certificates.  The business has a
> licence, an insurance policy and a fire inspection.  All of it expires,
> and all of it is the same problem: a date, a holder and a status.
>
> So it is one register.  The holder column says whether the row is
> about a child, a family, a member of staff or the business itself, and
> the name resolves itself from whichever register the ID belongs to.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Paperwork summary |

**Module constants:** `NAME` ('Document Tracker').

````python
"""`Document Tracker` - the paperwork, and the day each piece expires.

Children have enrollment agreements and immunisation records.  Staff
have background checks and first aid certificates.  The business has a
licence, an insurance policy and a fire inspection.  All of it expires,
and all of it is the same problem: a date, a holder and a status.

So it is one register.  The holder column says whether the row is
about a child, a family, a member of staff or the business itself, and
the name resolves itself from whichever register the ID belongs to.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from
from .common import (alert_cf, blank, id_col, status_cf)

NAME = "Document Tracker"


def spec(cfg: Config) -> SheetSpec:
    # The ID can belong to any of three registers, so the name lookup
    # tries each in turn and falls back to whatever was typed.
    who_name = (
        'IFERROR(INDEX([[Child Enrollment.name]],'
        'MATCH({who}{r},[[Child Enrollment.id]],0))&"",'
        'IFERROR(INDEX([[Staff Database.name]],'
        'MATCH({who}{r},[[Staff Database.id]],0))&"",'
        'IFERROR(INDEX([[Parent Database.family_name]],'
        'MATCH({who}{r},[[Parent Database.id]],0))&"",'
        '{who}{r})))')

    columns = [
        id_col("doc_type", "DocumentPrefix", digits=4, width=10,
               header="Doc ID"),
        Col("holder", "Belongs to", 13, "text",
            validate=list_from("DocumentHolderList"), cf=status_cf(),
            note="Child, family, staff or the business itself."),
        Col("who", "ID", 11, "id",
            note="The child ID, staff ID or family ID. For a business "
                 "document - a licence, an insurance policy - leave "
                 "it blank or type the setting's name."),
        Col("who_name", "Name", 22, "text",
            formula=blank("who", who_name),
            note="Looked up from whichever register the ID belongs "
                 "to. Nothing to type."),
        Col("doc_type", "Document", 26, "text",
            validate=list_from("DocumentTypeList"), cf=status_cf(),
            note="Pick from the list. Add your own at the bottom of "
                 "the list on Settings if your regulator wants "
                 "something this one does not have."),
        Col("required", "Required?", 10, "flag",
            validate=list_from("YesNoList"),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Yes", "fmt": "info"}],
            note="Only required documents are counted in the "
                 "'missing paperwork' figures on the child and staff "
                 "tabs."),
        Col("status", "Status", 13, "status",
            validate=list_from("DocumentStatusList"), cf=status_cf()),
        Col("requested", "Requested", 11, "date",
            validate=date_check(), hidden=True),
        Col("received", "Received", 11, "date", validate=date_check()),
        Col("expires", "Expires", 11, "date", validate=date_check(),
            note="Leave blank for anything that does not expire - an "
                 "enrollment agreement, a birth certificate sighting."),
        Col("days_left", "Days left", 9, "int",
            formula=blank("expires", '{expires}{r}-TodayDate'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bold_bad"},
                {"type": "cell", "criteria": "<=",
                 "value": "=DocUrgentDays", "fmt": "bad"},
                {"type": "cell", "criteria": "<=",
                 "value": "=DocWarningDays", "fmt": "warn"},
                {"type": "cell", "criteria": ">",
                 "value": "=DocWarningDays", "fmt": "good"}]),
        Col("state", "Where it stands", 14, "status",
            formula=blank(
                "doc_type",
                'IF({required}{r}<>"Yes","Not required",'
                'IF(({status}{r}="On file")+({status}{r}="Signed")+'
                '({status}{r}="Received")=0,"Missing",'
                'IF({expires}{r}="","On file",'
                'IF({days_left}{r}<0,"Expired",'
                'IF({days_left}{r}<=DocWarningDays,"Expiring",'
                '"On file")))))'),
            cf=status_cf(),
            note="One word for the whole row: Missing, Expired, "
                 "Expiring, On file or Not required. This is the "
                 "column the dashboard counts."),
        Col("renewed", "Last renewed", 11, "date",
            validate=date_check(), hidden=True),
        Col("renew_every", "Renew every (months)", 12, "int",
            hidden=True,
            note="For anything on a cycle. Only here so you can "
                 "filter by it - the workbook does not auto-renew."),
        Col("owner", "Chased by", 10, "id",
            validate=list_from("StaffList"), hidden=True,
            note="Whose job it is to get this one back."),
        Col("location", "Filed where", 20, "text",
            note="Cabinet 2, child files. Shared drive, Staff/HR. "
                 "A pointer, not the document."),
        Col("reference", "Reference", 14, "text", hidden=True,
            note="Policy number, certificate number, licence number."),
        Col("notes", "Notes", 26, "wrap"),
        Col("alert", "Action", 32, "status",
            formula=blank(
                "doc_type",
                'IF({required}{r}<>"Yes","",'
                'IF({state}{r}="Missing",'
                'IF({status}{r}="Chased","Chase - asked for it '
                'already, ask again",'
                'IF({status}{r}="Requested","Chase - requested, not '
                'back yet","Missing - request it")),'
                'IF({state}{r}="Expired","Expired - '
                '"&ABS({days_left}{r})&" days ago, renew now",'
                'IF(IF({expires}{r}="",FALSE,'
                '{days_left}{r}<=DocUrgentDays),'
                '"Expiring - "&{days_left}{r}&" days left",'
                'IF(IF({expires}{r}="",FALSE,'
                '{days_left}{r}<=DocWarningDays),'
                '"Due - book the renewal",'
                'IF(IF({who}{r}="",FALSE,{who_name}{r}={who}{r}),'
                '"Check - that ID is not in any register",'
                '"OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Document & Renewal Tracker",
        subtitle="Children's forms, staff certificates and the "
                 "setting's own licences - with the day each one "
                 "runs out.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=3,
        samples=sample(cfg).documents,
        help_text="One row per document. Say who it belongs to, "
                  "whether it is required and when it expires - the "
                  "status, the warnings and the missing-paperwork "
                  "counts all follow.",
        extra_notes=[
            "The name column resolves the ID against the children, "
            "the staff and the families in turn, so one register "
            "holds every kind of document without you having to "
            "choose a tab.",
            "Only rows marked 'Required - Yes' are counted in the "
            "missing paperwork figures on the child and staff "
            "records. Mark the optional ones No and they stop "
            "nagging.",
            "This tracks your paperwork. It does not make you "
            "compliant with anybody's licensing rules, and the "
            "document list is a starting point to edit rather than a "
            "legal checklist.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Paperwork summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Paperwork at a glance", last)
    sheet.table_head(row, 0, [
        "Documents tracked", "Required", "On file", "Missing",
        "Expired", "Expiring soon", "Child documents",
        "Staff documents", "Business documents", "Not required"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Document Tracker.doc_type]],"<>")',
        '=COUNTIFS([[Document Tracker.required]],"Yes")',
        '=COUNTIFS([[Document Tracker.state]],"On file")',
        '=COUNTIFS([[Document Tracker.state]],"Missing")',
        '=COUNTIFS([[Document Tracker.state]],"Expired")',
        '=COUNTIFS([[Document Tracker.state]],"Expiring")',
        '=COUNTIFS([[Document Tracker.holder]],"Child")',
        '=COUNTIFS([[Document Tracker.holder]],"Staff")',
        '=COUNTIFS([[Document Tracker.holder]],"Business")',
        '=COUNTIFS([[Document Tracker.state]],"Not required")',
    ], ["int", "int", "int", "int", "int", "int", "int", "int", "int",
        "int"])
````

---

## 9. The pages

The hand-built tabs: the dashboard, the welcome, the settings tab every formula reads, the monthly report with its printable invoice, and the annual report.

### `daycare_business_spreadsheet/sheets/dashboard.py`

*566 lines · 3 functions*

> `Dashboard` - the one page that answers "how is the setting doing?"
>
> Twenty numbers across the top, the jobs that need doing before you
> take your coat off, every room with its occupancy and its ratio, the
> families who owe you money, and twelve months of income against
> target.  Everything reads from the registers, so it is never out of
> date and there is nothing to refresh.
>
> Four questions, in the order a childcare owner actually asks them:
> are the rooms full, are the children here, are we getting paid, and
> is anything about to go wrong.

| Defines | What it does |
|---------|--------------|
| `_by_rank(sheet, rank_key, field, n)` | The nth row of a register by one of its hidden rank columns |
| `_nth(sheet, field, n, style)` | Row n of a register, blank rather than zero when it is empty |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Dashboard'), `LAST_COL` (12), `MONTHS` (12), `ROOM_ROWS` (8), `TOP_ROWS` (6).

````python
"""`Dashboard` - the one page that answers "how is the setting doing?"

Twenty numbers across the top, the jobs that need doing before you
take your coat off, every room with its occupancy and its ratio, the
families who owe you money, and twelve months of income against
target.  Everything reads from the registers, so it is never out of
date and there is nothing to refresh.

Four questions, in the order a childcare owner actually asks them:
are the rooms full, are the children here, are we getting paid, and
is anything about to go wrong.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter

NAME = "Dashboard"
LAST_COL = 12
MONTHS = 12
ROOM_ROWS = 8
TOP_ROWS = 6


def _by_rank(sheet: str, rank_key: str, field: str, n: int) -> str:
    """The nth row of a register by one of its hidden rank columns."""
    return (f'IFERROR(INDEX([[{sheet}.{field}]],MATCH({n},'
            f'[[{sheet}.{rank_key}]],0)),"")')


def _nth(sheet: str, field: str, n: int, style: str = "text") -> str:
    """Row n of a register, blank rather than zero when it is empty."""
    blank = '""' if style == "text" else '""'
    return (f'IFERROR(IF(INDEX([[{sheet}.name]],{n})="",{blank},'
            f'INDEX([[{sheet}.{field}]],{n})),{blank})')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws
    cur = cfg.currency_symbol

    s.banner("Childcare Dashboard",
             "Enrollment, attendance, money and the jobs that need "
             "doing - all worked out from the registers you have "
             "already filled in.", LAST_COL)
    s.set_col(0, 0, 26)
    for column in range(1, LAST_COL + 1):
        s.set_col(column, column, 13)

    # --------------------------------------------------------- the header --
    row = 4
    s.merge(row, 0, 0, "Setting", fmts.chrome("label_bold"))
    s.merge(row, 1, 3, None, fmts.calc("text"))
    ws.write_formula(row, 1, "=CentreName", fmts.calc("text"))
    s.merge(row, 4, 4, "Month", fmts.chrome("label_right"))
    ws.write_formula(row, 5, '=TEXT(TodayDate,"mmmm yyyy")',
                     fmts.calc("text"))
    s.merge(row, 6, 6, "Month key", fmts.chrome("label_right"))
    ws.write_formula(row, 7, '=TEXT(TodayDate,"yyyy-mm")',
                     fmts.calc("text"))
    book.define_name("DashKey", f"={quote_sheet(NAME)}!$H${row + 1}")
    s.merge(row, 8, 9, "Last month", fmts.chrome("label_right"))
    ws.write_formula(row, 10,
                     '=TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-1,1),'
                     '"yyyy-mm")', fmts.calc("text"))
    book.define_name("DashPrevKey", f"={quote_sheet(NAME)}!$K${row + 1}")
    s.merge(row, 11, 11, "Today", fmts.chrome("label_right"))
    s.merge(row, 12, LAST_COL, None, fmts.calc("date"))
    ws.write_formula(row, 12, "=TodayDate", fmts.calc("date"))
    row += 2

    # ----------------------------------------------------------- the KPIs --
    enrollment = [
        ("Children on the books",
         "=ChildrenEnrolled", "int",
         "active, on hold or on notice"),
        ("Occupancy",
         '=IF(RoomCapacity<=0,"",ChildrenEnrolled/RoomCapacity)',
         "percent", f"target {cfg.target_occupancy:.0%}"),
        ("Places free",
         "=MAX(0,RoomCapacity-ChildrenEnrolled)", "int",
         "across every room"),
        ("Waiting for a place",
         '=COUNTIFS([[Waitlist.closed]],"No")', "int",
         "live enquiries, not yet offered or lost"),
    ]
    attendance = [
        ("Here today",
         '=COUNTIFS([[Attendance.date]],TodayDate,'
         '[[Attendance.present]],1)', "int",
         "checked in on the attendance tab"),
        ("Attendance this month",
         '=IF(SUMIFS([[Attendance.session]],[[Attendance.month]],'
         'DashKey)<=0,"",SUMIFS([[Attendance.present]],'
         '[[Attendance.month]],DashKey)/'
         'SUMIFS([[Attendance.session]],[[Attendance.month]],DashKey))',
         "percent", f"target {cfg.target_attendance:.0%}"),
        ("Absences this month",
         '=SUMIFS([[Attendance.absent]],[[Attendance.month]],DashKey)',
         "int", "sick, family, holiday or unexplained"),
        ("Still to check out",
         '=COUNTIF([[Attendance.alert]],"Missing - no check-out*")',
         "int", "a session with no collection time"),
    ]
    financial = [
        ("Billed this month",
         '=ROUND(SUMIFS([[Invoices & Receipts.billed]],'
         '[[Invoices & Receipts.month]],DashKey),2)', "money0",
         f"target {cur}{cfg.revenue_target_month:,.0f}"),
        ("Against target",
         '=IF(RevenueTargetMonth<=0,"",'
         'SUMIFS([[Invoices & Receipts.billed]],'
         '[[Invoices & Receipts.month]],DashKey)/RevenueTargetMonth)',
         "percent0", "100% means the month is made"),
        ("Collected this month",
         '=ROUND(SUMIFS([[Invoices & Receipts.collected]],'
         '[[Invoices & Receipts.month]],DashKey),2)', "money0",
         "money that actually arrived"),
        ("Outstanding, all families",
         '=ROUND(SUM([[Parent Database.balance]]),2)', "money0",
         "billed and not yet paid"),
    ]
    operating = [
        ("Operating profit, month",
         "=ROUND(MonthBilled-MonthWages-MonthExpenses,2)", "money0",
         "billed less wages less costs"),
        ("Profit margin",
         '=IF(MonthBilled<=0,"",(MonthBilled-MonthWages-MonthExpenses)'
         '/MonthBilled)', "percent",
         f"target {cfg.profit_target_month / max(1.0, cfg.revenue_target_month):.0%}"
         " or better"),
        ("Wage cost of income",
         '=IF(MonthBilled<=0,"",MonthWages/MonthBilled)', "percent",
         "most settings land between 55% and 70%"),
        ("Income per child, week",
         '=IF(ChildrenEnrolled<=0,"",'
         'ROUND(SUM([[Child Enrollment.weekly_net]])/ChildrenEnrolled,'
         '2))', "money", "after discounts and subsidies"),
    ]
    people = [
        ("Staff on the rota today",
         '=COUNTIFS([[Staff Schedule.date]],TodayDate,'
         '[[Staff Schedule.counts]],"Yes")', "int",
         "counted towards ratio"),
        ("Rooms over ratio",
         '=COUNTIF([[Classroom Management.ratio_ok]],"No")', "int",
         "get cover into these first"),
        ("Documents expired or missing",
         '=COUNTIF([[Document Tracker.state]],"Expired")+'
         'COUNTIF([[Document Tracker.state]],"Missing")', "int",
         "required paperwork only"),
        ("Jobs waiting on you",
         '=COUNTIFS([[Parent Communication.status]],"Waiting on us")+'
         'COUNTIF([[Parent Communication.overdue]],"Yes")', "int",
         "parent messages you have not answered"),
    ]

    for band in (enrollment, attendance, financial, operating, people):
        column = 0
        bottom = row
        for label, formula, style, note in band:
            bottom = s.kpi(row, column, label, formula, style, span=3,
                           note=note)
            column += 3
        row = bottom
    row += 1

    # ------------------------------------------------- needs doing today ---
    row = s.section(row, "Needs doing today", LAST_COL)
    row = s.table_head(row, 0, ["What", "How many", "What it means",
                                "Where to fix it"],
                       [28, 11, 50, 18])
    alerts = [
        ("Rooms over ratio",
         '=COUNTIF([[Classroom Management.ratio_ok]],"No")',
         '="More children in the room than your own ratio allows. '
         'Move an adult in before anything else today"',
         "Classroom Management"),
        ("Rooms nearly full",
         '=COUNTIF([[Classroom Management.alert]],"Nearly full*")',
         '="At or above "&TEXT(RatioWarning,"0%")&" of licensed '
         'places. Start offering the waitlist"',
         "Classroom Management"),
        ("Children with no check-out",
         '=COUNTIF([[Attendance.alert]],"Missing - no check-out*")',
         '="A past session with a check-in and no check-out"',
         "Attendance"),
        ("Collected by somebody unexpected",
         '=COUNTIF([[Attendance.alert]],"Not authorised*")',
         '="The name in Collected by is not on that child'
         "'" 's authorised list"',
         "Attendance"),
        ("Unexplained absences to chase",
         '=COUNTIF([[Attendance.alert]],"Unexplained*")',
         '="Marked absent with no reason written down"',
         "Attendance"),
        ("Late pickups not charged",
         '=COUNTIF([[Attendance.alert]],"Invoice - late pickup*")',
         '="Collected after the grace period with no fee raised"',
         "Tuition & Fees"),
        ("Charges not on an invoice",
         '=COUNTIFS([[Tuition & Fees.billed]],"No")',
         '="Work you have done that nobody has been billed for"',
         "Tuition & Fees"),
        ("Invoices overdue",
         '=COUNTIFS([[Invoices & Receipts.status]],"Overdue")',
         '="Past the due date and still showing a balance"',
         "Invoices & Receipts"),
        ("Invoices still a draft",
         '=COUNTIFS([[Invoices & Receipts.status]],"Draft")',
         '="Raised but never marked as sent to the parent"',
         "Invoices & Receipts"),
        ("Families with a balance",
         '=COUNTIFS([[Parent Database.balance]],">0.004")',
         '="Owing something, overdue or not"', "Parent Database"),
        ("Parent messages waiting on you",
         '=COUNTIFS([[Parent Communication.status]],"Waiting on us")',
         '="You said you would come back to them"',
         "Parent Communication"),
        ("Follow-ups past their date",
         '=COUNTIF([[Parent Communication.overdue]],"Yes")',
         '="A follow-up date that has already gone"',
         "Parent Communication"),
        ("Incident forms not signed",
         '=COUNTIF([[Incident Reports.alert]],"Chase - parent*")',
         '="Parent told, but the form has not come back signed"',
         "Incident Reports"),
        ("Incidents with no parent told",
         '=COUNTIF([[Incident Reports.alert]],"Not notified*")',
         '="Somebody has to make that call today"',
         "Incident Reports"),
        ("Medication without written consent",
         '=COUNTIF([[Allergy & Medical.alert]],"Not authorised*")',
         '="A dose logged with no signed consent behind it"',
         "Allergy & Medical"),
        ("Children missing paperwork",
         '=COUNTIFS([[Child Enrollment.docs_missing]],">0")',
         '="Required documents not yet on file"',
         "Document Tracker"),
        ("Documents expiring soon",
         '=COUNTIF([[Document Tracker.state]],"Expiring")',
         '="Runs out inside "&DocWarningDays&" days"',
         "Document Tracker"),
        ("Documents already expired",
         '=COUNTIF([[Document Tracker.state]],"Expired")',
         '="Including your own licence, insurance and inspections"',
         "Document Tracker"),
        ("Staff certificates expired",
         '=COUNTIFS([[Staff Database.certs_ok]],"Expired")',
         '="First aid, CPR, safeguarding or background check"',
         "Staff Database"),
        ("Staff certificates expiring",
         '=COUNTIFS([[Staff Database.certs_ok]],"Expiring")',
         '="Book the course before it lapses"', "Staff Database"),
        ("Shifts double booked",
         '=COUNTIF([[Staff Schedule.alert]],"Duplicate*")',
         '="The same person in two places at once"',
         "Staff Schedule"),
        ("Pay runs not paid",
         '=COUNTIFS([[Payroll Tracker.paid]],"<>Yes",'
         '[[Payroll Tracker.period_end]],"<>")',
         '="Calculated and still open"', "Payroll Tracker"),
        ("Supplies out of stock",
         '=COUNTIF([[Supply Inventory.status]],"Out of stock")',
         '="Nothing left in the cupboard"', "Supply Inventory"),
        ("Supplies to reorder",
         '=COUNTIF([[Supply Inventory.status]],"Low stock")',
         '="Below the minimum or under "&StockWarningWeeks&" weeks '
         'of cover"', "Supply Inventory"),
        ("Birthdays coming up",
         '=COUNTIF([[Child Enrollment.alert]],"Birthday*")',
         '="Inside the next "&BirthdayWindow&" days"',
         "Child Enrollment"),
        ("Children moving room soon",
         '=COUNTIF([[Child Enrollment.alert]],"Moves up*")',
         '="They age out of their band within a month - tell the '
         'parents before they notice"', "Child Enrollment"),
        ("Waitlist enquiries going quiet",
         '=COUNTIF([[Waitlist.alert]],"Chase*")+'
         'COUNTIF([[Waitlist.alert]],"Follow up*")',
         '="Nobody has spoken to them in a fortnight"', "Waitlist"),
        ("Seats to offer the waitlist",
         '=COUNTIF([[Waitlist.alert]],"Space*")',
         '="A place has opened in the room they asked for"',
         "Waitlist"),
    ]
    first_alert = row
    for label, count, message, where in alerts:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        ws.write_formula(row, 1, s.reg.render(count), fmts.calc("int"))
        s.merge(row, 2, LAST_COL - 3, None, fmts.calc("text"))
        ws.write_formula(row, 2, s.reg.render(message),
                         fmts.calc("text"))
        s.merge(row, LAST_COL - 2, LAST_COL, where, fmts.chrome("label"))
        row += 1
    ws.conditional_format(first_alert, 1, row - 1, 1,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("warn")})
    ws.conditional_format(first_alert, 1, row - 1, 1,
                          {"type": "cell", "criteria": "=", "value": 0,
                           "format": fmts.cf("good")})
    row += 1

    # ------------------------------------------------------- the rooms ----
    row = s.section(row, "Every room, right now", LAST_COL)
    row = s.table_head(row, 0, [
        "Room", "Age band", "Places", "Enrolled", "Free", "Occupancy",
        "Here today", "Adults today", "Adults needed", "Ratio now",
        "Within ratio", "Weekly income", "Empty seats cost"],
        [18, 15, 9, 10, 8, 11, 11, 12, 13, 11, 12, 13, 14])
    rooms_first = row
    for n in range(1, ROOM_ROWS + 1):
        cells = [
            (0, _nth("Classroom Management", "name", n), "text"),
            (1, _nth("Classroom Management", "age_band", n), "text"),
            (2, _nth("Classroom Management", "capacity", n), "int"),
            (3, _nth("Classroom Management", "enrolled", n), "int"),
            (4, _nth("Classroom Management", "seats", n), "int"),
            (5, _nth("Classroom Management", "occupancy", n),
             "percent"),
            (6, _nth("Classroom Management", "present_today", n),
             "int"),
            (7, _nth("Classroom Management", "staff_today", n),
             "number1"),
            (8, _nth("Classroom Management", "staff_needed", n),
             "number1"),
            (9, _nth("Classroom Management", "ratio_now", n), "text"),
            (10, _nth("Classroom Management", "ratio_ok", n), "text"),
            (11, _nth("Classroom Management", "revenue", n), "money"),
            (12, _nth("Classroom Management", "empty_cost", n),
             "money"),
        ]
        for column, formula, style in cells:
            ws.write_formula(row, column, s.reg.render("=" + formula),
                             fmts.calc(style))
        row += 1
    ws.conditional_format(rooms_first, 5, row - 1, 5,
                          {"type": "cell", "criteria": ">=",
                           "value": "=RatioWarning",
                           "format": fmts.cf("warn")})
    ws.conditional_format(rooms_first, 10, row - 1, 10,
                          {"type": "text", "criteria": "containing",
                           "value": "No", "format": fmts.cf("bad")})
    ws.conditional_format(rooms_first, 10, row - 1, 10,
                          {"type": "text", "criteria": "containing",
                           "value": "Yes", "format": fmts.cf("good")})
    row += 1

    # --------------------------------------------------- money and people --
    row = s.section(row, "Who owes you, and who earns it", LAST_COL)
    row = s.table_head(row, 0, [
        "Family", "Children", "Weekly fees", "Billed", "Paid",
        "Balance", "Overdue", "Oldest bill", "Status",
        "Biggest fee payer", "A year", "Room"],
        [22, 9, 12, 12, 12, 12, 12, 12, 16, 22, 12, 13])
    money_first = row
    for n in range(1, TOP_ROWS + 1):
        cells = [
            (0, _by_rank("Parent Database", "weekly_fees_rank",
                         "family_name", n), "text"),
            (1, _by_rank("Parent Database", "weekly_fees_rank",
                         "children", n), "int"),
            (2, _by_rank("Parent Database", "weekly_fees_rank",
                         "weekly_fees", n), "money"),
            (3, _by_rank("Parent Database", "weekly_fees_rank",
                         "invoiced", n), "money"),
            (4, _by_rank("Parent Database", "weekly_fees_rank",
                         "paid", n), "money"),
            (5, _by_rank("Parent Database", "weekly_fees_rank",
                         "balance", n), "money"),
            (6, _by_rank("Parent Database", "weekly_fees_rank",
                         "overdue", n), "money"),
            (7, _by_rank("Parent Database", "weekly_fees_rank",
                         "oldest_due", n), "date"),
            (8, _by_rank("Parent Database", "weekly_fees_rank",
                         "status", n), "text"),
            (9, _by_rank("Child Enrollment", "annual_net_rank", "name",
                         n), "text"),
            (10, _by_rank("Child Enrollment", "annual_net_rank",
                          "annual_net", n), "money"),
            (11, _by_rank("Child Enrollment", "annual_net_rank", "room",
                          n), "text"),
        ]
        for column, formula, style in cells:
            ws.write_formula(row, column, s.reg.render("=" + formula),
                             fmts.calc(style))
        row += 1
    ws.conditional_format(money_first, 6, row - 1, 6,
                          {"type": "cell", "criteria": ">",
                           "value": 0.004, "format": fmts.cf("bad")})
    row += 1

    # ------------------------------------------------- twelve months ------
    row = s.section(row, "The last twelve months", LAST_COL)
    row = s.table_head(row, 0, [
        "Month", "Billed", "Collected", "Target", "Against target",
        "Wages", "Other costs", "Operating profit", "Margin",
        "Children billed", "Sessions", "Attendance", "Income a child"],
        [12, 13, 13, 12, 13, 12, 12, 15, 10, 13, 10, 11, 13])
    chart_first = row
    for index in range(MONTHS - 1, -1, -1):
        key = (f'TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{index},1),'
               f'"yyyy-mm")')
        label = (f'=TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{index},'
                 f'1),"mmm yy")')
        billed = (f'SUMIFS([[Invoices & Receipts.billed]],'
                  f'[[Invoices & Receipts.month]],{key})')
        wages = (f'SUMIFS([[Payroll Tracker.total_cost]],'
                 f'[[Payroll Tracker.month]],{key})')
        costs = (f'SUMIFS([[Expense Tracker.pl_amount]],'
                 f'[[Expense Tracker.month]],{key})')
        sessions = (f'SUMIFS([[Attendance.session]],'
                    f'[[Attendance.month]],{key})')
        present = (f'SUMIFS([[Attendance.present]],'
                   f'[[Attendance.month]],{key})')
        s.write(row, 0, None, fmts.calc("text"))
        ws.write_formula(row, 0, label, fmts.calc("text"))
        formulas = [
            (1, f"=ROUND({billed},2)", "money"),
            (2, f'=ROUND(SUMIFS([[Invoices & Receipts.collected]],'
                f'[[Invoices & Receipts.month]],{key}),2)', "money"),
            (3, "=ROUND(RevenueTargetMonth,2)", "money"),
            (4, f"=ROUND({billed}-RevenueTargetMonth,2)", "money"),
            (5, f"=ROUND({wages},2)", "money"),
            (6, f"=ROUND({costs},2)", "money"),
            (7, f"=ROUND({billed}-{wages}-{costs},2)", "money"),
            (8, f'=IF({billed}<=0,"",({billed}-{wages}-{costs})/'
                f'{billed})', "percent"),
            (9, f'=COUNTIFS([[Invoices & Receipts.month]],{key},'
                f'[[Invoices & Receipts.type]],"Invoice")', "int"),
            (10, f"={sessions}", "int"),
            (11, f'=IF({sessions}<=0,"",{present}/{sessions})',
             "percent"),
            (12, f'=IF(ChildrenEnrolled<=0,"",ROUND({billed}/'
                 f'ChildrenEnrolled,2))', "money"),
        ]
        for column, formula, style in formulas:
            ws.write_formula(row, column, s.reg.render(formula),
                             fmts.calc(style))
        row += 1
    last_month_row = row - 1
    ws.conditional_format(chart_first, 1, last_month_row, 1,
                          {"type": "data_bar",
                           "bar_color": Palette.primary_mid})
    ws.conditional_format(chart_first, 7, last_month_row, 7,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(chart_first, 4, last_month_row, 4,
                          {"type": "cell", "criteria": ">=", "value": 0,
                           "format": fmts.cf("good")})
    ws.conditional_format(chart_first, 11, last_month_row, 11,
                          {"type": "cell", "criteria": "<",
                           "value": "=TargetAttendance",
                           "format": fmts.cf("warn")})

    # sparklines beat a chart Excel 2010 might refuse to draw
    ws.add_sparkline(row, 1, {
        "range": (f"{quote_sheet(NAME)}!$B${chart_first + 1}:"
                  f"$B${last_month_row + 1}"),
        "type": "column", "series_color": Palette.primary_mid})
    s.merge(row, 2, 3, "Billed, twelve months", fmts.chrome("note"))
    ws.add_sparkline(row, 4, {
        "range": (f"{quote_sheet(NAME)}!$H${chart_first + 1}:"
                  f"$H${last_month_row + 1}"),
        "type": "column", "series_color": Palette.accent,
        "negative_points": True})
    s.merge(row, 5, 7, "Operating profit, twelve months",
            fmts.chrome("note"))
    ws.add_sparkline(row, 8, {
        "range": (f"{quote_sheet(NAME)}!$L${chart_first + 1}:"
                  f"$L${last_month_row + 1}"),
        "type": "line", "series_color": Palette.deep,
        "markers": True})
    s.merge(row, 9, LAST_COL, "Attendance, twelve months",
            fmts.chrome("note"))
    row += 2

    # ---------------------------------------------------- the month in one -
    row = s.section(row, "This month in one line", LAST_COL)
    row = s.table_head(row, 0, [
        "Billed", "Collected", "Still owed", "Wages", "Other costs",
        "Operating profit", "Margin", "Break-even income",
        "Cost a child, month"],
        [14, 14, 13, 13, 13, 15, 10, 16, 16])
    totals_row = row
    ws.write_formula(
        totals_row, 0,
        s.reg.render('=ROUND(SUMIFS([[Invoices & Receipts.billed]],'
                     '[[Invoices & Receipts.month]],DashKey),2)'),
        fmts.calc("money"))
    book.define_name("MonthBilled",
                     f"={quote_sheet(NAME)}!$A${totals_row + 1}")
    ws.write_formula(
        totals_row, 1,
        s.reg.render('=ROUND(SUMIFS([[Invoices & Receipts.collected]],'
                     '[[Invoices & Receipts.month]],DashKey),2)'),
        fmts.calc("money"))
    book.define_name("MonthCollected",
                     f"={quote_sheet(NAME)}!$B${totals_row + 1}")
    ws.write_formula(
        totals_row, 2,
        s.reg.render("=ROUND(SUM([[Parent Database.balance]]),2)"),
        fmts.calc("money"))
    ws.write_formula(
        totals_row, 3,
        s.reg.render('=ROUND(SUMIFS([[Payroll Tracker.total_cost]],'
                     '[[Payroll Tracker.month]],DashKey),2)'),
        fmts.calc("money"))
    book.define_name("MonthWages",
                     f"={quote_sheet(NAME)}!$D${totals_row + 1}")
    ws.write_formula(
        totals_row, 4,
        s.reg.render('=ROUND(SUMIFS([[Expense Tracker.pl_amount]],'
                     '[[Expense Tracker.month]],DashKey),2)'),
        fmts.calc("money"))
    book.define_name("MonthExpenses",
                     f"={quote_sheet(NAME)}!$E${totals_row + 1}")
    ws.write_formula(totals_row, 5,
                     "=ROUND(MonthBilled-MonthWages-MonthExpenses,2)",
                     fmts.calc("money"))
    book.define_name("MonthProfit",
                     f"={quote_sheet(NAME)}!$F${totals_row + 1}")
    ws.write_formula(totals_row, 6,
                     '=IF(MonthBilled<=0,"",MonthProfit/MonthBilled)',
                     fmts.calc("percent"))
    ws.write_formula(totals_row, 7,
                     "=ROUND(MonthWages+MonthExpenses,2)",
                     fmts.calc("money"))
    ws.write_formula(totals_row, 8,
                     '=IF(ChildrenEnrolled<=0,"",'
                     'ROUND((MonthWages+MonthExpenses)/'
                     'ChildrenEnrolled,2))', fmts.calc("money"))
    ws.conditional_format(totals_row, 5, totals_row, 5,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    row = totals_row + 2

    row = s.note(
        row,
        "Operating profit here is what you billed for the month, less "
        "the full cost of the pay runs, less the expenses logged "
        "against it. Wages are counted once - from the Payroll "
        "Tracker - and wage rows on the Expense Tracker are left out "
        "unless you switch that on in Settings. It is not a tax "
        "computation and it knows nothing about depreciation, "
        "drawings or loans: it is the number that tells you whether "
        "the month worked.",
        LAST_COL, height=46)
    row += 1
    row = s.note(
        row,
        "Ratios on this page come from the rota, not from where staff "
        "are usually based: a shift only counts towards a room when "
        "it has a room against it and is not marked Off, Leave, "
        "Training or Admin. The bands and the numbers themselves are "
        "yours to set on the Settings tab - they are not anybody's "
        "licensing rules and this workbook does not make you "
        "compliant with them.",
        LAST_COL, height=42)

    s.finish(freeze_row=6, gridlines=False, fit_width=True)
    s.protect()
    return s
````

### `daycare_business_spreadsheet/sheets/getting_started.py`

*337 lines · 1 function*

> `Getting Started` - what this is, what to type, and what it will not do.
>
> The first tab a buyer opens.  It sets the file up in seven steps,
> maps all twenty-three tabs, and is straight about the limits: this is
> a business management workbook, not a child development record, not a
> medical system, and not a licensing compliance tool.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Getting Started'), `LAST_COL` (9), `TABS` (23 items), `STEPS` (7 items), `RULES` (7 items), `LIMITS` (7 items), `SUMS` (12 items).

````python
"""`Getting Started` - what this is, what to type, and what it will not do.

The first tab a buyer opens.  It sets the file up in seven steps,
maps all twenty-three tabs, and is straight about the limits: this is
a business management workbook, not a child development record, not a
medical system, and not a licensing compliance tool.
"""

from __future__ import annotations

from ..config import DISCLAIMER, PRIVACY_NOTE, Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter

NAME = "Getting Started"
LAST_COL = 9

TABS = [
    ("Dashboard", "The setting at a glance: occupancy, attendance, "
                  "money in and out, every room's ratio and the jobs "
                  "that need doing today.", "Read it"),
    ("Getting Started", "This page.", "Read it"),
    ("Settings & Lists", "Your setting, your fees, your age bands and "
                         "ratios, your tax rules, your lists. Set "
                         "these first.", "Set it up"),
    ("Child Enrollment", "The register. Every child, their family, "
                         "their room, their schedule and the rate "
                         "they are on.", "Type here"),
    ("Parent Database", "One row per family, not per child - so "
                        "siblings share a bill, an address and an "
                        "emergency contact.", "Type here"),
    ("Waitlist", "Enquiries, tours, offers and where each one sits in "
                 "the queue for the room they want.", "Type here"),
    ("Attendance", "Daily check in and check out, with late arrivals, "
                   "early collections and missing sign-outs flagged.",
     "Type here"),
    ("Tuition & Fees", "One row per charge. Tuition, extras, "
                       "discounts and subsidies - the lines that make "
                       "up an invoice.", "Type here"),
    ("Invoices & Receipts", "The money ledger: invoices, payments, "
                            "credit notes, refunds and write-offs, "
                            "one row per movement.", "Type here"),
    ("Daily Reports", "The printable note that goes home: mood, meals, "
                      "naps, nappies and what they did today.",
     "Type here"),
    ("Meal Tracker", "What was served, how much was eaten, and who "
                     "needed something different.", "Type here"),
    ("Allergy & Medical", "Allergies, conditions, care plans, consent "
                          "and the medication administration log.",
     "Type here"),
    ("Incident Reports", "Accidents, injuries and near misses, with "
                         "the parent notification timed to the "
                         "minute.", "Type here"),
    ("Staff Database", "The team: roles, pay, certificates and when "
                       "each one expires.", "Type here"),
    ("Staff Schedule", "The rota and the clock. Who is in which room, "
                       "for how long, and what it costs.",
     "Type here"),
    ("Payroll Tracker", "Pay runs built from the rota - basic, "
                        "overtime, employer costs and net pay.",
     "Type here"),
    ("Classroom Management", "Every room: licensed places, children "
                             "enrolled, who is in today and whether "
                             "you are within ratio.", "Type here"),
    ("Expense Tracker", "Every cost, categorised the way a profit and "
                        "loss wants it, against your budget.",
     "Type here"),
    ("Supply Inventory", "Nappies, food, cleaning and craft - what is "
                         "left, how fast it goes, what to order.",
     "Type here"),
    ("Parent Communication", "Every conversation, what was agreed and "
                             "who is still waiting on a reply.",
     "Type here"),
    ("Document Tracker", "Children's forms, staff certificates and "
                         "the setting's own licences, with expiry "
                         "dates.", "Type here"),
    ("Monthly Reports", "One printable page for the month - "
                        "enrollment, attendance, money and staffing.",
     "Read it"),
    ("Annual Reports", "Twelve months side by side, plus the year's "
                       "totals and what changed.", "Read it"),
]

STEPS = [
    ("1. Settings first",
     "Open Settings & Lists and work down it: the setting's name, "
     "your currency, your opening hours, your age bands and the "
     "staff-to-child ratio for each one, your fees, your tax "
     "switches and your targets. Twenty minutes here and no other "
     "tab has to guess."),
    ("2. Set the age bands to your own licence",
     "The ratio table on Settings is the most country-specific thing "
     "in this workbook, which is why it is a table you edit and not "
     "a formula you cannot. Infant, toddler, preschool - the "
     "boundaries in months and the children per adult are all yours "
     "to change. Every room warning, every ratio check and every "
     "'moves up on' date follows that table."),
    ("3. Open the rooms",
     "On Classroom Management, name each room, give it its licensed "
     "places and pick the age band it runs. Everything on the "
     "dashboard that talks about occupancy, free seats or ratio "
     "reads these rows."),
    ("4. Add the families, then the children",
     "Families first, on Parent Database - one row per household, "
     "with the emergency contacts and the authorised pickup list. "
     "Then the children on Child Enrollment, each pointing at their "
     "family ID. Two children in one family share a bill "
     "automatically, and the second one picks up the sibling "
     "discount without you doing anything."),
    ("5. Take the register every day",
     "Attendance is the tab that drives the rest. Date, child, "
     "status, in and out. From it come the attendance percentages, "
     "the late pickup fees, the daily reports, the room counts and "
     "the ratio checks."),
    ("6. Bill once a month",
     "Add the charge lines on Tuition & Fees, raise one invoice per "
     "family on Invoices & Receipts, and tag the lines with that "
     "invoice number. When the money comes in, add a Payment row "
     "pointing at the same invoice. Partial payments, instalments "
     "and credit notes all work with no extra columns."),
    ("7. Once a month, in twenty minutes",
     "Run the payroll rows from the rota, log the bills on Expense "
     "Tracker, count the supplies, check the Document Tracker for "
     "anything expiring, and read the Monthly Report. That is the "
     "whole routine."),
]

RULES = [
    "Blue cells are yours. Grey cells are formulas - they are locked "
    "so a stray paste cannot break the workbook.",
    "Every sheet is protected with the password 'premium'. Review > "
    "Unprotect Sheet if you want to change the structure.",
    "IDs number themselves from the row. Filter and sort freely, but "
    "do not cut and paste rows around - copy the values instead.",
    "Type dates as real dates and times as real times. 7:30 am, not "
    "'half seven'.",
    "The family is the billing unit and the child is the care unit. "
    "Invoices go to a family ID; attendance, meals and reports go to "
    "a child ID. Keeping those straight is what makes siblings work.",
    "Never type over a formula to make a number look right. Find the "
    "row that is wrong instead - the Action column on every tab is "
    "there to point at it.",
    "Add rows by typing in the first empty row of a tab. Every tab "
    "has hundreds already formatted and waiting.",
]

LIMITS = [
    "It is a business management workbook. It is not a child "
    "development record, a learning journal, or an assessment "
    "framework, and it makes no claim to meet EYFS, Head Start, "
    "NAEYC or any other curriculum or quality standard.",
    "It is not a medical system. The allergy, medication and "
    "incident tabs are a written record of what you did and when - "
    "they are not clinical advice and they do not replace a care "
    "plan written with a health professional.",
    "It does not make you compliant with anybody's licensing rules. "
    "The ratios and age bands are the ones you type in. Check them "
    "against your own regulator, and check them again when the rules "
    "change.",
    "It does not send anything. No emails, no texts, no app "
    "notifications, no parent portal. Every record in here is typed "
    "in or pasted in by you.",
    "It is not accounting, payroll or tax software, and nothing in "
    "it is accounting, payroll or tax advice. It does not calculate "
    "income tax or social security on wages - it shows what a pay "
    "run costs the setting.",
    "It does not take payments and it does not store card details.",
    "Tax is whatever you set it to be. There is no country baked in, "
    "and childcare is exempt in a great many places - which is why "
    "the default rate is zero.",
]

SUMS = [
    ("Occupancy", "Children holding a place, divided by the licensed "
                  "places across your rooms. On hold and notice "
                  "given still hold a place; withdrawn does not."),
    ("Attendance %", "Sessions actually attended, divided by sessions "
                     "scheduled. Public holidays, closures and snow "
                     "days are taken out of both sides, because they "
                     "are nobody's fault."),
    ("Staff-to-child ratio", "Children present in a room today, "
                             "against the adults rota'd into that "
                             "room today. A shift only counts when it "
                             "has a room against it and is not Off, "
                             "Leave, Training or Admin."),
    ("Weekly fee", "The agreed rate converted to a week by the "
                   "child's own billing basis, less the sibling or "
                   "manual discount, less any subsidy."),
    ("Invoice total", "The sum of the charge lines tagged with that "
                      "invoice number - or the typed amount, if there "
                      "are no lines behind it."),
    ("Billed", "Invoices, less credit notes, less write-offs. This is "
               "revenue."),
    ("Collected", "Payments, less refunds. This is cash. It is a "
                  "different number from billed, on purpose."),
    ("Family balance", "Everything that put the account up, less "
                       "everything that brought it down - deposits "
                       "included, whether or not they are matched to "
                       "a bill."),
    ("Late pickup fee", "Minutes past the child's own contracted "
                        "pickup time, less the grace period, charged "
                        "as a first block and then per minute."),
    ("Pay run cost", "Basic hours plus overtime at the multiplier, "
                     "plus bonus and adjustments, plus employer costs "
                     "- and contractors and volunteers carry no "
                     "employer costs."),
    ("Operating profit", "Billed for the month, less the full cost of "
                         "the pay runs, less the expenses logged. "
                         "Wages are counted once, from the Payroll "
                         "Tracker."),
    ("Weeks of cover", "Stock on hand divided by what you have used a "
                       "week since tracking started - which is how "
                       "you find out that the wipes run out on "
                       "Thursday."),
]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.accent)
    ws = s.ws

    s.banner(cfg.product_name,
             "A complete childcare business system in one spreadsheet "
             "- enrollment, attendance, tuition, staff, rooms and "
             "parents.", LAST_COL)
    s.set_col(0, 0, 3)
    s.set_col(1, 1, 26)
    s.set_col(2, LAST_COL, 13)

    row = 4
    row = s.note(
        row,
        "Everything in this workbook is worked out from five tabs you "
        "fill in as you go: the children, the register, the charges, "
        "the rota and the costs. Fill those honestly and the "
        "dashboard, the invoices and the reports look after "
        "themselves.",
        LAST_COL, height=32)
    row += 1

    # ------------------------------------------------- seven steps ----
    row = s.section(row, "Setting up, in seven steps", LAST_COL)
    for title, text in STEPS:
        s.write(row, 1, title, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, text, fmts.chrome("note"))
        s.row_height(row, 42)
        row += 1
    row += 1

    # ---------------------------------------------------- the tabs ----
    row = s.section(row, "What is on each tab", LAST_COL)
    row = s.table_head(row, 1, ["Tab", "What it is for", "You"],
                       [24, 58, 12])
    for name, text, kind in TABS:
        ws.write_url(row, 1, f"internal:{quote_sheet(name)}!A1",
                     fmts.chrome("link"), name)
        s.merge(row, 2, LAST_COL - 1, text, fmts.chrome("note"))
        s.write(row, LAST_COL, kind, fmts.chrome("label"))
        s.row_height(row, 28)
        row += 1
    row += 1

    # ---------------------------------------------- the house rules ---
    row = s.section(row, "Seven rules that keep it working", LAST_COL)
    for rule in RULES:
        s.merge(row, 1, LAST_COL, "-  " + rule, fmts.chrome("note"))
        s.row_height(row, 28)
        row += 1
    row += 1

    # -------------------------------------------- the printable ones --
    row = s.section(row, "The three things you will print", LAST_COL)
    printables = [
        ("Invoice or receipt",
         "Invoices & Receipts builds the document for you: reference, "
         "family, period, the charge lines behind it, what has been "
         "paid and what is left. Filter to the family, print the "
         "row's block, or copy it into the printable layout at the "
         "bottom of the Monthly Reports tab."),
        ("Daily child report",
         "Daily Reports writes the parent-facing summary sentence "
         "itself from the mood, the meals, the naps and the nappies "
         "you logged. Filter to today, filter to the child, print."),
        ("Incident report",
         "Incident Reports composes the whole paragraph - what "
         "happened, where, the injury, the first aid, who was told "
         "and how many minutes it took. Print it, have the parent "
         "sign it, tick Signed."),
    ]
    for label, text in printables:
        s.write(row, 1, label, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, text, fmts.chrome("note"))
        s.row_height(row, 44)
        row += 1
    row += 1

    # ------------------------------------------------ what it is not --
    row = s.section(row, "What this workbook does not do", LAST_COL)
    for limit in LIMITS:
        s.merge(row, 1, LAST_COL, "-  " + limit, fmts.chrome("note"))
        s.row_height(row, 34)
        row += 1
    row += 1

    # --------------------------------------------- the numbers used ---
    row = s.section(row, "The sums behind the headline numbers",
                    LAST_COL)
    row = s.table_head(row, 1, ["Figure", "How it is worked out"],
                       [26, 68])
    for label, text in SUMS:
        s.write(row, 1, label, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, text, fmts.chrome("note"))
        s.row_height(row, 32)
        row += 1
    row += 1

    # ------------------------------------------------- the small -------
    row = s.section(row, "The small print", LAST_COL)
    row = s.note(row, PRIVACY_NOTE, LAST_COL, height=34)
    row += 1
    row = s.note(row, DISCLAIMER, LAST_COL, height=76)
    row += 1
    s.write(row, 1, "Version", fmts.chrome("label_bold"))
    s.write(row, 2, cfg.version, fmts.chrome("label"))
    s.write(row, 4, "Made by", fmts.chrome("label_bold"))
    s.merge(row, 5, LAST_COL, cfg.author, fmts.chrome("label"))
    row += 1
    s.write(row, 1, "Password", fmts.chrome("label_bold"))
    s.write(row, 2, cfg.password, fmts.chrome("label"))
    s.write(row, 4, "Built for", fmts.chrome("label_bold"))
    s.merge(row, 5, LAST_COL, "Excel 2010 and later, Google Sheets, "
                              "LibreOffice, Numbers",
            fmts.chrome("label"))

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    s.protect()
    return s
````

### `daycare_business_spreadsheet/sheets/settings.py`

*662 lines · 2 functions*

> `Settings & Lists` - every number and every list the workbook runs on.
>
> Nothing in this product is hard coded to one country, one licence or
> one price list.  The age bands and the staff-to-child ratios that go
> with them, the tax switches, the fee basis, the late collection
> charge, the overtime threshold, the budget and all forty-odd
> drop-downs are set here, once, and every other tab follows.
>
> The ratio table in particular is a starting point and not a legal
> opinion.  How many two-year-olds one adult may supervise is a local
> licensing decision; the workbook's job is to apply whatever you tell
> it, consistently, on every room and every day.

| Defines | What it does |
|---------|--------------|
| `_month_keys(today, back, ahead)` | `yyyy-mm` keys around today, oldest first |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Settings & Lists'), `LAST_COL` (11), `LISTS` (43 items).

````python
"""`Settings & Lists` - every number and every list the workbook runs on.

Nothing in this product is hard coded to one country, one licence or
one price list.  The age bands and the staff-to-child ratios that go
with them, the tax switches, the fee basis, the late collection
charge, the overtime threshold, the budget and all forty-odd
drop-downs are set here, once, and every other tab follows.

The ratio table in particular is a starting point and not a legal
opinion.  How many two-year-olds one adult may supervise is a local
licensing decision; the workbook's job is to apply whatever you tell
it, consistently, on every room and every day.
"""

from __future__ import annotations

import datetime as _dt
from typing import List, Tuple

from ..config import (AGE_BANDS, ALLERGY_SEVERITY, ATTENDANCE_STATUSES,
                      BILLING_BASES, CHARGE_SOURCES, COMM_CATEGORIES,
                      COMM_DIRECTIONS, COMM_STATUSES, CONTACT_METHODS,
                      DIET_TYPES, DISCLAIMER, DOCUMENT_HOLDERS,
                      DOCUMENT_STATUSES, DOCUMENT_TYPES,
                      EATEN_AMOUNTS, EMPLOYMENT_STATUSES,
                      ENROLLMENT_STATUSES, EXPENSE_CATEGORIES,
                      FEE_TYPES, FREQUENCIES, GENDERS,
                      ID_PREFIXES, INCIDENT_CATEGORIES,
                      INCIDENT_LOCATIONS, INCIDENT_SEVERITY,
                      INQUIRY_SOURCES, LEAVE_TYPES, LEAVING_REASONS,
                      LEDGER_TYPES, MEAL_SITTINGS,
                      MEDICAL_RECORD_TYPES, MOODS, PAY_PERIODS,
                      PAY_TYPES, PAYMENT_METHODS, PRIVACY_NOTE,
                      PROGRAM_TYPES, RELATIONSHIPS, ROOM_STATUSES,
                      SCHEDULE_TYPES, SHIFT_TYPES, STAFF_ROLES,
                      SUPPLY_CATEGORIES, TOILETING_OPTIONS,
                      WAITLIST_PRIORITY, WAITLIST_STATUSES, Config,
                      Palette)
from ..layout import Registry, col_letter, dynamic_name, quote_sheet
from ..writer import SheetWriter, list_from, number_check

NAME = "Settings & Lists"
LAST_COL = 11

#: name -> values.  Everything the workbook offers you in a drop-down.
LISTS: List[Tuple[str, List[str]]] = [
    ("YesNoList", ["Yes", "No"]),
    ("GenderList", list(GENDERS)),
    ("ProgramList", list(PROGRAM_TYPES)),
    ("EnrollmentStatusList", list(ENROLLMENT_STATUSES)),
    ("ScheduleTypeList", list(SCHEDULE_TYPES)),
    ("AttendanceStatusList", list(ATTENDANCE_STATUSES)),
    ("BillingBasisList", list(BILLING_BASES)),
    ("FeeTypeList", list(FEE_TYPES)),
    ("ChargeSourceList", list(CHARGE_SOURCES)),
    ("LedgerTypeList", list(LEDGER_TYPES)),
    ("PaymentMethodList", list(PAYMENT_METHODS)),
    ("RelationshipList", list(RELATIONSHIPS)),
    ("ContactMethodList", list(CONTACT_METHODS)),
    ("DirectionList", list(COMM_DIRECTIONS)),
    ("MoodList", list(MOODS)),
    ("MealSittingList", list(MEAL_SITTINGS)),
    ("EatenAmountList", list(EATEN_AMOUNTS)),
    ("DietList", list(DIET_TYPES)),
    ("ToiletingList", list(TOILETING_OPTIONS)),
    ("AllergySeverityList", list(ALLERGY_SEVERITY)),
    ("MedicalRecordList", list(MEDICAL_RECORD_TYPES)),
    ("IncidentCategoryList", list(INCIDENT_CATEGORIES)),
    ("IncidentLocationList", list(INCIDENT_LOCATIONS)),
    ("IncidentSeverityList", list(INCIDENT_SEVERITY)),
    ("RoleList", list(STAFF_ROLES)),
    ("EmploymentStatusList", list(EMPLOYMENT_STATUSES)),
    ("PayTypeList", list(PAY_TYPES)),
    ("ShiftTypeList", list(SHIFT_TYPES)),
    ("LeaveTypeList", list(LEAVE_TYPES)),
    ("PayPeriodList", list(PAY_PERIODS)),
    ("ExpenseCategoryList", list(EXPENSE_CATEGORIES)),
    ("FrequencyList", list(FREQUENCIES)),
    ("SupplyCategoryList", list(SUPPLY_CATEGORIES)),
    ("DocumentTypeList", list(DOCUMENT_TYPES)),
    ("DocumentHolderList", list(DOCUMENT_HOLDERS)),
    ("DocumentStatusList", list(DOCUMENT_STATUSES)),
    ("CommCategoryList", list(COMM_CATEGORIES)),
    ("CommStatusList", list(COMM_STATUSES)),
    ("InquirySourceList", list(INQUIRY_SOURCES)),
    ("WaitlistStatusList", list(WAITLIST_STATUSES)),
    ("WaitlistPriorityList", list(WAITLIST_PRIORITY)),
    ("LeavingReasonList", list(LEAVING_REASONS)),
    ("RoomStatusList", list(ROOM_STATUSES)),
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
             "Your setting, your fees, your ratios, your lists. Set "
             "these once and every tab in the workbook follows.",
             LAST_COL)
    s.set_col(0, 0, 3)
    s.set_col(1, 2, 23)
    s.set_col(3, 3, 20)
    s.set_col(4, LAST_COL, 17)

    row = 4
    row = s.note(
        row,
        "White cells are yours to change; everything else is locked "
        "and calculated. Nothing here is specific to one country - set "
        "your own age bands and ratios, your own tax switches, your "
        "own fees and currency, and the rest of the workbook follows.",
        LAST_COL, height=32)
    row += 1

    # ---------------------------------------------------- the setting ---
    row = s.section(row, "Your setting", LAST_COL)
    row = s.field(row, "Setting name", cfg.centre.name, name="CentreName",
                  value_end=5, help_col=7,
                  help_text="Printed at the top of every report and "
                            "invoice.")
    row = s.field(row, "Owner / director", cfg.centre.owner,
                  name="OwnerName", value_end=5, help_col=7)
    row = s.field(row, "Address", cfg.centre.address, value_end=5,
                  help_col=7)
    row = s.field(row, "Phone", cfg.centre.phone, value_end=5,
                  help_col=7)
    row = s.field(row, "Email", cfg.centre.email, value_end=5,
                  help_col=7)
    row = s.field(row, "Website", cfg.centre.website, value_end=5,
                  help_col=7)
    row = s.field(row, "Licence number", cfg.centre.licence,
                  value_end=5, help_col=7,
                  help_text="Your childcare registration or licence "
                            "reference, for your own records.")
    row = s.field(row, "Tax / business number", cfg.centre.tax_id,
                  value_end=5, help_col=7)
    row = s.field(row, "Opening hours", cfg.centre.opening_hours,
                  value_end=5, help_col=7)
    row = s.field(row, "Strapline", cfg.centre.tagline, value_end=5,
                  help_col=7)
    row = s.field(row, "Currency symbol", cfg.currency_symbol,
                  name="CurrencySign",
                  help_text="Shown on screen. The cell formats follow "
                            "the symbol you build the file with.")
    row = s.field(row, "Currency code", cfg.currency_code)
    row = s.field(row, "Tracking started",
                  cfg.today - _dt.timedelta(days=90), style="date",
                  name="TrackingStart",
                  help_text="The day you began logging in this file. "
                            "Supply cover and averages measure from "
                            "here.")
    row = s.field(row, "Today", None, style="date", name="TodayDate",
                  mode="calc", formula="=TODAY()",
                  help_text="Live. Everything that says 'overdue', "
                            "'expiring' or 'this month' reads this "
                            "cell.")
    row = s.field(row, "This month", None, style="text",
                  name="ThisMonthKey", mode="calc",
                  formula='=TEXT(TODAY(),"yyyy-mm")',
                  help_text="The month key every monthly total groups "
                            "on.")
    row = s.field(row, "Weeks tracked", None, style="number1",
                  name="TrackingWeeks", mode="calc",
                  formula="=MAX(4,ROUND((TodayDate-TrackingStart)/7,0))",
                  help_text="Used for the 'used a week' and weeks of "
                            "cover figures on the supply tab.")
    row += 1

    # ------------------------------------------- opening and capacity ---
    row = s.section(row, "Opening hours, capacity and the year",
                    LAST_COL)
    row = s.field(row, "Opens at", cfg.open_hour, style="number1",
                  name="OpenHour", validate=number_check(0, 24),
                  help_text="As a 24 hour number - 7 is 7am, 18 is 6pm.")
    row = s.field(row, "Closes at", cfg.close_hour, style="number1",
                  name="CloseHour", validate=number_check(0, 24))
    row = s.field(row, "Hours open a day", None, style="number1",
                  name="OpenHoursDay", mode="calc",
                  formula="=MAX(1,CloseHour-OpenHour)",
                  help_text="The ratio figures divide rota hours by "
                            "this to work out how many adults were "
                            "actually in a room.")
    row = s.field(row, "Days open a week", cfg.open_days, style="int",
                  name="OpenDays",
                  validate=number_check(1, 7, decimal=False))
    row = s.field(row, "Weeks open a year", cfg.weeks_open, style="int",
                  name="WeeksOpen",
                  validate=number_check(1, 53, decimal=False),
                  help_text="Trading weeks, after closures. What a "
                            "child's place is worth over a year uses "
                            "this.")
    row = s.field(row, "Weeks paid a year", 52, style="int",
                  name="WeeksPaid",
                  validate=number_check(1, 53, decimal=False),
                  help_text="Used to turn a salary into an hourly "
                            "rate. Usually 52 even when you close for "
                            "two of them.")
    row = s.field(row, "Weeks in a term", cfg.term_weeks, style="int",
                  name="TermWeeks",
                  validate=number_check(1, 30, decimal=False),
                  help_text="Only used when a child is on a term "
                            "rate.")
    row = s.field(row, "Licensed capacity", cfg.licensed_capacity,
                  style="int", name="LicensedCapacity",
                  validate=number_check(1, 1000, decimal=False),
                  help_text="The number on your licence, for the whole "
                            "setting. Room by room capacity is on the "
                            "Classroom Management tab.")
    row = s.field(row, "Warn when a room is this full",
                  cfg.ratio_warning_percent, style="percent",
                  name="RatioWarning", validate=number_check(0, 1),
                  help_text="Rooms at or above this go amber, so you "
                            "open the waitlist before the last place "
                            "goes.")
    row += 1

    # ------------------------------------------------------ age bands ---
    row = s.section(row, "Age bands and staff-to-child ratios", LAST_COL)
    row = s.note(
        row,
        "These are the most country-specific numbers in the workbook, "
        "which is exactly why they are a table you can edit rather "
        "than a formula you cannot. Set them to whatever your own "
        "licence says. Every room, every ratio warning and every "
        "'moves up on' date follows this table.",
        LAST_COL, height=38)
    row += 1
    band_head = row
    s.table_head(band_head, 1, [
        "Age band", "From (months)", "To (months)",
        "Children per adult", "Max group size"],
        widths=[22, 15, 15, 18, 16])
    first_band = band_head + 1
    band_rows = len(AGE_BANDS) + 2
    for index in range(band_rows):
        line = first_band + index
        if index < len(AGE_BANDS):
            label, m_from, m_to, ratio, group = AGE_BANDS[index]
            s.write(line, 1, label, fmts.input("text"))
            s.write(line, 2, m_from, fmts.input("int"))
            s.write(line, 3, m_to, fmts.input("int"))
            s.write(line, 4, ratio, fmts.input("int"))
            s.write(line, 5, group, fmts.input("int"))
        else:
            for col in range(1, 6):
                s.write(line, col, None,
                        fmts.input("text" if col == 1 else "int"))
    last_band = first_band + band_rows
    for key, column in (("AgeBandList", 1), ("AgeBandFrom", 2),
                        ("AgeBandTo", 3), ("AgeBandRatio", 4),
                        ("AgeBandMax", 5)):
        book.define_name(
            key,
            f"={sheet_ref}!${col_letter(column)}${first_band + 1}:"
            f"${col_letter(column)}${last_band}")
    row = last_band + 1

    # ---------------------------------------------------------- fees ----
    row = s.section(row, "Fees and billing", LAST_COL)
    row = s.note(
        row,
        "Each child's own basis and rate live on the Child Enrollment "
        "tab - these are the standard rates the waitlist and the "
        "quick calculations use.",
        LAST_COL, height=26)
    row += 1
    row = s.field(row, "Full-time week", cfg.full_time_week,
                  style="money", name="FullTimeWeek",
                  validate=number_check(0, 100000))
    row = s.field(row, "Part-time day", cfg.part_time_day,
                  style="money", name="PartTimeDay",
                  validate=number_check(0, 100000))
    row = s.field(row, "Hourly rate", cfg.hourly_rate_fee,
                  style="money", name="HourlyFee",
                  validate=number_check(0, 10000))
    row = s.field(row, "Registration fee", cfg.registration_fee,
                  style="money", name="RegistrationFee",
                  validate=number_check(0, 100000))
    row = s.field(row, "Enrollment deposit", cfg.enrollment_deposit,
                  style="money", name="EnrollmentDeposit",
                  validate=number_check(0, 100000),
                  help_text="A deposit is a payment with no invoice "
                            "behind it yet, so it can never be "
                            "counted as income twice.")
    row = s.field(row, "Meal fee a day", cfg.meal_fee_day,
                  style="money", name="MealFeeDay",
                  validate=number_check(0, 10000))
    row = s.field(row, "Transport fee a week", cfg.transport_fee_week,
                  style="money", name="TransportFeeWeek",
                  validate=number_check(0, 10000))
    row = s.field(row, "Sibling discount", cfg.sibling_discount,
                  style="percent", name="SiblingDiscount",
                  validate=number_check(0, 1),
                  help_text="Applied automatically to the second and "
                            "later child of a family.")
    row = s.field(row, "Staff discount", cfg.staff_discount,
                  style="percent", name="StaffDiscount",
                  validate=number_check(0, 1))
    row = s.field(row, "Payment terms, days", cfg.payment_terms_days,
                  style="int", name="PaymentTermsDays",
                  validate=number_check(0, 365, decimal=False),
                  help_text="How long after the invoice date it falls "
                            "due.")
    row = s.field(row, "Invoice on day of month", cfg.invoice_day,
                  style="int", name="InvoiceDay",
                  validate=number_check(1, 28, decimal=False))
    row = s.field(row, "Late payment fee", cfg.late_payment_fee,
                  style="money", name="LatePaymentFee",
                  validate=number_check(0, 10000))
    row += 1

    # ----------------------------------------------- late collection ----
    row = s.section(row, "Late collection charges", LAST_COL)
    row = s.note(
        row,
        "Measured against each child's own contracted pickup time on "
        "the enrollment tab, not against closing time. Free for the "
        "grace period, then the first block, then per minute - which "
        "is how most parent handbooks word it.",
        LAST_COL, height=32)
    row += 1
    row = s.field(row, "Grace, minutes", cfg.late_pickup_grace,
                  style="int", name="LatePickupGrace",
                  validate=number_check(0, 120, decimal=False))
    row = s.field(row, "First block, minutes",
                  cfg.late_pickup_block, style="int",
                  name="LatePickupBlock",
                  validate=number_check(0, 120, decimal=False))
    row = s.field(row, "First block charge", cfg.late_pickup_fee,
                  style="money", name="LatePickupFee",
                  validate=number_check(0, 10000))
    row = s.field(row, "Then per minute", cfg.late_pickup_per,
                  style="money", name="LatePickupPerMin",
                  validate=number_check(0, 1000))
    row += 1

    # ----------------------------------------------------- tax & fees ---
    row = s.section(row, "Tax and payment processing", LAST_COL)
    row = s.note(
        row,
        "Childcare is tax exempt in a great many places, so the "
        "default here is zero and nothing is assumed. Set the rate and "
        "say what it applies to - tuition, extras, both or neither.",
        LAST_COL, height=28)
    row += 1
    row = s.field(row, "Tax rate", cfg.tax_rate, style="percent",
                  name="TaxRate", validate=number_check(0, 1),
                  help_text="20% is 0.2 or 20% - both work.")
    row = s.field(row, "Tax on tuition?", cfg.tax_on_tuition,
                  name="TaxOnTuition", validate=list_from("YesNoList"))
    row = s.field(row, "Tax on extras?", cfg.tax_on_extras,
                  name="TaxOnExtras", validate=list_from("YesNoList"),
                  help_text="Meals, trips, late fees and everything "
                            "else that is not tuition.")
    row = s.field(row, "Tax registered?", cfg.tax_registered,
                  validate=list_from("YesNoList"),
                  help_text="For your own reference - it changes "
                            "nothing in the formulas.")
    row = s.field(row, "Card fee, percent", cfg.card_fee_percent,
                  style="percent", name="CardFeePercent",
                  validate=number_check(0, 1))
    row = s.field(row, "Card fee, fixed", cfg.card_fee_fixed,
                  style="money", name="CardFeeFixed",
                  validate=number_check(0, 100))
    row += 1

    # ------------------------------------------------------ staff pay ---
    row = s.section(row, "Staff pay and hours", LAST_COL)
    row = s.field(row, "Default hourly rate", cfg.default_hourly,
                  style="money", name="DefaultHourly",
                  validate=number_check(0, 1000))
    row = s.field(row, "Overtime after, hours a week",
                  cfg.overtime_after_hours, style="number1",
                  name="OvertimeAfter", validate=number_check(0, 100),
                  help_text="Hours above this in a single week are "
                            "paid at the multiplier below.")
    row = s.field(row, "Overtime multiplier",
                  cfg.overtime_multiplier, style="number1",
                  name="OvertimeMultiplier",
                  validate=number_check(1, 5))
    row = s.field(row, "Employer costs on wages", cfg.payroll_burden,
                  style="percent", name="PayrollBurden",
                  validate=number_check(0, 1),
                  help_text="Employer tax, insurance and cover as a "
                            "percentage of gross pay. Contractors and "
                            "volunteers carry none.")
    row = s.field(row, "Unpaid break, minutes", cfg.break_minutes,
                  style="int", name="BreakMinutes",
                  validate=number_check(0, 240, decimal=False))
    row = s.field(row, "Payroll sits in expenses?", "No",
                  name="PayrollInExpenses",
                  validate=list_from("YesNoList"),
                  help_text="Leave this No and wages come from the "
                            "Payroll Tracker only. Set it Yes and the "
                            "wage rows on the Expense Tracker are "
                            "counted instead. Never both.")
    row += 1

    # -------------------------------------------------------- targets ---
    row = s.section(row, "Targets the dashboard measures you against",
                    LAST_COL)
    row = s.field(row, "Revenue target, month",
                  cfg.revenue_target_month, style="money",
                  name="RevenueTargetMonth",
                  validate=number_check(0, 10000000))
    row = s.field(row, "Revenue target, year", cfg.revenue_target_year,
                  style="money", name="RevenueTargetYear",
                  validate=number_check(0, 100000000))
    row = s.field(row, "Profit target, month", cfg.profit_target_month,
                  style="money", name="ProfitTargetMonth",
                  validate=number_check(0, 10000000))
    row = s.field(row, "Children enrolled, target",
                  cfg.enrollment_goal, style="int",
                  name="EnrollmentTarget",
                  validate=number_check(0, 1000, decimal=False))
    row = s.field(row, "Occupancy target", cfg.target_occupancy,
                  style="percent", name="TargetOccupancy",
                  validate=number_check(0, 1))
    row = s.field(row, "Attendance target", cfg.target_attendance,
                  style="percent", name="TargetAttendance",
                  validate=number_check(0, 1))
    row = s.field(row, "Collection target", cfg.target_collection,
                  style="percent", name="TargetCollection",
                  validate=number_check(0, 1),
                  help_text="The share of what you bill that actually "
                            "arrives. Below 95% and something is "
                            "wrong with the chasing, not the "
                            "families.")
    row += 1

    # ------------------------------------------------------ reminders ---
    row = s.section(row, "Reminders and chase windows", LAST_COL)
    row = s.field(row, "Warn about expiry, days",
                  cfg.document_warning_days, style="int",
                  name="DocWarningDays",
                  validate=number_check(1, 365, decimal=False),
                  help_text="Amber this many days before a document "
                            "or certificate runs out.")
    row = s.field(row, "Urgent expiry, days", cfg.document_urgent_days,
                  style="int", name="DocUrgentDays",
                  validate=number_check(1, 365, decimal=False))
    row = s.field(row, "Birthday window, days",
                  cfg.birthday_window_days, style="int",
                  name="BirthdayWindow",
                  validate=number_check(1, 365, decimal=False))
    row = s.field(row, "Chase an absence after, days",
                  cfg.absence_chase_days, style="int",
                  name="AbsenceChaseDays",
                  validate=number_check(1, 30, decimal=False),
                  help_text="Unexplained absences inside this window "
                            "raise a prompt on the attendance tab.")
    row = s.field(row, "Chase an invoice after, days",
                  cfg.overdue_chase_days, style="int",
                  name="OverdueChaseDays",
                  validate=number_check(1, 180, decimal=False))
    row += 1

    # ------------------------------------------------------- supplies ---
    row = s.section(row, "Supplies and reordering", LAST_COL)
    row = s.field(row, "Order up to, weeks of cover",
                  cfg.reorder_weeks, style="number1",
                  name="ReorderWeeks", validate=number_check(0.5, 52))
    row = s.field(row, "Warn below, weeks of cover",
                  cfg.stock_warning_weeks, style="number1",
                  name="StockWarningWeeks",
                  validate=number_check(0, 52))
    row += 1

    # --------------------------------------------------- the budget ----
    row = s.section(row, "Monthly budget by category", LAST_COL)
    row = s.note(
        row,
        "What you expect to spend in a normal month. The Expense "
        "Tracker reads this table, compares it with what you actually "
        "spent and shows the variance in the month it happened. Leave "
        "a category at zero and it simply is not measured.",
        LAST_COL, height=32)
    row += 1
    budget_head = row
    s.table_head(budget_head, 1, ["Category", "Budget a month"],
                 widths=[28, 16])
    first_budget = budget_head + 1
    budget_rows = len(EXPENSE_CATEGORIES)
    #: a sensible opening budget so the variance column has something
    #: to say on day one - every figure editable
    seed = {
        "Rent": 3200, "Utilities": 620, "Food & groceries": 1850,
        "Nappies & hygiene": 420, "Cleaning supplies": 240,
        "Toys & learning materials": 300,
        "Educational supplies": 220, "Arts & crafts": 150,
        "Insurance": 310, "Licensing & registration": 90,
        "Repairs & maintenance": 260,
        "Marketing & advertising": 180,
        "Software & subscriptions": 95,
        "Telephone & internet": 110, "Office & admin": 95,
        "Training & qualifications": 150, "Bank & card fees": 85,
        "Professional fees": 140, "Waste & recycling": 70,
    }
    for index, category in enumerate(EXPENSE_CATEGORIES):
        line = first_budget + index
        s.write(line, 1, category, fmts.input("text"))
        s.write(line, 2, float(seed.get(category, 0)),
                fmts.input("money"))
    last_budget = first_budget + budget_rows
    book.define_name(
        "BudgetCategory",
        f"={sheet_ref}!$B${first_budget + 1}:$B${last_budget}")
    book.define_name(
        "BudgetAmount",
        f"={sheet_ref}!$C${first_budget + 1}:$C${last_budget}")
    total_line = last_budget
    s.write(total_line, 1, "Total budget a month",
            fmts.chrome("total_label"))
    s.write(total_line, 2, "=SUM(BudgetAmount)",
            fmts.chrome("total_value",
                        num_format=cfg.money_format))
    book.define_name(
        "BudgetMonthlyTotal",
        f"={sheet_ref}!$C${total_line + 1}")
    row = total_line + 2

    # ----------------------------------------------- workbook totals ----
    row = s.section(row, "Workbook totals", LAST_COL)
    row = s.note(
        row,
        "Calculated once here so several tabs can share them. Nothing "
        "to change.",
        LAST_COL, height=24)
    row += 1
    row = s.field(row, "Children holding a place", None, style="int",
                  name="ChildrenEnrolled", mode="calc",
                  formula='=COUNTIFS([[Child Enrollment.holds_place]],'
                          '"Yes")')
    row = s.field(row, "Children attending", None, style="int",
                  name="ChildrenActive", mode="calc",
                  formula='=COUNTIFS([[Child Enrollment.status]],'
                          '"Active")')
    row = s.field(row, "Staff on the books", None, style="int",
                  name="StaffCount", mode="calc",
                  formula='=COUNTIFS([[Staff Database.status]],'
                          '"Active")+'
                          'COUNTIFS([[Staff Database.status]],'
                          '"Probation")')
    row = s.field(row, "Places across all rooms", None, style="int",
                  name="RoomCapacity", mode="calc",
                  formula="=SUM([[Classroom Management.capacity]])")
    row = s.field(row, "Revenue, all time", None, style="money",
                  name="RevenueAll", mode="calc",
                  formula="=ROUND(SUM([[Invoices & Receipts.billed]]),"
                          "2)")
    row = s.field(row, "Collected, all time", None, style="money",
                  name="CollectedAll", mode="calc",
                  formula="=ROUND(SUM([[Invoices & Receipts.collected]])"
                          ",2)")
    row += 1

    # ------------------------------------------------------ prefixes ----
    row = s.section(row, "ID prefixes", LAST_COL)
    row = s.note(
        row,
        "Every record numbers itself from its row. Change a prefix "
        "here and the IDs on that tab change with it - but do it "
        "before you have typed three hundred attendance rows, not "
        "after.",
        LAST_COL, height=28)
    row += 1
    half = (len(ID_PREFIXES) + 1) // 2
    top = row
    for index, (label, default, name, _digits) in enumerate(ID_PREFIXES):
        column = 0 if index < half else 4
        line = top + (index if index < half else index - half)
        s.write(line, column + 1, label, fmts.chrome("label_bold"))
        s.write(line, column + 2, default, fmts.input("text"))
        book.define_name(
            name, f"={sheet_ref}!${col_letter(column + 2)}${line + 1}")
    row = top + half + 1

    # -------------------------------------------------------- lists -----
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

    # ---------------------------------------------------- month keys ----
    row = s.section(row, "Month keys", LAST_COL)
    row = s.note(
        row,
        "The months the reports and the budget pick from. They are "
        "plain text in yyyy-mm form so that every monthly total in "
        "the workbook groups on exactly the same thing.",
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

    row = s.note(row, PRIVACY_NOTE, LAST_COL, height=30)
    row += 1
    row = s.note(row, DISCLAIMER, LAST_COL, height=74)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    s.protect()
    return s
````

### `daycare_business_spreadsheet/sheets/monthly.py`

*542 lines · 1 function*

> `Monthly Reports` - one month, on one page you can print.
>
> Pick a month at the top and everything below it moves: enrollment in
> and out, attendance room by room, what you billed and what actually
> arrived, where the money went, and how the staffing stacked up
> against it.
>
> At the bottom sits the printable invoice.  Choose an invoice number
> and the whole document builds itself - your details, the family's,
> the charge lines behind it, what has been paid and what is left.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Monthly Reports'), `LAST_COL` (11), `ROOM_ROWS` (8), `LINE_ROWS` (14), `CAT_ROWS` (12).

````python
"""`Monthly Reports` - one month, on one page you can print.

Pick a month at the top and everything below it moves: enrollment in
and out, attendance room by room, what you billed and what actually
arrived, where the money went, and how the staffing stacked up
against it.

At the bottom sits the printable invoice.  Choose an invoice number
and the whole document builds itself - your details, the family's,
the charge lines behind it, what has been paid and what is left.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, col_letter, quote_sheet
from ..writer import SheetWriter, list_from

NAME = "Monthly Reports"
LAST_COL = 11
ROOM_ROWS = 8
LINE_ROWS = 14
CAT_ROWS = 12


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.deep)
    ws = s.ws
    sheet_ref = quote_sheet(NAME)

    s.banner("Monthly Report",
             "One month of your setting on a single printable page - "
             "children, attendance, money, costs and staffing.",
             LAST_COL)
    s.set_col(0, 0, 24)
    for column in range(1, LAST_COL + 1):
        s.set_col(column, column, 14)

    # ------------------------------------------------------- the picker --
    row = 4
    s.merge(row, 0, 0, "Month", fmts.chrome("label_bold"))
    s.write(row, 1, None, fmts.input("text"))
    ws.write_formula(row, 1, '=TEXT(TODAY(),"yyyy-mm")',
                     fmts.input("text"))
    ws.data_validation(row, 1, row, 1, list_from("MonthKeyList"))
    book.define_name("ReportMonth", f"={sheet_ref}!$B${row + 1}")
    s.merge(row, 2, 2, "Reading", fmts.chrome("label_right"))
    s.merge(row, 3, 4, None, fmts.calc("text"))
    ws.write_formula(
        row, 3,
        '=IFERROR(TEXT(DATEVALUE(ReportMonth&"-01"),"mmmm yyyy"),'
        '"Pick a month")', fmts.calc("text"))
    s.merge(row, 5, 6, "Compared with", fmts.chrome("label_right"))
    s.merge(row, 7, 8, None, fmts.calc("text"))
    ws.write_formula(
        row, 7,
        '=IFERROR(TEXT(DATE(VALUE(LEFT(ReportMonth,4)),'
        'VALUE(RIGHT(ReportMonth,2))-1,1),"yyyy-mm"),"")',
        fmts.calc("text"))
    book.define_name("ReportPrev", f"={sheet_ref}!$H${row + 1}")
    # Two helper cells, off to the right and hidden: the first and the
    # last day of the chosen month. Counting dated rows with COUNTIFS
    # over that window is faster than TEXT() across a whole column and
    # it never trips over an empty cell.
    s.set_col(13, 14, 11, None, hidden=True)
    ws.write_formula(
        row, 13,
        '=IFERROR(DATE(VALUE(LEFT(ReportMonth,4)),'
        'VALUE(RIGHT(ReportMonth,2)),1),TodayDate)',
        fmts.calc("date"))
    book.define_name("ReportFrom", f"={sheet_ref}!$N${row + 1}")
    ws.write_formula(
        row, 14,
        '=IFERROR(DATE(VALUE(LEFT(ReportMonth,4)),'
        'VALUE(RIGHT(ReportMonth,2))+1,1)-1,TodayDate)',
        fmts.calc("date"))
    book.define_name("ReportTo", f"={sheet_ref}!$O${row + 1}")
    s.merge(row, 9, 9, "Printed", fmts.chrome("label_right"))
    s.merge(row, 10, LAST_COL, None, fmts.calc("date"))
    ws.write_formula(row, 10, "=TodayDate", fmts.calc("date"))
    row += 2

    row = s.note(
        row,
        "Change the month in the blue cell and every figure on this "
        "page follows it. Months are plain text in yyyy-mm form, "
        "which is what every monthly total in the workbook groups on.",
        LAST_COL, height=28)
    row += 1

    billed = ('SUMIFS([[Invoices & Receipts.billed]],'
              '[[Invoices & Receipts.month]],ReportMonth)')
    billed_prev = ('SUMIFS([[Invoices & Receipts.billed]],'
                   '[[Invoices & Receipts.month]],ReportPrev)')
    collected = ('SUMIFS([[Invoices & Receipts.collected]],'
                 '[[Invoices & Receipts.month]],ReportMonth)')
    wages = ('SUMIFS([[Payroll Tracker.total_cost]],'
             '[[Payroll Tracker.month]],ReportMonth)')
    costs = ('SUMIFS([[Expense Tracker.pl_amount]],'
             '[[Expense Tracker.month]],ReportMonth)')
    sessions = ('SUMIFS([[Attendance.session]],'
                '[[Attendance.month]],ReportMonth)')
    present = ('SUMIFS([[Attendance.present]],'
               '[[Attendance.month]],ReportMonth)')

    # ----------------------------------------------------- the headlines --
    row = s.section(row, "The month in nine numbers", LAST_COL)
    tiles = [
        ("Billed", f"=ROUND({billed},2)", "money0",
         "invoices less credits and write-offs"),
        ("Collected", f"=ROUND({collected},2)", "money0",
         "payments less refunds"),
        ("Operating profit", f"=ROUND({billed}-{wages}-{costs},2)",
         "money0", "after wages and costs"),
        ("Margin",
         f'=IF({billed}<=0,"",({billed}-{wages}-{costs})/{billed})',
         "percent", "of everything billed"),
        ("Attendance",
         f'=IF({sessions}<=0,"",{present}/{sessions})', "percent",
         f"target {cfg.target_attendance:.0%}"),
        ("Sessions attended", f"={present}", "int",
         "closures and holidays excluded"),
        ("Children billed",
         '=COUNTIFS([[Invoices & Receipts.month]],ReportMonth,'
         '[[Invoices & Receipts.type]],"Invoice")', "int",
         "one invoice per family"),
        ("Collection rate",
         f'=IF({billed}<=0,"",{collected}/{billed})', "percent",
         f"target {cfg.target_collection:.0%}"),
        ("Against last month",
         f'=IF({billed_prev}<=0,"",{billed}/{billed_prev}-1)',
         "percent", "billed, month on month"),
    ]
    column = 0
    bottom = row
    for index, (label, formula, style, note) in enumerate(tiles):
        if index and index % 3 == 0:
            row = bottom
            column = 0
        bottom = s.kpi(row, column, label, formula, style, span=4,
                       note=note)
        column += 4
    row = bottom + 1

    # ---------------------------------------------------- enrollment -----
    row = s.section(row, "Children in and out", LAST_COL)
    row = s.table_head(row, 0, [
        "On the books", "Attending", "Started this month",
        "Left this month", "Net change", "Places free",
        "Occupancy", "On the waitlist", "Offers out",
        "Enquiries this month", "Converted"],
        [14, 12, 15, 14, 12, 12, 11, 14, 11, 16, 12])
    s.calc_row(row, 0, [
        "=ChildrenEnrolled",
        "=ChildrenActive",
        '=COUNTIFS([[Child Enrollment.start_date]],">="&ReportFrom,[[Child Enrollment.start_date]],"<="&ReportTo)',
        '=COUNTIFS([[Child Enrollment.end_date]],">="&ReportFrom,[[Child Enrollment.end_date]],"<="&ReportTo)',
        '=COUNTIFS([[Child Enrollment.start_date]],">="&ReportFrom,[[Child Enrollment.start_date]],"<="&ReportTo)-COUNTIFS([[Child Enrollment.end_date]],">="&ReportFrom,[[Child Enrollment.end_date]],"<="&ReportTo)',
        "=MAX(0,RoomCapacity-ChildrenEnrolled)",
        '=IF(RoomCapacity<=0,"",ChildrenEnrolled/RoomCapacity)',
        '=COUNTIFS([[Waitlist.closed]],"No")',
        '=COUNTIFS([[Waitlist.status]],"Place offered")',
        '=COUNTIFS([[Waitlist.enquiry_date]],">="&ReportFrom,[[Waitlist.enquiry_date]],"<="&ReportTo)',
        '=COUNTIFS([[Waitlist.converted]],"Yes")',
    ], ["int", "int", "int", "int", "int", "int", "percent", "int",
        "int", "int", "int"])
    row += 2

    # ------------------------------------------------- room by room ------
    row = s.section(row, "Room by room", LAST_COL)
    row = s.table_head(row, 0, [
        "Room", "Places", "Enrolled", "Occupancy", "Sessions",
        "Attended", "Absent", "Attendance", "Weekly income",
        "Income a place", "Empty seats cost"],
        [18, 10, 11, 11, 11, 11, 10, 12, 14, 14, 15])
    rooms_first = row
    for n in range(1, ROOM_ROWS + 1):
        name = f'INDEX([[Classroom Management.name]],{n})'
        guard = f'IFERROR(IF({name}="","",{{body}}),"")'
        sess = (f'SUMIFS([[Attendance.session]],[[Attendance.room]],'
                f'{name},[[Attendance.month]],ReportMonth)')
        here = (f'SUMIFS([[Attendance.present]],[[Attendance.room]],'
                f'{name},[[Attendance.month]],ReportMonth)')
        away = (f'SUMIFS([[Attendance.absent]],[[Attendance.room]],'
                f'{name},[[Attendance.month]],ReportMonth)')
        cells = [
            (0, f'IFERROR({name}&"","")', "text"),
            (1, guard.replace(
                "{body}", f'INDEX([[Classroom Management.capacity]],'
                          f'{n})'), "int"),
            (2, guard.replace(
                "{body}", f'INDEX([[Classroom Management.enrolled]],'
                          f'{n})'), "int"),
            (3, guard.replace(
                "{body}", f'INDEX([[Classroom Management.occupancy]],'
                          f'{n})'), "percent"),
            (4, guard.replace("{body}", sess), "int"),
            (5, guard.replace("{body}", here), "int"),
            (6, guard.replace("{body}", away), "int"),
            (7, guard.replace(
                "{body}", f'IF({sess}<=0,"",{here}/{sess})'),
             "percent"),
            (8, guard.replace(
                "{body}", f'INDEX([[Classroom Management.revenue]],'
                          f'{n})'), "money"),
            (9, guard.replace(
                "{body}", f'INDEX([[Classroom Management.revenue_seat]]'
                          f',{n})'), "money"),
            (10, guard.replace(
                "{body}", f'INDEX([[Classroom Management.empty_cost]],'
                          f'{n})'), "money"),
        ]
        for column, formula, style in cells:
            ws.write_formula(row, column, s.reg.render("=" + formula),
                             fmts.calc(style))
        row += 1
    ws.conditional_format(rooms_first, 7, row - 1, 7,
                          {"type": "cell", "criteria": "<",
                           "value": "=TargetAttendance",
                           "format": fmts.cf("warn")})
    ws.conditional_format(rooms_first, 3, row - 1, 3,
                          {"type": "data_bar",
                           "bar_color": Palette.primary_mid})
    row += 1

    # ------------------------------------------------------ the money ----
    row = s.section(row, "Money in", LAST_COL)
    row = s.table_head(row, 0, [
        "Invoiced", "Credit notes", "Written off", "Billed",
        "Payments", "Refunds", "Collected", "Card fees",
        "Still owed, all time", "Overdue invoices", "Collection rate"],
        [13, 13, 12, 13, 12, 11, 12, 11, 17, 15, 13])
    s.calc_row(row, 0, [
        '=ROUND(SUMIFS([[Invoices & Receipts.total]],'
        '[[Invoices & Receipts.month]],ReportMonth,'
        '[[Invoices & Receipts.type]],"Invoice"),2)',
        '=ROUND(SUMIFS([[Invoices & Receipts.total]],'
        '[[Invoices & Receipts.month]],ReportMonth,'
        '[[Invoices & Receipts.type]],"Credit note"),2)',
        '=ROUND(SUMIFS([[Invoices & Receipts.total]],'
        '[[Invoices & Receipts.month]],ReportMonth,'
        '[[Invoices & Receipts.type]],"Write-off"),2)',
        f"=ROUND({billed},2)",
        '=ROUND(SUMIFS([[Invoices & Receipts.total]],'
        '[[Invoices & Receipts.month]],ReportMonth,'
        '[[Invoices & Receipts.type]],"Payment"),2)',
        '=ROUND(SUMIFS([[Invoices & Receipts.total]],'
        '[[Invoices & Receipts.month]],ReportMonth,'
        '[[Invoices & Receipts.type]],"Refund"),2)',
        f"=ROUND({collected},2)",
        '=ROUND(SUMIFS([[Invoices & Receipts.fee]],'
        '[[Invoices & Receipts.month]],ReportMonth),2)',
        "=ROUND(SUM([[Parent Database.balance]]),2)",
        '=COUNTIFS([[Invoices & Receipts.status]],"Overdue")',
        f'=IF({billed}<=0,"",{collected}/{billed})',
    ], ["money", "money", "money", "money", "money", "money", "money",
        "money", "money", "int", "percent"])
    row += 2

    # -------------------------------------------------- what was charged --
    row = s.section(row, "What was charged, by type", LAST_COL)
    row = s.table_head(row, 0, [
        "Tuition", "Registration", "Late pickup", "Meals",
        "Transport", "Trips & activities", "Other charges",
        "Discounts given", "Subsidies", "Net charges",
        "Not yet invoiced"],
        [13, 13, 12, 11, 11, 15, 13, 14, 12, 12, 14])
    charge = ('SUMIFS([[Tuition & Fees.net]],[[Tuition & Fees.month]],'
              'ReportMonth,[[Tuition & Fees.fee_type]],"{t}")')
    others = ("+".join(
        charge.replace("{t}", kind) for kind in
        ("Snack fee", "Activity fee", "Supply fee",
         "Nappy & wipes fee", "Holiday club", "Extra day",
         "Extra hours", "Late payment fee", "Other charge")))
    discounts = ("+".join(
        charge.replace("{t}", kind) for kind in
        ("Sibling discount", "Staff discount", "Promotional discount",
         "Goodwill credit")))
    subsidies = ("+".join(
        charge.replace("{t}", kind) for kind in
        ("Subsidy credit", "Scholarship credit")))
    s.calc_row(row, 0, [
        "=ROUND(" + charge.replace("{t}", "Tuition") + ",2)",
        "=ROUND(" + charge.replace("{t}", "Registration fee")
        + "+" + charge.replace("{t}", "Enrollment deposit") + ",2)",
        "=ROUND(" + charge.replace("{t}", "Late pickup fee") + ",2)",
        "=ROUND(" + charge.replace("{t}", "Meal fee") + ",2)",
        "=ROUND(" + charge.replace("{t}", "Transport fee") + ",2)",
        "=ROUND(" + charge.replace("{t}", "Field trip") + ",2)",
        "=ROUND(" + others + ",2)",
        "=ROUND(" + discounts + ",2)",
        "=ROUND(" + subsidies + ",2)",
        '=ROUND(SUMIFS([[Tuition & Fees.net]],'
        '[[Tuition & Fees.month]],ReportMonth),2)',
        '=ROUND(SUMIFS([[Tuition & Fees.net]],'
        '[[Tuition & Fees.month]],ReportMonth,'
        '[[Tuition & Fees.billed]],"No"),2)',
    ], ["money"] * 11)
    row += 2

    # ----------------------------------------------------- money out -----
    row = s.section(row, "Money out", LAST_COL)
    row = s.table_head(row, 0, [
        "Wages & on-costs", "Rota hours", "Overtime hours",
        "Wage cost of income", "Other costs", "Budget",
        "Over or under", "Cost a child", "Operating profit",
        "Margin", "Profit a child"],
        [16, 12, 14, 17, 13, 12, 14, 12, 15, 10, 13])
    s.calc_row(row, 0, [
        f"=ROUND({wages},2)",
        '=ROUND(SUMIFS([[Payroll Tracker.total_hours]],'
        '[[Payroll Tracker.month]],ReportMonth),1)',
        '=ROUND(SUMIFS([[Payroll Tracker.ot_hours]],'
        '[[Payroll Tracker.month]],ReportMonth),1)',
        f'=IF({billed}<=0,"",{wages}/{billed})',
        f"=ROUND({costs},2)",
        "=ROUND(BudgetMonthlyTotal,2)",
        f"=ROUND({costs}-BudgetMonthlyTotal,2)",
        f'=IF(ChildrenEnrolled<=0,"",ROUND(({wages}+{costs})/'
        f'ChildrenEnrolled,2))',
        f"=ROUND({billed}-{wages}-{costs},2)",
        f'=IF({billed}<=0,"",({billed}-{wages}-{costs})/{billed})',
        f'=IF(ChildrenEnrolled<=0,"",ROUND(({billed}-{wages}-{costs})/'
        f'ChildrenEnrolled,2))',
    ], ["money", "number1", "number1", "percent", "money", "money",
        "money", "money", "money", "percent", "money"])
    row += 2

    # ------------------------------------------- expenses by category ----
    row = s.section(row, "Where the money went", LAST_COL)
    row = s.table_head(row, 0, [
        "Category", "Spent this month", "Budget", "Over or under",
        "Share of costs", "Category", "Spent this month", "Budget",
        "Over or under", "Share of costs"],
        [22, 15, 12, 14, 13, 22, 15, 12, 14, 13])
    cat_first = row
    for n in range(1, CAT_ROWS + 1):
        for half in (0, 1):
            index = n + half * CAT_ROWS
            base = half * 5
            cat = f'INDEX(BudgetCategory,{index})'
            spent = (f'SUMIFS([[Expense Tracker.pl_amount]],'
                     f'[[Expense Tracker.month]],ReportMonth,'
                     f'[[Expense Tracker.category]],{cat})')
            budget = f'INDEX(BudgetAmount,{index})'
            cells = [
                (base + 0, f'IFERROR({cat}&"","")', "text"),
                (base + 1, f'IFERROR(ROUND({spent},2),"")', "money"),
                (base + 2, f'IFERROR(ROUND({budget},2),"")', "money"),
                (base + 3, f'IFERROR(ROUND({spent}-{budget},2),"")',
                 "money"),
                (base + 4, f'IFERROR(IF({costs}<=0,"",{spent}/{costs}),'
                           f'"")', "percent"),
            ]
            for column, formula, style in cells:
                ws.write_formula(row, column,
                                 s.reg.render("=" + formula),
                                 fmts.calc(style))
        row += 1
    for column in (3, 8):
        ws.conditional_format(cat_first, column, row - 1, column,
                              {"type": "cell", "criteria": ">",
                               "value": 0.004,
                               "format": fmts.cf("bad")})
        ws.conditional_format(cat_first, column, row - 1, column,
                              {"type": "cell", "criteria": "<=",
                               "value": 0, "format": fmts.cf("good")})
    row += 1

    # ------------------------------------------------- care and safety ---
    row = s.section(row, "Care, safety and paperwork", LAST_COL)
    row = s.table_head(row, 0, [
        "Daily reports written", "Meals logged", "Meals refused",
        "Medication doses", "Incidents", "Of those, serious",
        "Average minutes to tell a parent", "Forms signed",
        "Parent conversations", "Complaints", "Compliments"],
        [18, 12, 13, 15, 11, 15, 22, 12, 17, 12, 13])
    s.calc_row(row, 0, [
        '=COUNTIFS([[Daily Reports.date]],">="&ReportFrom,[[Daily Reports.date]],"<="&ReportTo)',
        '=COUNTIFS([[Meal Tracker.month]],ReportMonth)',
        '=COUNTIFS([[Meal Tracker.month]],ReportMonth,'
        '[[Meal Tracker.refused]],"Yes")',
        '=COUNTIFS([[Allergy & Medical.given_date]],">="&ReportFrom,[[Allergy & Medical.given_date]],"<="&ReportTo)',
        '=COUNTIFS([[Incident Reports.month]],ReportMonth)',
        '=COUNTIFS([[Incident Reports.month]],ReportMonth,'
        '[[Incident Reports.severity]],"Serious")+'
        'COUNTIFS([[Incident Reports.month]],ReportMonth,'
        '[[Incident Reports.severity]],"Emergency services called")',
        '=IFERROR(ROUND(SUMIFS([[Incident Reports.notify_mins]],'
        '[[Incident Reports.month]],ReportMonth)/'
        'MAX(1,COUNTIFS([[Incident Reports.month]],ReportMonth,'
        '[[Incident Reports.notified]],"Yes")),0),0)',
        '=COUNTIFS([[Incident Reports.month]],ReportMonth,'
        '[[Incident Reports.signed]],"Yes")',
        '=COUNTIFS([[Parent Communication.month]],ReportMonth)',
        '=COUNTIFS([[Parent Communication.month]],ReportMonth,'
        '[[Parent Communication.category]],"Complaint")',
        '=COUNTIFS([[Parent Communication.month]],ReportMonth,'
        '[[Parent Communication.category]],"Compliment")',
    ], ["int", "int", "int", "int", "int", "int", "int", "int", "int",
        "int", "int"])
    row += 2

    # ------------------------------------------------ printable invoice --
    row = s.section(row, "Printable invoice or receipt", LAST_COL)
    row = s.note(
        row,
        "Pick a reference from the drop-down and the document builds "
        "itself from the ledger and the charge lines behind it. Print "
        "this block - the print area is set to end here.",
        LAST_COL, height=28)
    row += 1

    s.merge(row, 0, 0, "Reference", fmts.chrome("label_bold"))
    s.write(row, 1, None, fmts.input("text"))
    ws.data_validation(row, 1, row, 1, list_from("InvoiceList"))
    book.define_name("PrintInvoice", f"={sheet_ref}!$B${row + 1}")
    s.merge(row, 2, 3, "Choose an invoice number", fmts.chrome("note"))
    row += 2

    def look(field: str) -> str:
        return (f'IFERROR(INDEX([[Invoices & Receipts.{field}]],'
                f'MATCH(PrintInvoice,[[Invoices & Receipts.id]],0)),"")')

    s.merge(row, 0, 5, "", fmts.chrome("print_title"))
    ws.write_formula(row, 0, "=CentreName", fmts.chrome("print_title"))
    s.merge(row, 6, LAST_COL, None, fmts.chrome("print_sub"))
    ws.write_formula(
        row, 6,
        s.reg.render('=IF(PrintInvoice="","INVOICE",'
                     'IF(LEFT(PrintInvoice,LEN(ReceiptPrefix))'
                     '=ReceiptPrefix,"RECEIPT","INVOICE"))'),
        fmts.chrome("print_sub"))
    row += 1

    left = [
        ("From", "=CentreName"),
        ("", "=OwnerName"),
        ("Reference", "=PrintInvoice"),
        ("Date", f"={look('date')}"),
        ("Period", f"={look('period')}"),
    ]
    right = [
        ("Bill to", f"={look('family_name')}"),
        ("Family ID", f"={look('family')}"),
        ("Type", f"={look('type')}"),
        ("Due", f"={look('due_date')}"),
        ("Status", f"={look('status')}"),
    ]
    # the ruled lines are filled in by the ledger, not by hand, so they
    # stay locked even though they are drawn like a paper form
    filled = fmts.chrome("form_line", locked=True)
    for index in range(len(left)):
        label, formula = left[index]
        s.write(row, 0, label, fmts.chrome("form_label"))
        s.merge(row, 1, 4, None, filled)
        ws.write_formula(row, 1, s.reg.render(formula), filled)
        label, formula = right[index]
        s.write(row, 5, label, fmts.chrome("form_label"))
        s.merge(row, 6, LAST_COL, None, filled)
        ws.write_formula(row, 6, s.reg.render(formula), filled)
        row += 1
    row += 1

    row = s.table_head(row, 0, [
        "Line", "Child", "Description", "Period", "Quantity", "Rate",
        "Discount", "Tax", "Amount"],
        [7, 18, 30, 20, 10, 12, 11, 10, 13])
    for n in range(1, LINE_ROWS + 1):
        key = f'PrintInvoice&"|"&{n}'

        def line(field: str) -> str:
            return (f'IFERROR(INDEX([[Tuition & Fees.{field}]],'
                    f'MATCH({key},[[Tuition & Fees.key]],0)),"")')

        cells = [
            (0, f'IF({line("description")}="","",{n})', "int"),
            (1, line("child_name"), "text"),
            (2, line("description"), "text"),
            (3, f'IF({line("period_from")}="","",'
                f'TEXT({line("period_from")},"d mmm")&" - "&'
                f'TEXT({line("period_to")},"d mmm"))', "text"),
            (4, line("qty"), "number1"),
            (5, line("rate"), "money"),
            (6, line("discount"), "money"),
            (7, line("tax"), "money"),
            (8, line("net"), "money"),
        ]
        for column, formula, style in cells:
            ws.write_formula(row, column, s.reg.render("=" + formula),
                             fmts.calc(style))
        row += 1
    row += 1

    totals = [
        ("Charge lines", f"={look('charges')}", "money"),
        ("Invoice total", f"={look('total')}", "money"),
        ("Paid", f"={look('paid')}", "money"),
        ("Credited", f"={look('credited')}", "money"),
        ("Balance", f"={look('balance')}", "money"),
    ]
    for label, formula, style in totals:
        s.merge(row, 5, 7, label, fmts.chrome("total_label"))
        s.merge(row, 8, LAST_COL, None,
                fmts.chrome("total_value",
                            num_format=cfg.money_format))
        ws.write_formula(row, 8, s.reg.render(formula),
                         fmts.chrome("total_value",
                                     num_format=cfg.money_format))
        row += 1
    ws.conditional_format(row - 1, 8, row - 1, 8,
                          {"type": "cell", "criteria": ">",
                           "value": 0.004, "format": fmts.cf("bad")})
    row += 1

    s.merge(row, 0, LAST_COL, None, fmts.chrome("note_box"))
    ws.write_formula(
        row, 0,
        s.reg.render('="Payment terms "&PaymentTermsDays&" days. '
                     'Please quote "&IF(PrintInvoice="","the invoice '
                     'number",PrintInvoice)&" with your payment. '
                     'Questions about this bill? Call "&CentreName&'
                     '" and ask for "&OwnerName&"."'),
        fmts.chrome("note_box"))
    s.row_height(row, 26)
    print_last = row
    row += 2

    row = s.note(
        row,
        "Billed and collected are deliberately two different "
        "figures. Billed is revenue - what you charged, less credits "
        "and write-offs. Collected is cash - what actually arrived, "
        "less refunds. A month can be good on one and poor on the "
        "other, and knowing which is the whole point of keeping them "
        "apart.",
        LAST_COL, height=38)

    s.finish(freeze_row=6, gridlines=False, fit_width=True,
             print_area=f"A1:{col_letter(LAST_COL)}{print_last + 1}")
    s.protect()
    return s
````

### `daycare_business_spreadsheet/sheets/annual.py`

*469 lines · 3 functions*

> `Annual Reports` - twelve months side by side, and what they add up to.
>
> Pick a year and read across: what you billed and collected each
> month, what the staff cost, what everything else cost, what was left,
> how full the rooms were and how many children were on the books.
>
> Underneath sits the part that actually changes decisions - the fee
> mix for the year, the expense categories ranked by how much they took,
> the children and families who contribute most, and a straight
> comparison with the year before.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |
| `_rng(first, last, letter)` |  |
| `_rank(sheet, rank_key, field, n)` |  |

**Module constants:** `NAME` ('Annual Reports'), `LAST_COL` (12), `TOP_ROWS` (10), `CAT_ROWS` (15).

````python
"""`Annual Reports` - twelve months side by side, and what they add up to.

Pick a year and read across: what you billed and collected each
month, what the staff cost, what everything else cost, what was left,
how full the rooms were and how many children were on the books.

Underneath sits the part that actually changes decisions - the fee
mix for the year, the expense categories ranked by how much they took,
the children and families who contribute most, and a straight
comparison with the year before.
"""

from __future__ import annotations

from ..config import MONTH_NAMES, Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, number_check

NAME = "Annual Reports"
LAST_COL = 12
TOP_ROWS = 10
CAT_ROWS = 15


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.deep)
    ws = s.ws
    sheet_ref = quote_sheet(NAME)

    s.banner("Annual Report",
             "Twelve months of the setting in one table, with the "
             "totals, the averages and the comparison with last "
             "year.", LAST_COL)
    s.set_col(0, 0, 22)
    for column in range(1, LAST_COL + 1):
        s.set_col(column, column, 14)

    # -------------------------------------------------------- the picker --
    row = 4
    s.merge(row, 0, 0, "Year", fmts.chrome("label_bold"))
    s.write(row, 1, None, fmts.input("int"))
    ws.write_formula(row, 1, "=YEAR(TODAY())", fmts.input("int"))
    ws.data_validation(row, 1, row, 1,
                       number_check(2000, 2100, decimal=False))
    book.define_name("ReportYear", f"={sheet_ref}!$B${row + 1}")
    s.merge(row, 2, 3, "Compared with", fmts.chrome("label_right"))
    s.write(row, 4, None, fmts.calc("int"))
    ws.write_formula(row, 4, "=ReportYear-1", fmts.calc("int"))
    book.define_name("ReportYearPrev", f"={sheet_ref}!$E${row + 1}")
    s.merge(row, 5, 6, "Weeks open a year", fmts.chrome("label_right"))
    s.write(row, 7, None, fmts.calc("int"))
    ws.write_formula(row, 7, "=WeeksOpen", fmts.calc("int"))
    s.merge(row, 8, 9, "Printed", fmts.chrome("label_right"))
    s.merge(row, 10, LAST_COL, None, fmts.calc("date"))
    ws.write_formula(row, 10, "=TodayDate", fmts.calc("date"))
    row += 2

    row = s.note(
        row,
        "Type a year in the blue cell. Every figure on this page is "
        "summed from the month keys that begin with it, so a month "
        "you have not logged yet simply reads zero rather than "
        "breaking the total.",
        LAST_COL, height=28)
    row += 1

    def mk(index: int) -> str:
        """The yyyy-mm key for month `index` (1-12) of the chosen year."""
        return f'ReportYear&"-"&TEXT({index},"00")'

    def year_sum(sheet: str, value: str, key: str,
                 year: str = "ReportYear") -> str:
        """Sum one register column across the twelve months of a year.

        The result is wrapped in brackets: these totals get divided by
        a target or a headcount, and Excel divides before it adds.
        """
        return "(" + "+".join(
            f'SUMIFS([[{sheet}.{value}]],[[{sheet}.{key}]],'
            f'{year}&"-"&TEXT({m},"00"))' for m in range(1, 13)) + ")"

    billed_year = year_sum("Invoices & Receipts", "billed", "month")
    billed_prev = year_sum("Invoices & Receipts", "billed", "month",
                           "ReportYearPrev")
    collected_year = year_sum("Invoices & Receipts", "collected",
                              "month")
    wages_year = year_sum("Payroll Tracker", "total_cost", "month")
    wages_prev = year_sum("Payroll Tracker", "total_cost", "month",
                          "ReportYearPrev")
    costs_year = year_sum("Expense Tracker", "pl_amount", "month")
    costs_prev = year_sum("Expense Tracker", "pl_amount", "month",
                          "ReportYearPrev")
    sessions_year = year_sum("Attendance", "session", "month")
    present_year = year_sum("Attendance", "present", "month")

    # ------------------------------------------------------ the headlines -
    row = s.section(row, "The year in eight numbers", LAST_COL)
    tiles = [
        ("Billed", f"=ROUND({billed_year},2)", "money0",
         f"target {cfg.currency_symbol}"
         f"{cfg.revenue_target_year:,.0f}"),
        ("Against target",
         f'=IF(RevenueTargetYear<=0,"",{billed_year}/'
         f'RevenueTargetYear)', "percent0", "100% makes the year"),
        ("Collected", f"=ROUND({collected_year},2)", "money0",
         "cash that arrived"),
        ("Collection rate",
         f'=IF({billed_year}<=0,"",{collected_year}/{billed_year})',
         "percent", f"target {cfg.target_collection:.0%}"),
        ("Wages", f"=ROUND({wages_year},2)", "money0",
         "including employer costs"),
        ("Other costs", f"=ROUND({costs_year},2)", "money0",
         "everything that is not wages"),
        ("Operating profit",
         f"=ROUND({billed_year}-{wages_year}-{costs_year},2)",
         "money0", "what the year left you"),
        ("Margin",
         f'=IF({billed_year}<=0,"",({billed_year}-{wages_year}-'
         f'{costs_year})/{billed_year})', "percent",
         "of everything billed"),
    ]
    column = 0
    bottom = row
    for index, (label, formula, style, note) in enumerate(tiles):
        if index and index % 4 == 0:
            row = bottom
            column = 0
        bottom = s.kpi(row, column, label, formula, style, span=3,
                       note=note)
        column += 3
    row = bottom + 1

    # -------------------------------------------------- month by month ---
    row = s.section(row, "Month by month", LAST_COL)
    row = s.table_head(row, 0, [
        "Month", "Billed", "Collected", "Still to collect", "Wages",
        "Other costs", "Operating profit", "Margin",
        "Children billed", "Sessions", "Attendance", "Incidents",
        "New starters"],
        [12, 13, 13, 14, 12, 12, 15, 10, 13, 10, 11, 10, 12])
    grid_first = row
    for index in range(1, 13):
        key = mk(index)
        billed = (f'SUMIFS([[Invoices & Receipts.billed]],'
                  f'[[Invoices & Receipts.month]],{key})')
        collected = (f'SUMIFS([[Invoices & Receipts.collected]],'
                     f'[[Invoices & Receipts.month]],{key})')
        wages = (f'SUMIFS([[Payroll Tracker.total_cost]],'
                 f'[[Payroll Tracker.month]],{key})')
        costs = (f'SUMIFS([[Expense Tracker.pl_amount]],'
                 f'[[Expense Tracker.month]],{key})')
        sessions = (f'SUMIFS([[Attendance.session]],'
                    f'[[Attendance.month]],{key})')
        present = (f'SUMIFS([[Attendance.present]],'
                   f'[[Attendance.month]],{key})')
        s.write(row, 0, MONTH_NAMES[index - 1], fmts.chrome("label_bold"))
        formulas = [
            (1, f"=ROUND({billed},2)", "money"),
            (2, f"=ROUND({collected},2)", "money"),
            (3, f"=ROUND({billed}-{collected},2)", "money"),
            (4, f"=ROUND({wages},2)", "money"),
            (5, f"=ROUND({costs},2)", "money"),
            (6, f"=ROUND({billed}-{wages}-{costs},2)", "money"),
            (7, f'=IF({billed}<=0,"",({billed}-{wages}-{costs})/'
                f'{billed})', "percent"),
            (8, f'=COUNTIFS([[Invoices & Receipts.month]],{key},'
                f'[[Invoices & Receipts.type]],"Invoice")', "int"),
            (9, f"={sessions}", "int"),
            (10, f'=IF({sessions}<=0,"",{present}/{sessions})',
             "percent"),
            (11, f'=COUNTIFS([[Incident Reports.month]],{key})', "int"),
            (12, f'=COUNTIFS([[Child Enrollment.start_date]],">="&'
                 f'DATE(ReportYear,{index},1),'
                 f'[[Child Enrollment.start_date]],"<"&'
                 f'DATE(ReportYear,{index}+1,1))', "int"),
        ]
        for column, formula, style in formulas:
            ws.write_formula(row, column, s.reg.render(formula),
                             fmts.calc(style))
        row += 1
    grid_last = row - 1

    # the totals line
    s.write(row, 0, "Year", fmts.chrome("total_label"))
    totals = [
        (1, f"=ROUND({billed_year},2)", "money"),
        (2, f"=ROUND({collected_year},2)", "money"),
        (3, f"=ROUND({billed_year}-{collected_year},2)", "money"),
        (4, f"=ROUND({wages_year},2)", "money"),
        (5, f"=ROUND({costs_year},2)", "money"),
        (6, f"=ROUND({billed_year}-{wages_year}-{costs_year},2)",
         "money"),
        (7, f'=IF({billed_year}<=0,"",({billed_year}-{wages_year}-'
            f'{costs_year})/{billed_year})', "percent"),
        (8, f'=SUM({_rng(grid_first, grid_last, "I")})', "int"),
        (9, f"={sessions_year}", "int"),
        (10, f'=IF({sessions_year}<=0,"",{present_year}/'
             f'{sessions_year})', "percent"),
        (11, f'=SUM({_rng(grid_first, grid_last, "L")})', "int"),
        (12, f'=SUM({_rng(grid_first, grid_last, "M")})', "int"),
    ]
    for column, formula, style in totals:
        ws.write_formula(
            row, column, s.reg.render(formula),
            fmts.chrome("total_value",
                        num_format=s.num_format(style)))
    row += 1
    ws.conditional_format(grid_first, 1, grid_last, 1,
                          {"type": "data_bar",
                           "bar_color": Palette.primary_mid})
    ws.conditional_format(grid_first, 6, grid_last, 6,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(grid_first, 10, grid_last, 10,
                          {"type": "cell", "criteria": "<",
                           "value": "=TargetAttendance",
                           "format": fmts.cf("warn")})

    ws.add_sparkline(row, 1, {
        "range": f"{sheet_ref}!$B${grid_first + 1}:$B${grid_last + 1}",
        "type": "column", "series_color": Palette.primary_mid})
    s.merge(row, 2, 3, "Billed by month", fmts.chrome("note"))
    ws.add_sparkline(row, 4, {
        "range": f"{sheet_ref}!$G${grid_first + 1}:$G${grid_last + 1}",
        "type": "column", "series_color": Palette.accent,
        "negative_points": True})
    s.merge(row, 5, 6, "Operating profit by month",
            fmts.chrome("note"))
    ws.add_sparkline(row, 7, {
        "range": f"{sheet_ref}!$K${grid_first + 1}:$K${grid_last + 1}",
        "type": "line", "series_color": Palette.deep, "markers": True})
    s.merge(row, 8, LAST_COL, "Attendance by month",
            fmts.chrome("note"))
    row += 2

    # ------------------------------------------------- this year v last --
    row = s.section(row, "This year against last", LAST_COL)
    row = s.table_head(row, 0, [
        "Figure", "This year", "Last year", "Change", "Change %",
        "What it tells you"],
        [22, 15, 15, 14, 12, 46])
    compare = [
        ("Billed", billed_year, billed_prev,
         "Revenue. Growth here without growth in places means your "
         "fees went up."),
        ("Wages & on-costs", wages_year, wages_prev,
         "The biggest line in any setting. Watch it as a share of "
         "billing, not on its own."),
        ("Other costs", costs_year, costs_prev,
         "Rent, food, nappies, insurance and everything else."),
        ("Operating profit",
         f"({billed_year})-({wages_year})-({costs_year})",
         f"({billed_prev})-({wages_prev})-({costs_prev})",
         "Billed less wages less costs. The number that decides "
         "whether next year happens."),
    ]
    for label, this_year, last_year, meaning in compare:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        ws.write_formula(row, 1, s.reg.render(f"=ROUND({this_year},2)"),
                         fmts.calc("money"))
        ws.write_formula(row, 2, s.reg.render(f"=ROUND({last_year},2)"),
                         fmts.calc("money"))
        ws.write_formula(
            row, 3,
            s.reg.render(f"=ROUND(({this_year})-({last_year}),2)"),
            fmts.calc("money"))
        ws.write_formula(
            row, 4,
            s.reg.render(f'=IF(({last_year})<=0,"",({this_year})/'
                         f'({last_year})-1)'),
            fmts.calc("percent"))
        s.merge(row, 5, LAST_COL, meaning, fmts.chrome("note"))
        s.row_height(row, 28)
        row += 1
    row += 1

    # ------------------------------------------------------- the fee mix --
    row = s.section(row, "What the year was billed for", LAST_COL)
    row = s.table_head(row, 0, [
        "Tuition", "Registration & deposits", "Late pickup",
        "Meals & snacks", "Transport", "Trips & activities",
        "Other charges", "Discounts", "Subsidies", "Net charges",
        "Average a child"],
        [14, 18, 12, 14, 12, 15, 13, 12, 12, 13, 14])

    def fee_year(*kinds: str) -> str:
        parts = []
        for kind in kinds:
            parts.append("+".join(
                f'SUMIFS([[Tuition & Fees.net]],'
                f'[[Tuition & Fees.month]],ReportYear&"-"&'
                f'TEXT({m},"00"),[[Tuition & Fees.fee_type]],'
                f'"{kind}")' for m in range(1, 13)))
        return "+".join(parts)

    net_charges = year_sum("Tuition & Fees", "net", "month")
    s.calc_row(row, 0, [
        "=ROUND(" + fee_year("Tuition") + ",2)",
        "=ROUND(" + fee_year("Registration fee",
                             "Enrollment deposit") + ",2)",
        "=ROUND(" + fee_year("Late pickup fee") + ",2)",
        "=ROUND(" + fee_year("Meal fee", "Snack fee") + ",2)",
        "=ROUND(" + fee_year("Transport fee") + ",2)",
        "=ROUND(" + fee_year("Field trip", "Activity fee") + ",2)",
        "=ROUND(" + fee_year("Extra day", "Extra hours",
                             "Holiday club", "Supply fee",
                             "Nappy & wipes fee", "Late payment fee",
                             "Other charge") + ",2)",
        "=ROUND(" + fee_year("Sibling discount", "Staff discount",
                             "Promotional discount",
                             "Goodwill credit") + ",2)",
        "=ROUND(" + fee_year("Subsidy credit",
                             "Scholarship credit") + ",2)",
        f"=ROUND({net_charges},2)",
        f'=IF(ChildrenEnrolled<=0,"",ROUND({net_charges}/'
        f'ChildrenEnrolled,2))',
    ], ["money"] * 11)
    row += 2

    # --------------------------------------------- costs by category -----
    row = s.section(row, "Where the year's money went", LAST_COL)
    row = s.table_head(row, 0, [
        "Category", "Spent this year", "Budget a year",
        "Over or under", "Share of costs", "A month",
        "Category", "Spent this year", "Budget a year",
        "Over or under", "Share of costs", "A month"],
        [20, 14, 14, 13, 12, 11, 20, 14, 14, 13, 12, 11])
    cat_first = row
    for n in range(1, CAT_ROWS + 1):
        for half in (0, 1):
            index = n + half * CAT_ROWS
            base = half * 6
            cat = f'INDEX(BudgetCategory,{index})'
            spent = "+".join(
                f'SUMIFS([[Expense Tracker.pl_amount]],'
                f'[[Expense Tracker.month]],ReportYear&"-"&'
                f'TEXT({m},"00"),[[Expense Tracker.category]],{cat})'
                for m in range(1, 13))
            budget = f"INDEX(BudgetAmount,{index})*12"
            cells = [
                (base + 0, f'IFERROR({cat}&"","")', "text"),
                (base + 1, f'IFERROR(ROUND({spent},2),"")', "money"),
                (base + 2, f'IFERROR(ROUND({budget},2),"")', "money"),
                (base + 3, f'IFERROR(ROUND(({spent})-({budget}),2),"")',
                 "money"),
                (base + 4, f'IFERROR(IF({costs_year}<=0,"",({spent})/'
                           f'({costs_year})),"")', "percent"),
                (base + 5, f'IFERROR(ROUND(({spent})/12,2),"")',
                 "money"),
            ]
            for column, formula, style in cells:
                ws.write_formula(row, column,
                                 s.reg.render("=" + formula),
                                 fmts.calc(style))
        row += 1
    for column in (3, 9):
        ws.conditional_format(cat_first, column, row - 1, column,
                              {"type": "cell", "criteria": ">",
                               "value": 0.004,
                               "format": fmts.cf("bad")})
        ws.conditional_format(cat_first, column, row - 1, column,
                              {"type": "cell", "criteria": "<=",
                               "value": 0, "format": fmts.cf("good")})
    row += 1

    # ------------------------------------------------- who pays the bills -
    row = s.section(row, "Who pays the bills", LAST_COL)
    row = s.table_head(row, 0, [
        "Family", "Children", "Weekly fees", "A year", "Balance",
        "Child", "Room", "Weekly fee", "A year", "Attendance",
        "With us, months"],
        [22, 9, 13, 13, 12, 20, 14, 12, 13, 12, 14])
    for n in range(1, TOP_ROWS + 1):
        cells = [
            (0, _rank("Parent Database", "weekly_fees_rank",
                      "family_name", n), "text"),
            (1, _rank("Parent Database", "weekly_fees_rank",
                      "children", n), "int"),
            (2, _rank("Parent Database", "weekly_fees_rank",
                      "weekly_fees", n), "money"),
            (3, 'IFERROR(INDEX([[Parent Database.weekly_fees]],'
                f'MATCH({n},[[Parent Database.weekly_fees_rank]],0))'
                '*WeeksOpen,"")', "money"),
            (4, _rank("Parent Database", "weekly_fees_rank",
                      "balance", n), "money"),
            (5, _rank("Child Enrollment", "annual_net_rank", "name", n),
             "text"),
            (6, _rank("Child Enrollment", "annual_net_rank", "room", n),
             "text"),
            (7, _rank("Child Enrollment", "annual_net_rank",
                      "weekly_net", n), "money"),
            (8, _rank("Child Enrollment", "annual_net_rank",
                      "annual_net", n), "money"),
            (9, _rank("Child Enrollment", "annual_net_rank",
                      "attendance_pct", n), "percent"),
            (10, _rank("Child Enrollment", "annual_net_rank",
                       "months_with_us", n), "int"),
        ]
        for column, formula, style in cells:
            ws.write_formula(row, column, s.reg.render("=" + formula),
                             fmts.calc(style))
        row += 1
    row += 1

    # ------------------------------------------------ the year in people --
    row = s.section(row, "The year in people", LAST_COL)
    row = s.table_head(row, 0, [
        "Children started", "Children left", "Net change",
        "On the books now", "Enquiries", "Converted",
        "Conversion rate", "Staff on the books", "Staff hours",
        "Incidents", "Doses given", "Parent conversations"],
        [15, 13, 12, 15, 12, 12, 14, 15, 12, 11, 12, 17])
    starts = 'COUNTIFS([[Child Enrollment.start_date]],">="&DATE(ReportYear,1,1),[[Child Enrollment.start_date]],"<="&DATE(ReportYear+1,1,1)-1)'
    leaves = 'COUNTIFS([[Child Enrollment.end_date]],">="&DATE(ReportYear,1,1),[[Child Enrollment.end_date]],"<="&DATE(ReportYear+1,1,1)-1)'
    enquiries = 'COUNTIFS([[Waitlist.enquiry_date]],">="&DATE(ReportYear,1,1),[[Waitlist.enquiry_date]],"<="&DATE(ReportYear+1,1,1)-1)'
    converted = ('COUNTIFS([[Waitlist.enquiry_date]],">="&DATE(ReportYear,1,1),'
                 '[[Waitlist.enquiry_date]],"<="&DATE(ReportYear+1,1,1)-1,'
                 '[[Waitlist.converted]],"Yes")')
    s.calc_row(row, 0, [
        f"={starts}",
        f"={leaves}",
        f"={starts}-{leaves}",
        "=ChildrenEnrolled",
        f"={enquiries}",
        f"={converted}",
        f'=IF({enquiries}<=0,"",{converted}/{enquiries})',
        "=StaffCount",
        "=ROUND(" + year_sum("Payroll Tracker", "total_hours",
                             "month") + ",1)",
        "=" + "+".join(
            f'COUNTIFS([[Incident Reports.month]],ReportYear&"-"&'
            f'TEXT({m},"00"))' for m in range(1, 13)),
        '=COUNTIFS([[Allergy & Medical.given_date]],">="&DATE(ReportYear,1,1),[[Allergy & Medical.given_date]],"<="&DATE(ReportYear+1,1,1)-1)',
        "=" + "+".join(
            f'COUNTIFS([[Parent Communication.month]],ReportYear&"-"&'
            f'TEXT({m},"00"))' for m in range(1, 13)),
    ], ["int", "int", "int", "int", "int", "int", "percent", "int",
        "number1", "int", "int", "int"])
    row += 2

    row = s.note(
        row,
        "Everything on this page is summed from the month keys, so "
        "the year total and the twelve monthly figures can never "
        "disagree. A month with nothing logged reads zero rather "
        "than an error, which means a part-finished year still adds "
        "up.",
        LAST_COL, height=32)
    row += 1
    row = s.note(
        row,
        "Occupancy, fee income and the per-child figures use the "
        "children on the books today, not an average across the "
        "year. If your numbers moved a lot, read the month-by-month "
        "table rather than the year total.",
        LAST_COL, height=30)

    s.finish(freeze_row=6, gridlines=False, fit_width=True)
    s.protect()
    return s


def _rng(first: int, last: int, letter: str) -> str:
    return f"{quote_sheet(NAME)}!${letter}${first + 1}:${letter}${last + 1}"


def _rank(sheet: str, rank_key: str, field: str, n: int) -> str:
    return (f'IFERROR(INDEX([[{sheet}.{field}]],MATCH({n},'
            f'[[{sheet}.{rank_key}]],0)),"")')
````

---

## 10. The quality tools

Not shipped with the workbook, but the reason it can be trusted: a static validator, a bracket and placeholder checker, a demo data checker that rebuilds the arithmetic in Python, and a recalculation harness that really evaluates every formula.

### `tools/validate_daycare_workbook.py`

*892 lines · 15 functions*

> Static checks on a generated Daycare Business workbook.
>
> Run it with any Python that has openpyxl installed::
>
>     python tools/validate_daycare_workbook.py output/daycare.xlsx
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
| `_name_target(wb, name)` | The one cell a defined name points at, or None |
| `check_pages(wb, problems, notes)` | The summary pages, checked through their defined names |
| `_specs()` | Every register specification, keyed by sheet name |
| `check_chain(wb, problems, notes)` | The enquiry-to-enrolment-to-invoice chain only holds while these do |
| `check_ratio_engine(wb, problems, notes)` | The ratio table a buyer edits to match their own licensing |
| `check(path)` |  |

**Module constants:** `FUNCTION_RE` (re.compile(...)), `NAME_RE` (re.compile(...)), `BARE_COL_RE` (re.compile(...)), `KNOWN_FUNCTIONS` ({'IF', 'IFERROR', 'SUM', 'SUMIF', 'SUMIFS', 'COUNT', 'COU...), `MODERN_ONLY` ({'XLOOKUP', 'IFS', 'MAXIFS', 'MINIFS', 'TEXTJOIN', 'CONCA...), `PRECEDENCE_OK` (0 entries), `RANGE_RE` (re.compile(...)), `SEEDED_INPUTS` ({'Monthly Reports!B5', 'Annual Reports!B5'}), `CHAIN_CHECKS` (99 items), `PAGE_FORMULAS` (20 items), `PAGE_INPUTS` (25 items), `REQUIRED_LIST_NAMES` (22 items).

````python
"""Static checks on a generated Daycare Business workbook.

Run it with any Python that has openpyxl installed::

    python tools/validate_daycare_workbook.py output/daycare.xlsx

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
from daycare_business_spreadsheet.build import SHEET_ORDER as EXPECTED_SHEETS  # noqa: E402,E501

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


#: input cells that ship with a formula in them as a helpful default -
#: the buyer is meant to be able to type a fixed value straight over it
SEEDED_INPUTS = {
    "Monthly Reports!B5",       # the month the report opens on
    "Annual Reports!B5",        # the year the report opens on
}

#: the automation chain, link by link: (sheet, column key, what the
#: formula in that column has to still be pointing at)
CHAIN_CHECKS = [
    # ---- enrolment: who is on the books and what they are worth ---------
    ("Child Enrollment", "age_months", ("text", "DATEDIF")),
    ("Child Enrollment", "age_band", ("text", "AgeBandList")),
    ("Child Enrollment", "family_seq", ("column", "Child Enrollment",
                                        "family")),
    ("Child Enrollment", "auto_discount", ("text", "SiblingDiscount")),
    ("Child Enrollment", "gross_week", ("column", "Child Enrollment",
                                        "rate")),
    ("Child Enrollment", "weekly_net", ("column", "Child Enrollment",
                                        "gross_week")),
    ("Child Enrollment", "monthly_net", ("column", "Child Enrollment",
                                         "weekly_net")),
    ("Child Enrollment", "billed", ("sheet", "Tuition & Fees")),
    ("Child Enrollment", "balance", ("sheet", "Parent Database")),
    ("Child Enrollment", "days_in", ("sheet", "Attendance")),
    ("Child Enrollment", "attendance_pct", ("column", "Child Enrollment",
                                            "sessions")),
    ("Child Enrollment", "docs_ok", ("sheet", "Document Tracker")),
    ("Child Enrollment", "bday_in", ("column", "Child Enrollment",
                                     "next_bday")),
    # ---- families: one bill, several children ---------------------------
    ("Parent Database", "children", ("sheet", "Child Enrollment")),
    ("Parent Database", "weekly_fees", ("sheet", "Child Enrollment")),
    ("Parent Database", "invoiced", ("sheet", "Invoices & Receipts")),
    ("Parent Database", "paid", ("sheet", "Invoices & Receipts")),
    ("Parent Database", "balance", ("column", "Parent Database", "paid")),
    ("Parent Database", "overdue", ("text", "Overdue")),
    ("Parent Database", "messages", ("sheet", "Parent Communication")),
    # ---- the waiting list ------------------------------------------------
    ("Waitlist", "age_at_start", ("column", "Waitlist", "start_wanted")),
    ("Waitlist", "seats_now", ("sheet", "Classroom Management")),
    ("Waitlist", "converted", ("text", "Enrolled")),
    # ---- the daily register ---------------------------------------------
    ("Attendance", "present", ("text", "Present")),
    ("Attendance", "hours", ("column", "Attendance", "in_time")),
    ("Attendance", "due_in", ("sheet", "Child Enrollment")),
    ("Attendance", "late_in", ("column", "Attendance", "due_in")),
    ("Attendance", "late_fee", ("text", "LatePickupFee")),
    ("Attendance", "auth_list", ("sheet", "Child Enrollment")),
    ("Attendance", "pickup_ok", ("column", "Attendance", "auth_list")),
    ("Attendance", "daykey", ("column", "Attendance", "child")),
    # ---- what each child is charged -------------------------------------
    ("Tuition & Fees", "gross", ("column", "Tuition & Fees", "rate")),
    ("Tuition & Fees", "discount", ("column", "Tuition & Fees",
                                    "discount_pct")),
    ("Tuition & Fees", "net", ("column", "Tuition & Fees", "is_credit")),
    ("Tuition & Fees", "billed", ("column", "Tuition & Fees", "invoice")),
    ("Tuition & Fees", "key", ("column", "Tuition & Fees", "seq")),
    # ---- the family ledger ----------------------------------------------
    ("Invoices & Receipts", "charges", ("sheet", "Tuition & Fees")),
    ("Invoices & Receipts", "total", ("column", "Invoices & Receipts",
                                      "charges")),
    ("Invoices & Receipts", "paid", ("text", "Payment")),
    ("Invoices & Receipts", "balance", ("column", "Invoices & Receipts",
                                        "paid")),
    ("Invoices & Receipts", "status", ("column", "Invoices & Receipts",
                                       "balance")),
    ("Invoices & Receipts", "days_overdue", ("column",
                                             "Invoices & Receipts",
                                             "due_date")),
    ("Invoices & Receipts", "billed", ("text", "Credit note")),
    ("Invoices & Receipts", "collected", ("text", "Refund")),
    # ---- what the parent reads at the end of the day --------------------
    ("Daily Reports", "hours", ("sheet", "Attendance")),
    ("Daily Reports", "meals_logged", ("sheet", "Meal Tracker")),
    ("Daily Reports", "incidents", ("sheet", "Incident Reports")),
    ("Daily Reports", "nap_hours", ("column", "Daily Reports", "nap_from")),
    # ---- food, allergies, accidents -------------------------------------
    ("Meal Tracker", "diet", ("sheet", "Child Enrollment")),
    ("Meal Tracker", "allergy_level", ("sheet", "Child Enrollment")),
    ("Meal Tracker", "refused", ("text", "Refused")),
    ("Allergy & Medical", "days_left", ("text", "TodayDate")),
    ("Allergy & Medical", "doses_logged", ("column", "Allergy & Medical",
                                           "detail")),
    ("Incident Reports", "age_at", ("sheet", "Child Enrollment")),
    ("Incident Reports", "notify_mins", ("column", "Incident Reports",
                                         "notified_time")),
    ("Incident Reports", "complete", ("column", "Incident Reports",
                                      "notified")),
    # ---- the team --------------------------------------------------------
    ("Staff Database", "next_expiry", ("column", "Staff Database",
                                       "first_aid_exp")),
    ("Staff Database", "days_to_expiry", ("column", "Staff Database",
                                          "next_expiry")),
    ("Staff Database", "certs_ok", ("text", "DocUrgentDays")),
    ("Staff Database", "hours_week", ("sheet", "Staff Schedule")),
    ("Staff Database", "pay_month", ("sheet", "Payroll Tracker")),
    ("Staff Database", "children_room", ("sheet", "Child Enrollment")),
    ("Staff Schedule", "hours", ("text", "BreakMinutes")),
    ("Staff Schedule", "actual_hours", ("column", "Staff Schedule",
                                        "clock_in")),
    ("Staff Schedule", "week_to_date", ("column", "Staff Schedule",
                                        "actual_hours")),
    ("Staff Schedule", "ot_hours", ("text", "OvertimeAfter")),
    ("Staff Schedule", "cost", ("text", "OvertimeMultiplier")),
    ("Staff Schedule", "clash", ("column", "Staff Schedule", "finish")),
    ("Payroll Tracker", "basic_hours", ("sheet", "Staff Schedule")),
    ("Payroll Tracker", "ot_hours", ("sheet", "Staff Schedule")),
    ("Payroll Tracker", "gross", ("column", "Payroll Tracker", "ot_pay")),
    ("Payroll Tracker", "employer_cost", ("text", "PayrollBurden")),
    ("Payroll Tracker", "total_cost", ("column", "Payroll Tracker",
                                       "employer_cost")),
    ("Payroll Tracker", "net_pay", ("column", "Payroll Tracker",
                                    "deductions")),
    # ---- rooms, ratios and capacity -------------------------------------
    ("Classroom Management", "enrolled", ("sheet", "Child Enrollment")),
    ("Classroom Management", "seats", ("column", "Classroom Management",
                                       "capacity")),
    ("Classroom Management", "occupancy", ("column",
                                           "Classroom Management",
                                           "enrolled")),
    ("Classroom Management", "waiting", ("sheet", "Waitlist")),
    ("Classroom Management", "present_today", ("sheet", "Attendance")),
    ("Classroom Management", "staff_today", ("sheet", "Staff Schedule")),
    ("Classroom Management", "staff_needed", ("column",
                                              "Classroom Management",
                                              "ratio")),
    ("Classroom Management", "ratio_ok", ("column",
                                          "Classroom Management",
                                          "staff_needed")),
    ("Classroom Management", "revenue", ("sheet", "Child Enrollment")),
    ("Classroom Management", "empty_cost", ("column",
                                            "Classroom Management",
                                            "revenue")),
    # ---- costs, stock, paperwork ----------------------------------------
    ("Expense Tracker", "is_wages", ("text", "Staff wages")),
    ("Expense Tracker", "in_pl", ("text", "PayrollInExpenses")),
    ("Expense Tracker", "pl_amount", ("column", "Expense Tracker",
                                      "in_pl")),
    ("Expense Tracker", "monthly", ("text", "Quarterly")),
    ("Expense Tracker", "budget", ("text", "BudgetAmount")),
    ("Expense Tracker", "child_cost", ("text", "ChildrenEnrolled")),
    ("Supply Inventory", "stock", ("column", "Supply Inventory",
                                   "received")),
    ("Supply Inventory", "weeks_cover", ("column", "Supply Inventory",
                                         "weekly_use")),
    ("Supply Inventory", "order_qty", ("text", "ReorderWeeks")),
    ("Supply Inventory", "status", ("text", "StockWarningWeeks")),
    ("Parent Communication", "days_open", ("column",
                                           "Parent Communication",
                                           "closed_date")),
    ("Parent Communication", "overdue", ("column",
                                         "Parent Communication",
                                         "follow_date")),
    ("Document Tracker", "who_name", ("sheet", "Staff Database")),
    ("Document Tracker", "days_left", ("text", "TodayDate")),
    ("Document Tracker", "state", ("column", "Document Tracker",
                                   "days_left")),
]

#: cells reached by a defined name that have to hold a live formula -
#: the summary pages are where the value of the file lives, and a name
#: pointing at an empty cell means a page quietly stopped working.
PAGE_FORMULAS = [
    ("TodayDate", "settings: the live date every alert reads"),
    ("ThisMonthKey", "settings: the month key the totals group on"),
    ("TrackingWeeks", "settings: weeks of history, for the averages"),
    ("OpenHoursDay", "settings: hours the doors are open each day"),
    ("BudgetMonthlyTotal", "settings: the monthly running cost budget"),
    ("ChildrenEnrolled", "settings: children on the books"),
    ("ChildrenActive", "settings: children actually attending"),
    ("StaffCount", "settings: how many staff are on the payroll"),
    ("RoomCapacity", "settings: places across all the rooms"),
    ("RevenueAll", "settings: everything billed, all time"),
    ("CollectedAll", "settings: everything collected, all time"),
    ("DashKey", "dashboard: the month the tiles report on"),
    ("DashPrevKey", "dashboard: last month, for the comparisons"),
    ("MonthBilled", "dashboard: billed this month"),
    ("MonthCollected", "dashboard: collected this month"),
    ("MonthWages", "dashboard: the pay runs this month"),
    ("MonthExpenses", "dashboard: the bills this month"),
    ("MonthProfit", "dashboard: what is left after the bills"),
    ("ReportPrev", "monthly reports: last month, for the comparisons"),
    ("ReportYearPrev", "annual reports: last year, for the comparisons"),
]

#: cells a buyer has to be able to type in
PAGE_INPUTS = [
    ("CentreName", "settings: the nursery's name"),
    ("OwnerName", "settings: the owner or director"),
    ("CurrencySign", "settings: the currency symbol"),
    ("TrackingStart", "settings: the day tracking began"),
    ("OpenHour", "settings: opening time"),
    ("CloseHour", "settings: closing time"),
    ("OpenDays", "settings: days open a week"),
    ("WeeksOpen", "settings: trading weeks a year"),
    ("LicensedCapacity", "settings: licensed places"),
    ("RatioWarning", "settings: the occupancy a room warns at"),
    ("FullTimeWeek", "settings: the standard full-time weekly fee"),
    ("SiblingDiscount", "settings: the sibling discount"),
    ("PaymentTermsDays", "settings: how long a family has to pay"),
    ("LatePickupFee", "settings: the late collection charge"),
    ("TaxRate", "settings: the tax rate"),
    ("TaxOnTuition", "settings: does tax apply to tuition"),
    ("PayrollBurden", "settings: employer costs on wages"),
    ("OvertimeAfter", "settings: hours a week before overtime"),
    ("PayrollInExpenses", "settings: are wages typed as expenses too"),
    ("RevenueTargetMonth", "settings: the monthly revenue target"),
    ("TargetOccupancy", "settings: the occupancy you are aiming at"),
    ("DocWarningDays", "settings: how early paperwork warns"),
    ("ReportMonth", "monthly reports: the month being reported on"),
    ("ReportYear", "annual reports: the year being reported on"),
    ("PrintInvoice", "monthly reports: the invoice being printed"),
]

#: every drop down on the workbook reads one of these - a missing one
#: means a blank list and a column nobody can fill in
REQUIRED_LIST_NAMES = [
    "ChildList", "FamilyList", "ClassroomList", "StaffList",
    "InvoiceList", "SupplyList", "EnrollmentStatusList",
    "AttendanceStatusList", "ProgramList", "ScheduleTypeList",
    "BillingBasisList", "FeeTypeList", "LedgerTypeList",
    "PaymentMethodList", "ExpenseCategoryList", "DocumentTypeList",
    "IncidentCategoryList", "AllergySeverityList", "DietList",
    "YesNoList", "MonthKeyList", "AgeBandList",
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
    from daycare_business_spreadsheet.build import REGISTER_MODULES
    from daycare_business_spreadsheet.config import Config
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


def check_ratio_engine(wb, problems: list, notes: list) -> None:
    """The ratio table a buyer edits to match their own licensing."""
    ws = wb["Settings & Lists"]
    bands = 0
    for row in ws.iter_rows():
        for cell in row:
            if cell.value in (None, ""):
                continue
            if isinstance(cell.value, str) and cell.value.startswith("="):
                continue
            if cell.column_letter == "B" and isinstance(cell.value, str):
                pass
        bands += 0
    for name in ("AgeBandList", "AgeBandFrom", "AgeBandTo", "AgeBandRatio",
                 "AgeBandMax"):
        defn = wb.defined_names.get(name)
        if defn is None:
            problems.append(f"the editable ratio table has lost {name}")
            continue
        try:
            sheet, ref = next(defn.destinations)
        except Exception:                        # noqa: B902
            problems.append(f"{name} does not point anywhere")
            continue
        if ":" not in ref:
            problems.append(f"{name} should be a range, got {ref}")
            continue
        bands = max(bands, 1)
    locked = 0
    defn = wb.defined_names.get("AgeBandRatio")
    if defn is not None:
        sheet, ref = next(defn.destinations)
        for row in wb[sheet][ref.replace("$", "")]:
            for cell in row:
                if cell.protection and cell.protection.locked:
                    locked += 1
    if locked:
        problems.append(f"{locked} staff-to-child ratio cells are locked - "
                        f"buyers have to be able to match local licensing")
    notes.append("ratio table: editable, five named ranges present"
                 if bands else "ratio table: NOT FOUND")


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
    check_ratio_engine(wb, problems, notes)

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
              else "output/Daycare Business Management Spreadsheet.xlsx")
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

### `tools/check_daycare_sample.py`

*438 lines · 7 functions*

> Does the demo nursery actually hang together?
>
> The validator checks the formulas; this checks the *story* they are
> told. It reads only the literal cells a user would have typed - never
> a formula column - rebuilds the arithmetic in Python, and complains if
> the example data would make the workbook look wrong on the first open:
> children in rooms that do not exist, attendance for children who had
> left, charges on invoices that were never raised, payments bigger than
> the bill, shifts that overlap, stock that goes negative, IDs that
> point at nothing.
>
>     python tools/check_daycare_sample.py output/daycare.xlsx

| Defines | What it does |
|---------|--------------|
| `_specs(cfg)` |  |
| `_rows(ws, spec, key_column)` | Every filled data row as {column key: literal value} |
| `_month(value)` |  |
| `_day(value)` |  |
| `_minutes(value)` |  |
| `_num(value)` |  |
| `check(path)` |  |

**Module constants:** `MONEY` (0.015), `ANCHORS` (18 entries), `CREDIT_TYPES` (2 items).

````python
#!/usr/bin/env python3
"""Does the demo nursery actually hang together?

The validator checks the formulas; this checks the *story* they are
told. It reads only the literal cells a user would have typed - never
a formula column - rebuilds the arithmetic in Python, and complains if
the example data would make the workbook look wrong on the first open:
children in rooms that do not exist, attendance for children who had
left, charges on invoices that were never raised, payments bigger than
the bill, shifts that overlap, stock that goes negative, IDs that
point at nothing.

    python tools/check_daycare_sample.py output/daycare.xlsx
"""

from __future__ import annotations

import datetime as _dt
import os
import sys
from collections import defaultdict

try:
    import openpyxl
except ImportError:                                   # pragma: no cover
    print("openpyxl is required: pip install openpyxl")
    raise SystemExit(2)

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from daycare_business_spreadsheet.build import REGISTER_MODULES  # noqa: E402
from daycare_business_spreadsheet.config import Config, rid      # noqa: E402

MONEY = 0.015           # a cent and a half of rounding slack

#: the column every register is keyed on - a row with this cell empty
#: is an empty row, whatever else is in it
ANCHORS = {
    "Classroom Management": "name",
    "Parent Database": "parent1",
    "Child Enrollment": "name",
    "Waitlist": "child_name",
    "Attendance": "date",
    "Tuition & Fees": "date",
    "Invoices & Receipts": "date",
    "Daily Reports": "date",
    "Meal Tracker": "date",
    "Allergy & Medical": "child",
    "Incident Reports": "date",
    "Staff Database": "name",
    "Staff Schedule": "date",
    "Payroll Tracker": "period_end",
    "Expense Tracker": "date",
    "Supply Inventory": "name",
    "Parent Communication": "date",
    "Document Tracker": "doc_type",
}

#: ledger rows that reduce what a family owes
CREDIT_TYPES = ("Credit note", "Write-off")


def _specs(cfg: Config):
    return {module.NAME: module.spec(cfg) for module in REGISTER_MODULES}


def _rows(ws, spec, key_column: str):
    """Every filled data row as {column key: literal value}."""
    first = spec.first_data_row                      # 1 based
    last = spec.last_data_row
    key_index = spec.index(key_column) + 1
    out = []
    for excel_row in range(first, last + 1):
        anchor = ws.cell(row=excel_row, column=key_index).value
        if anchor in (None, ""):
            continue
        record = {"_row": excel_row - first + 1}
        for index, col in enumerate(spec.columns):
            value = ws.cell(row=excel_row, column=index + 1).value
            if isinstance(value, str) and value.startswith("="):
                continue                              # never trust a formula
            record[col.key] = value
        out.append(record)
    return out


def _month(value) -> str:
    if isinstance(value, _dt.datetime):
        value = value.date()
    return f"{value.year:04d}-{value.month:02d}" if value else ""


def _day(value):
    return value.date() if isinstance(value, _dt.datetime) else value


def _minutes(value) -> int:
    if isinstance(value, _dt.datetime):
        return value.hour * 60 + value.minute
    if isinstance(value, _dt.time):
        return value.hour * 60 + value.minute
    if isinstance(value, (int, float)):
        return int(round(float(value) * 24 * 60))
    return 0


def _num(value) -> float:
    return float(value) if isinstance(value, (int, float)) else 0.0


def check(path: str) -> int:                          # noqa: C901
    cfg = Config()
    specs = _specs(cfg)
    wb = openpyxl.load_workbook(path, data_only=False)
    problems: list = []
    notes: list = []

    data = {}
    for name, anchor in ANCHORS.items():
        if name not in wb.sheetnames:
            problems.append(f"missing sheet {name}")
            continue
        data[name] = _rows(wb[name], specs[name], anchor)
    if problems:
        for line in problems:
            print("    - " + line)
        return 1

    counts = ", ".join(f"{name} {len(rows)}"
                       for name, rows in data.items() if rows)
    notes.append("demo rows: " + counts)
    if not data["Child Enrollment"]:
        print(f"--- {path}\n  . no example data in this build (blank "
              f"version) - nothing to check")
        return 0

    # ---------------------------------------------------- the ID spine --
    room_by_id = {rid("Classroom", row["_row"]): row
                  for row in data["Classroom Management"]}
    family_by_id = {rid("Family", row["_row"]): row
                    for row in data["Parent Database"]}
    child_by_id = {rid("Child", row["_row"]): row
                   for row in data["Child Enrollment"]}
    staff_by_id = {rid("Staff", row["_row"]): row
                   for row in data["Staff Database"]}
    supply_by_id = {rid("Supply", row["_row"]): row
                    for row in data["Supply Inventory"]}
    ledger_by_id = {rid("Invoice", row["_row"]): row
                    for row in data["Invoices & Receipts"]}

    room_names = {row.get("name") for row in data["Classroom Management"]}

    def known(table, value, where, kind):
        if value in (None, ""):
            return False
        if value not in table:
            problems.append(f"{where}: {kind} '{value}' does not exist")
            return False
        return True

    for row in data["Child Enrollment"]:
        where = f"Child Enrollment row {row['_row']}"
        known(family_by_id, row.get("family"), where, "family")
        if row.get("room") and row["room"] not in room_names:
            problems.append(f"{where}: room '{row['room']}' is not a room")
    for row in data["Attendance"]:
        where = f"Attendance row {row['_row']}"
        known(child_by_id, row.get("child"), where, "child")
    for row in data["Tuition & Fees"]:
        where = f"Tuition & Fees row {row['_row']}"
        known(child_by_id, row.get("child"), where, "child")
        if row.get("invoice"):
            known(ledger_by_id, row["invoice"], where, "invoice")
    for row in data["Invoices & Receipts"]:
        where = f"Invoices & Receipts row {row['_row']}"
        known(family_by_id, row.get("family"), where, "family")
        if row.get("child"):
            known(child_by_id, row["child"], where, "child")
        if row.get("against"):
            known(ledger_by_id, row["against"], where, "invoice")
    for name, key in (("Daily Reports", "child"), ("Meal Tracker", "child"),
                      ("Allergy & Medical", "child"),
                      ("Incident Reports", "child")):
        for row in data[name]:
            known(child_by_id, row.get(key), f"{name} row {row['_row']}",
                  "child")
    for row in data["Staff Schedule"]:
        where = f"Staff Schedule row {row['_row']}"
        known(staff_by_id, row.get("staff"), where, "staff member")
        if row.get("room") and row["room"] not in room_names:
            problems.append(f"{where}: room '{row['room']}' is not a room")
    for row in data["Payroll Tracker"]:
        known(staff_by_id, row.get("staff"),
              f"Payroll Tracker row {row['_row']}", "staff member")
    for row in data["Parent Communication"]:
        where = f"Parent Communication row {row['_row']}"
        known(family_by_id, row.get("family"), where, "family")
        if row.get("child"):
            known(child_by_id, row["child"], where, "child")
    for row in data["Document Tracker"]:
        where = f"Document Tracker row {row['_row']}"
        holder, who = row.get("holder"), row.get("who")
        if holder == "Child" and who:
            known(child_by_id, who, where, "child")
        if holder == "Staff" and who:
            known(staff_by_id, who, where, "staff member")
    notes.append(f"ID spine: {len(child_by_id)} children, "
                 f"{len(family_by_id)} families, {len(staff_by_id)} staff, "
                 f"{len(room_by_id)} rooms, {len(supply_by_id)} supplies")

    # --------------------------------------------- multi-child families --
    per_family = defaultdict(list)
    for key, row in child_by_id.items():
        if row.get("family"):
            per_family[row["family"]].append(key)
    siblings = sum(1 for kids in per_family.values() if len(kids) > 1)
    if siblings < 2:
        problems.append("the demo needs at least two families with more "
                        "than one child - that is the headline feature")
    notes.append(f"families: {siblings} with siblings, "
                 f"{len(per_family)} with a child on the books")

    # ------------------------------------------------ rooms and capacity --
    in_room = defaultdict(int)
    for row in child_by_id.values():
        if row.get("status") in ("Active", "On hold", "Starting soon"):
            in_room[row.get("room")] += 1
    for row in data["Classroom Management"]:
        places = _num(row.get("capacity"))
        taken = in_room.get(row.get("name"), 0)
        if places and taken > places:
            problems.append(
                f"room '{row['name']}' holds {places:.0f} but "
                f"{taken} children are enrolled in it")
    notes.append("rooms: " + ", ".join(
        f"{row['name']} {in_room.get(row['name'], 0)}/"
        f"{_num(row.get('capacity')):.0f}"
        for row in data["Classroom Management"]))

    # -------------------------------------------------------- attendance --
    seen = set()
    for row in data["Attendance"]:
        day, child = _day(row.get("date")), row.get("child")
        if (day, child) in seen:
            problems.append(f"Attendance row {row['_row']}: {child} is "
                            f"marked twice on {day}")
        seen.add((day, child))
        enrolled = child_by_id.get(child, {})
        start = _day(enrolled.get("start_date"))
        end = _day(enrolled.get("end_date"))
        if start and day and day < start:
            problems.append(f"Attendance row {row['_row']}: {child} "
                            f"attended on {day}, before they started")
        if end and day and day > end:
            problems.append(f"Attendance row {row['_row']}: {child} "
                            f"attended on {day}, after they left")
        in_time, out_time = row.get("in_time"), row.get("out_time")
        if in_time and out_time and _minutes(out_time) <= _minutes(in_time):
            problems.append(f"Attendance row {row['_row']}: collected "
                            f"before they arrived")
    present = sum(1 for row in data["Attendance"]
                  if row.get("status") in ("Present", "Late",
                                           "Early pickup"))
    sessions = sum(1 for row in data["Attendance"]
                   if row.get("status") != "Holiday / closed")
    if sessions:
        rate = present / sessions
        if not 0.7 <= rate <= 0.99:
            problems.append(f"attendance runs at {rate:.0%} - between "
                            f"70% and 99% looks like a real nursery")
        notes.append(f"attendance: {len(data['Attendance'])} rows, "
                     f"{rate:.0%} present")
    missing_out = sum(1 for row in data["Attendance"]
                      if row.get("in_time") and not row.get("out_time"))
    if not missing_out:
        problems.append("no child is left without a check-out time - the "
                        "missing check-out alert has nothing to show")
    notes.append(f"check-out alerts: {missing_out} open")

    # ------------------------------------------------- charges and money --
    charge_total = defaultdict(float)
    for row in data["Tuition & Fees"]:
        qty = row.get("qty")
        qty = 1.0 if qty in (None, "") else float(qty)
        gross = round(qty * _num(row.get("rate")), 2)
        net = round(gross * (1 - _num(row.get("discount_pct"))), 2)
        if row.get("invoice"):
            charge_total[row["invoice"]] += net
    invoiced = paid = credited = 0.0
    for row in data["Invoices & Receipts"]:
        kind = row.get("type")
        amount = _num(row.get("amount"))
        key = rid("Invoice", row["_row"])
        if kind == "Invoice":
            lines = charge_total.get(key, 0.0)
            invoiced += lines if lines else amount
            if lines and amount and abs(lines - amount) > MONEY:
                problems.append(
                    f"invoice {key} totals {lines:,.2f} from the charge "
                    f"lines but {amount:,.2f} was typed in")
        elif kind == "Payment":
            paid += amount
        elif kind == "Refund":
            paid -= amount
        elif kind in CREDIT_TYPES:
            credited += amount
        label = f"{row.get('period') or ''} {row.get('description') or ''}"
        deposit = "deposit" in label.lower()
        if (kind in ("Payment", "Refund") and not row.get("against")
                and not deposit):
            # a deposit is paid before there is anything to pay it
            # against - everything else has to name its invoice
            problems.append(f"Invoices & Receipts row {row['_row']}: a "
                            f"{kind.lower()} with no invoice to match")
    owed = invoiced - credited - paid
    if invoiced <= 0:
        problems.append("nothing has been invoiced - the money pages "
                        "will all read zero")
    else:
        collection = paid / invoiced
        if not 0.6 <= collection <= 1.0:
            problems.append(
                f"the demo has collected {collection:.0%} of what it "
                f"billed - between 60% and 100% tells a better story")
        notes.append(f"money: billed {invoiced:,.0f}, collected "
                     f"{paid:,.0f} ({paid / invoiced:.0%}), credited "
                     f"{credited:,.0f}, outstanding {owed:,.0f}")
    over = defaultdict(float)
    for row in data["Invoices & Receipts"]:
        if row.get("type") == "Payment" and row.get("against"):
            over[row["against"]] += _num(row.get("amount"))
    for key, amount in over.items():
        target = charge_total.get(key)
        if target is None:
            invoice = ledger_by_id.get(key, {})
            target = _num(invoice.get("amount"))
        if target and amount - target > MONEY:
            problems.append(f"invoice {key} is overpaid: {amount:,.2f} "
                            f"against {target:,.2f}")
    unbilled = sum(1 for row in data["Tuition & Fees"]
                   if not row.get("invoice"))
    notes.append(f"charges: {len(data['Tuition & Fees'])} lines, "
                 f"{unbilled} not yet on an invoice")

    # ------------------------------------------------------- the rota ----
    rota = defaultdict(list)
    for row in data["Staff Schedule"]:
        if not row.get("start") or not row.get("finish"):
            continue
        start, finish = _minutes(row["start"]), _minutes(row["finish"])
        if finish <= start:
            problems.append(f"Staff Schedule row {row['_row']}: the shift "
                            f"finishes before it starts")
            continue
        rota[(row.get("staff"), _day(row.get("date")))].append(
            (start, finish, row["_row"]))
    clashes = 0
    for (who, day), shifts in rota.items():
        shifts.sort()
        for first, second in zip(shifts, shifts[1:]):
            if second[0] < first[1]:
                clashes += 1
                if clashes < 6:
                    problems.append(f"double shift: {who} on {day}, rows "
                                    f"{first[2]} and {second[2]}")
    hours = sum((f - s) / 60.0 for shifts in rota.values()
                for s, f, _ in shifts)
    notes.append(f"rota: {len(data['Staff Schedule'])} shifts, "
                 f"{hours:,.0f} hours, {clashes} clashes")

    # ------------------------------------------------------- the team ----
    expiring = 0
    for row in data["Staff Database"]:
        for key in ("first_aid_exp", "cpr_exp", "safeguard_exp",
                    "check_exp"):
            when = _day(row.get(key))
            if when and when <= cfg.today:
                expiring += 1
    if not expiring:
        problems.append("every certificate is in date - the expiry "
                        "reminder has nothing to show")
    paid_staff = [row for row in data["Staff Database"]
                  if row.get("status") == "Active"]
    notes.append(f"team: {len(paid_staff)} active, {expiring} lapsed "
                 f"certificates on file")

    # --------------------------------------------------------- the stock -
    for row in data["Supply Inventory"]:
        stock = (_num(row.get("opening")) + _num(row.get("received"))
                 - _num(row.get("used")) - _num(row.get("waste"))
                 + _num(row.get("adjust")))
        if stock < 0:
            problems.append(f"supply '{row.get('name')}' goes to "
                            f"{stock:,.0f} - stock cannot be negative")

    # ------------------------------------------------ the monthly shape --
    by_month = defaultdict(float)
    for row in data["Invoices & Receipts"]:
        if row.get("type") == "Invoice":
            key = rid("Invoice", row["_row"])
            amount = charge_total.get(key) or _num(row.get("amount"))
            by_month[_month(row.get("date"))] += amount
    if by_month:
        average = sum(by_month.values()) / len(by_month)
        target = cfg.revenue_target_month
        if average < target * 0.5 or average > target * 1.6:
            problems.append(
                f"the demo bills {average:,.0f} a month against a target "
                f"of {target:,.0f} - the dashboard will look odd")
        notes.append("billing by month: " + ", ".join(
            f"{key} {value:,.0f}" for key, value in sorted(by_month.items())))

    # ------------------------------------------------------- paperwork ---
    overdue_docs = sum(1 for row in data["Document Tracker"]
                       if _day(row.get("expires"))
                       and _day(row["expires"]) < cfg.today)
    if not overdue_docs:
        problems.append("no document has expired - the renewal alert has "
                        "nothing to show")
    notes.append(f"paperwork: {len(data['Document Tracker'])} documents, "
                 f"{overdue_docs} past their renewal date")

    print(f"--- {path}")
    for note in notes:
        print(f"  . {note}")
    if problems:
        print(f"  ! {len(problems)} problem(s):")
        for line in problems[:30]:
            print("    - " + line)
        return 1
    print("  OK - the example nursery adds up")
    return 0


if __name__ == "__main__":
    target = (sys.argv[1] if len(sys.argv) > 1
              else "output/Daycare Business Management Spreadsheet.xlsx")
    raise SystemExit(check(target))
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

#### Child Enrollment

Every child, their family, their room, their fee and the paperwork still outstanding.

73 columns, 150 rows of capacity, 30 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Child ID | `id` | id | formula |
| B | Child's full name | `name` | text | you type it |
| C | Known as | `preferred` | text | you type it |
| D | Date of birth | `dob` | date | you type it (checked) |
| E | Age (months) | `age_months` | int | formula, hidden |
| F | Age | `age` | text | formula |
| G | Age band | `age_band` | text | formula |
| H | Gender | `gender` | text | drop down `GenderList` |
| I | Family ID | `family` | id | drop down `FamilyList` |
| J | Family | `family_name` | text | formula |
| K | Main contact | `parent_name` | text | formula |
| L | Phone | `parent_phone` | text | formula |
| M | Email | `parent_email` | text | formula, hidden |
| N | Sibling no. | `family_seq` | int | formula, hidden |
| O | Family key | `family_key` | text | formula, hidden |
| P | Programme | `program` | text | drop down `ProgramList` |
| Q | Classroom | `room` | text | drop down `ClassroomList` |
| R | Key teacher | `teacher` | text | formula |
| S | Schedule | `schedule` | text | drop down `ScheduleTypeList` |
| T | Days a week | `days_week` | number1 | you type it (checked) |
| U | Hours a day | `hours_day` | number1 | you type it (checked) |
| V | Usual drop-off | `usual_in` | time | you type it (checked) |
| W | Usual pickup | `usual_out` | time | you type it (checked) |
| X | Enrolled on | `enrolled_date` | date | you type it (checked) |
| Y | Starts | `start_date` | date | you type it (checked) |
| Z | Last day | `end_date` | date | you type it (checked) |
| AA | Leaving because | `leaving_reason` | text | drop down `LeavingReasonList`, hidden |
| AB | Status | `status` | status | drop down `EnrollmentStatusList` |
| AC | Holds a place? | `holds_place` | flag | formula, hidden |
| AD | Months with us | `months_with_us` | int | formula, hidden |
| AE | Billing basis | `basis` | text | drop down `BillingBasisList` |
| AF | Agreed rate | `rate` | money | you type it (checked) |
| AG | Sibling % | `auto_discount` | percent | formula, hidden |
| AH | Discount | `discount_pct` | percent | you type it (checked) |
| AI | Discount applied | `discount_used` | percent | formula |
| AJ | Subsidy a week | `subsidy_week` | money | you type it (checked) |
| AK | Full rate a week | `gross_week` | money | formula |
| AL | Weekly fee | `weekly_net` | money | formula |
| AM | Monthly fee | `monthly_net` | money | formula |
| AN | A year | `annual_net` | money | formula, hidden |
| AO | Billed to date | `billed` | money | formula |
| AP | Family balance | `balance` | money | formula |
| AQ | Sessions booked | `sessions` | int | formula, hidden |
| AR | Days attended | `days_in` | int | formula |
| AS | Days absent | `days_out` | int | formula |
| AT | Sick days | `days_sick` | int | formula, hidden |
| AU | Attendance | `attendance_pct` | percent | formula |
| AV | Last attended | `last_in` | date | formula |
| AW | Days since | `days_since` | int | formula, hidden |
| AX | Late arrivals | `late_count` | int | formula, hidden |
| AY | Late pickups | `late_pickups` | int | formula |
| AZ | Diet | `diet` | text | drop down `DietList` |
| BA | Allergy risk | `allergy_level` | status | formula |
| BB | Allergies | `allergy_count` | int | formula, hidden |
| BC | Main allergy | `first_allergy` | text | formula, hidden |
| BD | Care notes | `medical_notes` | wrap | you type it |
| BE | Authorised to collect | `pickup_list` | wrap | you type it |
| BF | Pickup password | `pickup_password` | text | you type it |
| BG | Photos OK? | `photo_ok` | flag | you type it (checked) |
| BH | Sunscreen OK? | `sunscreen_ok` | flag | you type it (checked), hidden |
| BI | Outings OK? | `outings_ok` | flag | you type it (checked), hidden |
| BJ | Docs required | `docs_required` | int | formula, hidden |
| BK | Docs on file | `docs_ok` | int | formula, hidden |
| BL | Docs missing | `docs_missing` | int | formula |
| BM | Agreement | `contract` | status | drop down `DocumentStatusList` |
| BN | Next birthday | `next_bday` | date | formula |
| BO | Days to birthday | `bday_in` | int | formula, hidden |
| BP | Turning | `turning` | int | formula, hidden |
| BQ | Years with us | `anniversary` | number1 | formula, hidden |
| BR | Moves up on | `moves_up` | date | formula |
| BS | Notes | `notes` | wrap | you type it |
| BT | Value rank | `annual_net_rank` | int | formula, hidden |
| BU | Action | `alert` | status | formula |

#### Parent Database

One row per family. Siblings share it, invoices follow it, and the balance is the family's, not the child's.

51 columns, 200 rows of capacity, 28 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Family ID | `id` | id | formula |
| B | Family name | `family_name` | text | you type it |
| C | Parent / guardian 1 | `parent1` | text | you type it |
| D | Relationship | `relation1` | text | drop down `RelationshipList` |
| E | Phone | `phone1` | text | you type it |
| F | Email | `email1` | text | you type it |
| G | Employer | `employer1` | text | you type it, hidden |
| H | Work phone | `work_phone1` | text | you type it, hidden |
| I | Parent / guardian 2 | `parent2` | text | you type it |
| J | Relationship | `relation2` | text | drop down `RelationshipList` |
| K | Phone | `phone2` | text | you type it |
| L | Email | `email2` | text | you type it, hidden |
| M | Address | `address` | text | you type it |
| N | Town / city | `city` | text | you type it |
| O | Postcode / ZIP | `postcode` | text | you type it |
| P | Contact by | `contact_by` | text | drop down `ContactMethodList` |
| Q | Main language | `language` | text | you type it, hidden |
| R | Emergency contact 1 | `em1_name` | text | you type it |
| S | Phone | `em1_phone` | text | you type it |
| T | Relationship | `em1_relation` | text | drop down `RelationshipList`, hidden |
| U | Emergency contact 2 | `em2_name` | text | you type it, hidden |
| V | Phone | `em2_phone` | text | you type it, hidden |
| W | Family doctor | `doctor` | text | you type it, hidden |
| X | Doctor's phone | `doctor_phone` | text | you type it, hidden |
| Y | Authorised to collect | `pickup_list` | wrap | you type it |
| Z | Pickup list signed? | `pickup_ok` | flag | drop down `YesNoList` |
| AA | Custody / legal notes | `custody_notes` | wrap | you type it, hidden |
| AB | Children | `children` | int | formula |
| AC | In now | `children_active` | int | formula |
| AD | First child | `child1` | text | formula |
| AE | Second child | `child2` | text | formula |
| AF | Third child | `child3` | text | formula, hidden |
| AG | And others | `more_children` | text | formula, hidden |
| AH | Fees a week | `weekly_fees` | money | formula |
| AI | Fees a month | `monthly_fees` | money | formula, hidden |
| AJ | Billed | `billed` | money | formula, hidden |
| AK | Invoiced | `invoiced` | money | formula |
| AL | Paid | `paid` | money | formula |
| AM | Balance | `balance` | money | formula |
| AN | Overdue | `overdue` | money | formula |
| AO | Oldest unpaid | `oldest_due` | date | formula, hidden |
| AP | Last payment | `last_paid` | date | formula |
| AQ | Days since | `days_since_paid` | int | formula, hidden |
| AR | On subsidy? | `subsidy` | flag | formula, hidden |
| AS | With us since | `joined` | date | you type it (checked) |
| AT | Messages logged | `messages` | int | formula, hidden |
| AU | Open items | `open_items` | int | formula |
| AV | Family status | `status` | status | formula |
| AW | Notes | `notes` | wrap | you type it |
| AX | Value rank | `weekly_fees_rank` | int | formula, hidden |
| AY | Action | `alert` | status | formula |

#### Waitlist

Every enquiry, where it got to, how long they have waited and which of them you can say yes to today.

36 columns, 120 rows of capacity, 21 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Enquiry ID | `id` | id | formula |
| B | Enquiry date | `enquiry_date` | date | you type it (checked) |
| C | Child's name | `child_name` | text | you type it |
| D | Date of birth | `dob` | date | you type it (checked) |
| E | Age (months) | `age_months` | int | formula, hidden |
| F | Age now | `age` | text | formula |
| G | Age at start | `age_at_start` | text | formula, hidden |
| H | Age band then | `band_wanted` | text | formula |
| I | Parent / guardian | `parent_name` | text | you type it |
| J | Phone | `phone` | text | you type it |
| K | Email | `email` | text | you type it |
| L | Heard about you via | `source` | text | drop down `InquirySourceList` |
| M | Programme wanted | `program` | text | drop down `ProgramList` |
| N | Room wanted | `room` | text | drop down `ClassroomList` |
| O | Days a week | `days_wanted` | number1 | you type it (checked) |
| P | Wants to start | `start_wanted` | date | you type it (checked) |
| Q | Priority | `priority` | text | drop down `WaitlistPriorityList` |
| R | Priority order | `priority_rank` | int | formula, hidden |
| S | Sibling already here | `sibling` | id | drop down `ChildList` |
| T | Sibling | `sibling_name` | text | formula, hidden |
| U | Status | `status` | status | drop down `WaitlistStatusList` |
| V | Info sent | `info_sent` | date | you type it (checked), hidden |
| W | Tour | `tour_date` | date | you type it (checked) |
| X | Offered | `offered_date` | date | you type it (checked) |
| Y | They replied | `response_date` | date | you type it (checked), hidden |
| Z | Became child | `enrolled_child` | id | drop down `ChildList` |
| AA | Closed? | `closed` | flag | formula, hidden |
| AB | Enrolled? | `converted` | flag | formula |
| AC | Days waiting | `days_waiting` | int | formula |
| AD | Days since contact | `days_quiet` | int | formula |
| AE | Queue position | `position` | int | formula |
| AF | Seats free now | `seats_now` | int | formula |
| AG | Worth a week | `value_week` | money | formula |
| AH | Last follow-up | `followed_up` | date | you type it (checked), hidden |
| AI | Notes | `notes` | wrap | you type it |
| AJ | Action | `alert` | status | formula |

#### Attendance

Who was in, when they arrived, when they left, who collected them - and what that late pickup is worth.

35 columns, 2000 rows of capacity, 12 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Entry ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Day | `weekday` | text | formula |
| D | Month | `month` | text | formula, hidden |
| E | Week beginning | `week_start` | date | formula, hidden |
| F | Child | `child` | id | drop down `ChildList` |
| G | Child's name | `child_name` | text | formula |
| H | Room | `room` | text | formula |
| I | Programme | `program` | text | formula, hidden |
| J | Day key | `daykey` | text | formula, hidden |
| K | Status | `status` | status | drop down `AttendanceStatusList` |
| L | Session? | `session` | int | formula, hidden |
| M | Present? | `present` | int | formula, hidden |
| N | Absent? | `absent` | int | formula, hidden |
| O | Checked in | `in_time` | time | you type it (checked) |
| P | Checked out | `out_time` | time | you type it (checked) |
| Q | Hours | `hours` | number1 | formula |
| R | Due in | `due_in` | time | formula, hidden |
| S | Due out | `due_out` | time | formula, hidden |
| T | Late in (min) | `late_in` | int | formula |
| U | Left early (min) | `early_out` | int | formula, hidden |
| V | Late pickup (min) | `late_out` | int | formula |
| W | Late fee due | `late_fee` | money | formula |
| X | Fee charged? | `fee_raised` | flag | drop down `YesNoList` |
| Y | Dropped off by | `dropped_by` | text | you type it |
| Z | Collected by | `collected_by` | text | you type it |
| AA | Authorised list | `auth_list` | text | formula, hidden |
| AB | Authorised? | `pickup_ok` | flag | formula |
| AC | Absence reason | `reason` | text | you type it |
| AD | Recorded by | `recorded_by` | id | drop down `StaffList` |
| AE | Staff | `recorded_name` | text | formula, hidden |
| AF | Temp / health check | `temp_check` | text | you type it, hidden |
| AG | Notes | `notes` | wrap | you type it |
| AH | Absent in a row | `run_absent` | int | formula, hidden |
| AI | Action | `alert` | status | formula |

#### Tuition & Fees

Every charge and every credit, ready to be pulled onto an invoice. Discounts and subsidies are typed positive and come off automatically.

29 columns, 700 rows of capacity, 13 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Charge ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Month | `month` | text | formula, hidden |
| D | Child | `child` | id | drop down `ChildList` |
| E | Child's name | `child_name` | text | formula |
| F | Family | `family` | id | formula |
| G | Family override | `family_set` | id | drop down `FamilyList`, hidden |
| H | Billed to | `family_used` | id | formula, hidden |
| I | Charge type | `fee_type` | text | drop down `FeeTypeList` |
| J | Credit? | `is_credit` | flag | formula, hidden |
| K | Description | `description` | text | you type it |
| L | Covers from | `period_from` | date | you type it (checked), hidden |
| M | Covers to | `period_to` | date | you type it (checked), hidden |
| N | Quantity | `qty` | number1 | you type it (checked) |
| O | Rate | `rate` | money | you type it (checked) |
| P | Gross | `gross` | money | formula |
| Q | Discount % | `discount_pct` | percent | you type it (checked), hidden |
| R | Discount | `discount` | money | formula, hidden |
| S | Subtotal | `subtotal` | money | formula |
| T | Tax | `tax` | money | formula |
| U | Amount | `net` | money | formula |
| V | Invoice | `invoice` | text | drop down `InvoiceList` |
| W | Billed? | `billed` | flag | formula |
| X | Invoice found? | `invoice_ok` | flag | formula, hidden |
| Y | Line | `seq` | int | formula, hidden |
| Z | Key | `key` | text | formula, hidden |
| AA | Raised by | `auto_source` | text | drop down `ChargeSourceList`, hidden |
| AB | Notes | `notes` | wrap | you type it |
| AC | Action | `alert` | status | formula |

#### Invoices & Receipts

One ledger for invoices, payments, credits, refunds and write-offs. Revenue and cash kept firmly apart.

31 columns, 700 rows of capacity, 13 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Reference | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Month | `month` | text | formula, hidden |
| D | Type | `type` | status | drop down `LedgerTypeList` |
| E | Family | `family` | id | drop down `FamilyList` |
| F | Family name | `family_name` | text | formula |
| G | Child (optional) | `child` | id | drop down `ChildList`, hidden |
| H | Period | `period` | text | you type it |
| I | Description | `description` | text | you type it |
| J | From charge lines | `charges` | money | formula |
| K | Lines | `lines` | int | formula, hidden |
| L | Amount typed | `amount` | money | you type it (checked) |
| M | Total | `total` | money | formula |
| N | Pays invoice | `against` | text | drop down `InvoiceList` |
| O | Method | `method` | text | drop down `PaymentMethodList` |
| P | Reference | `reference` | text | you type it, hidden |
| Q | Terms (days) | `terms` | int | you type it (checked), hidden |
| R | Due | `due_date` | date | formula |
| S | Due (number) | `due_num` | number | formula, hidden |
| T | Paid so far | `paid` | money | formula |
| U | Credited | `credited` | money | formula, hidden |
| V | Still owed | `balance` | money | formula |
| W | Status | `status` | status | formula |
| X | Sent to parent? | `sent` | flag | drop down `YesNoList` |
| Y | Days overdue | `days_overdue` | int | formula |
| Z | Revenue effect | `billed` | money | formula |
| AA | Cash effect | `collected` | money | formula |
| AB | Account effect | `owed` | money | formula, hidden |
| AC | Processing fee | `fee` | money | formula, hidden |
| AD | Notes | `notes` | wrap | you type it |
| AE | Action | `alert` | status | formula |

#### Daily Reports

Mood, activities, naps, nappies, meals and the end-of-day summary - written for you from the columns beside it.

41 columns, 600 rows of capacity, 20 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Report ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Day | `weekday` | text | formula, hidden |
| D | Month | `month` | text | formula, hidden |
| E | Child | `child` | id | drop down `ChildList` |
| F | Child's name | `child_name` | text | formula |
| G | Known as | `preferred` | text | formula, hidden |
| H | Room | `room` | text | formula |
| I | Day key | `daykey` | text | formula, hidden |
| J | Arrived | `arrived` | time | formula |
| K | Left | `departed` | time | formula |
| L | Hours | `hours` | number1 | formula, hidden |
| M | Mood | `mood` | text | drop down `MoodList` |
| N | Behaviour note | `behaviour` | wrap | you type it, hidden |
| O | What we did | `activities` | wrap | you type it |
| P | Learning focus | `learning` | wrap | you type it |
| Q | Outdoors (min) | `outdoor_min` | int | you type it (checked) |
| R | Nap from | `nap_from` | time | you type it (checked) |
| S | Nap to | `nap_to` | time | you type it (checked) |
| T | Slept | `nap_hours` | number1 | formula |
| U | Second nap (min) | `nap2_min` | int | you type it (checked), hidden |
| V | Total sleep | `sleep_total` | number1 | formula, hidden |
| W | Nappies wet | `nappies_wet` | int | you type it (checked) |
| X | Nappies soiled | `nappies_soiled` | int | you type it (checked) |
| Y | Nappy changes | `nappies` | int | formula, hidden |
| Z | Toileting | `toileting` | text | drop down `ToiletingList` |
| AA | Potty training | `potty_notes` | wrap | you type it, hidden |
| AB | Meals logged | `meals_logged` | int | formula |
| AC | Ate well | `ate_well` | int | formula, hidden |
| AD | Refused | `refused` | int | formula, hidden |
| AE | Eating | `eating` | text | formula |
| AF | Drinks | `water` | int | you type it (checked) |
| AG | Bottles (ml) | `bottles` | int | you type it (checked), hidden |
| AH | Written by | `teacher` | id | drop down `StaffList` |
| AI | Staff | `teacher_name` | text | formula, hidden |
| AJ | Teacher's observations | `observations` | wrap | you type it |
| AK | Message to parents | `parent_notes` | wrap | you type it |
| AL | Incidents today | `incidents` | int | formula |
| AM | End of day summary | `summary` | wrap | formula |
| AN | Sent to parent? | `shared` | flag | drop down `YesNoList` |
| AO | Action | `alert` | status | formula |

#### Meal Tracker

Every plate, with the child's diet and allergy level on the same line as the food.

22 columns, 600 rows of capacity, 11 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Meal ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Month | `month` | text | formula, hidden |
| D | Sitting | `sitting` | text | drop down `MealSittingList` |
| E | Child | `child` | id | drop down `ChildList` |
| F | Child's name | `child_name` | text | formula |
| G | Room | `room` | text | formula, hidden |
| H | Day key | `daykey` | text | formula, hidden |
| I | Diet on file | `diet` | text | formula |
| J | Allergy risk | `allergy_level` | status | formula |
| K | Allergic to | `allergy_main` | text | formula |
| L | Food served | `food` | text | you type it |
| M | Amount eaten | `portion` | text | drop down `EatenAmountList` |
| N | Refused? | `refused` | flag | formula, hidden |
| O | Alternative given | `alternative` | text | you type it, hidden |
| P | Drink | `drink` | text | you type it |
| Q | Fluid (ml) | `water_ml` | int | you type it (checked) |
| R | Brought own? | `own_food` | flag | drop down `YesNoList`, hidden |
| S | Logged by | `logged_by` | id | drop down `StaffList` |
| T | Staff | `logged_name` | text | formula, hidden |
| U | Notes | `notes` | wrap | you type it |
| V | Action | `alert` | status | formula |

#### Allergy & Medical

Conditions, allergy plans, consents, expiry dates and the dose log - one child's record in one place.

33 columns, 180 rows of capacity, 22 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Record ID | `id` | id | formula |
| B | Child | `child` | id | drop down `ChildList` |
| C | Child's name | `child_name` | text | formula |
| D | Room | `room` | text | formula, hidden |
| E | Line | `seq` | int | formula, hidden |
| F | Key | `key` | text | formula, hidden |
| G | Record type | `record_type` | text | drop down `MedicalRecordList` |
| H | Dose log? | `is_event` | flag | formula, hidden |
| I | Detail | `detail` | text | you type it |
| J | Severity | `severity` | status | drop down `AllergySeverityList` |
| K | What happens | `reaction` | wrap | you type it |
| L | What to do | `action` | wrap | you type it |
| M | Medication | `medication` | text | you type it |
| N | Dose & route | `dose` | text | you type it |
| O | Stored where | `stored` | text | you type it |
| P | Staff trained | `trained` | text | you type it, hidden |
| Q | Parent consent | `consent` | flag | drop down `YesNoList` |
| R | Consent dated | `consent_date` | date | you type it (checked), hidden |
| S | Prescribed by | `prescriber` | text | you type it, hidden |
| T | Doctor / clinic | `doctor` | text | you type it |
| U | Doctor's phone | `doctor_phone` | text | you type it |
| V | Plan dated | `reviewed` | date | you type it (checked) |
| W | Expires | `expires` | date | you type it (checked) |
| X | Days left | `days_left` | int | formula |
| Y | Given on | `given_date` | date | you type it (checked) |
| Z | Given at | `given_time` | time | you type it (checked) |
| AA | Given by | `given_by` | id | drop down `StaffList` |
| AB | Staff | `given_name` | text | formula |
| AC | Witnessed by | `witness` | text | you type it, hidden |
| AD | Doses logged | `doses_logged` | int | formula, hidden |
| AE | Last dose | `last_dose` | date | formula, hidden |
| AF | Notes | `notes` | wrap | you type it |
| AG | Action | `alert` | status | formula |

#### Incident Reports

What happened, who dealt with it, when the parent was told and what changed afterwards.

36 columns, 150 rows of capacity, 25 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Incident ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Month | `month` | text | formula, hidden |
| D | Time | `time` | time | you type it (checked) |
| E | Child | `child` | id | drop down `ChildList` |
| F | Child's name | `child_name` | text | formula |
| G | Age then | `age_at` | text | formula, hidden |
| H | Room | `room` | text | formula, hidden |
| I | Where | `location` | text | drop down `IncidentLocationList` |
| J | What happened | `category` | text | drop down `IncidentCategoryList` |
| K | Severity | `severity` | status | drop down `IncidentSeverityList` |
| L | Description | `description` | wrap | you type it |
| M | Injury & where | `injury` | wrap | you type it |
| N | First aid given | `first_aid` | wrap | you type it |
| O | First aid by | `first_aider` | id | drop down `StaffList` |
| P | First aider | `first_aider_name` | text | formula |
| Q | Staff present | `staff` | id | drop down `StaffList` |
| R | Supervising | `staff_name` | text | formula, hidden |
| S | Witness | `witness` | text | you type it |
| T | Parent told? | `notified` | flag | drop down `YesNoList` |
| U | Told on | `notified_date` | date | you type it (checked) |
| V | Told at | `notified_time` | time | you type it (checked) |
| W | Told by | `notified_how` | text | drop down `ContactMethodList` |
| X | Told by whom | `notified_by` | id | drop down `StaffList`, hidden |
| Y | Minutes to tell | `notify_mins` | int | formula |
| Z | Parent signed? | `signed` | flag | drop down `YesNoList` |
| AA | Signed on | `signed_date` | date | you type it (checked), hidden |
| AB | Follow-up action | `follow_up` | wrap | you type it |
| AC | Follow up by | `follow_by` | date | you type it (checked) |
| AD | Done? | `follow_done` | flag | drop down `YesNoList` |
| AE | Reported to | `reported_out` | text | you type it |
| AF | Medical attention | `medical_sought` | text | you type it, hidden |
| AG | Form complete? | `complete` | flag | formula |
| AH | Printable write-up | `writeup` | wrap | formula |
| AI | Internal notes | `notes` | wrap | you type it |
| AJ | Action | `alert` | status | formula |

#### Staff Database

The team, what they are paid, and exactly how many days until a certificate runs out.

40 columns, 60 rows of capacity, 25 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Staff ID | `id` | id | formula |
| B | Name | `name` | text | you type it |
| C | Role | `role` | text | drop down `RoleList` |
| D | Main room | `room` | text | drop down `ClassroomList` |
| E | Phone | `phone` | text | you type it |
| F | Email | `email` | text | you type it |
| G | Address | `address` | text | you type it, hidden |
| H | Date of birth | `dob` | date | you type it (checked), hidden |
| I | Started | `hire_date` | date | you type it (checked) |
| J | Status | `status` | status | drop down `EmploymentStatusList` |
| K | Left on | `end_date` | date | you type it (checked), hidden |
| L | Years here | `years` | number1 | formula, hidden |
| M | Paid as | `pay_type` | text | drop down `PayTypeList` |
| N | Hourly rate | `hourly` | money | you type it (checked) |
| O | Salary a year | `salary` | money | you type it (checked) |
| P | Hours a week | `contract_hours` | number1 | you type it (checked) |
| Q | Day rate | `day_rate` | money | you type it (checked), hidden |
| R | Effective hourly | `rate` | money | formula |
| S | Cost an hour | `cost_hour` | money | formula, hidden |
| T | First aid expires | `first_aid_exp` | date | you type it (checked) |
| U | CPR expires | `cpr_exp` | date | you type it (checked) |
| V | Safeguarding expires | `safeguard_exp` | date | you type it (checked) |
| W | Background check expires | `check_exp` | date | you type it (checked) |
| X | Check completed | `check_done` | date | you type it (checked), hidden |
| Y | Qualification | `qualification` | text | you type it |
| Z | Other training | `training` | wrap | you type it, hidden |
| AA | Next expiry | `next_expiry` | date | formula |
| AB | Days left | `days_to_expiry` | int | formula |
| AC | Certificates | `certs_ok` | status | formula |
| AD | Docs missing | `docs_missing` | int | formula |
| AE | Rota'd this week | `hours_week` | number1 | formula |
| AF | Hours this month | `hours_month` | number1 | formula, hidden |
| AG | Over contract | `over_contract` | number1 | formula, hidden |
| AH | Paid this month | `pay_month` | money | formula, hidden |
| AI | Children in room | `children_room` | int | formula, hidden |
| AJ | Leave taken | `leave_days` | number1 | formula, hidden |
| AK | Emergency contact | `em_name` | text | you type it |
| AL | Emergency phone | `em_phone` | text | you type it |
| AM | Notes | `notes` | wrap | you type it |
| AN | Action | `alert` | status | formula |

#### Staff Schedule

The rota, the clock, the overtime and what each shift costs - with the ratio counts feeding the classroom tab.

32 columns, 750 rows of capacity, 12 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Shift ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Day | `weekday` | text | formula |
| D | Week beginning | `week_start` | date | formula, hidden |
| E | Month | `month` | text | formula, hidden |
| F | Staff | `staff` | id | drop down `StaffList` |
| G | Name | `staff_name` | text | formula |
| H | Role | `role` | text | formula, hidden |
| I | Paid as | `pay_type` | text | formula, hidden |
| J | Room | `room` | text | drop down `ClassroomList` |
| K | Shift | `shift` | text | drop down `ShiftTypeList` |
| L | Counts in ratio? | `counts` | flag | formula |
| M | Leave type | `leave_type` | text | drop down `LeaveTypeList`, hidden |
| N | Starts | `start` | time | you type it (checked) |
| O | Ends | `finish` | time | you type it (checked) |
| P | Break (min) | `break_min` | int | you type it (checked) |
| Q | Rota'd hours | `hours` | number1 | formula |
| R | Clocked in | `clock_in` | time | you type it (checked) |
| S | Clocked out | `clock_out` | time | you type it (checked) |
| T | Actual hours | `actual_hours` | number1 | formula |
| U | Over / under | `variance` | number1 | formula, hidden |
| V | Week to date | `week_to_date` | number1 | formula, hidden |
| W | Overtime | `ot_hours` | number1 | formula |
| X | Basic hours | `basic_hours` | number1 | formula, hidden |
| Y | Rate | `rate` | money | formula, hidden |
| Z | Shift cost | `cost` | money | formula |
| AA | With on-costs | `cost_full` | money | formula, hidden |
| AB | Covered by | `substitute` | id | drop down `StaffList` |
| AC | Cover | `sub_name` | text | formula, hidden |
| AD | Double booked | `clash` | int | formula, hidden |
| AE | Notes | `notes` | wrap | you type it |
| AF | Action | `alert` | status | formula |

#### Payroll Tracker

Hours from the rota, pay from the staff tab, employer costs on top - one row per person per run.

29 columns, 350 rows of capacity, 11 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Pay run ID | `id` | id | formula |
| B | Period ends | `period_end` | date | you type it (checked) |
| C | Month | `month` | text | formula, hidden |
| D | Period starts | `period_start` | date | you type it (checked) |
| E | Run type | `pay_period` | text | drop down `PayPeriodList` |
| F | Staff | `staff` | id | drop down `StaffList` |
| G | Name | `staff_name` | text | formula |
| H | Paid as | `pay_type` | text | formula |
| I | Rate | `rate` | money | formula |
| J | Basic hours | `basic_hours` | number1 | formula |
| K | Overtime hours | `ot_hours` | number1 | formula |
| L | Total hours | `total_hours` | number1 | formula, hidden |
| M | Leave days | `leave_days` | int | formula, hidden |
| N | Runs a year | `periods_year` | int | formula, hidden |
| O | Basic pay | `basic_pay` | money | formula |
| P | Overtime pay | `ot_pay` | money | formula |
| Q | Bonus | `bonus` | money | you type it (checked) |
| R | Adjustment | `adjust` | money | you type it (checked) |
| S | Gross pay | `gross` | money | formula |
| T | Employer costs | `employer_cost` | money | formula |
| U | Total cost | `total_cost` | money | formula |
| V | Deductions | `deductions` | money | you type it (checked), hidden |
| W | Net to pay | `net_pay` | money | formula |
| X | Cost an hour | `cost_hour` | money | formula, hidden |
| Y | Paid? | `paid` | flag | drop down `YesNoList` |
| Z | Paid on | `paid_date` | date | you type it (checked) |
| AA | Method | `method` | text | drop down `PaymentMethodList`, hidden |
| AB | Notes | `notes` | wrap | you type it |
| AC | Action | `alert` | status | formula |

#### Classroom Management

Capacity, occupancy, staff-to-child ratios and what every empty place is costing you a week.

33 columns, 30 rows of capacity, 8 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Room ID | `id` | id | formula |
| B | Room name | `name` | text | you type it |
| C | Age band | `age_band` | text | drop down `AgeBandList` |
| D | From (months) | `age_from` | int | formula |
| E | To (months) | `age_to` | int | formula |
| F | Licensed places | `capacity` | int | you type it (checked) |
| G | Band ratio | `ratio_default` | int | formula, hidden |
| H | Ratio override | `ratio_set` | int | you type it (checked) |
| I | Children per adult | `ratio` | int | formula |
| J | Max group size | `max_group` | int | formula, hidden |
| K | Lead teacher | `teacher` | id | drop down `StaffList` |
| L | Teacher | `teacher_name` | text | formula |
| M | Assistant | `assistant` | id | drop down `StaffList` |
| N | Assistant name | `assistant_name` | text | formula, hidden |
| O | Room status | `room_status` | status | drop down `RoomStatusList` |
| P | Children enrolled | `enrolled` | int | formula |
| Q | Active only | `active` | int | formula, hidden |
| R | Starting soon | `starting` | int | formula |
| S | Leaving soon | `leaving` | int | formula, hidden |
| T | Seats free | `seats` | int | formula |
| U | Occupancy | `occupancy` | percent | formula |
| V | On the waitlist | `waiting` | int | formula |
| W | In today | `present_today` | int | formula |
| X | Adults in today | `staff_today` | number1 | formula |
| Y | Adults needed | `staff_needed` | number1 | formula |
| Z | Adults at full | `staff_needed_full` | number1 | formula, hidden |
| AA | Ratio today | `ratio_now` | text | formula |
| AB | Ratio met? | `ratio_ok` | flag | formula |
| AC | Room revenue a week | `revenue` | money | formula |
| AD | Per place a week | `revenue_seat` | money | formula, hidden |
| AE | Empty places cost | `empty_cost` | money | formula |
| AF | Notes | `notes` | wrap | you type it |
| AG | Action | `alert` | status | formula |

#### Expense Tracker

Everything that goes out, categorised, converted to a monthly equivalent and measured against budget.

23 columns, 500 rows of capacity, 11 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Expense ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Month | `month` | text | formula, hidden |
| D | Category | `category` | text | drop down `ExpenseCategoryList` |
| E | What it was | `description` | text | you type it |
| F | Paid to | `vendor` | text | you type it |
| G | Amount | `amount` | money | you type it (checked) |
| H | Tax | `tax` | money | you type it (checked), hidden |
| I | Total paid | `total` | money | formula |
| J | Paid by | `method` | text | drop down `PaymentMethodList` |
| K | Receipt? | `receipt` | flag | drop down `YesNoList` |
| L | Receipt ref | `receipt_ref` | text | you type it, hidden |
| M | Recurring | `frequency` | text | drop down `FrequencyList` |
| N | Monthly equivalent | `monthly` | money | formula |
| O | Payroll category? | `is_wages` | flag | formula, hidden |
| P | Counts in profit? | `in_pl` | flag | formula |
| Q | In the P&L | `pl_amount` | money | formula, hidden |
| R | Monthly budget | `budget` | money | formula, hidden |
| S | Category this month | `cat_month` | money | formula, hidden |
| T | Budget variance | `variance` | money | formula |
| U | Cost per child | `child_cost` | money | formula, hidden |
| V | Notes | `notes` | wrap | you type it |
| W | Action | `alert` | status | formula |

#### Supply Inventory

What is on the shelf, how fast it goes, and what to order before somebody has to drive to the shop.

27 columns, 180 rows of capacity, 17 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Item ID | `id` | id | formula |
| B | Item | `name` | text | you type it |
| C | Category | `category` | text | drop down `SupplyCategoryList` |
| D | Unit | `unit` | text | you type it |
| E | Items per unit | `pack_size` | int | you type it (checked), hidden |
| F | Supplier | `supplier` | text | you type it |
| G | Supplier phone | `supplier_phone` | text | you type it, hidden |
| H | Lead time (days) | `lead_days` | int | you type it (checked), hidden |
| I | Opening stock | `opening` | number1 | you type it (checked) |
| J | Received | `received` | number1 | you type it (checked) |
| K | Used | `used` | number1 | you type it (checked) |
| L | Waste / damaged | `waste` | number1 | you type it (checked) |
| M | Adjustment | `adjust` | number1 | you type it (checked) |
| N | On the shelf | `stock` | number1 | formula |
| O | Cost a unit | `cost` | money | you type it (checked) |
| P | Stock value | `value` | money | formula |
| Q | Spent so far | `spend` | money | formula, hidden |
| R | Used a week | `weekly_use` | number1 | formula |
| S | Weeks left | `weeks_cover` | number1 | formula |
| T | Minimum level | `min_level` | number1 | you type it (checked) |
| U | Order this much | `order_qty` | number1 | formula |
| V | Order value | `order_cost` | money | formula, hidden |
| W | Last ordered | `last_ordered` | date | you type it (checked) |
| X | Last counted | `last_counted` | date | you type it (checked), hidden |
| Y | Stock status | `status` | status | formula |
| Z | Notes | `notes` | wrap | you type it |
| AA | Action | `alert` | status | formula |

#### Parent Communication

Every conversation, what was agreed and whether anybody is still waiting on you.

24 columns, 350 rows of capacity, 15 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Message ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Month | `month` | text | formula, hidden |
| D | Time | `time` | time | you type it (checked), hidden |
| E | Family | `family` | id | drop down `FamilyList` |
| F | Family name | `family_name` | text | formula |
| G | Spoke to | `parent` | text | formula, hidden |
| H | About child | `child` | id | drop down `ChildList` |
| I | Child | `child_name` | text | formula |
| J | In or out | `direction` | text | drop down `DirectionList` |
| K | How | `method` | text | drop down `ContactMethodList` |
| L | About | `category` | text | drop down `CommCategoryList` |
| M | Subject | `subject` | text | you type it |
| N | What was said | `notes` | wrap | you type it |
| O | Logged by | `staff` | id | drop down `StaffList` |
| P | Staff | `staff_name` | text | formula, hidden |
| Q | Follow up by | `follow_date` | date | you type it (checked) |
| R | Action agreed | `follow_action` | wrap | you type it |
| S | Status | `status` | status | drop down `CommStatusList` |
| T | Closed on | `closed_date` | date | you type it (checked), hidden |
| U | Days open | `days_open` | int | formula |
| V | Follow-up late? | `overdue` | flag | formula, hidden |
| W | Meeting booked | `meeting` | date | you type it (checked), hidden |
| X | Action | `alert` | status | formula |

#### Document Tracker

Children's forms, staff certificates and the setting's own licences - with the day each one runs out.

19 columns, 500 rows of capacity, 14 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | Doc ID | `id` | id | formula |
| B | Belongs to | `holder` | text | drop down `DocumentHolderList` |
| C | ID | `who` | id | you type it |
| D | Name | `who_name` | text | formula |
| E | Document | `doc_type` | text | drop down `DocumentTypeList` |
| F | Required? | `required` | flag | drop down `YesNoList` |
| G | Status | `status` | status | drop down `DocumentStatusList` |
| H | Requested | `requested` | date | you type it (checked), hidden |
| I | Received | `received` | date | you type it (checked) |
| J | Expires | `expires` | date | you type it (checked) |
| K | Days left | `days_left` | int | formula |
| L | Where it stands | `state` | status | formula |
| M | Last renewed | `renewed` | date | you type it (checked), hidden |
| N | Renew every (months) | `renew_every` | int | you type it, hidden |
| O | Chased by | `owner` | id | drop down `StaffList`, hidden |
| P | Filed where | `location` | text | you type it |
| Q | Reference | `reference` | text | you type it, hidden |
| R | Notes | `notes` | wrap | you type it |
| S | Action | `alert` | status | formula |

---

*13,598 lines of Python in 40 files.*
