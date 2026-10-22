# Salon Business Management Spreadsheet — source code

Complete source code of the generator that writes the `.xlsx` product.

Writes a 20 tab `.xlsx` workbook with over 122,000 formulas: an appointment planner with a double-booking alarm, a client CRM that knows who is drifting away, a service menu priced by the chair hour, a till that separates service from retail, deposits that cannot be counted twice, eight different commission methods, payroll with employer costs, expenses with recurring bills spread into a real monthly overhead, a twelve month profit and loss, stock and purchasing with reorder alerts, campaigns measured against the discount they gave away, a loyalty scheme, goals and budgets, analytics by category, day, hour and source, and printable weekly and monthly reports. Document author is **Novality Store**; every formula cell is locked and every sheet protected with the password **`premium`**, while the cells you are meant to type in stay open.

Nothing is hard coded to one country or one pay model. The tax rate and the two switches that say what it applies to, the commission basis, the deposit rule, the rounding rule, the loyalty scheme and every drop-down list are settings, not assumptions.

Pure Python, one dependency. Tested on XlsxWriter 3.2.0 and on the much older 0.3.0, so it does not rely on anything recent. The workbook itself is Excel 2010 safe — no dynamic arrays, no `XLOOKUP`, no `IFS`, nothing that needs array entry — and it opens in Google Sheets, LibreOffice and Numbers.

```bash
python build_salon_workbook.py --help

  -o, --output PATH     where to write the file
  --blank               no example data, just the empty workbook
  --no-protect          leave the sheets unlocked
  --password TEXT       protection password (default: premium)
  --author TEXT         document author  (default: Novality Store)
  --salon TEXT          salon name on the cover and the reports
  --owner TEXT          the owner's name
  --currency SYMBOL     currency symbol  (default: $)
  --currency-code CODE  three letter code (default: USD)
  --tax PERCENT         sales tax / VAT rate
  --tax-on WHAT         both, retail, services or neither
  --commission PERCENT  default service commission
  --commission-method   the method new stylists start on
  --hourly AMOUNT       default hourly rate
  --burden PERCENT      employer costs on top of wages
  --deposit PERCENT     deposit as a share of the price
  --chairs N            chairs or stations in the salon
  --open-days N         trading days a week
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
  * [`build_salon_workbook.py`](#build_salon_workbookpy)
* [2. The package](#2-the-package)
  * [`salon_business_spreadsheet/__init__.py`](#salon_business_spreadsheet__init__py)
  * [`salon_business_spreadsheet/__main__.py`](#salon_business_spreadsheet__main__py)
  * [`salon_business_spreadsheet/cli.py`](#salon_business_spreadsheetclipy)
  * [`salon_business_spreadsheet/config.py`](#salon_business_spreadsheetconfigpy)
  * [`salon_business_spreadsheet/layout.py`](#salon_business_spreadsheetlayoutpy)
  * [`salon_business_spreadsheet/formats.py`](#salon_business_spreadsheetformatspy)
  * [`salon_business_spreadsheet/writer.py`](#salon_business_spreadsheetwriterpy)
  * [`salon_business_spreadsheet/names.py`](#salon_business_spreadsheetnamespy)
  * [`salon_business_spreadsheet/build.py`](#salon_business_spreadsheetbuildpy)
  * [`salon_business_spreadsheet/sample_data.py`](#salon_business_spreadsheetsample_datapy)
* [3. The diary and the client book](#3-the-diary-and-the-client-book)
  * [`salon_business_spreadsheet/sheets/__init__.py`](#salon_business_spreadsheetsheets__init__py)
  * [`salon_business_spreadsheet/sheets/common.py`](#salon_business_spreadsheetsheetscommonpy)
  * [`salon_business_spreadsheet/sheets/services.py`](#salon_business_spreadsheetsheetsservicespy)
  * [`salon_business_spreadsheet/sheets/clients.py`](#salon_business_spreadsheetsheetsclientspy)
  * [`salon_business_spreadsheet/sheets/appointments.py`](#salon_business_spreadsheetsheetsappointmentspy)
* [4. The money](#4-the-money)
  * [`salon_business_spreadsheet/sheets/sales.py`](#salon_business_spreadsheetsheetssalespy)
  * [`salon_business_spreadsheet/sheets/payments.py`](#salon_business_spreadsheetsheetspaymentspy)
* [5. The team and what they are paid](#5-the-team-and-what-they-are-paid)
  * [`salon_business_spreadsheet/sheets/staff.py`](#salon_business_spreadsheetsheetsstaffpy)
  * [`salon_business_spreadsheet/sheets/commissions.py`](#salon_business_spreadsheetsheetscommissionspy)
  * [`salon_business_spreadsheet/sheets/payroll.py`](#salon_business_spreadsheetsheetspayrollpy)
* [6. Costs, stock and suppliers](#6-costs-stock-and-suppliers)
  * [`salon_business_spreadsheet/sheets/expenses.py`](#salon_business_spreadsheetsheetsexpensespy)
  * [`salon_business_spreadsheet/sheets/inventory.py`](#salon_business_spreadsheetsheetsinventorypy)
  * [`salon_business_spreadsheet/sheets/purchases.py`](#salon_business_spreadsheetsheetspurchasespy)
* [7. Growth](#7-growth)
  * [`salon_business_spreadsheet/sheets/marketing.py`](#salon_business_spreadsheetsheetsmarketingpy)
  * [`salon_business_spreadsheet/sheets/loyalty.py`](#salon_business_spreadsheetsheetsloyaltypy)
* [8. The pages](#8-the-pages)
  * [`salon_business_spreadsheet/sheets/start_here.py`](#salon_business_spreadsheetsheetsstart_herepy)
  * [`salon_business_spreadsheet/sheets/dashboard.py`](#salon_business_spreadsheetsheetsdashboardpy)
  * [`salon_business_spreadsheet/sheets/profit_loss.py`](#salon_business_spreadsheetsheetsprofit_losspy)
  * [`salon_business_spreadsheet/sheets/goals.py`](#salon_business_spreadsheetsheetsgoalspy)
  * [`salon_business_spreadsheet/sheets/analytics.py`](#salon_business_spreadsheetsheetsanalyticspy)
  * [`salon_business_spreadsheet/sheets/reports.py`](#salon_business_spreadsheetsheetsreportspy)
  * [`salon_business_spreadsheet/sheets/settings.py`](#salon_business_spreadsheetsheetssettingspy)
* [9. The quality tools](#9-the-quality-tools)
  * [`tools/validate_salon_workbook.py`](#toolsvalidate_salon_workbookpy)
  * [`tools/check_parens.py`](#toolscheck_parenspy)
  * [`tools/check_salon_sample.py`](#toolscheck_salon_samplepy)
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

`build_salon_workbook.py` parses the command line into a `Config`, then
`salon_business_spreadsheet.build.build_workbook(path, cfg)` does six things in order:

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
6. **Activate the first tab** and close the file - the buyer should land on Start Here, not on Settings.

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
| 1 | **Start Here** | page | `salon_business_spreadsheet/sheets/start_here.py` | - | - |
| 2 | **Dashboard** | page | `salon_business_spreadsheet/sheets/dashboard.py` | - | - |
| 3 | **Appointments** | register | `salon_business_spreadsheet/sheets/appointments.py` | 34 | 1200 |
| 4 | **Clients** | register | `salon_business_spreadsheet/sheets/clients.py` | 43 | 500 |
| 5 | **Services** | register | `salon_business_spreadsheet/sheets/services.py` | 39 | 200 |
| 6 | **Sales** | register | `salon_business_spreadsheet/sheets/sales.py` | 38 | 1200 |
| 7 | **Payments** | register | `salon_business_spreadsheet/sheets/payments.py` | 23 | 1200 |
| 8 | **Staff** | register | `salon_business_spreadsheet/sheets/staff.py` | 41 | 40 |
| 9 | **Commissions** | register | `salon_business_spreadsheet/sheets/commissions.py` | 27 | 240 |
| 10 | **Payroll** | register | `salon_business_spreadsheet/sheets/payroll.py` | 30 | 360 |
| 11 | **Expenses** | register | `salon_business_spreadsheet/sheets/expenses.py` | 25 | 600 |
| 12 | **Profit & Loss** | page | `salon_business_spreadsheet/sheets/profit_loss.py` | - | - |
| 13 | **Inventory** | register | `salon_business_spreadsheet/sheets/inventory.py` | 36 | 300 |
| 14 | **Purchases** | register | `salon_business_spreadsheet/sheets/purchases.py` | 25 | 300 |
| 15 | **Marketing** | register | `salon_business_spreadsheet/sheets/marketing.py` | 25 | 100 |
| 16 | **Loyalty** | register | `salon_business_spreadsheet/sheets/loyalty.py` | 24 | 500 |
| 17 | **Goals & Budget** | page | `salon_business_spreadsheet/sheets/goals.py` | - | - |
| 18 | **Analytics** | page | `salon_business_spreadsheet/sheets/analytics.py` | - | - |
| 19 | **Reports** | page | `salon_business_spreadsheet/sheets/reports.py` | - | - |
| 20 | **Settings** | page | `salon_business_spreadsheet/sheets/settings.py` | - | - |

### Every file at a glance

| File | Lines | What it is |
|------|------:|------------|
| [`build_salon_workbook.py`](#build_salon_workbookpy) | 17 | Build the Salon Business Management Spreadsheet workbook |
| [`salon_business_spreadsheet/__init__.py`](#salon_business_spreadsheet__init__py) | 19 | Salon Business Management Spreadsheet |
| [`salon_business_spreadsheet/__main__.py`](#salon_business_spreadsheet__main__py) | 6 | ``python -m salon_business_spreadsheet`` |
| [`salon_business_spreadsheet/cli.py`](#salon_business_spreadsheetclipy) | 187 | Command line front end |
| [`salon_business_spreadsheet/config.py`](#salon_business_spreadsheetconfigpy) | 478 | Configuration for the Salon Business Management Spreadsheet |
| [`salon_business_spreadsheet/layout.py`](#salon_business_spreadsheetlayoutpy) | 253 | Layout engine |
| [`salon_business_spreadsheet/formats.py`](#salon_business_spreadsheetformatspy) | 202 | Cell format factory |
| [`salon_business_spreadsheet/writer.py`](#salon_business_spreadsheetwriterpy) | 458 | Sheet writing helpers shared by every tab of the workbook |
| [`salon_business_spreadsheet/names.py`](#salon_business_spreadsheetnamespy) | 58 | Workbook level defined names that point at the register sheets |
| [`salon_business_spreadsheet/build.py`](#salon_business_spreadsheetbuildpy) | 159 | Workbook builder: puts the twenty tabs together in the right order |
| [`salon_business_spreadsheet/sample_data.py`](#salon_business_spreadsheetsample_datapy) | 1,263 | Demo data: ten weeks at Gloss & Grace Salon |
| [`salon_business_spreadsheet/sheets/__init__.py`](#salon_business_spreadsheetsheets__init__py) | 0 |  |
| [`salon_business_spreadsheet/sheets/common.py`](#salon_business_spreadsheetsheetscommonpy) | 341 | Small building blocks shared by the register sheets |
| [`salon_business_spreadsheet/sheets/services.py`](#salon_business_spreadsheetsheetsservicespy) | 276 | `Services` - the menu, with what each treatment really earns |
| [`salon_business_spreadsheet/sheets/clients.py`](#salon_business_spreadsheetsheetsclientspy) | 310 | `Clients` - the client book, and who is quietly drifting away |
| [`salon_business_spreadsheet/sheets/appointments.py`](#salon_business_spreadsheetsheetsappointmentspy) | 266 | `Appointments` - the diary, with a double-booking alarm |
| [`salon_business_spreadsheet/sheets/sales.py`](#salon_business_spreadsheetsheetssalespy) | 271 | `Sales` - the till |
| [`salon_business_spreadsheet/sheets/payments.py`](#salon_business_spreadsheetsheetspaymentspy) | 189 | `Payments` - money in and money back out, with the deposit rule |
| [`salon_business_spreadsheet/sheets/staff.py`](#salon_business_spreadsheetsheetsstaffpy) | 298 | `Staff` - the team, and what each chair is actually bringing in |
| [`salon_business_spreadsheet/sheets/commissions.py`](#salon_business_spreadsheetsheetscommissionspy) | 258 | `Commissions` - one row per stylist per month, eight ways to pay |
| [`salon_business_spreadsheet/sheets/payroll.py`](#salon_business_spreadsheetsheetspayrollpy) | 220 | `Payroll` - what each pay run actually costs the salon |
| [`salon_business_spreadsheet/sheets/expenses.py`](#salon_business_spreadsheetsheetsexpensespy) | 180 | `Expenses` - everything the salon spends, in the order a P&L wants it |
| [`salon_business_spreadsheet/sheets/inventory.py`](#salon_business_spreadsheetsheetsinventorypy) | 267 | `Inventory` - retail stock and backbar, counted the same way |
| [`salon_business_spreadsheet/sheets/purchases.py`](#salon_business_spreadsheetsheetspurchasespy) | 240 | `Purchases` - stock orders, and the supplier book underneath them |
| [`salon_business_spreadsheet/sheets/marketing.py`](#salon_business_spreadsheetsheetsmarketingpy) | 205 | `Marketing` - campaigns, offers and what they actually brought in |
| [`salon_business_spreadsheet/sheets/loyalty.py`](#salon_business_spreadsheetsheetsloyaltypy) | 190 | `Loyalty` - points, rewards and referrals, one row per client |
| [`salon_business_spreadsheet/sheets/start_here.py`](#salon_business_spreadsheetsheetsstart_herepy) | 259 | `Start Here` - what this is, what to type, and what it will not do |
| [`salon_business_spreadsheet/sheets/dashboard.py`](#salon_business_spreadsheetsheetsdashboardpy) | 438 | `Dashboard` - the one page that answers "how is the salon doing?" Sixteen numbers across the top, the jobs... |
| [`salon_business_spreadsheet/sheets/profit_loss.py`](#salon_business_spreadsheetsheetsprofit_losspy) | 316 | `Profit & Loss` - twelve months, in the order an accountant reads it |
| [`salon_business_spreadsheet/sheets/goals.py`](#salon_business_spreadsheetsheetsgoalspy) | 364 | `Goals & Budget` - what you said you would do, against what happened |
| [`salon_business_spreadsheet/sheets/analytics.py`](#salon_business_spreadsheetsheetsanalyticspy) | 351 | `Analytics` - the patterns you cannot see from behind the desk |
| [`salon_business_spreadsheet/sheets/reports.py`](#salon_business_spreadsheetsheetsreportspy) | 401 | `Reports` - the week and the month, on one printable page |
| [`salon_business_spreadsheet/sheets/settings.py`](#salon_business_spreadsheetsheetssettingspy) | 495 | `Settings` - every number and every list the workbook runs on |
| [`tools/validate_salon_workbook.py`](#toolsvalidate_salon_workbookpy) | 654 | Static checks on a generated Salon Business workbook |
| [`tools/check_parens.py`](#toolscheck_parenspy) | 121 | Parenthesis and quote balance check for every formula in a workbook |
| [`tools/check_salon_sample.py`](#toolscheck_salon_samplepy) | 451 | Does the demo salon actually hang together? The validator checks the formulas; this checks the *story*... |
| [`tools/recalc_check.py`](#toolsrecalc_checkpy) | 216 | Really calculate every formula in a generated workbook and report any cell that ends up as an Excel error... |

---

## 1. Run it

Save the files below keeping these exact paths, then run the one liner. Nothing else is needed - no project file, no build step.

```bash
pip install XlsxWriter
python build_salon_workbook.py
```

The workbook lands in `output/Salon Business Management Spreadsheet.xlsx`.

### `build_salon_workbook.py`

*17 lines*

> Build the Salon Business Management Spreadsheet workbook.
>
>     python3 build_salon_workbook.py
>     python3 build_salon_workbook.py -o "output/My Salon.xlsx" --blank
>     python3 build_salon_workbook.py --currency "£" --tax 20 --rows 400
>
> Run `--help` for every option. The generator itself lives in the
> `salon_business_spreadsheet` package next to this file.

````python
#!/usr/bin/env python3
"""Build the Salon Business Management Spreadsheet workbook.

    python3 build_salon_workbook.py
    python3 build_salon_workbook.py -o "output/My Salon.xlsx" --blank
    python3 build_salon_workbook.py --currency "£" --tax 20 --rows 400

Run `--help` for every option. The generator itself lives in the
`salon_business_spreadsheet` package next to this file.
"""

import sys

from salon_business_spreadsheet.cli import main

if __name__ == "__main__":
    sys.exit(main())
````

---

## 2. The package

Configuration, the formula template engine, cell formats, the sheet writer, the cross sheet names and the demo salon. These are the parts every tab leans on.

### `salon_business_spreadsheet/__init__.py`

*19 lines*

> Salon Business Management Spreadsheet.
>
> Build the workbook with::
>
>     python -m salon_business_spreadsheet -o "output/Salon.xlsx"
>
> or from Python::
>
>     from salon_business_spreadsheet import build_workbook
>     build_workbook("Salon Business Management Spreadsheet.xlsx")

````python
"""Salon Business Management Spreadsheet.

Build the workbook with::

    python -m salon_business_spreadsheet -o "output/Salon.xlsx"

or from Python::

    from salon_business_spreadsheet import build_workbook
    build_workbook("Salon Business Management Spreadsheet.xlsx")
"""

from .build import build_workbook
from .config import (AUTHOR, PRODUCT_LISTING_TITLE, PRODUCT_NAME,
                     PRODUCT_SUBTITLE, PRODUCT_VERSION, Config)

__all__ = ["build_workbook", "Config", "AUTHOR", "PRODUCT_NAME",
           "PRODUCT_SUBTITLE", "PRODUCT_LISTING_TITLE", "PRODUCT_VERSION"]
__version__ = PRODUCT_VERSION
````

### `salon_business_spreadsheet/__main__.py`

*6 lines*

> ``python -m salon_business_spreadsheet``.

````python
"""``python -m salon_business_spreadsheet``."""

from .cli import main

if __name__ == "__main__":
    raise SystemExit(main())
````

### `salon_business_spreadsheet/cli.py`

*187 lines · 3 functions*

> Command line front end.
>
> python -m salon_business_spreadsheet
> python -m salon_business_spreadsheet -o "output/Salon.xlsx" --blank
> python -m salon_business_spreadsheet --currency "£" --tax 20 --rows 400

| Defines | What it does |
|---------|--------------|
| `build_parser()` |  |
| `config_from_args(args)` |  |
| `main(argv)` |  |

**Module constants:** `DEFAULT_OUTPUT` (os.path.join(...)).

````python
"""Command line front end.

    python -m salon_business_spreadsheet
    python -m salon_business_spreadsheet -o "output/Salon.xlsx" --blank
    python -m salon_business_spreadsheet --currency "£" --tax 20 --rows 400
"""

from __future__ import annotations

import argparse
import datetime as _dt
import os
import sys
from typing import List, Optional

from .build import SHEET_ORDER, build_workbook
from .config import (AUTHOR, COMMISSION_METHODS, PRODUCT_LISTING_TITLE,
                     PRODUCT_NAME, ROUNDING_RULES, Config)

DEFAULT_OUTPUT = os.path.join(
    "output", "Salon Business Management Spreadsheet.xlsx")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="salon_business_spreadsheet",
        description=f"Build the {PRODUCT_NAME} workbook.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument("-o", "--output", default=DEFAULT_OUTPUT,
                        help="where to write the .xlsx file")
    parser.add_argument("--blank", action="store_true",
                        help="leave the registers empty (no demo salon)")
    parser.add_argument("--no-protect", action="store_true",
                        help="do not protect the sheets (for development)")
    parser.add_argument("--password", default=None,
                        help="sheet protection password")
    parser.add_argument("--author", default=AUTHOR,
                        help="document author / creator")
    parser.add_argument("--salon", default=None,
                        help="pre-fill the salon name on Settings")
    parser.add_argument("--owner", default=None,
                        help="pre-fill the owner's name")
    parser.add_argument("--currency", default=None,
                        help="currency symbol, e.g. $ £ € A$")
    parser.add_argument("--currency-code", default=None,
                        help="three letter currency code, e.g. USD")
    parser.add_argument("--year", type=int, default=None,
                        help="the year the reports open on")
    parser.add_argument("--tax", type=float, default=None,
                        help="sales tax / VAT rate as a percentage, "
                             "e.g. 20 for 20%%")
    parser.add_argument("--tax-on", choices=["both", "retail",
                                             "services", "neither"],
                        default=None,
                        help="what the tax rate applies to")
    parser.add_argument("--commission", type=float, default=None,
                        help="default service commission as a "
                             "percentage, e.g. 40")
    parser.add_argument("--commission-method", choices=COMMISSION_METHODS,
                        default=None,
                        help="the method new staff rows start on")
    parser.add_argument("--hourly", type=float, default=None,
                        help="default hourly rate")
    parser.add_argument("--burden", type=float, default=None,
                        help="employer costs on wages as a percentage")
    parser.add_argument("--deposit", type=float, default=None,
                        help="deposit as a percentage of the price")
    parser.add_argument("--rounding", choices=ROUNDING_RULES,
                        default=None,
                        help="how calculated prices are rounded")
    parser.add_argument("--target", type=float, default=None,
                        help="monthly revenue target")
    parser.add_argument("--chairs", type=int, default=None,
                        help="how many chairs or stations the salon has")
    parser.add_argument("--open-days", type=int, default=None,
                        help="trading days a week")
    parser.add_argument("--rows", type=int, default=None,
                        help="scale every register (the default row "
                             "counts are scaled to fit this many "
                             "appointment rows)")
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
        kwargs["tax_on_retail"] = ("Yes" if args.tax_on in
                                   ("both", "retail") else "No")
        kwargs["tax_on_services"] = ("Yes" if args.tax_on in
                                     ("both", "services") else "No")
    if args.commission is not None:
        kwargs["service_commission"] = round(args.commission / 100.0, 4)
    if args.commission_method:
        kwargs["default_commission_method"] = args.commission_method
    if args.hourly is not None:
        kwargs["hourly_rate"] = args.hourly
    if args.burden is not None:
        kwargs["payroll_burden"] = round(args.burden / 100.0, 4)
    if args.deposit is not None:
        kwargs["deposit_percent"] = round(args.deposit / 100.0, 4)
    if args.rounding:
        kwargs["rounding_rule"] = args.rounding
    if args.target is not None:
        kwargs["monthly_revenue_target"] = args.target
        kwargs["sample_monthly_revenue_target"] = args.target
    if args.chairs is not None:
        kwargs["chairs"] = args.chairs
    if args.open_days is not None:
        kwargs["open_days"] = args.open_days
    if args.today:
        kwargs["today"] = _dt.date.fromisoformat(args.today)
    cfg = Config(**kwargs)
    if args.salon:
        cfg.salon.name = args.salon
    if args.owner:
        cfg.salon.owner = args.owner
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
    if cfg.tax_on_services == "Yes":
        tax_on.append("services")
    if cfg.tax_on_retail == "Yes":
        tax_on.append("retail")

    print(f"{cfg.product_name}")
    print(f"  listing title : {PRODUCT_LISTING_TITLE}")
    print(f"  author        : {cfg.author}")
    print(f"  sheets        : {len(SHEET_ORDER)}")
    print(f"  year          : {cfg.year}")
    print(f"  salon         : {cfg.salon.name}")
    print(f"  chairs        : {cfg.chairs} over {cfg.open_days} days "
          f"({cfg.chair_hours_week:,.0f} sellable hours a week)")
    print(f"  tax           : {cfg.tax_rate:.2%} on "
          f"{' and '.join(tax_on) if tax_on else 'nothing'}")
    print(f"  commission    : {cfg.service_commission:.0%} default, "
          f"method '{cfg.default_commission_method}'")
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

### `salon_business_spreadsheet/config.py`

*478 lines · 3 classes · 1 function*

> Configuration for the Salon Business Management Spreadsheet.
>
> Everything a salon owner may want to re-brand lives here, so the
> workbook can be regenerated for a different salon, currency, country or
> commission model without touching any of the sheet code.
>
> Author of the generated workbooks: Novality Store.

| Defines | What it does |
|---------|--------------|
| `rid(kind, number)` | Build the ID the workbook itself will generate on row `number` |
| `class Palette` |  |
| `class SalonDefaults` | Sample values written into the business profile on Settings |
| `class Config` | Everything the generator needs to know to write one workbook |

**Module constants:** `AUTHOR` ('Novality Store'), `PRODUCT_NAME` ('Salon Business Management Spreadsheet'), `PRODUCT_SUBTITLE` ('Appointments, Clients, Services, Sales, Staff Commission...), `PRODUCT_VERSION` ('1.0.0'), `PRODUCT_LISTING_TITLE` ('Salon Business Management Spreadsheet, Appointment and C...), `FORMULA_PASSWORD` ('premium'), `DISCLAIMER` ("This workbook is a business organising tool. It is not a...), `SERVICE_CATEGORIES` (17 items), `BOOKING_SOURCES` (13 items), `APPOINTMENT_STATUSES` (8 items), `FINAL_STATUSES` (3 items), `PAYMENT_METHODS` (8 items), `PAYMENT_KINDS` (8 items), `COMMISSION_METHODS` (8 items) and 22 more.

````python
"""
Configuration for the Salon Business Management Spreadsheet.

Everything a salon owner may want to re-brand lives here, so the
workbook can be regenerated for a different salon, currency, country or
commission model without touching any of the sheet code.

Author of the generated workbooks: Novality Store.
"""

from __future__ import annotations

from dataclasses import dataclass, field, replace
from datetime import date


# ---------------------------------------------------------------------------
# Product / branding constants
# ---------------------------------------------------------------------------

AUTHOR = "Novality Store"
PRODUCT_NAME = "Salon Business Management Spreadsheet"
PRODUCT_SUBTITLE = (
    "Appointments, Clients, Services, Sales, Staff Commission, "
    "Inventory, Expenses and Profit in one workbook"
)
PRODUCT_VERSION = "1.0.0"

#: Title suggested for the shop listing.
PRODUCT_LISTING_TITLE = (
    "Salon Business Management Spreadsheet, Appointment and Client "
    "Tracker, Staff Commission Calculator, Salon Income Expense and "
    "Profit Template"
)

#: Password that protects every sheet so formula cells cannot be edited
#: by accident.  Input cells stay unlocked, formula cells are locked.
FORMULA_PASSWORD = "premium"

#: The workbook is a business organiser, not an accountant.  This line
#: appears on Start Here, on the dashboard and in the help tab.
DISCLAIMER = (
    "This workbook is a business organising tool. It is not accounting, "
    "payroll, legal or tax advice, and it does not file anything for "
    "you. Tax, commission and payroll settings are yours to set and "
    "check against your own country's rules, and nothing here replaces "
    "an accountant or a booking system that sends confirmations."
)

#: Service categories offered in the drop downs - broad enough for hair,
#: nails, lashes, skin and barbering.
SERVICE_CATEGORIES = [
    "Hair - cutting", "Hair - colour", "Hair - treatment",
    "Hair - styling", "Barbering", "Nails - manicure",
    "Nails - pedicure", "Nails - extensions", "Lashes", "Brows",
    "Waxing", "Facials & skin", "Massage & body", "Make-up",
    "Tanning", "Package", "Other service",
]

#: Where a booking came from.
BOOKING_SOURCES = [
    "Repeat client", "Walk-in", "Instagram", "Facebook", "Google",
    "TikTok", "Referral", "Booking site", "Phone", "Website",
    "Gift certificate", "Promotion", "Other",
]

#: An appointment moves through these, in this order.
APPOINTMENT_STATUSES = [
    "Enquiry", "Booked", "Confirmed", "Arrived", "Completed",
    "Cancelled", "No show", "Rescheduled",
]

#: Statuses that mean the appointment reached a final answer - the
#: completion and cancellation rates divide by these only.
FINAL_STATUSES = ["Completed", "Cancelled", "No show"]

#: How a client pays.
PAYMENT_METHODS = [
    "Cash", "Card", "Bank transfer", "Mobile wallet", "Gift certificate",
    "Package / prepaid", "Online deposit", "Other",
]

#: What a payment is for - keeps deposits out of the revenue figures
#: until the appointment they belong to is completed.
PAYMENT_KINDS = [
    "Deposit", "Balance", "Full payment", "Retail only", "Refund",
    "Gift certificate sale", "Package sale", "Tip",
]

#: How staff are paid.  Every one of these is supported by the
#: commission tab - the owner picks per person.
COMMISSION_METHODS = [
    "None", "Percent of service", "Percent of service and retail",
    "Percent after product cost", "Tiered on monthly service revenue",
    "Flat amount per service", "Hourly only", "Rent a chair",
]

#: Where a team member is up to.
STAFF_STATUSES = ["Active", "On leave", "Trial", "Inactive", "Left"]

#: Employment types.
STAFF_TYPES = ["Employed", "Self-employed", "Chair renter", "Apprentice",
               "Owner", "Contractor"]

#: Expense categories, in the order a profit and loss statement wants
#: them.
EXPENSE_CATEGORIES = [
    "Rent", "Utilities", "Product & stock", "Backbar supplies",
    "Salaries & wages", "Commission", "Payroll taxes", "Insurance",
    "Equipment", "Repairs & maintenance", "Marketing & advertising",
    "Software & subscriptions", "Card processing fees", "Laundry",
    "Cleaning", "Training & education", "Licences & memberships",
    "Accounting & legal", "Telephone & internet", "Travel & parking",
    "Bank charges", "Loan repayment", "Other",
]

#: Product categories for the stock room.
PRODUCT_CATEGORIES = [
    "Retail - hair", "Retail - skin", "Retail - nails",
    "Retail - accessories", "Backbar - colour", "Backbar - developer",
    "Backbar - shampoo & conditioner", "Backbar - treatment",
    "Backbar - nails", "Backbar - lashes", "Backbar - waxing",
    "Disposables", "Tools & equipment", "Other stock",
]

#: Marketing channels a campaign can run on.
MARKETING_CHANNELS = [
    "Instagram", "Facebook", "TikTok", "Google Ads", "Google Business",
    "Email", "SMS", "Flyers & print", "Local press", "Referral scheme",
    "Loyalty scheme", "Influencer", "Event", "Window & signage", "Other",
]

#: Loyalty reward types.
LOYALTY_REWARDS = [
    "Free service", "Percent off", "Amount off", "Free add-on",
    "Free retail product", "Birthday treat", "Referral credit", "None",
]

#: Client status, worked out from the last visit but overridable.
CLIENT_STATUSES = ["New", "Active", "Regular", "VIP", "Lapsing",
                   "Inactive", "Archived"]

#: What a till line can be.  Keeping services and retail apart is what
#: makes the retail share, the commission split and the stock figures
#: work.
SALE_KINDS = ["Service", "Retail", "Package", "Gift certificate",
              "Course", "Other"]

#: Why an appointment fell over - useful once you have fifty of them.
CANCEL_REASONS = ["Client cancelled", "Client rescheduled",
                  "Client unwell", "No show", "Salon cancelled",
                  "Stylist unavailable", "Weather", "Double booked",
                  "Price", "Other"]

#: How often a recurring cost comes round.
FREQUENCIES = ["Weekly", "Fortnightly", "Four weekly", "Monthly",
               "Quarterly", "Twice a year", "Yearly", "One off"]

#: Pay run frequency on the payroll tab.
PAY_PERIODS = ["Weekly", "Fortnightly", "Monthly", "Four weekly",
               "One off"]

#: Purchase order progress.
PURCHASE_STATUSES = ["Draft", "Ordered", "Part received", "Received",
                     "Cancelled", "Returned"]

#: Campaign progress.
CAMPAIGN_STATUSES = ["Planned", "Running", "Finished", "Paused",
                     "Cancelled"]

#: How a client prefers to be contacted.
CONTACT_METHODS = ["Text", "Phone call", "Email", "WhatsApp",
                   "Instagram DM", "Do not contact"]

#: Stylist levels - who is allowed to take which service, and the
#: price band it sits in.
STYLIST_LEVELS = ["Any stylist", "Apprentice", "Stylist", "Senior stylist",
                  "Master stylist", "Owner", "Specialist"]

#: Service, campaign and product lifecycle.
SERVICE_STATUSES = ["Active", "New", "Seasonal", "On hold", "Discontinued"]

#: How a price is rounded when a service price is calculated.
ROUNDING_RULES = ["None", "Nearest 0.99", "Nearest 0.95", "Nearest 0.50",
                  "Nearest 1", "Nearest 5", "Up to whole 1",
                  "Up to whole 5"]

#: Tiered commission bands: monthly service revenue from, rate.  Shown
#: on the Settings tab and editable there.
COMMISSION_TIERS = [
    (0.0, 0.35),
    (2000.0, 0.40),
    (3500.0, 0.45),
    (5000.0, 0.50),
]

#: ID prefixes.  Each one is an input field on the Settings tab with a
#: defined name, so a salon can renumber its own records without
#: touching a formula: (label, default, defined name, digits).
ID_PREFIXES = [
    ("Client IDs", "CLI-", "ClientPrefix", 4),
    ("Service IDs", "SRV-", "ServicePrefix", 3),
    ("Stylist IDs", "STF-", "StaffPrefix", 3),
    ("Appointment IDs", "APT-", "AppointmentPrefix", 4),
    ("Sale IDs", "TXN-", "SalePrefix", 4),
    ("Payment IDs", "PMT-", "PaymentPrefix", 4),
    ("Commission IDs", "COM-", "CommissionPrefix", 4),
    ("Pay run IDs", "RUN-", "PayrollPrefix", 4),
    ("Expense IDs", "EXP-", "ExpensePrefix", 4),
    ("Product IDs", "PRD-", "ProductPrefix", 3),
    ("Purchase order IDs", "PO-", "PurchasePrefix", 4),
    ("Campaign IDs", "CMP-", "CampaignPrefix", 3),
    ("Loyalty IDs", "LOY-", "LoyaltyPrefix", 4),
]

#: defined name -> (prefix, digits), for the demo data generator.
PREFIX_BY_NAME = {name: (prefix, digits)
                  for _label, prefix, name, digits in ID_PREFIXES}


def rid(kind: str, number: int) -> str:
    """Build the ID the workbook itself will generate on row `number`.

    ``rid("Client", 7)`` -> ``"CLI-0007"``.  The demo data has to agree
    with the ID formulas exactly, because the IDs are what tie the
    tabs together.
    """
    prefix, digits = PREFIX_BY_NAME[f"{kind}Prefix"]
    return f"{prefix}{number:0{digits}d}"


#: The twelve month names used by every monthly roll-up.
MONTH_NAMES = ["January", "February", "March", "April", "May", "June",
               "July", "August", "September", "October", "November",
               "December"]


# ---------------------------------------------------------------------------
# Colours - a salon palette: plum, rose, soft gold, warm grey
# ---------------------------------------------------------------------------

class Palette:
    ink = "#2E2430"            # near black plum - titles
    deep = "#4C2A42"           # deep plum - staff & money tabs
    primary = "#6D3B5E"        # plum headers
    primary_mid = "#9A5C86"    # section bands
    primary_soft = "#F3E6EF"   # light band / alternating rows
    accent = "#C08A5E"         # warm gold accent
    accent_soft = "#F7ECE1"
    input_bg = "#FFFFFF"       # editable cell background
    input_border = "#E2D3DD"
    calc_bg = "#FAF5F8"        # locked formula cell background
    calc_border = "#EADFE7"
    white = "#FFFFFF"
    grey_text = "#6B5C66"
    line = "#DCCBD6"
    good_bg = "#DCEBD8"
    good_text = "#2F5A2B"
    warn_bg = "#FDEDD2"
    warn_text = "#8A5A00"
    bad_bg = "#F8DAD9"
    bad_text = "#9C2A24"
    info_bg = "#E3E4F2"
    info_text = "#3A4578"
    note_bg = "#FBF7FA"


# ---------------------------------------------------------------------------
# Row capacity of every register
# ---------------------------------------------------------------------------

DEFAULT_ROWS = {
    "Start Here": 0,
    "Dashboard": 0,
    "Appointments": 1200,
    "Clients": 500,
    "Services": 200,
    "Sales": 1200,
    "Payments": 1200,
    "Staff": 40,
    "Commissions": 240,
    "Payroll": 360,
    "Expenses": 600,
    "Profit & Loss": 0,
    "Inventory": 300,
    "Purchases": 300,
    "Marketing": 100,
    "Loyalty": 500,
    "Goals & Budget": 0,
    "Analytics": 0,
    "Reports": 0,
    "Settings": 0,
}


@dataclass
class SalonDefaults:
    """Sample values written into the business profile on Settings."""

    name: str = "Gloss & Grace Salon"
    owner: str = "Nadia Ellison"
    brand: str = "Gloss & Grace"
    email: str = "hello@glossandgrace.co"
    phone: str = "(512) 555 0184"
    address: str = "118 Juniper Street, Austin, TX 78702"
    website: str = "glossandgrace.co"
    instagram: str = "@glossandgracesalon"
    booking_link: str = "glossandgrace.co/book"
    tagline: str = "Colour, cuts and care since 2019"
    tax_id: str = "TX-SAL-2019-8820"
    opening_hours: str = "Tue-Fri 9:00-19:00, Sat 9:00-17:00"


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

    # Salon profile pre-filled on the Settings tab --------------------------
    salon: SalonDefaults = field(default_factory=SalonDefaults)
    plan_year: int = 0                   # 0 = the year of `today`

    # Opening hours used by the planner and the utilisation figures ---------
    open_hour: float = 9.0
    close_hour: float = 19.0
    open_days: int = 5                   # trading days a week
    chairs: int = 4                      # stations available

    # Money -----------------------------------------------------------------
    tax_rate: float = 0.0825             # sales tax on retail, editable
    tax_on_services: str = "No"          # many places do not tax services
    tax_on_retail: str = "Yes"
    tax_registered: str = "Yes"
    card_fee_percent: float = 0.0275     # what the card machine takes
    card_fee_fixed: float = 0.25
    deposit_percent: float = 0.25        # asked for on big bookings
    deposit_threshold: float = 80.00     # bookings above this want one
    cancellation_hours: int = 24
    no_show_fee: float = 25.00
    rounding_rule: str = "Nearest 0.50"

    # Staff pay -------------------------------------------------------------
    default_commission_method: str = "Percent of service"
    service_commission: float = 0.40
    retail_commission: float = 0.10
    apprentice_commission: float = 0.15
    hourly_rate: float = 16.00           # assistant / front desk hourly
    payroll_burden: float = 0.12         # employer tax, insurance, holiday
    tip_handling: str = "Paid to staff in full"
    commission_on: str = "Net of discounts"
    chair_rent_week: float = 180.00

    # Targets the dashboard measures against --------------------------------
    monthly_revenue_target: float = 28000.0
    annual_revenue_target: float = 336000.0
    monthly_profit_target: float = 6000.0
    target_average_ticket: float = 78.00
    target_retail_percent: float = 0.12   # retail as a share of revenue
    target_rebook_rate: float = 0.60
    target_retention: float = 0.70
    target_utilisation: float = 0.75
    # the example salon is smaller than the defaults assume, so the demo
    # file carries targets it can actually be measured against
    sample_monthly_revenue_target: float = 18000.0
    sample_annual_revenue_target: float = 216000.0
    sample_monthly_profit_target: float = 2400.0

    # Stock -----------------------------------------------------------------
    retail_markup: float = 2.00          # retail price = cost x this
    default_reorder_days: int = 14
    stock_warning_weeks: float = 2.0
    product_cost_percent: float = 0.08   # backbar cost as a share of a
    #                                      service, when nothing is logged
    target_chair_rate: float = 70.00     # what an hour of chair time has
    #                                      to earn before commission
    price_review_days: int = 365         # nag when a price is older
    patch_test_days: int = 180           # colour patch test validity

    # Client care -----------------------------------------------------------
    vip_spend_year: float = 1200.00      # spend in 12m that makes a VIP
    rebook_weeks: int = 6                # the usual gap between visits
    lapsing_days: int = 90               # after this a client is lapsing
    inactive_days: int = 180
    loyalty_points_per: float = 1.0      # points per unit of spend
    loyalty_reward_at: int = 200         # points needed for a reward
    loyalty_reward_value: float = 20.00
    referral_reward: float = 10.00
    review_period_days: int = 30

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
    def labor_cost_per_hour(self) -> float:
        """An hour of paid time once employer on-costs are added."""
        return round(self.hourly_rate * (1 + self.payroll_burden), 2)

    @property
    def chair_hours_week(self) -> float:
        """Sellable hours in a week - the denominator of utilisation."""
        return round((self.close_hour - self.open_hour) * self.open_days
                     * self.chairs, 1)

    def scale_rows(self, main_rows: int) -> None:
        """Scale every register in place so Appointments holds N rows."""
        base = DEFAULT_ROWS["Appointments"]
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

### `salon_business_spreadsheet/layout.py`

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

### `salon_business_spreadsheet/formats.py`

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

### `salon_business_spreadsheet/writer.py`

*458 lines · 1 class · 8 functions*

> Sheet writing helpers shared by every tab of the workbook.
>
> :class:`SheetWriter` wraps an XlsxWriter worksheet and adds the house style:
> title band, navigation strip, section headers, KPI tiles, label/value settings
> rows and - most importantly - :meth:`SheetWriter.grid`, which turns a
> :class:`~handmade_pricing_spreadsheet.layout.SheetSpec` into a finished register
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
:class:`~handmade_pricing_spreadsheet.layout.SheetSpec` into a finished register
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
    ("Start Here", "Start Here"),
    ("Dashboard", "Dashboard"),
    ("Diary", "Appointments"),
    ("Clients", "Clients"),
    ("Services", "Services"),
    ("Sales", "Sales"),
    ("Staff", "Staff"),
    ("Profit", "Profit & Loss"),
    ("Settings", "Settings"),
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

### `salon_business_spreadsheet/names.py`

*58 lines · 1 function*

> Workbook level defined names that point at the register sheets.
>
> The drop-down lists on the `Settings` tab define their own names.
> These are the ones that have to follow live data - client IDs,
> service IDs, the team, the product list - so they are built from the
> registry once every register exists. Each one is an ``OFFSET`` range
> that grows and shrinks with the number of filled rows, which is why a
> new client appears in the appointment drop-down the moment you type
> their name.

| Defines | What it does |
|---------|--------------|
| `define_cross_sheet_names(book, reg)` | Create the live lookup names |

**Module constants:** `CROSS_SHEET_NAMES` (24 entries).

````python
"""Workbook level defined names that point at the register sheets.

The drop-down lists on the `Settings` tab define their own names.
These are the ones that have to follow live data - client IDs,
service IDs, the team, the product list - so they are built from the
registry once every register exists. Each one is an ``OFFSET`` range
that grows and shrinks with the number of filled rows, which is why a
new client appears in the appointment drop-down the moment you type
their name.
"""

from __future__ import annotations

from typing import Dict, Tuple

from .layout import Registry, dynamic_name

#: defined name -> (sheet, column key)
CROSS_SHEET_NAMES: Dict[str, Tuple[str, str]] = {
    # the spine of the workbook - every drop-down that must follow data
    "ClientList": ("Clients", "id"),
    "ClientNames": ("Clients", "name"),
    "ServiceList": ("Services", "id"),
    "ServiceNames": ("Services", "name"),
    "StaffList": ("Staff", "id"),
    "StaffNames": ("Staff", "name"),
    "ProductList": ("Inventory", "id"),
    "ProductNames": ("Inventory", "name"),
    "AppointmentIDs": ("Appointments", "id"),
    "SaleIDs": ("Sales", "id"),
    "PromoCodes": ("Marketing", "code"),
    # columns the dashboard, the analytics and the reports scan
    "ServiceCategories": ("Services", "category"),
    "ServiceRevenueCol": ("Services", "revenue"),
    "SaleDates": ("Sales", "date"),
    "SaleMonths": ("Sales", "month"),
    "SaleNet": ("Sales", "net"),
    "SaleStaff": ("Sales", "staff"),
    "SaleKinds": ("Sales", "kind"),
    "AppointmentDates": ("Appointments", "date"),
    "AppointmentStatus": ("Appointments", "status"),
    "ClientSegments": ("Clients", "segment"),
    "ExpenseMonths": ("Expenses", "month"),
    "ExpenseCategoriesCol": ("Expenses", "category"),
    "ExpenseAmounts": ("Expenses", "amount"),
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

### `salon_business_spreadsheet/build.py`

*159 lines · 3 functions*

> Workbook builder: puts the twenty tabs together in the right order.
>
> Typical use::
>
>     from salon_business_spreadsheet import build_workbook
>     build_workbook("Salon Business Management Spreadsheet.xlsx")

| Defines | What it does |
|---------|--------------|
| `build_register(book, fmts, cfg, reg, module)` | Create one register sheet from its :class:`SheetSpec` |
| `build_workbook(path, cfg)` | Write the whole product to *path* and return the path |
| `_set_properties(book, cfg)` |  |

**Module constants:** `REGISTER_MODULES` (13 items), `SHEET_ORDER` (20 items), `PAGE_BUILDERS` (7 entries), `REGISTER_BY_NAME` ({module.NAME: module for module in REGISTER_MODULES}).

````python
"""Workbook builder: puts the twenty tabs together in the right order.

Typical use::

    from salon_business_spreadsheet import build_workbook
    build_workbook("Salon Business Management Spreadsheet.xlsx")
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
from .sheets import (analytics, appointments, clients, commissions,
                     dashboard, expenses, goals, inventory, loyalty,
                     marketing, payments, payroll, profit_loss,
                     purchases, reports, sales, services, settings,
                     staff, start_here)
from .writer import SheetWriter

#: register modules, in the order their specs are registered
REGISTER_MODULES = [
    services, clients, staff, appointments, sales, payments,
    commissions, payroll, expenses, inventory, purchases, marketing,
    loyalty,
]

#: tab order of the finished workbook
SHEET_ORDER: List[str] = [
    "Start Here",
    "Dashboard",
    "Appointments",
    "Clients",
    "Services",
    "Sales",
    "Payments",
    "Staff",
    "Commissions",
    "Payroll",
    "Expenses",
    "Profit & Loss",
    "Inventory",
    "Purchases",
    "Marketing",
    "Loyalty",
    "Goals & Budget",
    "Analytics",
    "Reports",
    "Settings",
]

#: sheets that are not registers - name -> builder(book, fmts, cfg, reg)
PAGE_BUILDERS: Dict[str, Callable] = {
    "Start Here": start_here.build,
    "Dashboard": dashboard.build,
    "Profit & Loss": profit_loss.build,
    "Goals & Budget": goals.build,
    "Analytics": analytics.build,
    "Reports": reports.build,
    "Settings": settings.build,
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


def build_workbook(path: str = "Salon Business Spreadsheet.xlsx",
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

    # 3. Live lookup names (client IDs, service IDs, the team).
    define_cross_sheet_names(book, reg)

    # 4. Final touches: open on the first page.
    first = sheets[SHEET_ORDER[0]].ws
    first.activate()
    first.set_first_sheet()

    book.close()
    return path


def _set_properties(book, cfg: Config) -> None:
    book.set_properties({
        "title": cfg.product_name,
        "subject": "Salon appointments, clients, pricing, sales, staff "
                   "commission, stock and profit",
        "author": cfg.author,
        "manager": cfg.author,
        "company": cfg.author,
        "category": "Small business spreadsheet template",
        "keywords": ("salon spreadsheet, hair salon business planner, "
                     "appointment tracker, client CRM, salon pricing "
                     "calculator, stylist commission calculator, "
                     "salon profit and loss, beauty business "
                     "spreadsheet, nail tech planner, lash artist "
                     "booking tracker, spreadsheet template"),
        "comments": (f"{cfg.product_name}. Created by {cfg.author}. "
                     f"Version {cfg.version}. Formula cells are locked "
                     f"and every sheet is protected with the password "
                     f"'{cfg.password}'."),
        "status": "Final",
        "created": _dt.datetime(cfg.today.year, cfg.today.month,
                                cfg.today.day),
    })
````

### `salon_business_spreadsheet/sample_data.py`

*1,263 lines · 1 class · 7 functions*

> Demo data: ten weeks at Gloss & Grace Salon.
>
> A four-chair salon in Austin: an owner who still works the floor, a
> commission-only colourist, a stylist, a self-employed nail and lash
> tech, an apprentice, a front desk and one chair renter who bills her
> own clients. Every tab opens with numbers that hang together - the
> diary feeds the till, the till feeds the commissions, the commissions
> feed the payroll, and the payroll lands in the profit and loss.
>
> Everything is dated relative to the build date, so the workbook never
> looks stale, and `--blank` strips the lot out for your own data.

| Defines | What it does |
|---------|--------------|
| `d(base, offset)` |  |
| `t(hour, minute)` |  |
| `money(value)` |  |
| `month_key(day)` |  |
| `class Sample` | fields: services, clients, appointments, sales, payments, staff, commissions, payroll |
| `get(cfg)` | Build (and cache) the demo data for this configuration |
| `_client_pool(rng, count)` |  |
| `_build(cfg)` |  |

**Module constants:** `SEED` (20261022), `STAFF` (8 items), `SERVICES` (34 items), `SERVICE_NOTES` (4 entries), `SERVICE_WEIGHT` (34 items), `PRODUCTS` (32 items), `SUPPLIERS` (6 items), `FIRST_NAMES` (70 items), `LAST_NAMES` (40 items), `PREFERENCE_NOTES` (10 items), `TAGS` (9 items), `OPEN_WEEKDAYS` (5 items), `PER_WEEKDAY` (5 entries), `HISTORY_DAYS` (60) and 1 more.

````python
"""Demo data: ten weeks at Gloss & Grace Salon.

A four-chair salon in Austin: an owner who still works the floor, a
commission-only colourist, a stylist, a self-employed nail and lash
tech, an apprentice, a front desk and one chair renter who bills her
own clients. Every tab opens with numbers that hang together - the
diary feeds the till, the till feeds the commissions, the commissions
feed the payroll, and the payroll lands in the profit and loss.

Everything is dated relative to the build date, so the workbook never
looks stale, and `--blank` strips the lot out for your own data.
"""

from __future__ import annotations

import datetime as _dt
import random
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple

from .config import COMMISSION_TIERS, Config, rid

SEED = 20261022


def d(base: _dt.date, offset: int) -> _dt.date:
    return base + _dt.timedelta(days=offset)


def t(hour: int, minute: int = 0) -> _dt.time:
    return _dt.time(hour, minute)


def money(value: float) -> float:
    return round(value + 1e-9, 2)


def month_key(day: _dt.date) -> str:
    return f"{day.year:04d}-{day.month:02d}"


@dataclass
class Sample:
    services: List[Dict[str, Any]] = field(default_factory=list)
    clients: List[Dict[str, Any]] = field(default_factory=list)
    appointments: List[Dict[str, Any]] = field(default_factory=list)
    sales: List[Dict[str, Any]] = field(default_factory=list)
    payments: List[Dict[str, Any]] = field(default_factory=list)
    staff: List[Dict[str, Any]] = field(default_factory=list)
    commissions: List[Dict[str, Any]] = field(default_factory=list)
    payroll: List[Dict[str, Any]] = field(default_factory=list)
    expenses: List[Dict[str, Any]] = field(default_factory=list)
    inventory: List[Dict[str, Any]] = field(default_factory=list)
    purchases: List[Dict[str, Any]] = field(default_factory=list)
    marketing: List[Dict[str, Any]] = field(default_factory=list)
    loyalty: List[Dict[str, Any]] = field(default_factory=list)
    suppliers: List[Dict[str, Any]] = field(default_factory=list)


_CACHE: Dict[int, Sample] = {}


def get(cfg: Config) -> Sample:
    """Build (and cache) the demo data for this configuration."""
    key = id(cfg)
    if key not in _CACHE:
        _CACHE[key] = _build(cfg) if cfg.include_sample_data else Sample()
    return _CACHE[key]


# ---------------------------------------------------------------------------
# The team.  Between them they use every commission method the
# workbook supports, which is the point: nothing is assumed.
# ---------------------------------------------------------------------------

STAFF: List[Dict[str, Any]] = [
    dict(name="Nadia Ellison", role="Owner & master stylist",
         level="Owner", type="Owner", status="Active",
         phone="512-555-0118", email="nadia@glossandgrace.example",
         hours_week=32, days="Tue, Wed, Thu, Fri, Sat",
         method="None", rate_service=None, rate_retail=None,
         flat_rate=None, hourly=None, rent_week=None,
         target_month=5800,
         notes="Takes the profit rather than a commission - her column "
               "is the salon's."),
    dict(name="Marcus Webb", role="Senior stylist & colourist",
         level="Senior stylist", type="Employed", status="Active",
         phone="512-555-0146", email="marcus@glossandgrace.example",
         hours_week=38, days="Tue, Wed, Thu, Fri, Sat",
         method="Tiered on monthly service revenue", rate_service=None,
         rate_retail=0.10, flat_rate=None, hourly=None, rent_week=None,
         target_month=5200,
         notes="Commission only, on the tiered ladder. A strong month "
               "pays him 50%."),
    dict(name="Priya Raman", role="Stylist", level="Stylist",
         type="Employed", status="Active", phone="512-555-0172",
         email="priya@glossandgrace.example", hours_week=34,
         days="Tue, Wed, Thu, Fri, Sat",
         method="Percent of service and retail", rate_service=0.38,
         rate_retail=0.10, flat_rate=None, hourly=None, rent_week=None,
         target_month=4000,
         notes="Best retail seller in the salon - the 10% is earned."),
    dict(name="Elena Ruiz", role="Nail & lash technician",
         level="Specialist", type="Self-employed", status="Active",
         phone="512-555-0193", email="elena@glossandgrace.example",
         hours_week=30, days="Wed, Thu, Fri, Sat",
         method="Percent after product cost", rate_service=0.45,
         rate_retail=0.10, flat_rate=None, hourly=None, rent_week=None,
         target_month=3000,
         notes="Self-employed, so no employer costs. Her gels and lash "
               "trays come off the top before the 45%."),
    dict(name="Jade Cooper", role="Apprentice stylist",
         level="Apprentice", type="Apprentice", status="Active",
         phone="512-555-0121", email="jade@glossandgrace.example",
         hours_week=30, days="Tue, Wed, Thu, Fri, Sat",
         method="Flat amount per service", rate_service=None,
         rate_retail=0.05, flat_rate=8.00, hourly=13.00, rent_week=None,
         target_month=1400,
         notes="Hourly plus a flat $8 a service while she trains."),
    dict(name="Tom Fenwick", role="Front desk & salon coordinator",
         level="Any stylist", type="Employed", status="Active",
         phone="512-555-0160", email="tom@glossandgrace.example",
         hours_week=24, days="Thu, Fri, Sat",
         method="Hourly only", rate_service=None, rate_retail=0.05,
         flat_rate=None, hourly=15.00, rent_week=None,
         target_month=None,
         notes="Books, rebooks and chases the gaps. Paid by the hour."),
    dict(name="Ruby Shah", role="Stylist, part time",
         level="Stylist", type="Part time", status="Active",
         phone="512-555-0205", email="ruby@glossandgrace.example",
         hours_week=16, days="Fri, Sat",
         method="Percent of service", rate_service=0.35,
         rate_retail=0.08, flat_rate=None, hourly=None, rent_week=None,
         target_month=2200,
         notes="Two long days a week on a straight 35% of what she "
               "takes."),
    dict(name="Sofia Marchetti", role="Barber (chair renter)",
         level="Specialist", type="Chair renter", status="Active",
         phone="512-555-0188", email="sofia@glossandgrace.example",
         hours_week=28, days="Wed, Thu, Fri, Sat",
         method="Rent a chair", rate_service=None, rate_retail=None,
         flat_rate=None, hourly=None, rent_week=180.00,
         target_month=None,
         notes="Rents the fourth chair for $180 a week and bills her "
               "own clients - her bookings are in the diary at zero "
               "so the room is planned properly."),
]

(OWNER, COLOURIST, STYLIST, NAILS, APPRENTICE, DESK,
 PART_TIME, RENTER) = range(8)


# ---------------------------------------------------------------------------
# The menu
# ---------------------------------------------------------------------------
# name, category, minutes, turnaround, price, product cost, level,
# commissionable, add-on, who usually does it
SERVICES: List[Tuple] = [
    ("Cut & finish", "Hair - cutting", 45, 10, 48, 1.20,
     "Stylist", "Yes", "No", [OWNER, COLOURIST, STYLIST]),
    ("Restyle & consultation", "Hair - cutting", 60, 15, 65, 1.50,
     "Senior stylist", "Yes", "No", [OWNER, COLOURIST]),
    ("Dry cut", "Hair - cutting", 30, 5, 35, 0.60,
     "Stylist", "Yes", "No", [STYLIST, COLOURIST]),
    ("Fringe trim", "Hair - cutting", 10, 5, 8, 0.20,
     "Any stylist", "No", "Yes", [APPRENTICE, STYLIST]),
    ("Children's cut", "Hair - cutting", 25, 5, 24, 0.40,
     "Any stylist", "Yes", "No", [APPRENTICE, STYLIST]),
    ("Blow dry", "Hair - styling", 30, 10, 32, 0.90,
     "Any stylist", "Yes", "No", [APPRENTICE, STYLIST, OWNER]),
    ("Blow dry, long or thick", "Hair - styling", 45, 10, 40, 1.30,
     "Any stylist", "Yes", "No", [APPRENTICE, STYLIST]),
    ("Hair up & occasion styling", "Hair - styling", 60, 15, 68, 2.40,
     "Senior stylist", "Yes", "No", [OWNER, STYLIST]),
    ("Root tint", "Hair - colour", 60, 15, 58, 8.50,
     "Stylist", "Yes", "No", [COLOURIST, OWNER, STYLIST]),
    ("Full head colour", "Hair - colour", 90, 15, 88, 14.00,
     "Stylist", "Yes", "No", [COLOURIST, OWNER]),
    ("Half head foils", "Hair - colour", 105, 15, 102, 16.50,
     "Senior stylist", "Yes", "No", [COLOURIST, OWNER]),
    ("Full head foils", "Hair - colour", 135, 20, 138, 24.00,
     "Senior stylist", "Yes", "No", [COLOURIST, OWNER]),
    ("Balayage", "Hair - colour", 150, 20, 168, 28.00,
     "Senior stylist", "Yes", "No", [COLOURIST, OWNER]),
    ("Toner & gloss", "Hair - colour", 30, 10, 34, 6.00,
     "Stylist", "Yes", "Yes", [COLOURIST, STYLIST]),
    ("Colour correction", "Hair - colour", 240, 30, 295, 46.00,
     "Master stylist", "Yes", "No", [COLOURIST]),
    ("Bond building treatment", "Hair - treatment", 25, 5, 35, 9.00,
     "Any stylist", "Yes", "Yes", [COLOURIST, STYLIST, OWNER]),
    ("Keratin smoothing", "Hair - treatment", 120, 20, 185, 38.00,
     "Senior stylist", "Yes", "No", [COLOURIST]),
    ("Scalp treatment", "Hair - treatment", 30, 5, 42, 7.00,
     "Any stylist", "Yes", "Yes", [STYLIST, APPRENTICE]),
    ("Skin fade", "Barbering", 35, 5, 30, 0.60,
     "Specialist", "Yes", "No", [RENTER]),
    ("Beard trim & hot towel", "Barbering", 25, 5, 20, 0.80,
     "Specialist", "Yes", "Yes", [RENTER]),
    ("Gel manicure", "Nails - manicure", 45, 10, 40, 4.20,
     "Specialist", "Yes", "No", [NAILS]),
    ("Luxury pedicure", "Nails - pedicure", 60, 10, 52, 5.60,
     "Specialist", "Yes", "No", [NAILS]),
    ("Acrylic full set", "Nails - extensions", 90, 15, 68, 9.50,
     "Specialist", "Yes", "No", [NAILS]),
    ("Acrylic infills", "Nails - extensions", 60, 10, 46, 6.00,
     "Specialist", "Yes", "No", [NAILS]),
    ("Classic lash set", "Lashes", 90, 15, 88, 12.50,
     "Specialist", "Yes", "No", [NAILS]),
    ("Lash infill", "Lashes", 60, 10, 50, 7.00,
     "Specialist", "Yes", "No", [NAILS]),
    ("Lash lift & tint", "Lashes", 45, 10, 58, 8.00,
     "Specialist", "Yes", "No", [NAILS]),
    ("Brow shape & tint", "Brows", 30, 5, 28, 3.00,
     "Any stylist", "Yes", "Yes", [NAILS, APPRENTICE]),
    ("Leg wax", "Waxing", 30, 5, 32, 4.00,
     "Any stylist", "Yes", "No", [NAILS]),
    ("Facial wax", "Waxing", 15, 5, 16, 1.80,
     "Any stylist", "Yes", "Yes", [NAILS, APPRENTICE]),
    ("Express facial", "Facials & skin", 30, 10, 45, 6.50,
     "Specialist", "Yes", "No", [NAILS]),
    ("Signature facial", "Facials & skin", 60, 10, 78, 11.00,
     "Specialist", "Yes", "No", [NAILS]),
    ("Colour & cut package", "Package", 150, 20, 138, 16.00,
     "Senior stylist", "Yes", "No", [COLOURIST, OWNER]),
    ("Bridal trial", "Package", 90, 15, 95, 8.00,
     "Master stylist", "Yes", "No", [OWNER]),
]

SERVICE_NOTES = {
    "Colour correction": "Always consult and patch test first. Priced "
                         "after the consultation, never over the phone.",
    "Balayage": "Our signature service and the best earner on the menu.",
    "Fringe trim": "Free to colour clients - kept on the menu so the "
                   "time is still planned.",
    "Keratin smoothing": "Needs the extractor on and the room to "
                         "itself for two hours.",
}

#: service index -> how often it is booked (relative weight)
SERVICE_WEIGHT = [
    18, 5, 6, 7, 5, 14, 7, 3,      # cutting and styling
    11, 8, 7, 6, 6, 5, 1,          # colour
    6, 1, 3,                       # treatments
    9, 5,                          # barbering (chair renter)
    9, 5, 4, 5,                    # nails
    5, 5, 4, 7,                    # lashes and brows
    4, 4,                          # waxing
    3, 3,                          # facials
    3, 1,                          # packages
]


# ---------------------------------------------------------------------------
# The stock room
# ---------------------------------------------------------------------------
# name, brand, category, size, cost, retail, opening, retail?, supplier
PRODUCTS: List[Tuple] = [
    ("Repair shampoo", "Nova Hair", "Retail - hair", "300ml", 9.40,
     22.00, 14, "Yes", "Lumiere Wholesale"),
    ("Repair conditioner", "Nova Hair", "Retail - hair", "300ml", 9.80,
     23.00, 12, "Yes", "Lumiere Wholesale"),
    ("Colour-safe shampoo", "Nova Hair", "Retail - hair", "300ml", 10.20,
     24.00, 11, "Yes", "Lumiere Wholesale"),
    ("Blonde toning shampoo", "Nova Hair", "Retail - hair", "250ml",
     11.00, 26.00, 9, "Yes", "Lumiere Wholesale"),
    ("Leave-in bond serum", "Nova Hair", "Retail - hair", "100ml", 14.50,
     34.00, 8, "Yes", "Lumiere Wholesale"),
    ("Heat protect spray", "Studio Six", "Retail - hair", "200ml", 8.20,
     19.00, 13, "Yes", "Lumiere Wholesale"),
    ("Texture spray", "Studio Six", "Retail - hair", "200ml", 9.00,
     21.00, 10, "Yes", "Lumiere Wholesale"),
    ("Dry shampoo", "Studio Six", "Retail - hair", "200ml", 7.40, 17.00,
     12, "Yes", "Lumiere Wholesale"),
    ("Styling cream", "Studio Six", "Retail - hair", "100ml", 8.60,
     20.00, 7, "Yes", "Lumiere Wholesale"),
    ("Hair oil", "Nova Hair", "Retail - hair", "50ml", 12.00, 28.00, 9,
     "Yes", "Lumiere Wholesale"),
    ("Cuticle oil pen", "Pollen Nails", "Retail - nails", "4ml", 4.20,
     11.00, 16, "Yes", "Pollen Nail Supply"),
    ("Hand cream", "Pollen Nails", "Retail - skin", "75ml", 5.80, 14.00,
     14, "Yes", "Pollen Nail Supply"),
    ("Lash cleanser", "Lume Lash", "Retail - skin", "50ml", 6.40, 16.00,
     11, "Yes", "Lume Lash Co"),
    ("Brow pencil", "Lume Lash", "Retail - accessories", "1g", 5.20,
     13.00, 9, "Yes", "Lume Lash Co"),
    ("Silk scrunchie set", "Gloss & Grace", "Retail - accessories",
     "3 pack", 4.00, 12.00, 18, "Yes", "Austin Salon Supplies"),
    ("Detangle brush", "Studio Six", "Retail - accessories", "each",
     6.00, 16.00, 8, "Yes", "Austin Salon Supplies"),
    ("Permanent colour tube", "Nova Colour", "Backbar - colour", "60ml",
     6.80, None, 48, "No", "Lumiere Wholesale"),
    ("High lift colour tube", "Nova Colour", "Backbar - colour", "60ml",
     7.60, None, 22, "No", "Lumiere Wholesale"),
    ("Developer 20 vol", "Nova Colour", "Backbar - developer", "1L",
     7.20, None, 9, "No", "Lumiere Wholesale"),
    ("Developer 30 vol", "Nova Colour", "Backbar - developer", "1L",
     7.20, None, 6, "No", "Lumiere Wholesale"),
    ("Bleach powder", "Nova Colour", "Backbar - colour", "500g", 19.50,
     None, 7, "No", "Lumiere Wholesale"),
    ("Backbar shampoo", "Nova Hair", "Backbar - shampoo & conditioner",
     "5L", 28.00, None, 4, "No", "Lumiere Wholesale"),
    ("Backbar conditioner", "Nova Hair",
     "Backbar - shampoo & conditioner", "5L", 30.00, None, 3, "No",
     "Lumiere Wholesale"),
    ("Bond builder salon kit", "Nova Hair", "Backbar - treatment", "1L",
     46.00, None, 2, "No", "Lumiere Wholesale"),
    ("Gel polish", "Pollen Nails", "Backbar - nails", "15ml", 7.90,
     None, 26, "No", "Pollen Nail Supply"),
    ("Acrylic powder", "Pollen Nails", "Backbar - nails", "250g", 16.00,
     None, 4, "No", "Pollen Nail Supply"),
    ("Lash tray, classic", "Lume Lash", "Backbar - lashes", "16 rows",
     9.50, None, 12, "No", "Lume Lash Co"),
    ("Lash adhesive", "Lume Lash", "Backbar - lashes", "5ml", 18.00,
     None, 3, "No", "Lume Lash Co"),
    ("Strip wax", "Wax Works", "Backbar - waxing", "800ml", 11.00, None,
     5, "No", "Austin Salon Supplies"),
    ("Foils, pre-cut", "Studio Six", "Disposables", "500 sheets", 14.00,
     None, 6, "No", "Austin Salon Supplies"),
    ("Gloves, nitrile", "Studio Six", "Disposables", "box of 100", 9.00,
     None, 11, "No", "Austin Salon Supplies"),
    ("Towels, black", "Studio Six", "Disposables", "12 pack", 22.00,
     None, 4, "No", "Austin Salon Supplies"),
]

SUPPLIERS: List[Dict[str, Any]] = [
    dict(supplier="Lumiere Wholesale", contact="Dana Pryce",
         phone="512-555-0410", email="orders@lumiere.example",
         account_number="GG-4471", minimum_order=150.00, lead_days=3,
         pays_by="Card on order",
         notes="Colour and retail haircare. Free delivery over $150."),
    dict(supplier="Pollen Nail Supply", contact="Rhea Dawson",
         phone="512-555-0433", email="hello@pollennails.example",
         account_number="PN-2290", minimum_order=80.00, lead_days=2,
         pays_by="Card on order", notes="Gel, acrylic and nail retail."),
    dict(supplier="Lume Lash Co", contact="Bea Tanaka",
         phone="512-555-0456", email="trade@lumelash.example",
         account_number="LL-0917", minimum_order=60.00, lead_days=4,
         pays_by="30 days",
         notes="Lash trays, adhesive, aftercare. Adhesive is "
               "date-stamped - check it in."),
    dict(supplier="Austin Salon Supplies", contact="Will Brady",
         phone="512-555-0478", email="sales@austinsalon.example",
         account_number="ASS-1182", minimum_order=50.00, lead_days=2,
         pays_by="30 days",
         notes="Towels, foils, gloves, sundries. Cheapest on "
               "disposables."),
    dict(supplier="Wax Works", contact="Nina Alvarez",
         phone="512-555-0491", email="orders@waxworks.example",
         account_number="WW-0330", minimum_order=40.00, lead_days=5,
         pays_by="Card on order", notes="Waxing consumables only."),
    dict(supplier="Barker Print", contact="Joe Barker",
         phone="512-555-0502", email="studio@barkerprint.example",
         account_number="BP-7781", minimum_order=0.00, lead_days=7,
         pays_by="Card on order",
         notes="Loyalty cards, price lists, window vinyl."),
]


FIRST_NAMES = [
    "Amara", "Beth", "Carla", "Dana", "Elise", "Farrah", "Gemma",
    "Hana", "Imani", "Jess", "Kiara", "Lena", "Maya", "Nadine",
    "Orla", "Paige", "Quinn", "Rosa", "Sienna", "Tessa", "Ursula",
    "Vera", "Willa", "Xenia", "Yara", "Zoe", "Alice", "Bridget",
    "Chiara", "Delphine", "Erin", "Freya", "Greta", "Heidi", "Isla",
    "Juno", "Kerry", "Lucia", "Marnie", "Nova", "Opal", "Petra",
    "Raina", "Sana", "Thea", "Una", "Vita", "Wren", "Xia", "Yuki",
    "Adam", "Ben", "Caleb", "Dev", "Eli", "Felix", "Gus", "Hugo",
    "Isaac", "Jonah", "Kofi", "Liam", "Mateo", "Noah", "Omar",
    "Pablo", "Reuben", "Sam", "Theo", "Vince",
]

LAST_NAMES = [
    "Alvarez", "Brennan", "Castillo", "Doyle", "Ellison", "Fischer",
    "Garcia", "Hughes", "Iverson", "Jackson", "Keller", "Lombardi",
    "Mehta", "Nolan", "Okafor", "Pereira", "Quigley", "Rossi",
    "Sandoval", "Thorne", "Upton", "Vance", "Whitfield", "Yates",
    "Zimmer", "Barlow", "Chen", "Delgado", "Esposito", "Farrell",
    "Gallagher", "Hartley", "Ingram", "Jenkins", "Kowalski", "Lind",
    "Moreau", "Nakamura", "Osei", "Patel",
]

PREFERENCE_NOTES = [
    "Half head foils, 7/1 and 8/3, toner at the basin. Coffee, oat "
    "milk, no sugar.",
    "Hates a round brush finish - rough dry and texture spray only.",
    "Colour formula on file: 6/0 roots, 30 minutes, no heat.",
    "Books the 9am slot, always early, never stays for a treatment.",
    "Sensitive scalp - cool water, barrier cream at the hairline.",
    "Growing out a fringe. Do not offer to trim it.",
    "Gel on hands, acrylic on toes. Prefers Elena, will wait for her.",
    "Lashes: classic, 9-11mm, C curl. Allergic to the strong adhesive.",
    "Wedding in the spring - trial booked, keeping the length.",
    "Quiet appointment please. No small talk, headphones in.",
]

TAGS = ["Colour client", "Blow dry regular", "Bridal", "Student",
        "Lash client", "Nails", "Barbering", "Treatment lover", ""]


def _client_pool(rng: random.Random, count: int) -> List[str]:
    names: List[str] = []
    used = set()
    while len(names) < count:
        name = f"{rng.choice(FIRST_NAMES)} {rng.choice(LAST_NAMES)}"
        if name in used:
            continue
        used.add(name)
        names.append(name)
    return names


# ---------------------------------------------------------------------------
# The generator
# ---------------------------------------------------------------------------

OPEN_WEEKDAYS = (1, 2, 3, 4, 5)        # Tuesday to Saturday
PER_WEEKDAY = {1: 11, 2: 13, 3: 15, 4: 17, 5: 19}
HISTORY_DAYS = 60
FUTURE_DAYS = 16


def _build(cfg: Config) -> Sample:          # noqa: C901 - demo data
    rng = random.Random(SEED)
    today = cfg.today
    sample = Sample()

    # ------------------------------------------------------------- staff --
    for index, person in enumerate(STAFF):
        row = dict(person)
        row["start_date"] = d(today, -(900 - index * 120))
        sample.staff.append(row)

    staff_id = {index: rid("Staff", index + 1) for index in range(len(STAFF))}

    # ---------------------------------------------------------- services --
    for index, (name, category, minutes, buffer_, price, cost, level,
                commissionable, addon, _who) in enumerate(SERVICES):
        sample.services.append(dict(
            name=name, category=category, duration=minutes,
            buffer=buffer_, price=price, product_cost=cost,
            addon=addon, commissionable=commissionable,
            online="No" if name in ("Colour correction", "Bridal trial")
                   else "Yes",
            level=level, status="Active",
            reviewed=d(today, -rng.randint(40, 330)),
            description="", notes=SERVICE_NOTES.get(name, ""),
        ))
    service_id = {index: rid("Service", index + 1)
                  for index in range(len(SERVICES))}

    # --------------------------------------------------------- inventory --
    product_rows: List[Dict[str, Any]] = []
    for index, (name, brand, category, size, cost, retail, opening,
                is_retail, supplier) in enumerate(PRODUCTS):
        used = 0
        if category.startswith("Backbar") or category == "Disposables":
            used = rng.randint(3, 22)
        product_rows.append(dict(
            name=name, brand=brand, category=category, size=size,
            sku=f"{brand[:2].upper()}-{1000 + index * 7}",
            supplier=supplier, unit_cost=cost, retail_price=retail,
            opening=opening, used=used,
            waste=rng.choice([0, 0, 0, 1]),
            adjust=rng.choice([0, 0, 0, 0, -1, 1]),
            reorder_level=None, reorder_qty=None,
            counted_on=d(today, -rng.randint(3, 48)),
            expires=(d(today, rng.randint(20, 400))
                     if "Developer" in name or "adhesive" in name
                     or "Bleach" in name else None),
            retail_item=is_retail, notes="",
        ))
    sample.inventory = product_rows
    product_id = {index: rid("Product", index + 1)
                  for index in range(len(PRODUCTS))}
    retail_indexes = [i for i, p in enumerate(PRODUCTS) if p[7] == "Yes"]

    sample.suppliers = [dict(row) for row in SUPPLIERS]

    # ----------------------------------------------------------- clients --
    names = _client_pool(rng, 132)
    sources = (["Repeat client"] * 4 + ["Instagram"] * 7 +
               ["Referral"] * 5 + ["Walk-in"] * 4 + ["Google"] * 4 +
               ["Booking site"] * 3 + ["Facebook"] * 2 + ["Website"] * 2 +
               ["TikTok"] + ["Gift certificate"] + ["Promotion"])
    for index, name in enumerate(names):
        joined = d(today, -rng.randint(20, 1400))
        colour_client = rng.random() < 0.45
        sample.clients.append(dict(
            name=name,
            phone=f"512-555-{rng.randint(1000, 9989):04d}",
            email=(f"{name.split()[0].lower()}."
                   f"{name.split()[1].lower()}@example.com"),
            contact=rng.choice(["Text", "Text", "Email", "Phone call",
                                "WhatsApp"]),
            joined=joined,
            source=rng.choice(sources),
            referred_by="",
            staff=staff_id[rng.choice(
                [OWNER, COLOURIST, COLOURIST, STYLIST, STYLIST, NAILS,
                 NAILS, APPRENTICE])],
            birthday=_dt.date(1965 + rng.randint(0, 40),
                              rng.randint(1, 12), rng.randint(1, 28)),
            preferences=(rng.choice(PREFERENCE_NOTES)
                         if rng.random() < 0.45 else ""),
            allergies=("Patch tested - no reaction"
                       if colour_client and rng.random() < 0.3 else ""),
            patch_test=(d(today, -rng.randint(10, 260))
                        if colour_client else None),
            marketing_ok="Yes" if rng.random() < 0.82 else "No",
            photos_ok="Yes" if rng.random() < 0.6 else "No",
            tags=rng.choice(TAGS),
            set_status="Active",
            notes="",
        ))
    client_id = {index: rid("Client", index + 1)
                 for index in range(len(names))}

    # a few referrals, so the loyalty tab has something to pay out on
    for index in range(len(names)):
        if rng.random() < 0.14:
            sender = rng.randrange(len(names))
            if sender != index:
                sample.clients[index]["referred_by"] = client_id[sender]
                sample.clients[index]["source"] = "Referral"

    # ------------------------------------------------------- the campaigns
    campaigns = [
        dict(name="New client, 20% off first visit", channel="Instagram",
             offer="20% off any first service", code="HELLO20",
             start=d(today, -112), end=d(today, 30), spend=180.00,
             target_rev=3000, status="Running",
             notes="Steady trickle. Most of them rebook, which is the "
                   "only reason it pays."),
        dict(name="Refer a friend", channel="Referral scheme",
             offer="$10 credit each when a friend books",
             code="FRIEND10", start=d(today, -240), end=None,
             spend=0.00, target_rev=2400, status="Running",
             notes="Costs nothing up front and brings the best "
                   "clients. Credit tracked on the loyalty tab."),
        dict(name="Tuesday blow dry club", channel="Email",
             offer="$25 blow dry, Tuesdays only", code="TUES25",
             start=d(today, -84), end=d(today, -14), spend=40.00,
             target_rev=900, status="Finished",
             notes="Filled the quiet morning but at a thin margin. "
                   "Worth repeating only in January."),
        dict(name="Balayage season", channel="Instagram",
             offer="Free bond treatment with any balayage",
             code="BONDFREE", start=d(today, -56), end=d(today, 14),
             spend=120.00, target_rev=3600, status="Running",
             notes="The treatment costs $9 and sells a $168 service. "
                   "Best offer we run."),
        dict(name="Lash launch", channel="TikTok",
             offer="$15 off a first classic lash set", code="LASH15",
             start=d(today, -42), end=d(today, -7), spend=95.00,
             target_rev=1200, status="Finished",
             notes="Brought in eight new lash clients. Five have "
                   "rebooked for infills."),
        dict(name="Birthday treat", channel="Email",
             offer="Free bond treatment in your birthday month",
             code="BIRTHDAY", start=d(today, -300), end=None,
             spend=0.00, target_rev=600, status="Running",
             notes="Automatic from the birthday column on the "
                   "clients tab. Costs product, buys goodwill."),
        dict(name="Google local ads", channel="Google Ads",
             offer="No offer - visibility only", code="GOOGLE",
             start=d(today, -70), end=d(today, 20), spend=240.00,
             target_rev=2000, status="Running",
             notes="Expensive per click. Watch the cost per new "
                   "client before renewing."),
        dict(name="Christmas gift vouchers", channel="Window & signage",
             offer="$100 voucher for $90", code="GIFT90",
             start=d(today, 40), end=d(today, 80), spend=60.00,
             target_rev=1500, status="Planned",
             notes="Print the cards in November. Last year sold 22."),
    ]
    sample.marketing = campaigns

    # --------------------------------------------------- the diary & till --
    appointments: List[Dict[str, Any]] = []
    sales: List[Dict[str, Any]] = []
    payments: List[Dict[str, Any]] = []

    service_weight = list(SERVICE_WEIGHT)
    service_indexes = list(range(len(SERVICES)))
    regulars = list(range(0, 86))        # the ones who come back

    # who can do what
    doers: Dict[int, List[int]] = {}
    for index, row in enumerate(SERVICES):
        doers[index] = list(row[9])
        if (row[1].startswith("Hair") and row[6] != "Master stylist"
                and PART_TIME not in doers[index]):
            doers[index].append(PART_TIME)

    def pick_service() -> int:
        return rng.choices(service_indexes, weights=service_weight)[0]

    last_seen: Dict[int, _dt.date] = {}

    for offset in range(-HISTORY_DAYS, FUTURE_DAYS + 1):
        day = d(today, offset)
        if day.weekday() not in OPEN_WEEKDAYS:
            continue
        future = day > today
        planned = PER_WEEKDAY[day.weekday()]
        if future:
            planned = max(3, int(planned * (0.78 if offset <= 7 else 0.45)))
        cursors = {index: 9.0 for index in range(len(STAFF))}
        cursors[NAILS] = 10.0
        cursors[RENTER] = 10.0

        for _slot in range(planned):
            s_index = pick_service()
            who_can = [w for w in doers[s_index] if w != DESK]
            who = rng.choice(who_can)
            minutes = SERVICES[s_index][2] + SERVICES[s_index][3]
            start_hour = cursors[who]
            if start_hour + minutes / 60 > 18.5:
                continue
            cursors[who] = start_hour + minutes / 60 + rng.choice(
                [0, 0, 0.25])
            hour = int(start_hour)
            minute = int(round((start_hour - hour) * 60 / 5) * 5)
            if minute >= 60:
                hour, minute = hour + 1, 0

            # who is in the chair
            if rng.random() < 0.72:
                client_index = rng.choice(regulars)
            else:
                client_index = rng.randrange(len(names))

            price = SERVICES[s_index][4]
            renter = who == RENTER
            if future:
                status = rng.choices(
                    ["Booked", "Confirmed", "Enquiry"],
                    weights=[6, 5, 1])[0]
            else:
                status = rng.choices(
                    ["Completed", "Cancelled", "No show", "Rescheduled"],
                    weights=[88, 6, 3, 3])[0]

            source = (rng.choice(["Repeat client", "Repeat client",
                                  "Repeat client", "Phone", "Instagram",
                                  "Walk-in", "Booking site", "Referral",
                                  "Google", "Website"])
                      if client_index in last_seen
                      else rng.choice(sources))

            appointment_no = len(appointments) + 1
            appointment = dict(
                date=day, start=t(hour, minute),
                client=client_id[client_index],
                service=service_id[s_index],
                staff=staff_id[who],
                minutes=None,
                price=0 if renter else None,
                status=status,
                source=source,
                deposit_due=None,
                notes=("Chair renter - bills her own clients"
                       if renter else ""),
                cancel_reason="",
                fee_charged=None,
                logged=d(day, -rng.randint(1, 24)),
            )
            if status in ("Cancelled", "Rescheduled"):
                appointment["cancel_reason"] = rng.choice(
                    ["Client cancelled", "Client rescheduled",
                     "Client unwell", "Stylist unavailable"])
            if status == "No show":
                appointment["cancel_reason"] = "No show"
                if rng.random() < 0.4:
                    appointment["fee_charged"] = cfg.no_show_fee
            appointments.append(appointment)
            appointment_ref = rid("Appointment", appointment_no)

            if status == "Completed":
                last_seen[client_index] = day
            if status != "Completed" or renter:
                continue

            # ---------------------------------------------- the sale --
            promo = ""
            discount_pct = None
            discount_amt = None
            category = SERVICES[s_index][1]
            offers = ["HELLO20", "FRIEND10"]
            if category == "Hair - styling" and price <= 45:
                offers.append("TUES25")
            if category == "Lashes":
                offers.append("LASH15")
            if category == "Hair - colour" and price >= 85:
                offers.append("BONDFREE")
            if price >= 55:
                offers.append("BIRTHDAY")
            if rng.random() < 0.13:
                promo = rng.choice(offers)
                if promo == "HELLO20":
                    discount_pct = 0.20
                elif promo == "TUES25":
                    discount_amt = money(max(0, price - 25))
                elif promo == "LASH15":
                    discount_amt = 15.00
                elif promo == "FRIEND10":
                    discount_amt = 10.00
                elif promo == "BONDFREE":
                    discount_amt = 35.00
                elif promo == "BIRTHDAY":
                    discount_amt = 20.00
            refund = None
            if rng.random() < 0.006:
                refund = money(price * 0.5)

            sale_no = len(sales) + 1
            sales.append(dict(
                date=day, appointment=appointment_ref,
                client=client_id[client_index], staff=staff_id[who],
                kind="Service", item=service_id[s_index],
                description="", qty=None, unit_price=None,
                discount_pct=discount_pct, discount_amt=discount_amt,
                promo=promo, refund=refund,
                notes=("Refunded - colour did not take, redo booked"
                       if refund else ""),
            ))
            gross = price
            discount = money((gross * (discount_pct or 0)) +
                             (discount_amt or 0))
            net_service = money(gross - discount - (refund or 0))
            sale_ref = rid("Sale", sale_no)
            service_tax = (money(net_service * cfg.tax_rate)
                           if cfg.tax_on_services == "Yes" else 0.0)
            service_total = money(net_service + service_tax)

            # ------------------------------------------- retail with it --
            retail_total = 0.0
            retail_sale_ref: Optional[str] = None
            attach = 0.30 if who == STYLIST else 0.17
            if rng.random() < attach:
                product_index = rng.choice(retail_indexes)
                qty = 1 if rng.random() < 0.85 else 2
                retail_no = len(sales) + 1
                sales.append(dict(
                    date=day, appointment=appointment_ref,
                    client=client_id[client_index], staff=staff_id[who],
                    kind="Retail", item=product_id[product_index],
                    description="", qty=qty if qty > 1 else None,
                    unit_price=None, discount_pct=None,
                    discount_amt=None, promo="", refund=None, notes="",
                ))
                retail_net = money(PRODUCTS[product_index][5] * qty)
                retail_tax = (money(retail_net * cfg.tax_rate)
                              if cfg.tax_on_retail == "Yes" else 0.0)
                retail_total = money(retail_net + retail_tax)
                retail_sale_ref = rid("Sale", retail_no)

            # --------------------------------------------- the money in --
            banked = "No" if (today - day).days <= 4 else "Yes"
            method = rng.choices(
                ["Card", "Card", "Card", "Cash", "Mobile wallet",
                 "Bank transfer"], weights=[40, 20, 12, 14, 10, 4])[0]
            deposit = 0.0
            if price >= cfg.deposit_threshold and rng.random() < 0.75:
                deposit = money(min(price * cfg.deposit_percent,
                                    service_total))
                payments.append(dict(
                    date=d(day, -rng.randint(2, 12)),
                    client=client_id[client_index], kind="Deposit",
                    method="Online deposit", amount=deposit, sale="",
                    appointment=appointment_ref, staff="",
                    reference=f"DEP{appointment_no:04d}", banked="Yes",
                    notes="",
                ))
            owing = money(service_total - deposit)
            unpaid = rng.random() < 0.012
            if not unpaid and owing > 0:
                payments.append(dict(
                    date=day, client=client_id[client_index],
                    kind="Balance" if deposit else "Full payment",
                    method=method, amount=owing, sale=sale_ref,
                    appointment=appointment_ref, staff=staff_id[who],
                    reference="", banked=banked, notes="",
                ))
            elif unpaid:
                appointments[-1]["notes"] = (
                    "Left without paying the balance - invoice sent")
            if retail_sale_ref and not unpaid:
                payments.append(dict(
                    date=day, client=client_id[client_index],
                    kind="Retail only", method=method,
                    amount=retail_total, sale=retail_sale_ref,
                    appointment=appointment_ref, staff=staff_id[who],
                    reference="", banked=banked, notes="",
                ))
            if refund:
                payments.append(dict(
                    date=d(day, rng.randint(1, 4)),
                    client=client_id[client_index], kind="Refund",
                    method=method, amount=refund, sale=sale_ref,
                    appointment=appointment_ref, staff=staff_id[who],
                    reference="", banked="Yes",
                    notes="Goodwill refund, redo booked"))
            if rng.random() < 0.34:
                tip = money(max(2, round(net_service * rng.uniform(
                    0.08, 0.18) * 2) / 2))
                payments.append(dict(
                    date=day, client=client_id[client_index], kind="Tip",
                    method=method, amount=tip, sale="",
                    appointment=appointment_ref, staff=staff_id[who],
                    reference="", banked=banked, notes=""))

    # -------------------------------------- chair rent, vouchers, courses --
    rent_day = today
    while rent_day.weekday() != 4:
        rent_day = d(rent_day, -1)
    for week in range(10):
        day = d(rent_day, -7 * week)
        sale_no = len(sales) + 1
        sales.append(dict(
            date=day, appointment="", client="", staff="", kind="Other",
            item="", description="Chair rent - Sofia Marchetti",
            qty=None, unit_price=cfg.chair_rent_week, discount_pct=None,
            discount_amt=None, promo="", refund=None,
            notes="Weekly chair rent. Income to the salon, not a "
                  "commission-earning sale."))
        payments.append(dict(
            date=day, client="", kind="Full payment",
            method="Bank transfer", amount=money(cfg.chair_rent_week),
            sale=rid("Sale", sale_no), appointment="", staff="",
            reference="RENT", banked="Yes", notes=""))

    for index in range(7):
        day = d(today, -rng.randint(4, 64))
        client_index = rng.randrange(len(names))
        value = rng.choice([50, 75, 100, 100, 150])
        sale_no = len(sales) + 1
        sales.append(dict(
            date=day, appointment="", client=client_id[client_index],
            staff=staff_id[DESK], kind="Gift certificate", item="",
            description=f"Gift certificate, ${value}", qty=None,
            unit_price=float(value), discount_pct=None,
            discount_amt=None, promo="", refund=None,
            notes="Sold at the desk. Counted as income when sold; "
                  "redeem it as a payment method later."))
        payments.append(dict(
            date=day, client=client_id[client_index],
            kind="Gift certificate sale", method="Card",
            amount=float(value), sale=rid("Sale", sale_no),
            appointment="", staff=staff_id[DESK], reference="",
            banked="Yes", notes=""))

    for index in range(4):
        day = d(today, -rng.randint(10, 60))
        client_index = rng.choice(regulars)
        value = rng.choice([180, 240, 300])
        sale_no = len(sales) + 1
        sales.append(dict(
            date=day, appointment="", client=client_id[client_index],
            staff=staff_id[OWNER], kind="Package", item="",
            description="Course of six blow dries", qty=None,
            unit_price=float(value), discount_pct=None,
            discount_amt=None, promo="", refund=None,
            notes="Paid up front, redeemed visit by visit."))
        payments.append(dict(
            date=day, client=client_id[client_index], kind="Package sale",
            method="Card", amount=float(value), sale=rid("Sale", sale_no),
            appointment="", staff=staff_id[OWNER], reference="",
            banked="Yes", notes=""))

    sample.appointments = appointments
    sample.sales = sales
    sample.payments = payments

    # ------------------------------------------------------- commissions --
    months = sorted({month_key(row["date"]) for row in sales})[-3:]
    totals: Dict[Tuple[str, int], Dict[str, float]] = {}
    for row in sales:
        key = month_key(row["date"])
        who = row["staff"]
        if not who:
            continue
        index = int(who.split("-")[1]) - 1
        bucket = totals.setdefault((key, index), dict(
            service=0.0, retail=0.0, cost=0.0, count=0))
        price = 0.0
        cost = 0.0
        if row["kind"] == "Service":
            s_index = int(row["item"].split("-")[1]) - 1
            price = SERVICES[s_index][4]
            cost = SERVICES[s_index][5]
        elif row["kind"] == "Retail":
            p_index = int(row["item"].split("-")[1]) - 1
            qty = row["qty"] or 1
            price = PRODUCTS[p_index][5] * qty
            cost = PRODUCTS[p_index][4] * qty
        else:
            price = row["unit_price"] or 0.0
        discount = money((price * (row["discount_pct"] or 0)) +
                         (row["discount_amt"] or 0))
        net = money(price - discount - (row["refund"] or 0))
        if row["kind"] == "Service":
            bucket["service"] += net
            bucket["cost"] += money(cost)
            bucket["count"] += 1
        elif row["kind"] == "Retail":
            bucket["retail"] += net

    def tier_rate(base: float) -> float:
        rate = COMMISSION_TIERS[0][1]
        for floor, value in COMMISSION_TIERS:
            if base >= floor:
                rate = value
        return rate

    commission_due: Dict[Tuple[str, int], float] = {}
    for key in months:
        for index, person in enumerate(STAFF):
            bucket = totals.get((key, index))
            if not bucket:
                continue
            method = person["method"]
            service = money(bucket["service"])
            retail = money(bucket["retail"])
            cost = money(bucket["cost"])
            rate = person["rate_service"] or cfg.service_commission
            retail_rate = person["rate_retail"] or cfg.retail_commission
            if method == "Percent of service":
                value = money(service * rate)
            elif method == "Percent of service and retail":
                value = money(money(service * rate) +
                              money(retail * retail_rate))
            elif method == "Percent after product cost":
                value = money(max(0.0, service - cost) * rate)
            elif method == "Tiered on monthly service revenue":
                value = money(service * tier_rate(service))
            elif method == "Flat amount per service":
                value = money(bucket["count"] * (person["flat_rate"] or 0))
            else:
                value = 0.0
            if value <= 0 and method in ("None", "Hourly only",
                                         "Rent a chair"):
                continue
            adjust = None
            if rng.random() < 0.12:
                adjust = rng.choice([-25.0, -15.0, 20.0, 35.0])
            current = key == month_key(today)
            paid_day = None
            if not current:
                year, month = (int(part) for part in key.split("-"))
                last = _dt.date(year + (month == 12),
                                1 if month == 12 else month + 1, 1)
                paid_day = d(last, -1)
            commission_due[(key, index)] = money(
                value + (adjust or 0))
            sample.commissions.append(dict(
                month=key, staff=staff_id[index],
                adjust=adjust, bonus=(50.0 if rng.random() < 0.08
                                      else None),
                approved="No" if current else "Yes",
                paid_on=paid_day,
                notes=("Current month, still running"
                       if current else ""),
            ))

    # ----------------------------------------------------------- payroll --
    tips_by: Dict[Tuple[_dt.date, int], float] = {}
    for row in payments:
        if row["kind"] != "Tip" or not row["staff"]:
            continue
        index = int(row["staff"].split("-")[1]) - 1
        week_end = row["date"]
        while week_end.weekday() != 4:
            week_end = d(week_end, 1)
        tips_by[(week_end, index)] = money(
            tips_by.get((week_end, index), 0.0) + row["amount"])

    friday = today
    while friday.weekday() != 4:
        friday = d(friday, -1)
    weekly_people = [(APPRENTICE, 30), (DESK, 24)]
    for week in range(10):
        period_end = d(friday, -7 * week)
        for index, hours in weekly_people:
            person = STAFF[index]
            worked = hours + rng.choice([-4, -2, 0, 0, 2, 3])
            commission = None
            if index == APPRENTICE and period_end.day >= 22:
                key = month_key(period_end)
                commission = commission_due.get((key, index))
            sample.payroll.append(dict(
                period_end=period_end, period="Weekly",
                staff=staff_id[index], hours=worked,
                rate=person["hourly"], commission=commission,
                tips=tips_by.get((period_end, index)),
                bonus=None, holiday=None,
                deductions=(12.0 if rng.random() < 0.08 else None),
                paid="Yes", paid_on=period_end, method="Bank transfer",
                notes="",
            ))

    for key in months:
        year, month = (int(part) for part in key.split("-"))
        last_day = _dt.date(year + (month == 12),
                            1 if month == 12 else month + 1, 1)
        period_end = min(d(last_day, -1), today)
        current = key == month_key(today)
        for index in (COLOURIST, STYLIST, NAILS, PART_TIME):
            value = commission_due.get((key, index))
            if value is None:
                continue
            tips = money(sum(
                row["amount"] for row in payments
                if row["kind"] == "Tip" and row["staff"] == staff_id[index]
                and month_key(row["date"]) == key))
            if current:
                continue        # nothing released yet - no row to file
            sample.payroll.append(dict(
                period_end=period_end, period="Monthly",
                staff=staff_id[index], hours=None, rate=None,
                commission=value,
                tips=tips or None, bonus=None, holiday=None,
                deductions=None, paid="Yes", paid_on=period_end,
                method="Bank transfer",
                notes="Commission released at the end of the month",
            ))

    # ---------------------------------------------------------- expenses --
    recurring = [
        ("Rent", "Juniper Street Holdings", "Salon rent, 118 Juniper St",
         2400.00, "Monthly", "Bank transfer"),
        ("Utilities", "Lone Star Power", "Electricity and water",
         310.00, "Monthly", "Bank transfer"),
        ("Insurance", "Shield Business Cover",
         "Public liability and treatment cover", 95.00, "Monthly",
         "Card"),
        ("Software & subscriptions", "BookEase",
         "Online booking and reminders", 79.00, "Monthly", "Card"),
        ("Telephone & internet", "Austin Fibre", "Salon line and wifi",
         65.00, "Monthly", "Bank transfer"),
        ("Accounting & legal", "Reed & Co Accountants",
         "Monthly bookkeeping", 120.00, "Monthly", "Bank transfer"),
        ("Loan repayment", "Capital Union Bank", "Fit-out loan",
         420.00, "Monthly", "Bank transfer"),
        ("Marketing & advertising", "Meta & Google", "Local ads",
         180.00, "Monthly", "Card"),
        ("Bank charges", "Capital Union Bank", "Business account fee",
         18.00, "Monthly", "Bank transfer"),
    ]
    weekly = [
        ("Laundry", "Fresh Fold Laundry", "Towels and gowns", 38.00),
        ("Cleaning", "Bright Spark Cleaning", "Evening clean, 3 nights",
         55.00),
        ("Other", "Juniper Street Market", "Coffee, milk, biscuits",
         24.00),
    ]
    for months_back in range(3):
        base_month = _dt.date(today.year, today.month, 1)
        for _ in range(months_back):
            base_month = d(base_month, -1).replace(day=1)
        for category, payee, description, amount, frequency, method \
                in recurring:
            day = min(base_month.replace(
                day=min(28, 3 + (hash(category) % 20))), today)
            paid = day <= today
            sample.expenses.append(dict(
                date=day, category=category, payee=payee,
                description=description,
                amount=amount + (rng.choice([0, 0, 0, 8, -6])
                                 if category == "Utilities" else 0),
                tax=None, method=method, recurring="Yes",
                frequency=frequency, due=day,
                paid="Yes" if paid else "No",
                paid_on=day if paid else None,
                reference=f"{category[:3].upper()}-{day:%y%m}",
                notes="",
            ))
    week_start = today
    while week_start.weekday() != 0:
        week_start = d(week_start, -1)
    for week in range(10):
        day = d(week_start, -7 * week)
        for category, payee, description, amount in weekly:
            sample.expenses.append(dict(
                date=day, category=category, payee=payee,
                description=description, amount=amount, tax=None,
                method="Card", recurring="Yes", frequency="Weekly",
                due=day, paid="Yes", paid_on=day,
                reference="", notes="",
            ))
    one_offs = [
        ("Training & education", "Nova Colour Academy",
         "Balayage masterclass, two stylists", 340.00, -38),
        ("Equipment", "Salon Fit Direct", "Replacement backwash tap",
         185.00, -52),
        ("Repairs & maintenance", "Ace Electrical",
         "Lighting repair, colour bar", 120.00, -21),
        ("Licences & memberships", "City of Austin",
         "Salon licence renewal", 240.00, -64),
        ("Card processing fees", "SwiftPay",
         "Card processing, last month", 318.00, -12),
        ("Card processing fees", "SwiftPay",
         "Card processing, month before", 296.00, -42),
        ("Marketing & advertising", "Barker Print",
         "Loyalty cards and price lists", 142.00, -30),
        ("Equipment", "Salon Fit Direct", "Second lash bed", 760.00, -7),
        ("Training & education", "Lume Lash Co",
         "Lash lift certification", 295.00, -25),
        ("Other", "Blooms on Fifth", "Reception flowers, monthly",
         45.00, -9),
    ]
    for category, payee, description, amount, offset in one_offs:
        day = d(today, offset)
        sample.expenses.append(dict(
            date=day, category=category, payee=payee,
            description=description, amount=amount, tax=None,
            method=rng.choice(["Card", "Bank transfer"]),
            recurring="No", frequency="One off", due=day,
            paid="Yes", paid_on=day, reference="", notes="",
        ))
    # two that are still waiting, so the chase list is not empty
    sample.expenses.append(dict(
        date=d(today, -5), category="Training & education",
        payee="Nova Colour Academy",
        description="Autumn colour course, two places",
        amount=486.00, tax=None, method="Bank transfer", recurring="No",
        frequency="One off", due=d(today, 9), paid="No", paid_on=None,
        reference="NCA-8841",
        notes="Booked, invoice due in nine days."))
    sample.expenses.append(dict(
        date=d(today, -16), category="Repairs & maintenance",
        payee="Ace Electrical", description="Emergency callout, dryer",
        amount=95.00, tax=None, method="Bank transfer", recurring="No",
        frequency="One off", due=d(today, -2), paid="No", paid_on=None,
        reference="AE-2214", notes="Chase - invoice is past due."))

    # --------------------------------------------------------- purchases --
    order_plan = [
        (0, 6, "Lumiere Wholesale", -62, "Received"),
        (16, 12, "Lumiere Wholesale", -62, "Received"),
        (18, 4, "Lumiere Wholesale", -62, "Received"),
        (20, 2, "Lumiere Wholesale", -62, "Received"),
        (1, 6, "Lumiere Wholesale", -48, "Received"),
        (4, 6, "Lumiere Wholesale", -48, "Received"),
        (24, 10, "Pollen Nail Supply", -45, "Received"),
        (25, 2, "Pollen Nail Supply", -45, "Received"),
        (26, 6, "Lume Lash Co", -40, "Received"),
        (27, 2, "Lume Lash Co", -40, "Received"),
        (2, 6, "Lumiere Wholesale", -34, "Received"),
        (5, 6, "Lumiere Wholesale", -34, "Received"),
        (29, 4, "Austin Salon Supplies", -30, "Received"),
        (30, 6, "Austin Salon Supplies", -30, "Received"),
        (31, 4, "Austin Salon Supplies", -30, "Received"),
        (17, 6, "Lumiere Wholesale", -24, "Received"),
        (19, 3, "Lumiere Wholesale", -24, "Received"),
        (21, 2, "Lumiere Wholesale", -24, "Received"),
        (3, 6, "Lumiere Wholesale", -18, "Received"),
        (9, 6, "Lumiere Wholesale", -18, "Received"),
        (28, 3, "Austin Salon Supplies", -16, "Received"),
        (10, 12, "Pollen Nail Supply", -12, "Received"),
        (11, 10, "Pollen Nail Supply", -12, "Received"),
        (22, 2, "Lumiere Wholesale", -9, "Part received"),
        (23, 1, "Lumiere Wholesale", -9, "Received"),
        (6, 6, "Lumiere Wholesale", -6, "Received"),
        (7, 6, "Lumiere Wholesale", -6, "Received"),
        (12, 6, "Lume Lash Co", -4, "Ordered"),
        (13, 6, "Lume Lash Co", -4, "Ordered"),
        (14, 12, "Austin Salon Supplies", -2, "Ordered"),
        (8, 6, "Lumiere Wholesale", -1, "Draft"),
    ]
    for p_index, qty, supplier, offset, status in order_plan:
        day = d(today, offset)
        cost = PRODUCTS[p_index][4]
        bumped = rng.random() < 0.18
        received_on = None
        qty_received = None
        if status == "Received":
            received_on = d(day, rng.randint(1, 5))
        elif status == "Part received":
            received_on = d(day, 3)
            qty_received = max(1, qty - 1)
        sample.purchases.append(dict(
            date=day, supplier=supplier, product=product_id[p_index],
            qty=qty,
            unit_cost=money(cost * 1.06) if bumped else None,
            shipping=(0.00 if supplier == "Lumiere Wholesale"
                      else rng.choice([0, 6.50, 9.00])),
            tax=None, status=status,
            expected=d(day, 3), received_on=received_on,
            qty_received=qty_received,
            invoice=(f"INV-{8000 + p_index * 13 + abs(offset)}"
                     if status != "Draft" else ""),
            paid="Yes" if offset < -12 else "No",
            notes=("Price up 6% on the last order" if bumped else ""),
        ))

    # ----------------------------------------------------------- loyalty --
    spend_by_client: Dict[str, float] = {}
    for row in sales:
        if not row["client"]:
            continue
        price = 0.0
        if row["kind"] == "Service":
            price = SERVICES[int(row["item"].split("-")[1]) - 1][4]
        elif row["kind"] == "Retail":
            p_index = int(row["item"].split("-")[1]) - 1
            price = PRODUCTS[p_index][5] * (row["qty"] or 1)
        else:
            price = row["unit_price"] or 0.0
        spend_by_client[row["client"]] = money(
            spend_by_client.get(row["client"], 0.0) + price)
    ranked = sorted(spend_by_client.items(), key=lambda item: -item[1])
    for client_ref, spend in ranked[:46]:
        points = int(spend * cfg.loyalty_points_per)
        redeemed = 0
        if points > cfg.loyalty_reward_at * 1.3 and rng.random() < 0.55:
            redeemed = cfg.loyalty_reward_at
        sample.loyalty.append(dict(
            client=client_ref,
            joined=d(today, -rng.randint(30, 500)),
            card=f"GG{rng.randint(10000, 99999)}",
            bonus_points=rng.choice([None, None, None, 25, 50]),
            redeemed=redeemed or None,
            last_reward=(d(today, -rng.randint(10, 120))
                         if redeemed else None),
            reward_type=rng.choice(["Amount off", "Free add-on",
                                    "Free service", "Percent off"]),
            referral_paid=None,
            active="Yes",
            notes="",
        ))

    # ----------------------------------- make the stock room add up --
    # Opening stock is set last, from what actually moved, so no
    # product on the demo tab finishes the period at a nonsense
    # number.  Three lines are left deliberately short so the reorder
    # alerts have something to say.
    short = {"Lash adhesive": 0, "Bleach powder": 1,
             "Backbar conditioner": 1, "Blonde toning shampoo": 2}
    received_qty: Dict[str, float] = {}
    for row in sample.purchases:
        if row["status"] == "Received":
            qty = row["qty_received"] or row["qty"] or 0
        elif row["status"] == "Part received":
            qty = row["qty_received"] or 0
        else:
            qty = 0
        received_qty[row["product"]] = received_qty.get(
            row["product"], 0) + qty
    sold_qty: Dict[str, float] = {}
    for row in sample.sales:
        if row["kind"] == "Retail":
            sold_qty[row["item"]] = sold_qty.get(row["item"], 0) + (
                row["qty"] or 1)
    for index, row in enumerate(product_rows):
        ref = product_id[index]
        went_out = ((row["used"] or 0) + sold_qty.get(ref, 0) +
                    (row["waste"] or 0) - (row["adjust"] or 0))
        cover = short.get(row["name"], rng.randint(4, 11))
        needed = went_out - received_qty.get(ref, 0) + cover
        row["opening"] = max(0, int(round(max(row["opening"], needed))))

    return sample
````

---

## 3. The diary and the client book

The menu with its minutes and its product cost, the client record that remembers the formula and the patch test, and the appointment planner with the double-booking alarm. The shared helpers in `common.py` are where the alert colours, the lookups and the rounding rules live.

### `salon_business_spreadsheet/sheets/__init__.py`

*0 lines*

````python

````

### `salon_business_spreadsheet/sheets/common.py`

*341 lines · 30 functions*

> Small building blocks shared by the register sheets.

| Defines | What it does |
|---------|--------------|
| `id_col(trigger, prefix, digits, width, header, note)` | Auto numbered ID column (blank until the trigger column is filled) |
| `month_col(date_key, key, header, width)` | `yyyy-mm` text column - the key every monthly roll-up groups on |
| `weekday_col(date_key, key, header, width)` |  |
| `week_col(date_key, key, header, width)` | Week number of the year |
| `alert_cf()` | Traffic light colouring for the 'Alert' / 'Check' columns |
| `status_cf()` | Colouring for the booking, payment, stock and staff statuses |
| `bar_cf(color)` |  |
| `scale_cf()` |  |
| `money_cf()` |  |
| `loss_cf()` | Red when a profit column goes negative |
| `owing_cf()` | Amber on anything still owed, green on a cleared balance |
| `dup_cf()` |  |
| `flag_yes_no_cf()` |  |
| `target_cf(target_name)` | Green at or above a named target, amber below it |
| `yesno(note)` |  |
| `score_validation(maximum)` |  |
| `blank(trigger, body)` | `=IF(trigger="","",body)` with the standard row placeholders |
| `lookup(sheet, target, key_col, key, missing)` | Generic `INDEX/MATCH` against any register, blank row safe |
| `lookup_client(target, key, missing)` |  |
| `lookup_service(target, key, missing)` |  |
| `lookup_staff(target, key, missing)` |  |
| `lookup_product(target, key, missing)` |  |
| `lookup_appointment(target, key, missing)` |  |
| `round_price(cell, rule)` | Apply the rounding rule chosen on Settings to a price `cell` |
| `rank_desc(key, guard, header, width, note)` | Hidden dense rank, 1 = the biggest value in `key` |
| `rank_asc(key, guard, header, width, note)` | Hidden dense rank, 1 = the smallest value in `key` |
| `min_number(value_range, cond, cap)` | Smallest positive number in `value_range` where `cond` holds |
| `seq_key(parent, trigger)` | Hidden `1, 2, 3 ...` within a parent, plus a `parent|n` key |
| `tax_on(amount, kind_is_retail)` | `amount` x the tax rate, if that kind of line is taxable |
| `card_fee(amount, method_cell)` |  |

**Module constants:** `ROW_OFFSET` (HEADER_ROW), `ROUNDING_BRANCHES` (8 items).

````python
"""Small building blocks shared by the register sheets."""

from __future__ import annotations

from typing import Any, Dict, List

from ..layout import HEADER_ROW, Col
from ..writer import cf_text, yes_no

# Everything below assumes the standard geometry: headers on row 4, data from
# row 5 - ROW()-4 therefore gives a 1 based record number.
ROW_OFFSET = HEADER_ROW


def id_col(trigger: str, prefix: str, digits: int = 4, width: float = 11,
           header: str = "ID", note: str = "") -> Col:
    """Auto numbered ID column (blank until the trigger column is filled).

    `prefix` is a formula fragment - either a defined name such as
    ``ClientPrefix`` or a quoted literal such as ``'"APT-"'``.
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
              width: float = 10) -> Col:
    """`yyyy-mm` text column - the key every monthly roll-up groups on."""
    return Col(key, header, width, "text",
               formula=(f'=IF({{{date_key}}}{{r}}="","",'
                        f'TEXT({{{date_key}}}{{r}},"yyyy-mm"))'),
               hidden=True,
               note="Automatic. Every monthly report groups on this.")


def weekday_col(date_key: str, key: str = "weekday", header: str = "Day",
                width: float = 8) -> Col:
    return Col(key, header, width, "text",
               formula=(f'=IF({{{date_key}}}{{r}}="","",'
                        f'TEXT({{{date_key}}}{{r}},"ddd"))'))


def week_col(date_key: str, key: str = "week", header: str = "Week",
             width: float = 7) -> Col:
    """Week number of the year."""
    return Col(key, header, width, "int",
               formula=(f'=IF({{{date_key}}}{{r}}="","",'
                        f'WEEKNUM({{{date_key}}}{{r}},2))'),
               hidden=True)


def alert_cf() -> List[Dict[str, Any]]:
    """Traffic light colouring for the 'Alert' / 'Check' columns."""
    return [
        cf_text("OK", "good", "begins with"),
        cf_text("Overdue", "bad"),
        cf_text("Late", "bad"),
        cf_text("Missing", "bad"),
        cf_text("Clash", "bad"),
        cf_text("Conflict", "bad"),
        cf_text("Duplicate", "bad"),
        cf_text("Unpaid", "bad"),
        cf_text("Loss", "bad"),
        cf_text("Below", "bad"),
        cf_text("Out of stock", "bad"),
        cf_text("No show", "bad"),
        cf_text("Action", "warn"),
        cf_text("Due", "warn"),
        cf_text("Check", "warn"),
        cf_text("Chase", "warn"),
        cf_text("Assign", "warn"),
        cf_text("Confirm", "warn"),
        cf_text("Order", "warn"),
        cf_text("Rebook", "warn"),
        cf_text("Add", "warn"),
        cf_text("Today", "info"),
        cf_text("Tomorrow", "info"),
        cf_text("Upcoming", "info"),
    ]


def status_cf() -> List[Dict[str, Any]]:
    """Colouring for the booking, payment, stock and staff statuses."""
    return [
        cf_text("Completed", "good"),
        cf_text("Complete", "good"),
        cf_text("Paid", "good"),
        cf_text("Settled", "good"),
        cf_text("Arrived", "good"),
        cf_text("Active", "good"),
        cf_text("Regular", "good"),
        cf_text("VIP", "good"),
        cf_text("In stock", "good"),
        cf_text("Received", "good"),
        cf_text("Approved", "good"),
        cf_text("Yes", "good"),
        cf_text("Booked", "info"),
        cf_text("Confirmed", "info"),
        cf_text("Scheduled", "info"),
        cf_text("New", "info"),
        cf_text("Ordered", "info"),
        cf_text("Running", "info"),
        cf_text("Part paid", "warn"),
        cf_text("Pending", "warn"),
        cf_text("Enquiry", "warn"),
        cf_text("Rescheduled", "warn"),
        cf_text("Low stock", "warn"),
        cf_text("Lapsing", "warn"),
        cf_text("Planned", "warn"),
        cf_text("On hold", "warn"),
        cf_text("Draft", "warn"),
        cf_text("Cancelled", "bad"),
        cf_text("No show", "bad"),
        cf_text("Refunded", "bad"),
        cf_text("Out of stock", "bad"),
        cf_text("Overdue", "bad"),
        cf_text("Written off", "bad"),
        cf_text("Inactive", "muted"),
        cf_text("Archived", "muted"),
        cf_text("Finished", "muted"),
        cf_text("Discontinued", "muted"),
    ]


def bar_cf(color: str = "#9A5C86") -> List[Dict[str, Any]]:
    return [{"type": "data_bar", "bar_color": color}]


def scale_cf() -> List[Dict[str, Any]]:
    return [{"type": "3_color_scale", "min_color": "#F8DAD9",
             "mid_color": "#FDEDD2", "max_color": "#DCEBD8"}]


def money_cf() -> List[Dict[str, Any]]:
    return [{"type": "cell", "criteria": ">", "value": 0, "fmt": "bold_good"}]


def loss_cf() -> List[Dict[str, Any]]:
    """Red when a profit column goes negative."""
    return [{"type": "cell", "criteria": "<", "value": 0, "fmt": "bold_bad"},
            {"type": "cell", "criteria": ">", "value": 0, "fmt": "good"}]


def owing_cf() -> List[Dict[str, Any]]:
    """Amber on anything still owed, green on a cleared balance."""
    return [{"type": "cell", "criteria": ">", "value": 0.004, "fmt": "warn"},
            {"type": "cell", "criteria": "<", "value": -0.004, "fmt": "info"}]


def dup_cf() -> List[Dict[str, Any]]:
    return [cf_text("Duplicate", "bad"), cf_text("OK", "good")]


def flag_yes_no_cf() -> List[Dict[str, Any]]:
    return [cf_text("No", "warn"), cf_text("Yes", "good")]


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


def lookup_client(target: str, key: str = "client",
                  missing: str = "Not in Clients") -> str:
    return lookup("Clients", target, "id", key, missing)


def lookup_service(target: str, key: str = "service",
                   missing: str = "Not in Services") -> str:
    return lookup("Services", target, "id", key, missing)


def lookup_staff(target: str, key: str = "staff",
                 missing: str = "Not in Staff") -> str:
    return lookup("Staff", target, "id", key, missing)


def lookup_product(target: str, key: str = "product",
                   missing: str = "Not in Inventory") -> str:
    return lookup("Inventory", target, "id", key, missing)


def lookup_appointment(target: str, key: str = "appointment",
                       missing: str = "Not in Appointments") -> str:
    return lookup("Appointments", target, "id", key, missing)


#: how each rounding rule on the Settings tab is applied.  Every branch
#: rounds *up* to the nearest charm price so a rounded price can never
#: drop below the price the maths asked for.
ROUNDING_BRANCHES = [
    ('"None"', "{x}"),
    ('"Nearest 0.99"', 'MAX(0.99,CEILING({x},1)-0.01)'),
    ('"Nearest 0.95"', 'MAX(0.95,CEILING({x},1)-0.05)'),
    ('"Nearest 0.50"', 'CEILING({x},0.5)'),
    ('"Nearest 1"', 'CEILING({x},1)'),
    ('"Nearest 5"', 'CEILING({x},5)'),
    ('"Up to whole 1"', 'CEILING({x},1)'),
    ('"Up to whole 5"', 'CEILING({x},5)'),
]


def round_price(cell: str, rule: str = "RoundingRule") -> str:
    """Apply the rounding rule chosen on Settings to a price `cell`.

    Pass a single cell reference, not an expression - the reference is
    repeated once per rule, and a long expression would be copied eight
    times into every row of the service menu.
    """
    out = f"ROUND({cell},2)"
    for test, body in reversed(ROUNDING_BRANCHES):
        out = f'IF({rule}={test},{body.format(x=cell)},{out})'
    return out


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
    to pull the nth line of one order without an array formula.
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

#: Sales tax on a line, honouring the two switches on Settings: tax is
#: added to retail, to services, to both or to neither.
def tax_on(amount: str, kind_is_retail: str) -> str:
    """`amount` x the tax rate, if that kind of line is taxable."""
    return (f'ROUND({amount}*IF({kind_is_retail},'
            f'IF(TaxOnRetail="Yes",TaxRate,0),'
            f'IF(TaxOnServices="Yes",TaxRate,0)),2)')


#: What the card machine keeps out of a payment.
def card_fee(amount: str, method_cell: str) -> str:
    return (f'IF({method_cell}="Card",ROUND({amount}*CardFeePercent+'
            f'CardFeeFixed,2),0)')
````

### `salon_business_spreadsheet/sheets/services.py`

*276 lines · 2 functions*

> `Services` - the menu, with what each treatment really earns.
>
> A salon price list is usually a list of numbers somebody typed once.
> This one costs every service: the backbar product it burns, the
> commission it pays, the chair time it occupies - and then says what it
> makes an hour, which is the only number that compares a $35 blow dry
> with a $220 balayage honestly.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Menu summary under the grid |

**Module constants:** `NAME` ('Services').

````python
"""`Services` - the menu, with what each treatment really earns.

A salon price list is usually a list of numbers somebody typed once.
This one costs every service: the backbar product it burns, the
commission it pays, the chair time it occupies - and then says what it
makes an hour, which is the only number that compares a $35 blow dry
with a $220 balayage honestly.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import list_from, number_check, date_check
from .common import (alert_cf, blank, id_col, loss_cf, rank_desc, round_price,
                     status_cf, yesno)

NAME = "Services"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "ServicePrefix", digits=3, width=10,
               note="Automatic. Bookings and sales point at this ID, so "
                    "renaming a service never breaks your history."),
        Col("name", "Service", 30, "text",
            note="Exactly as it appears on your menu and on the booking "
                 "page."),
        Col("category", "Category", 22, "text",
            validate=list_from("ServiceCategoryList"),
            note="Drives the revenue-by-category figures on the "
                 "dashboard and the analytics tab."),
        Col("description", "Description", 34, "wrap",
            note="What the client is buying. Used on the price list "
                 "print-out."),
        Col("duration", "Minutes", 9, "int",
            validate=number_check(0, 1440, decimal=False),
            note="Chair time, start to finish. Include the time you "
                 "cannot sell to anybody else."),
        Col("buffer", "Turnaround", 10, "int",
            validate=number_check(0, 240, decimal=False),
            note="Clean down and reset minutes after the client leaves. "
                 "Counted in the chair time, not charged to them."),
        Col("price", "Price", 11, "money",
            validate=number_check(0, 100000),
            note="What you charge today. The suggested price column "
                 "says what the maths would ask for."),
        Col("product_cost", "Product cost", 11, "money",
            validate=number_check(0, 100000),
            note="Backbar cost of one service - colour, developer, "
                 "wax, lash adhesive. Leave blank and the workbook "
                 "estimates it from the percentage on Settings."),
        Col("addon", "Add-on?", 9, "flag", validate=yesno(),
            note="Yes for treatments, toners and finishes sold "
                 "alongside something else."),
        Col("commissionable", "Pays commission?", 13, "flag",
            validate=yesno(),
            note="No for anything the owner keeps whole - a patch "
                 "test, a fringe trim, a redo."),
        Col("online", "Bookable online?", 13, "flag", validate=yesno()),
        Col("level", "Stylist level", 14, "text",
            validate=list_from("StylistLevelList"),
            note="Who is allowed to do it. Staff specific pricing lives "
                 "in the price override on the appointment."),
        Col("status", "Status", 12, "text",
            validate=list_from("ServiceStatusList"), cf=status_cf()),
        Col("reviewed", "Price set on", 11, "date", validate=date_check(),
            note="The last time this price changed. The alert column "
                 "nags when it is older than the review period."),
        Col("notes", "Notes", 26, "wrap"),
        # ------------------------------------------------- what it costs --
        Col("hours", "Chair hours", 10, "number",
            formula=blank(
                "name",
                'ROUND((IF({duration}{r}="",0,{duration}{r})+'
                'IF({buffer}{r}="",0,{buffer}{r}))/60,3)'),
            hidden=True),
        Col("cost_product", "Product cost used", 12, "money",
            formula=blank(
                "name",
                'IF({product_cost}{r}<>"",{product_cost}{r},'
                'IF({price}{r}="",0,'
                'ROUND({price}{r}*ProductCostPercent,2)))'),
            note="Yours if you typed one, otherwise the percentage "
                 "estimate from Settings."),
        Col("cost_commission", "Commission", 11, "money",
            formula=blank(
                "name",
                'IF({commissionable}{r}="No",0,IF({price}{r}="",0,'
                'ROUND({price}{r}*ServiceCommission,2)))'),
            note="At the default service rate on Settings. The real "
                 "figure per stylist is on the Commissions tab."),
        Col("cost_total", "Direct cost", 11, "money",
            formula=blank(
                "name",
                'ROUND({cost_product}{r}+{cost_commission}{r},2)'),
            note="Product plus commission. Rent, power and the front "
                 "desk are overheads - they are in the profit and loss, "
                 "not here."),
        Col("profit", "Gross profit", 11, "money",
            formula=blank("name",
                          'IF({price}{r}="","",'
                          'ROUND({price}{r}-{cost_total}{r},2))'),
            cf=loss_cf(),
            note="Price less the cost of delivering it. This is the "
                 "money that has to cover the rent."),
        Col("margin", "Margin", 9, "percent",
            formula=blank(
                "name",
                'IF({price}{r}="","",IF({price}{r}=0,"",'
                '{profit}{r}/{price}{r}))'),
            cf=[{"type": "cell", "criteria": "<", "value": 0, "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": 0.4,
                 "fmt": "warn"},
                {"type": "cell", "criteria": ">=", "value": 0.4,
                 "fmt": "good"}]),
        Col("price_hour", "Price an hour", 11, "money",
            formula=blank(
                "name",
                'IF({price}{r}="","",IF({hours}{r}=0,"",'
                'ROUND({price}{r}/{hours}{r},2)))'),
            cf=[{"type": "cell", "criteria": "<", "value": "=TargetChairRate",
                 "fmt": "warn"},
                {"type": "cell", "criteria": ">=",
                 "value": "=TargetChairRate", "fmt": "good"}],
            note="The number that compares services fairly. A cheap "
                 "service in half the time can beat an expensive one."),
        Col("profit_hour", "Profit an hour", 11, "money",
            formula=blank(
                "name",
                'IF({profit}{r}="","",IF({hours}{r}=0,"",'
                'ROUND({profit}{r}/{hours}{r},2)))'),
            cf=loss_cf()),
        Col("suggested_raw", "Suggested, raw", 11, "money",
            formula=blank(
                "name",
                'IF({hours}{r}=0,"",IF(ServiceCommission>=1,"",'
                '({cost_product}{r}+TargetChairRate*{hours}{r})/'
                'IF({commissionable}{r}="No",1,1-ServiceCommission)))'),
            hidden=True),
        Col("suggested", "Suggested price", 12, "money",
            formula=blank(
                "name",
                'IF({suggested_raw}{r}="","",ROUND(' +
                round_price("{suggested_raw}{r}") + ',2))'),
            note="(Product cost + your target chair rate x the hours) / "
                 "(1 - commission), rounded by the rule on Settings. "
                 "Change the target, not the formula."),
        Col("gap", "Price gap", 10, "money",
            formula=blank(
                "name",
                'IF({suggested}{r}="","",IF({price}{r}="","",'
                'ROUND({price}{r}-{suggested}{r},2)))'),
            cf=loss_cf(),
            note="Positive means you charge more than the maths asks "
                 "for. Negative is money you are leaving behind."),
        # ----------------------------------------------- how it is doing --
        Col("bookings", "Booked", 9, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Appointments.service]],{id}{r})'),
            note="Every appointment that names this service, whatever "
                 "happened to it."),
        Col("done", "Completed", 10, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Appointments.service]],{id}{r},'
                '[[Appointments.status]],"Completed")')),
        Col("units", "Sold", 8, "int",
            formula=blank(
                "name",
                'SUMIFS([[Sales.qty]],[[Sales.item]],{id}{r},'
                '[[Sales.kind]],"Service")'),
            hidden=True),
        Col("revenue", "Revenue", 12, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.net]],[[Sales.item]],{id}{r},'
                '[[Sales.kind]],"Service"),2)'),
            cf=[{"type": "data_bar", "bar_color": Palette.primary_mid}],
            note="Net of discounts, from the Sales tab."),
        Col("revenue_share", "Share of sales", 11, "percent",
            formula=blank(
                "name",
                'IF(ServiceRevenueAll=0,"",{revenue}{r}/ServiceRevenueAll)'),
            hidden=True),
        Col("profit_total", "Profit earned", 12, "money",
            formula=blank(
                "name",
                'ROUND({units}{r}*{profit}{r},2)'),
            hidden=True),
        Col("hours_total", "Chair hours used", 12, "number1",
            formula=blank("name", 'ROUND({units}{r}*{hours}{r},1)'),
            hidden=True),
        Col("retail_attach", "Retail attached", 11, "percent",
            formula=blank(
                "name",
                'IF({done}{r}=0,"",COUNTIFS([[Sales.item]],{id}{r},'
                '[[Sales.retail_with]],"Yes")/{done}{r})'),
            note="How often a client who had this service also bought "
                 "something to take home."),
        Col("last_done", "Last done", 11, "date",
            formula=blank(
                "name",
                'IFERROR(IF(COUNTIFS([[Appointments.service]],{id}{r})=0,'
                '"",SUMPRODUCT(MAX(([[Appointments.service]]={id}{r})*'
                'IFERROR([[Appointments.date]]*1,0)))),"")')),
        rank_desc("revenue", header="Revenue rank"),
        rank_desc("profit_hour", header="Rate rank"),
        rank_desc("done", header="Popularity"),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "name",
                'IF({price}{r}="","Add - no price yet",'
                'IF({duration}{r}="","Add - how long does it take?",'
                'IF({profit}{r}<0,"Loss - this service costs you money",'
                'IF({price_hour}{r}<TargetChairRate*0.7,'
                '"Check - well under your chair rate",'
                'IF(IF({reviewed}{r}="",FALSE,'
                'TodayDate-{reviewed}{r}>PriceReviewDays),'
                '"Due - price has not been reviewed",'
                'IF(IF({status}{r}="Active",{done}{r}=0,FALSE),'
                '"Check - on the menu but never booked","OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Services & Pricing",
        subtitle="The menu, the chair time, the product cost and what "
                 "each treatment actually earns an hour.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).services,
        help_text="Price, minutes and product cost are the three things "
                  "to get right. Everything to the right of them is "
                  "worked out for you.",
        extra_notes=[
            "Minutes matter as much as the price. Two hours of colour at "
            "$150 earns less an hour than a $45 cut in forty minutes, "
            "and the price an hour column is where that shows up.",
            "Leave the product cost blank and the workbook estimates it "
            "from the percentage on Settings - good enough to start, "
            "worth replacing with real numbers for your colour lines.",
            "The suggested price is built from your target chair rate, "
            "not from a markup. Raise the target on Settings and every "
            "service re-prices itself.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Menu summary under the grid."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1        # 0 based row under the data
    last = spec_.last_col_index
    row = sheet.section(row, "The menu at a glance", last)
    sheet.table_head(row, 0, [
        "Services listed", "Active", "Average price", "Average minutes",
        "Average price an hour", "Best rate an hour", "Never booked",
        "Priced under target"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Services.name]],"<>")',
        '=COUNTIFS([[Services.status]],"Active")',
        '=IFERROR(ROUND(AVERAGEIF([[Services.price]],">0"),2),0)',
        '=IFERROR(ROUND(AVERAGEIF([[Services.duration]],">0"),0),0)',
        '=IFERROR(ROUND(SUMPRODUCT(IFERROR([[Services.price_hour]]*1,0))/'
        'MAX(1,COUNTIF([[Services.name]],"<>")),2),0)',
        '=ROUND(MAX(IFERROR([[Services.price_hour]]*1,0)),2)',
        '=COUNTIFS([[Services.name]],"<>",[[Services.done]],0)',
        '=SUMPRODUCT((IFERROR([[Services.price_hour]]*1,0)>0)*'
        '(IFERROR([[Services.price_hour]]*1,0)<TargetChairRate))',
    ], ["int", "int", "money", "int", "money", "money", "int", "int"])
````

### `salon_business_spreadsheet/sheets/clients.py`

*310 lines · 2 functions*

> `Clients` - the client book, and who is quietly drifting away.
>
> Every salon knows its regulars. What a salon rarely knows is which of
> last spring's regulars has not been in since, and that is the column
> this tab exists for: last visit, days since, due to rebook, lapsing,
> gone. Spend, visit count and average ticket come along for the ride.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Client book summary |

**Module constants:** `NAME` ('Clients'), `NEVER` (2958465).

````python
"""`Clients` - the client book, and who is quietly drifting away.

Every salon knows its regulars. What a salon rarely knows is which of
last spring's regulars has not been in since, and that is the column
this tab exists for: last visit, days since, due to rebook, lapsing,
gone. Spend, visit count and average ticket come along for the ride.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec, max_if, min_if
from ..sample_data import get as sample
from ..writer import list_from, date_check
from .common import (alert_cf, blank, id_col, lookup_staff, owing_cf,
                     rank_desc, status_cf, yesno)

NAME = "Clients"

#: 31-Dec-9999 as a serial - what `min_if` hands back when nothing matched.
NEVER = 2958465


def spec(cfg: Config) -> SheetSpec:
    appt_client = "[[Appointments.client]]"
    appt_date = "[[Appointments.date]]"
    appt_status = "[[Appointments.status]]"

    columns = [
        id_col("name", "ClientPrefix", digits=4, width=11,
               note="Automatic. Bookings, sales and payments all point "
                    "at this ID."),
        Col("name", "Client", 24, "text",
            note="First and last name. Keep one row per person - "
                 "duplicates split their history in two, and the check "
                 "column tells you when you have made one."),
        Col("phone", "Phone", 16, "text"),
        Col("email", "Email", 26, "text"),
        Col("contact", "Contact by", 12, "text",
            validate=list_from("ContactList")),
        Col("joined", "First seen", 11, "date", validate=date_check(),
            note="When they first came in. Leave it blank and the "
                 "workbook uses their first completed appointment."),
        Col("source", "Found you by", 18, "text",
            validate=list_from("SourceList"),
            note="Where this client came from. The marketing tab "
                 "counts new clients by this column."),
        Col("referred_by", "Referred by", 18, "text",
            note="The client ID of whoever sent them. The loyalty tab "
                 "pays the referral credit from it."),
        Col("staff", "Usual stylist", 12, "id",
            validate=list_from("StaffList"),
            note="Staff ID. Used for the rebooking reminders and the "
                 "'my clients' counts on the staff tab."),
        Col("staff_name", "Stylist", 16, "text",
            formula=lookup_staff("name"), hidden=True),
        Col("birthday", "Birthday", 11, "date", validate=date_check(),
            note="Day and month are what matter - the year is up to "
                 "the client. Drives the birthday list on the "
                 "marketing tab."),
        Col("preferences", "Formula & preferences", 34, "wrap",
            note="Colour formula, timings, how they take their coffee, "
                 "the fringe conversation. The thing that makes a "
                 "client feel known."),
        Col("allergies", "Allergies & notes", 26, "wrap"),
        Col("patch_test", "Patch test", 11, "date", validate=date_check(),
            note="Last patch test date. The action column flags it "
                 "when it is older than the period on Settings."),
        Col("marketing_ok", "Marketing OK?", 12, "flag", validate=yesno(),
            note="Only email or text the clients who said yes."),
        Col("photos_ok", "Photos OK?", 11, "flag", validate=yesno()),
        Col("tags", "Tags", 18, "text",
            note="Colour client, bridal, student, blow-dry regular - "
                 "whatever you want to filter on."),
        Col("set_status", "Status", 12, "text",
            validate=list_from("ClientStatusList"), cf=status_cf(),
            note="Leave as Active for a normal client. Archived keeps "
                 "the history but drops them out of the chase lists."),
        Col("notes", "Notes", 26, "wrap"),
        # ----------------------------------------------- visit behaviour --
        Col("visits", "Visits", 8, "int",
            formula=blank(
                "name",
                f'COUNTIFS({appt_client},{{id}}{{r}},'
                f'{appt_status},"Completed")'),
            note="Completed appointments. Cancellations and no shows "
                 "are counted separately."),
        Col("visits_year", "Visits, 12m", 10, "int",
            formula=blank(
                "name",
                f'COUNTIFS({appt_client},{{id}}{{r}},'
                f'{appt_status},"Completed",'
                f'{appt_date},">="&TodayDate-365)'),
            hidden=True),
        Col("cancels", "Cancelled", 10, "int",
            formula=blank(
                "name",
                f'COUNTIFS({appt_client},{{id}}{{r}},'
                f'{appt_status},"Cancelled")'),
            hidden=True),
        Col("no_shows", "No shows", 10, "int",
            formula=blank(
                "name",
                f'COUNTIFS({appt_client},{{id}}{{r}},'
                f'{appt_status},"No show")'),
            cf=[{"type": "cell", "criteria": ">=", "value": 2,
                 "fmt": "bad"}],
            note="Two or more is the point at which most salons start "
                 "asking for a deposit."),
        Col("first_visit", "First visit", 11, "date",
            formula=blank(
                "name",
                'IF({joined}{r}<>"",{joined}{r},'
                'IF({visits}{r}=0,"",' +
                min_if(appt_date, f"{appt_client}={{id}}{{r}}",
                       f'{appt_status}="Completed"') + "))")),
        Col("last_visit", "Last visit", 11, "date",
            formula=blank(
                "name",
                'IF({visits}{r}=0,"",' +
                max_if(appt_date, f"{appt_client}={{id}}{{r}}",
                       f'{appt_status}="Completed"') + ")")),
        Col("days_since", "Days since", 10, "int",
            formula=blank(
                "name",
                'IF({last_visit}{r}="","",TodayDate-{last_visit}{r})'),
            cf=[{"type": "cell", "criteria": ">=",
                 "value": "=InactiveDays", "fmt": "bad"},
                {"type": "cell", "criteria": ">=",
                 "value": "=LapsingDays", "fmt": "warn"}]),
        Col("next_raw", "Next raw", 11, "number",
            formula=blank(
                "name",
                min_if(appt_date, f"{appt_client}={{id}}{{r}}",
                       f'{appt_status}<>"Cancelled"',
                       f'{appt_status}<>"Completed"',
                       f'{appt_status}<>"No show"',
                       f"{appt_date}>=TodayDate")),
            hidden=True),
        Col("next_visit", "Next booked", 11, "date",
            formula=blank(
                "name",
                f'IF({{next_raw}}{{r}}>={NEVER},"",{{next_raw}}{{r}})'),
            note="Their next appointment that has not been cancelled. "
                 "Blank here is the whole point of the rebooking list."),
        Col("gap_days", "Usual gap", 10, "int",
            formula=blank(
                "name",
                'IF({visits}{r}<2,"",ROUND(({last_visit}{r}-'
                '{first_visit}{r})/MAX(1,{visits}{r}-1),0))'),
            note="Average days between visits. A client whose gap is "
                 "stretching is a client on the way out."),
        Col("due_date", "Rebook by", 11, "date",
            formula=blank(
                "name",
                'IF({last_visit}{r}="","",{last_visit}{r}+'
                'IF({gap_days}{r}="",RebookWeeks*7,'
                'MAX(14,{gap_days}{r})))')),
        # ------------------------------------------------------- the money --
        Col("spend", "Lifetime spend", 13, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.net]],[[Sales.client]],'
                '{id}{r}),2)'),
            cf=[{"type": "data_bar", "bar_color": Palette.primary_mid}],
            note="Net of discounts and refunds, services and retail "
                 "together."),
        Col("spend_service", "On services", 12, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.net]],[[Sales.client]],{id}{r},'
                '[[Sales.kind]],"Service"),2)'),
            hidden=True),
        Col("spend_retail", "On retail", 11, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.net]],[[Sales.client]],{id}{r},'
                '[[Sales.kind]],"Retail"),2)'),
            hidden=True),
        Col("spend_year", "Spend, 12m", 12, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.net]],[[Sales.client]],{id}{r},'
                '[[Sales.date]],">="&TodayDate-365),2)')),
        Col("avg_ticket", "Average ticket", 12, "money",
            formula=blank(
                "name",
                'IF({visits}{r}=0,"",ROUND({spend}{r}/{visits}{r},2))'),
            cf=[{"type": "cell", "criteria": ">=",
                 "value": "=TargetTicket", "fmt": "good"}]),
        Col("yearly_value", "A year is worth", 12, "money",
            formula=blank(
                "name",
                'IF({avg_ticket}{r}="","",IF({gap_days}{r}="","",'
                'ROUND({avg_ticket}{r}*(365/MAX(14,{gap_days}{r})),2)))'),
            note="Average ticket x how many times a year they come. "
                 "This is what losing them costs."),
        Col("balance", "Balance owed", 12, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.balance]],[[Sales.client]],'
                '{id}{r}),2)'),
            cf=owing_cf(),
            note="Everything billed less everything paid, deposits "
                 "included. Positive means they owe you."),
        Col("points", "Loyalty points", 12, "int",
            formula=blank(
                "name",
                'IFERROR(INDEX([[Loyalty.balance]],'
                'MATCH({id}{r},[[Loyalty.client]],0))*1,0)'),
            note="From the loyalty tab. Zero means they are not on the "
                 "scheme yet."),
        # -------------------------------------------------- where they sit --
        Col("segment", "Segment", 13, "status",
            formula=blank(
                "name",
                'IF({set_status}{r}="Archived","Archived",'
                'IF({visits}{r}=0,"Not been yet",'
                'IF({days_since}{r}>=InactiveDays,"Inactive",'
                'IF({days_since}{r}>=LapsingDays,"Lapsing",'
                'IF({visits}{r}=1,"New",'
                'IF({spend_year}{r}>=VipSpend,"VIP",'
                'IF({visits_year}{r}>=4,"Regular","Active")))))))'),
            cf=status_cf(),
            note="Worked out from the visits and the spend, using the "
                 "thresholds on Settings. Nothing to type."),
        Col("retained", "Came back?", 11, "flag",
            formula=blank("name",
                          'IF({visits}{r}>=2,"Yes","No")'),
            cf=[{"type": "text", "criteria": "containing", "value": "Yes",
                 "fmt": "good"}],
            hidden=True),
        rank_desc("spend", header="Spend rank"),
        rank_desc("visits", header="Visit rank"),
        Col("alert", "Action", 32, "status",
            formula=blank(
                "name",
                'IF(COUNTIF({name_first}:{name_last},{name}{r})>1,'
                '"Duplicate - same name twice",'
                'IF({set_status}{r}="Archived","",'
                'IF({balance}{r}>0.004,"Chase - balance outstanding",'
                'IF({segment}{r}="Inactive","Gone - win-back offer",'
                'IF({segment}{r}="Lapsing","Chase - lapsing, call them",'
                'IF(IF({patch_test}{r}="",FALSE,'
                'TodayDate-{patch_test}{r}>PatchTestDays),'
                '"Due - patch test out of date",'
                'IF(IF({due_date}{r}="",FALSE,'
                'IF({next_visit}{r}<>"",FALSE,'
                '{due_date}{r}<=TodayDate+14)),'
                '"Rebook - due and nothing booked",'
                'IF(IF({birthday}{r}="",FALSE,'
                'ABS(DATE(YEAR(TodayDate),MONTH({birthday}{r}),'
                'DAY({birthday}{r}))-TodayDate)<=7),'
                '"Birthday - send the treat","OK"))))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Client CRM",
        subtitle="Who they are, what they have spent, when they were "
                 "last in - and which of them you are about to lose.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).clients,
        help_text="Type the name and the contact details. Visits, spend, "
                  "last visit and the chase list all build themselves "
                  "from the appointments and sales tabs.",
        extra_notes=[
            "The action column is the one to filter on each Monday. "
            "Lapsing, due a rebook, balance outstanding, patch test out "
            "of date - that is your call list for the week.",
            "A year is worth is the number to quote when somebody says "
            "losing one client does not matter. It is the average "
            "ticket times how often they come.",
            "Segment is worked out, not typed. Change the lapsing and "
            "VIP thresholds on Settings and every client re-sorts.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Client book summary."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "The client book", last)
    sheet.table_head(row, 0, [
        "Clients on file", "Active", "New this year", "Lapsing",
        "Inactive", "Average lifetime spend", "Average visits",
        "Owing you money", "Marketing opt-ins"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Clients.name]],"<>")',
        '=COUNTIFS([[Clients.segment]],"Active")+'
        'COUNTIFS([[Clients.segment]],"Regular")+'
        'COUNTIFS([[Clients.segment]],"VIP")+'
        'COUNTIFS([[Clients.segment]],"New")',
        '=COUNTIFS([[Clients.first_visit]],">="&TodayDate-365)',
        '=COUNTIFS([[Clients.segment]],"Lapsing")',
        '=COUNTIFS([[Clients.segment]],"Inactive")',
        '=IFERROR(ROUND(SUM([[Clients.spend]])/'
        'MAX(1,COUNTIF([[Clients.name]],"<>")),2),0)',
        '=IFERROR(ROUND(SUM([[Clients.visits]])/'
        'MAX(1,COUNTIF([[Clients.name]],"<>")),1),0)',
        '=COUNTIFS([[Clients.balance]],">0.004")',
        '=COUNTIFS([[Clients.marketing_ok]],"Yes")',
    ], ["int", "int", "int", "int", "int", "money", "number1", "int",
        "int"])
````

### `salon_business_spreadsheet/sheets/appointments.py`

*266 lines · 2 functions*

> `Appointments` - the diary, with a double-booking alarm.
>
> One row per booking. Type the date, the time, the client, the service
> and the stylist; the workbook fills in the name, the price, the
> duration, the finish time, the deposit position and - the column that
> earns its keep - whether that stylist is already booked at that moment.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Diary summary under the grid |

**Module constants:** `NAME` ('Appointments').

````python
"""`Appointments` - the diary, with a double-booking alarm.

One row per booking. Type the date, the time, the client, the service
and the stylist; the workbook fills in the name, the price, the
duration, the finish time, the deposit position and - the column that
earns its keep - whether that stylist is already booked at that moment.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check, time_check
from .common import (alert_cf, blank, id_col, lookup_client, lookup_service,
                     lookup_staff, month_col, status_cf, week_col,
                     weekday_col)

NAME = "Appointments"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "AppointmentPrefix", digits=4, width=11,
               note="Automatic. Payments and sales can quote this ID so "
                    "a deposit lands against the right booking."),
        Col("date", "Date", 11, "date", validate=date_check(),
            note="The day of the appointment. Everything on this row "
                 "waits for it."),
        weekday_col("date"),
        Col("start", "Start", 9, "time", validate=time_check(),
            note="Type 9:30 or 2:15 PM - Excel stores it as a time."),
        Col("client", "Client ID", 11, "id",
            validate=list_from("ClientList"),
            note="Pick the client ID. A walk-in with no record can be "
                 "left blank - the sale still counts."),
        Col("client_name", "Client", 20, "text",
            formula=lookup_client("name", missing="Walk-in"),
            note="Looked up from the Clients tab."),
        Col("service", "Service ID", 11, "id",
            validate=list_from("ServiceList"),
            note="Pick from the service menu. The price and the "
                 "minutes come with it."),
        Col("service_name", "Service", 24, "text",
            formula=lookup_service("name")),
        Col("staff", "Stylist ID", 11, "id",
            validate=list_from("StaffList")),
        Col("staff_name", "Stylist", 16, "text",
            formula=lookup_staff("name", missing="Unassigned")),
        Col("minutes", "Minutes", 9, "int",
            validate=number_check(0, 1440, decimal=False),
            note="Only fill this in when the booking is longer or "
                 "shorter than the menu says. Otherwise leave it "
                 "blank."),
        Col("price", "Price", 10, "money",
            validate=number_check(0, 100000),
            note="Overrides the menu price - a senior stylist "
                 "surcharge, a quoted colour correction, a friend "
                 "rate. Blank means use the menu."),
        Col("status", "Status", 13, "text",
            validate=list_from("AppointmentStatusList"), cf=status_cf(),
            note="Enquiry, Booked, Confirmed, Arrived, Completed, "
                 "Cancelled, No show, Rescheduled. Only the last three "
                 "and Completed count as a final answer."),
        Col("source", "Booked via", 16, "text",
            validate=list_from("SourceList")),
        Col("deposit_due", "Deposit asked", 12, "money",
            validate=number_check(0, 100000),
            note="Leave blank to use the deposit rule on Settings. "
                 "Type 0 to waive it."),
        Col("notes", "Notes", 28, "wrap"),
        Col("cancel_reason", "If cancelled, why", 22, "text",
            validate=list_from("CancelReasonList")),
        Col("fee_charged", "Late fee charged", 13, "money",
            validate=number_check(0, 10000),
            note="A no-show or late cancellation fee you actually "
                 "collected. Log the money itself on Payments."),
        # ------------------------------------------------- worked out ----
        month_col("date"),
        week_col("date"),
        Col("minutes_used", "Minutes used", 11, "int",
            formula=blank(
                "date",
                'IF({minutes}{r}<>"",{minutes}{r},IFERROR(INDEX('
                '[[Services.duration]],MATCH({service}{r},'
                '[[Services.id]],0))*1,0)+IFERROR(INDEX('
                '[[Services.buffer]],MATCH({service}{r},'
                '[[Services.id]],0))*1,0))'),
            hidden=True),
        Col("finish", "Finishes", 9, "time",
            formula=blank(
                "date",
                'IF({start}{r}="","",{start}{r}+{minutes_used}{r}/1440)'),
            note="Start plus the service time and the turnaround."),
        Col("hour", "Hour", 7, "int",
            formula=blank(
                "date",
                'IF({start}{r}="","",HOUR({start}{r}))'),
            hidden=True,
            note="The hour the appointment starts, for the "
                 "busiest-time analysis."),
        Col("hours", "Chair hours", 10, "number1",
            formula=blank("date", 'ROUND({minutes_used}{r}/60,2)'),
            hidden=True),
        Col("price_used", "Price charged", 11, "money",
            formula=blank(
                "date",
                'IF({price}{r}<>"",{price}{r},IFERROR(INDEX('
                '[[Services.price]],MATCH({service}{r},'
                '[[Services.id]],0))*1,0))'),
            note="Your override if there is one, otherwise the menu "
                 "price. This is what the diary is worth."),
        Col("deposit_wanted", "Deposit wanted", 12, "money",
            formula=blank(
                "date",
                'IF({deposit_due}{r}<>"",{deposit_due}{r},'
                'IF({price_used}{r}>=DepositThreshold,'
                'ROUND({price_used}{r}*DepositPercent,2),0))'),
            hidden=True),
        Col("deposit_paid", "Deposit paid", 12, "money",
            formula=blank(
                "date",
                'ROUND(SUMIFS([[Payments.amount]],'
                '[[Payments.appointment]],{id}{r},'
                '[[Payments.kind]],"Deposit"),2)'),
            note="Deposits taken against this booking. They are money "
                 "received, not revenue - the revenue appears when the "
                 "sale is rung up."),
        Col("billed", "Billed", 11, "money",
            formula=blank(
                "date",
                'ROUND(SUMIFS([[Sales.net]],[[Sales.appointment]],'
                '{id}{r}),2)'),
            note="What was actually rung through the till against this "
                 "appointment, retail included."),
        Col("is_final", "Final?", 8, "flag",
            formula=blank(
                "date",
                'IF({status}{r}="Completed","Yes",'
                'IF({status}{r}="Cancelled","Yes",'
                'IF({status}{r}="No show","Yes","No")))'),
            hidden=True,
            note="Completed, Cancelled and No show are final. "
                 "Everything else is still in play, so it is left out "
                 "of the completion and cancellation rates."),
        Col("lead_days", "Booked ahead", 11, "int",
            formula=blank(
                "date",
                'IF({logged}{r}="","",{date}{r}-{logged}{r})'),
            hidden=True),
        Col("logged", "Booked on", 11, "date", validate=date_check(),
            note="When the booking was made. Optional - it feeds the "
                 "lead-time figure on the analytics tab."),
        Col("rebooked", "Rebooked?", 11, "flag",
            formula=blank(
                "date",
                'IF({status}{r}<>"Completed","",'
                'IF({client}{r}="","No",'
                'IF(COUNTIFS([[Appointments.client]],{client}{r},'
                '[[Appointments.date]],">"&{date}{r},'
                '[[Appointments.status]],"<>Cancelled")>0,"Yes","No")))'),
            cf=[{"type": "text", "criteria": "containing", "value": "Yes",
                 "fmt": "good"},
                {"type": "text", "criteria": "containing", "value": "No",
                 "fmt": "warn"}],
            note="Did this client have anything else in the book after "
                 "this visit? The rebooking rate on the dashboard is "
                 "built from this column."),
        Col("clash", "Clash?", 9, "flag",
            formula=blank(
                "date",
                'IF({staff}{r}="","",IF({start}{r}="","",'
                'IF(COUNTIFS({staff_first}:{staff_last},{staff}{r},'
                '{date_first}:{date_last},{date}{r},'
                '{start_first}:{start_last},"<"&{finish}{r},'
                '{finish_first}:{finish_last},">"&{start}{r},'
                '{status_first}:{status_last},"<>Cancelled")>1,'
                '"Clash","")))'),
            cf=[{"type": "text", "criteria": "containing",
                 "value": "Clash", "fmt": "bad"}],
            note="Counts any other booking for the same stylist on the "
                 "same day whose time overlaps this one. Cancelled "
                 "bookings are ignored."),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "date",
                'IF({clash}{r}="Clash","Clash - stylist double booked",'
                'IF({staff}{r}="","Assign - no stylist yet",'
                'IF({service}{r}="","Add - no service chosen",'
                'IF({status}{r}="No show","No show - charge the fee?",'
                'IF({status}{r}="Cancelled","",'
                'IF(IF({status}{r}="Completed",'
                '{billed}{r}<{price_used}{r}-0.004,FALSE),'
                '"Unpaid - nothing rung up for this visit",'
                'IF(IF({date}{r}<TodayDate,{status}{r}="Booked",FALSE),'
                '"Overdue - past date, still only Booked",'
                'IF(IF({date}{r}=TodayDate,TRUE,FALSE),"Today",'
                'IF(IF({deposit_wanted}{r}>0,'
                '{deposit_paid}{r}<{deposit_wanted}{r}-0.004,FALSE),'
                '"Due - deposit not taken",'
                'IF(IF({date}{r}>TodayDate,'
                '{date}{r}<=TodayDate+2,FALSE),'
                '"Confirm - coming up","OK"))))))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Appointment Planner",
        subtitle="The diary: who, with whom, how long, how much - and "
                 "a double-booking alarm that checks every row.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary,
        freeze_cols=2,
        samples=sample(cfg).appointments,
        help_text="Date, start time, client, service, stylist, status. "
                  "The finish time, the price and the clash check look "
                  "after themselves.",
        extra_notes=[
            "Status is the engine of this workbook. Completed means it "
            "happened and can be charged; Cancelled and No show are "
            "the two ways it did not. Enquiry, Booked, Confirmed and "
            "Arrived are all still open, so they are left out of the "
            "completion and cancellation rates rather than counted as "
            "failures.",
            "A deposit is money received, never revenue. It sits on "
            "the Payments tab as kind Deposit, shows here as deposit "
            "paid, and is subtracted from the balance when the sale is "
            "rung up - so it can never be counted twice.",
            "Sorting the diary is safe: every formula finds its own "
            "row. Sort by date and stylist whenever the week gets "
            "messy.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Diary summary under the grid."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "The diary in numbers", last)
    sheet.table_head(row, 0, [
        "Appointments logged", "Completed", "Cancelled", "No shows",
        "Completion rate", "Cancellation rate", "No-show rate",
        "Booked value, completed", "Chair hours, completed",
        "Clashes to fix"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Appointments.date]],"<>")',
        '=COUNTIFS([[Appointments.status]],"Completed")',
        '=COUNTIFS([[Appointments.status]],"Cancelled")',
        '=COUNTIFS([[Appointments.status]],"No show")',
        '=IFERROR(COUNTIFS([[Appointments.status]],"Completed")/'
        'MAX(1,COUNTIFS([[Appointments.is_final]],"Yes")),0)',
        '=IFERROR(COUNTIFS([[Appointments.status]],"Cancelled")/'
        'MAX(1,COUNTIFS([[Appointments.is_final]],"Yes")),0)',
        '=IFERROR(COUNTIFS([[Appointments.status]],"No show")/'
        'MAX(1,COUNTIFS([[Appointments.is_final]],"Yes")),0)',
        '=ROUND(SUMIFS([[Appointments.price_used]],'
        '[[Appointments.status]],"Completed"),2)',
        '=ROUND(SUMIFS([[Appointments.hours]],'
        '[[Appointments.status]],"Completed"),1)',
        '=COUNTIFS([[Appointments.clash]],"Clash")',
    ], ["int", "int", "int", "int", "percent", "percent", "percent",
        "money", "number1", "int"])
````

---

## 4. The money

The till and the payments against it. Net sale is gross less discounts less refunds; a deposit is held against the booking and applied to the first line of that visit, so it can never be counted as revenue twice.

### `salon_business_spreadsheet/sheets/sales.py`

*271 lines · 2 functions*

> `Sales` - the till. One row per thing sold.
>
> Gross, discount, refund, net, tax, total, cost and profit, on every
> line. Net sales is gross less discounts less refunds, and that is the
> number every other tab in the workbook calls revenue - so a discount
> can never quietly inflate your takings and a refund can never be
> forgotten.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Till summary under the grid |

**Module constants:** `NAME` ('Sales').

````python
"""`Sales` - the till. One row per thing sold.

Gross, discount, refund, net, tax, total, cost and profit, on every
line. Net sales is gross less discounts less refunds, and that is the
number every other tab in the workbook calls revenue - so a discount
can never quietly inflate your takings and a refund can never be
forgotten.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, lookup_client, lookup_staff,
                     month_col, owing_cf, rank_desc, status_cf, week_col)

NAME = "Sales"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "SalePrefix", digits=4, width=11),
        Col("date", "Date", 11, "date", validate=date_check(),
            note="The day the money was earned, not the day it landed "
                 "in the bank."),
        Col("appointment", "Appointment", 12, "id",
            note="Optional. The APT ID this sale belongs to - it links "
                 "the till back to the diary and brings any deposit "
                 "with it."),
        Col("client", "Client ID", 11, "id",
            validate=list_from("ClientList"),
            note="Blank is fine for a walk-in who is not on the "
                 "client book."),
        Col("client_name", "Client", 20, "text",
            formula=lookup_client("name", missing="Walk-in"),
            hidden=True),
        Col("staff", "Stylist ID", 11, "id",
            validate=list_from("StaffList"),
            note="Who did the work, or who sold the product. This is "
                 "the column the commission is calculated from."),
        Col("staff_name", "Stylist", 16, "text",
            formula=lookup_staff("name", missing="House"), hidden=True),
        Col("kind", "Kind", 13, "text", validate=list_from("SaleKindList"),
            note="Service, Retail, Package, Gift certificate, Course "
                 "or Other. Keeping these apart is what makes the "
                 "retail share and the commission split honest."),
        Col("item", "Service / product", 12, "id",
            note="The service ID for a treatment, the product ID for "
                 "retail. Leave blank for anything that is neither."),
        Col("description", "Description", 26, "text",
            note="Only needed when there is no ID - 'gift voucher', "
                 "'course of six', 'fringe trim'."),
        Col("qty", "Qty", 7, "number1", validate=number_check(0, 10000),
            note="Blank counts as 1."),
        Col("unit_price", "Unit price", 11, "money",
            validate=number_check(0, 100000),
            note="Leave blank to take the menu or retail price. Fill "
                 "it in to override for this sale only."),
        Col("discount_pct", "Discount %", 10, "percent",
            validate=number_check(0, 1),
            note="As a percentage of the line - 10% is 0.1 or 10%."),
        Col("discount_amt", "Discount $", 11, "money",
            validate=number_check(0, 100000),
            note="A straight amount off, on top of any percentage."),
        Col("promo", "Promo code", 13, "text",
            note="Matches the code on the Marketing tab so a campaign "
                 "can be measured."),
        Col("refund", "Refunded", 11, "money",
            validate=number_check(0, 100000),
            note="Money handed back on this line. It comes straight "
                 "off net sales - no separate correction needed."),
        Col("notes", "Notes", 22, "wrap"),
        # --------------------------------------------------- the maths ----
        month_col("date"),
        week_col("date"),
        Col("qty_used", "Qty used", 9, "number1",
            formula=blank("date", 'IF({qty}{r}="",1,{qty}{r})'),
            hidden=True),
        Col("item_name", "Item", 24, "text",
            formula=blank(
                "date",
                'IF({description}{r}<>"",{description}{r},'
                'IF({kind}{r}="Retail",'
                'IFERROR(INDEX([[Inventory.name]],MATCH({item}{r},'
                '[[Inventory.id]],0))&"","Not in Inventory"),'
                'IFERROR(INDEX([[Services.name]],MATCH({item}{r},'
                '[[Services.id]],0))&"","Not in Services")))'),
            note="Your description if you typed one, otherwise the "
                 "name behind the ID."),
        Col("price_used", "Price used", 11, "money",
            formula=blank(
                "date",
                'IF({unit_price}{r}<>"",{unit_price}{r},'
                'IF({kind}{r}="Retail",'
                'IFERROR(INDEX([[Inventory.retail_price]],'
                'MATCH({item}{r},[[Inventory.id]],0))*1,0),'
                'IFERROR(INDEX([[Services.price]],MATCH({item}{r},'
                '[[Services.id]],0))*1,0)))'),
            hidden=True),
        Col("gross", "Gross", 11, "money",
            formula=blank(
                "date",
                'ROUND({qty_used}{r}*{price_used}{r},2)'),
            note="Quantity x price, before anything is taken off."),
        Col("discount", "Discount", 11, "money",
            formula=blank(
                "date",
                'ROUND(IF({discount_pct}{r}="",0,'
                '{gross}{r}*{discount_pct}{r})+'
                'IF({discount_amt}{r}="",0,{discount_amt}{r}),2)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}]),
        Col("net", "Net sale", 11, "money",
            formula=blank(
                "date",
                'ROUND({gross}{r}-{discount}{r}-'
                'IF({refund}{r}="",0,{refund}{r}),2)'),
            cf=[{"type": "data_bar", "bar_color": Palette.primary_mid}],
            note="Gross less discounts less refunds. This is the "
                 "revenue figure the whole workbook uses."),
        Col("tax", "Tax", 10, "money",
            formula=blank(
                "date",
                'IF({kind}{r}="Retail",'
                'IF(TaxOnRetail="Yes",ROUND({net}{r}*TaxRate,2),0),'
                'IF({kind}{r}="Service",'
                'IF(TaxOnServices="Yes",ROUND({net}{r}*TaxRate,2),0),'
                '0))'),
            note="Worked out from the two switches and the rate on "
                 "Settings. Set both to No if sales tax is not your "
                 "problem, or change the rate to your own."),
        Col("total", "Total due", 11, "money",
            formula=blank("date", 'ROUND({net}{r}+{tax}{r},2)'),
            note="What the client owes for this line, tax included."),
        Col("cost", "Cost", 10, "money",
            formula=blank(
                "date",
                'ROUND({qty_used}{r}*IF({kind}{r}="Retail",'
                'IFERROR(INDEX([[Inventory.unit_cost]],MATCH({item}{r},'
                '[[Inventory.id]],0))*1,0),'
                'IFERROR(INDEX([[Services.cost_product]],'
                'MATCH({item}{r},[[Services.id]],0))*1,0)),2)'),
            note="Trade cost for retail, backbar cost for a service. "
                 "Wages and commission are handled on the staff tabs, "
                 "not here."),
        Col("profit", "Gross profit", 11, "money",
            formula=blank("date", 'ROUND({net}{r}-{cost}{r},2)'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}]),
        # ---------------------------------------------------- the money ----
        Col("first_of_appt", "First line?", 10, "flag",
            formula=blank(
                "date",
                'IF({appointment}{r}="","No",'
                'IF(COUNTIFS({appointment_first}:{appointment}{r},'
                '{appointment}{r})=1,"Yes","No"))'),
            hidden=True,
            note="Only the first sale line of an appointment picks up "
                 "that appointment's deposit, so a deposit can never "
                 "be applied twice."),
        Col("paid_direct", "Paid to this sale", 12, "money",
            formula=blank(
                "date",
                'ROUND(SUMIFS([[Payments.amount]],[[Payments.sale]],'
                '{id}{r},[[Payments.counts]],"Yes"),2)'),
            hidden=True),
        Col("paid_deposit", "Deposit applied", 12, "money",
            formula=blank(
                "date",
                'IF({first_of_appt}{r}="No",0,'
                'ROUND(SUMIFS([[Payments.amount]],'
                '[[Payments.appointment]],{appointment}{r},'
                '[[Payments.sale]],"",[[Payments.counts]],"Yes"),2))'),
            hidden=True),
        Col("paid", "Paid", 11, "money",
            formula=blank(
                "date",
                'ROUND({paid_direct}{r}+{paid_deposit}{r},2)'),
            note="Payments booked against this sale, plus any deposit "
                 "taken on the appointment that has not already been "
                 "applied somewhere else."),
        Col("balance", "Balance", 11, "money",
            formula=blank("date", 'ROUND({total}{r}-{paid}{r},2)'),
            cf=owing_cf(),
            note="Total due less everything paid. Positive is money "
                 "still owed to you."),
        Col("status", "Paid?", 12, "status",
            formula=blank(
                "date",
                'IF({total}{r}<=0,"No charge",'
                'IF({balance}{r}<=0.004,"Paid",'
                'IF({paid}{r}<=0.004,"Unpaid","Part paid")))'),
            cf=status_cf()),
        Col("retail_with", "Retail too?", 11, "flag",
            formula=blank(
                "date",
                'IF({kind}{r}<>"Service","",IF({client}{r}="","No",'
                'IF(COUNTIFS({client_first}:{client_last},{client}{r},'
                '{date_first}:{date_last},{date}{r},'
                '{kind_first}:{kind_last},"Retail")>0,"Yes","No")))'),
            hidden=True,
            note="Did this client also take something home that day? "
                 "The retail attachment rate is built from it."),
        rank_desc("net", guard='{date}{r}=""', header="Sale rank"),
        Col("alert", "Check", 28, "status",
            formula=blank(
                "date",
                'IF({kind}{r}="","Add - what kind of sale is this?",'
                'IF(IF({price_used}{r}<=0,{refund}{r}="",FALSE),'
                '"Check - no price found for this item",'
                'IF({staff}{r}="","Assign - nobody credited",'
                'IF({balance}{r}>0.004,"Unpaid - balance outstanding",'
                'IF({profit}{r}<0,"Loss - sold below cost",'
                'IF({discount}{r}>{gross}{r}*0.5,'
                '"Check - more than half off","OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Sales & Transactions",
        subtitle="Every service, product and voucher sold - gross, "
                 "discount, refund, net, tax and the profit left over.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).sales,
        help_text="One row per thing sold. Two lines for a client who "
                  "had a colour and bought shampoo - that is how the "
                  "retail figures stay separate from the service "
                  "figures.",
        extra_notes=[
            "Net sale = gross - discounts - refunds. Nothing else in "
            "the workbook calls anything revenue, so the dashboard, "
            "the profit and loss and the commission calculator can "
            "never disagree with each other.",
            "Refunds go on the line they relate to, as a positive "
            "number in the refunded column. No negative sales, no "
            "correcting entries, no double subtraction.",
            "Tax is driven by the two switches on Settings. Tax on "
            "retail only, on services only, on both or on neither - "
            "set it once to match where you trade.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Till summary under the grid."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "The till, all time", last)
    sheet.table_head(row, 0, [
        "Lines", "Gross", "Discounts", "Refunds", "Net sales", "Tax",
        "Cost of sales", "Gross profit", "Paid", "Outstanding"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Sales.date]],"<>")',
        '=ROUND(SUM([[Sales.gross]]),2)',
        '=ROUND(SUM([[Sales.discount]]),2)',
        '=ROUND(SUM([[Sales.refund]]),2)',
        '=ROUND(SUM([[Sales.net]]),2)',
        '=ROUND(SUM([[Sales.tax]]),2)',
        '=ROUND(SUM([[Sales.cost]]),2)',
        '=ROUND(SUM([[Sales.profit]]),2)',
        '=ROUND(SUM([[Sales.paid]]),2)',
        '=ROUND(SUM([[Sales.balance]]),2)',
    ], ["int", "money", "money", "money", "money", "money", "money",
        "money", "money", "money"])
````

### `salon_business_spreadsheet/sheets/payments.py`

*189 lines · 2 functions*

> `Payments` - money in and money back out, with the deposit rule.
>
> A deposit is cash in the bank and nothing more. It is logged here the
> day it is taken, it shows against the appointment, and it is applied to
> the bill when the sale is rung up. It is never revenue, and it is never
> counted twice: once a payment carries a sale ID it stops being an
> unapplied deposit.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Money summary under the grid |

**Module constants:** `NAME` ('Payments').

````python
"""`Payments` - money in and money back out, with the deposit rule.

A deposit is cash in the bank and nothing more. It is logged here the
day it is taken, it shows against the appointment, and it is applied to
the bill when the sale is rung up. It is never revenue, and it is never
counted twice: once a payment carries a sale ID it stops being an
unapplied deposit.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, lookup_client, lookup_staff,
                     month_col, status_cf, week_col, yesno)

NAME = "Payments"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "PaymentPrefix", digits=4, width=11),
        Col("date", "Date", 11, "date", validate=date_check(),
            note="The day the money moved."),
        Col("client", "Client ID", 11, "id",
            validate=list_from("ClientList")),
        Col("client_name", "Client", 20, "text",
            formula=lookup_client("name", missing="Walk-in"),
            hidden=True),
        Col("kind", "What for", 16, "text",
            validate=list_from("PaymentKindList"),
            note="Deposit, Balance, Full payment, Retail only, Refund, "
                 "Gift certificate sale, Package sale or Tip. This "
                 "column decides how the money is treated everywhere "
                 "else."),
        Col("method", "Method", 15, "text",
            validate=list_from("PaymentMethodList"),
            note="Card payments pick up the processing fee from "
                 "Settings automatically."),
        Col("amount", "Amount", 11, "money",
            validate=number_check(0, 1000000),
            note="Always a positive number, even for a refund - the "
                 "kind column says which way it went."),
        Col("sale", "Sale ID", 11, "id",
            note="The till line this settles. Fill it in and the "
                 "payment is applied to that sale."),
        Col("appointment", "Appointment", 12, "id",
            note="For a deposit taken before there is a sale. When the "
                 "client comes in, the deposit is applied to the first "
                 "sale line of that appointment on its own."),
        Col("staff", "Stylist ID", 11, "id",
            validate=list_from("StaffList"),
            note="Who the tip belongs to, or who took the payment."),
        Col("staff_name", "Stylist", 16, "text",
            formula=lookup_staff("name", missing=""), hidden=True),
        Col("reference", "Reference", 16, "text",
            note="Receipt number, last four digits, transfer "
                 "reference - whatever you would need to find it "
                 "again."),
        Col("banked", "Reconciled?", 12, "flag", validate=yesno(),
            note="Tick it off when you have seen it on the bank "
                 "statement or in the till count."),
        Col("notes", "Notes", 24, "wrap"),
        # ---------------------------------------------------- worked out --
        month_col("date"),
        week_col("date"),
        Col("counts", "Settles a bill?", 12, "flag",
            formula=blank(
                "date",
                'IF({kind}{r}="Tip","No",IF({kind}{r}="Refund","No",'
                '"Yes"))'),
            hidden=True,
            note="Tips belong to the stylist and refunds go the other "
                 "way, so neither of them pays off a bill."),
        Col("direction", "In or out", 9, "int",
            formula=blank(
                "date",
                'IF({kind}{r}="Refund",-1,1)'),
            hidden=True),
        Col("fee", "Processing fee", 12, "money",
            formula=blank(
                "date",
                'IF({amount}{r}="",0,IF({method}{r}="Card",'
                'ROUND({amount}{r}*CardFeePercent+CardFeeFixed,2),'
                'IF({method}{r}="Online deposit",'
                'ROUND({amount}{r}*CardFeePercent+CardFeeFixed,2),0)))'),
            note="What the card machine keeps. Set the percentage and "
                 "the fixed part on Settings; it never touches cash or "
                 "bank transfers."),
        Col("net_cash", "Into the bank", 12, "money",
            formula=blank(
                "date",
                'ROUND(IF({amount}{r}="",0,{amount}{r})*'
                '{direction}{r}-{fee}{r},2)'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}],
            note="Amount, minus the fee, negative for a refund. Add "
                 "this column up and you have your cash flow."),
        Col("is_tip", "Tip?", 8, "flag",
            formula=blank("date", 'IF({kind}{r}="Tip","Yes","No")'),
            hidden=True),
        Col("applied_to", "Applied to", 14, "text",
            formula=blank(
                "date",
                'IF({sale}{r}<>"","Sale "&{sale}{r},'
                'IF({appointment}{r}<>"","Booking "&{appointment}{r},'
                'IF({kind}{r}="Tip","Stylist","Unapplied")))'),
            cf=status_cf(),
            note="Unapplied is not wrong - a gift certificate sale or "
                 "a general payment on account belongs here."),
        Col("alert", "Check", 28, "status",
            formula=blank(
                "date",
                'IF({amount}{r}="","Add - no amount yet",'
                'IF({kind}{r}="","Add - what was this for?",'
                'IF(IF({sale}{r}="",FALSE,'
                'COUNTIF([[Sales.id]],{sale}{r})=0),'
                '"Missing - no sale with that ID",'
                'IF(IF({appointment}{r}="",FALSE,'
                'COUNTIF([[Appointments.id]],{appointment}{r})=0),'
                '"Missing - no booking with that ID",'
                'IF(IF({kind}{r}="Tip",{staff}{r}="",FALSE),'
                '"Assign - whose tip is it?",'
                'IF(IF({kind}{r}="Deposit",'
                'IF({sale}{r}="",{appointment}{r}="",FALSE),FALSE),'
                '"Check - deposit not tied to a booking",'
                'IF({banked}{r}="No","Due - not reconciled yet",'
                '"OK")))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Payments & Deposits",
        subtitle="Every payment, deposit, tip and refund - what it "
                 "settles, what the card machine took, what reached "
                 "the bank.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).payments,
        help_text="Log the money here; log what was sold on the Sales "
                  "tab. Keeping them apart is what lets a deposit sit "
                  "in the bank for three weeks without pretending to "
                  "be revenue.",
        extra_notes=[
            "Deposits: log one with kind Deposit and the appointment "
            "ID. It shows on the booking straight away and is applied "
            "to the first sale line of that appointment when the "
            "client comes in. Once a payment has a sale ID it is no "
            "longer treated as an unapplied deposit, so nothing is "
            "ever counted twice.",
            "Tips are the stylist's money. They are kept out of sales, "
            "out of the profit and loss, and carried through to the "
            "payroll tab instead.",
            "Refunds are logged as a positive amount with kind Refund. "
            "The 'into the bank' column turns them negative; the sale "
            "they relate to should also carry the refunded amount so "
            "that net sales drops.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Money summary under the grid."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Money in, money out", last)
    sheet.table_head(row, 0, [
        "Payments logged", "Taken", "Refunded", "Tips", "Card fees",
        "Into the bank", "Deposits held", "Not reconciled"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Payments.date]],"<>")',
        '=ROUND(SUMIFS([[Payments.amount]],[[Payments.kind]],'
        '"<>Refund"),2)',
        '=ROUND(SUMIFS([[Payments.amount]],[[Payments.kind]],'
        '"Refund"),2)',
        '=ROUND(SUMIFS([[Payments.amount]],[[Payments.kind]],"Tip"),2)',
        '=ROUND(SUM([[Payments.fee]]),2)',
        '=ROUND(SUM([[Payments.net_cash]]),2)',
        '=ROUND(SUMIFS([[Payments.amount]],[[Payments.kind]],"Deposit",'
        '[[Payments.sale]],""),2)',
        '=COUNTIFS([[Payments.banked]],"No")',
    ], ["int", "money", "money", "money", "money", "money", "money",
        "int"])
````

---

## 5. The team and what they are paid

Eight commission methods, chosen per person, and a payroll tab that knows the difference between what somebody is paid and what they cost.

### `salon_business_spreadsheet/sheets/staff.py`

*298 lines · 2 functions*

> `Staff` - the team, and what each chair is actually bringing in.
>
> Hours available against hours booked gives utilisation; revenue
> against hours gives the rate per chair hour. Those two numbers,
> side by side, are the whole argument for or against taking somebody
> on, moving a shift, or putting the prices up.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Team summary under the grid |

**Module constants:** `NAME` ('Staff').

````python
"""`Staff` - the team, and what each chair is actually bringing in.

Hours available against hours booked gives utilisation; revenue
against hours gives the rate per chair hour. Those two numbers,
side by side, are the whole argument for or against taking somebody
on, moving a shift, or putting the prices up.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, rank_desc, status_cf)

NAME = "Staff"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "StaffPrefix", digits=3, width=10,
               note="Automatic. Appointments, sales, commissions and "
                    "payroll all point at this ID."),
        Col("name", "Name", 20, "text"),
        Col("role", "Role", 18, "text",
            note="Stylist, colourist, nail tech, lash artist, front "
                 "desk, apprentice - whatever you call it."),
        Col("level", "Level", 14, "text",
            validate=list_from("StylistLevelList")),
        Col("type", "Employment", 15, "text",
            validate=list_from("StaffTypeList"),
            note="Employed, self-employed, chair renter, apprentice, "
                 "owner or contractor. Chair renters pay you rent "
                 "instead of earning commission."),
        Col("start_date", "Started", 11, "date", validate=date_check()),
        Col("status", "Status", 11, "text",
            validate=list_from("StaffStatusList"), cf=status_cf()),
        Col("phone", "Phone", 15, "text"),
        Col("email", "Email", 24, "text"),
        Col("hours_week", "Hours a week", 12, "number1",
            validate=number_check(0, 100),
            note="Contracted or usual hours on the floor. The "
                 "utilisation figure divides booked hours by these."),
        Col("days", "Days worked", 18, "text",
            note="Mon, Tue, Thu, Fri, Sat - free text, for the rota at "
                 "a glance."),
        Col("method", "Commission method", 24, "text",
            validate=list_from("CommissionMethodList"),
            note="Yours to choose, per person: none, percent of "
                 "service, service and retail, percent after product "
                 "cost, tiered, flat per service, hourly only, or "
                 "rent a chair. The commissions tab follows whatever "
                 "you pick here."),
        Col("rate_service", "Service rate", 11, "percent",
            validate=number_check(0, 1),
            note="Leave blank to use the salon default on Settings."),
        Col("rate_retail", "Retail rate", 11, "percent",
            validate=number_check(0, 1)),
        Col("flat_rate", "Flat per service", 12, "money",
            validate=number_check(0, 1000),
            note="Only used by the 'flat amount per service' method."),
        Col("hourly", "Hourly rate", 11, "money",
            validate=number_check(0, 1000),
            note="Blank uses the default hourly rate on Settings."),
        Col("rent_week", "Chair rent", 11, "money",
            validate=number_check(0, 10000),
            note="Weekly rent for a chair renter. It is income to the "
                 "salon, not a cost - log the money itself on the "
                 "Sales tab as kind Other."),
        Col("target_month", "Monthly target", 13, "money",
            validate=number_check(0, 1000000),
            note="What you expect this person to bill in a month."),
        Col("notes", "Notes", 24, "wrap"),
        # -------------------------------------------------- performance ---
        Col("appts", "Appointments", 12, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Appointments.staff]],{id}{r})'),
            hidden=True),
        Col("done", "Completed", 10, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Appointments.staff]],{id}{r},'
                '[[Appointments.status]],"Completed")')),
        Col("no_shows", "No shows", 9, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Appointments.staff]],{id}{r},'
                '[[Appointments.status]],"No show")'),
            hidden=True),
        Col("clients", "Clients seen", 11, "int",
            formula=blank(
                "name",
                'COUNTIFS([[Clients.staff]],{id}{r})'),
            hidden=True,
            note="Clients who have this stylist down as their usual."),
        Col("hours_booked", "Hours booked", 12, "number1",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Appointments.hours]],'
                '[[Appointments.staff]],{id}{r},'
                '[[Appointments.status]],"Completed",'
                '[[Appointments.date]],">="&TodayDate-28),1)'),
            note="Chair hours actually worked in the last four weeks."),
        Col("hours_avail", "Hours available", 12, "number1",
            formula=blank(
                "name",
                'IF({hours_week}{r}="",0,ROUND({hours_week}{r}*4,1))'),
            hidden=True),
        Col("utilisation", "Utilisation", 11, "percent",
            formula=blank(
                "name",
                'IF({hours_avail}{r}=0,"",'
                '{hours_booked}{r}/{hours_avail}{r})'),
            cf=[{"type": "cell", "criteria": ">=",
                 "value": "=TargetUtilisation", "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0.5,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<",
                 "value": "=TargetUtilisation", "fmt": "warn"}],
            note="Hours booked against hours available, last four "
                 "weeks. Below half means an empty chair you are "
                 "paying for."),
        Col("rev_service", "Service sales", 12, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.net]],[[Sales.staff]],{id}{r},'
                '[[Sales.kind]],"Service"),2)'),
            cf=[{"type": "data_bar", "bar_color": Palette.primary_mid}]),
        Col("rev_retail", "Retail sales", 11, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.net]],[[Sales.staff]],{id}{r},'
                '[[Sales.kind]],"Retail"),2)')),
        Col("rev_total", "Total sales", 12, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.net]],[[Sales.staff]],'
                '{id}{r}),2)')),
        Col("rev_month", "Sales this month", 13, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.net]],[[Sales.staff]],{id}{r},'
                '[[Sales.month]],ThisMonthKey),2)'),
            cf=[{"type": "data_bar", "bar_color": Palette.accent}]),
        Col("target_gap", "Against target", 12, "money",
            formula=blank(
                "name",
                'IF({target_month}{r}="","",'
                'ROUND({rev_month}{r}-{target_month}{r},2))'),
            cf=[{"type": "cell", "criteria": ">=", "value": 0,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "warn"}]),
        Col("avg_ticket", "Average ticket", 12, "money",
            formula=blank(
                "name",
                'IF({done}{r}=0,"",ROUND({rev_total}{r}/{done}{r},2))'),
            cf=[{"type": "cell", "criteria": ">=", "value": "=TargetTicket",
                 "fmt": "good"}]),
        Col("retail_share", "Retail share", 11, "percent",
            formula=blank(
                "name",
                'IF({rev_total}{r}=0,"",'
                '{rev_retail}{r}/{rev_total}{r})'),
            cf=[{"type": "cell", "criteria": ">=",
                 "value": "=TargetRetailShare", "fmt": "good"},
                {"type": "cell", "criteria": "<",
                 "value": "=TargetRetailShare", "fmt": "warn"}],
            note="How much of this stylist's money comes from product. "
                 "The easiest number in the salon to improve."),
        Col("rate_hour", "Earned an hour", 12, "money",
            formula=blank(
                "name",
                'IF({hours_booked}{r}=0,"",ROUND(SUMIFS('
                '[[Sales.net]],[[Sales.staff]],{id}{r},'
                '[[Sales.date]],">="&TodayDate-28)/'
                '{hours_booked}{r},2))'),
            cf=[{"type": "cell", "criteria": ">=",
                 "value": "=TargetChairRate", "fmt": "good"},
                {"type": "cell", "criteria": "<",
                 "value": "=TargetChairRate", "fmt": "warn"}]),
        Col("rebook_rate", "Rebook rate", 11, "percent",
            formula=blank(
                "name",
                'IF(COUNTIFS([[Appointments.staff]],{id}{r},'
                '[[Appointments.status]],"Completed")=0,"",'
                'COUNTIFS([[Appointments.staff]],{id}{r},'
                '[[Appointments.rebooked]],"Yes")/'
                'COUNTIFS([[Appointments.staff]],{id}{r},'
                '[[Appointments.status]],"Completed"))'),
            cf=[{"type": "cell", "criteria": ">=",
                 "value": "=TargetRebook", "fmt": "good"},
                {"type": "cell", "criteria": "<",
                 "value": "=TargetRebook", "fmt": "warn"}],
            note="How often their clients leave with the next "
                 "appointment already in the book."),
        Col("tips", "Tips", 10, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Payments.amount]],[[Payments.staff]],'
                '{id}{r},[[Payments.kind]],"Tip"),2)'),
            hidden=True),
        Col("commission_paid", "Commission", 12, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Commissions.commission]],'
                '[[Commissions.staff]],{id}{r}),2)'),
            hidden=True,
            note="Everything calculated for this person on the "
                 "commissions tab."),
        Col("cost_month", "Pay this month", 12, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Payroll.total_cost]],[[Payroll.staff]],'
                '{id}{r},[[Payroll.month]],ThisMonthKey),2)'),
            note="From the payroll tab - wages, commission and the "
                 "employer costs on top."),
        Col("margin_month", "Leaves behind", 12, "money",
            formula=blank(
                "name",
                'ROUND({rev_month}{r}-{cost_month}{r},2)'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "good"}],
            note="This month's sales less this month's pay. What is "
                 "left to cover the rent and everything else."),
        rank_desc("rev_total", header="Sales rank"),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "name",
                'IF({status}{r}<>"Active","",'
                'IF({method}{r}="","Add - pick a commission method",'
                'IF(IF({hours_week}{r}="",TRUE,{hours_week}{r}=0),'
                '"Add - no hours, no utilisation",'
                'IF({utilisation}{r}<0.5,"Check - chair is half empty",'
                'IF({margin_month}{r}<0,"Loss - pay is ahead of sales",'
                'IF(IF({target_month}{r}="",FALSE,'
                '{rev_month}{r}<{target_month}{r}*0.8),'
                '"Below - behind the monthly target",'
                'IF({retail_share}{r}<TargetRetailShare,'
                '"Action - barely selling any product","OK")))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Staff Management",
        subtitle="The team, their pay basis, and what each chair "
                 "brings in an hour.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.deep,
        freeze_cols=2,
        samples=sample(cfg).staff,
        help_text="Set the commission method per person. Everything "
                  "from 'hours booked' rightwards is worked out from "
                  "the diary, the till and the payroll tab.",
        extra_notes=[
            "There is no house commission rate baked into this "
            "workbook. Each person carries their own method and their "
            "own rate, and the commission tab does what that method "
            "says - including paying nothing at all to a chair renter.",
            "Utilisation uses the last four weeks, so it reacts to a "
            "quiet month rather than being flattered by a busy year.",
            "Leaves behind is the test that matters: this month's "
            "sales less this month's full cost of employing them. If "
            "it is negative, nothing else on the row makes up for it.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Team summary under the grid."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "The team at a glance", last)
    sheet.table_head(row, 0, [
        "On the team", "Active", "Hours a week", "Completed visits",
        "Service sales", "Retail sales", "Average utilisation",
        "Average rebook rate", "Pay this month"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Staff.name]],"<>")',
        '=COUNTIFS([[Staff.status]],"Active")',
        '=ROUND(SUM([[Staff.hours_week]]),1)',
        '=SUM([[Staff.done]])',
        '=ROUND(SUM([[Staff.rev_service]]),2)',
        '=ROUND(SUM([[Staff.rev_retail]]),2)',
        '=IFERROR(SUMPRODUCT(IFERROR([[Staff.utilisation]]*1,0))/'
        'MAX(1,COUNTIFS([[Staff.status]],"Active")),0)',
        '=IFERROR(SUMPRODUCT(IFERROR([[Staff.rebook_rate]]*1,0))/'
        'MAX(1,COUNTIFS([[Staff.status]],"Active")),0)',
        '=ROUND(SUM([[Staff.cost_month]]),2)',
    ], ["int", "int", "number1", "int", "money", "money", "percent",
        "percent", "money"])
````

### `salon_business_spreadsheet/sheets/commissions.py`

*258 lines · 2 functions*

> `Commissions` - one row per stylist per month, eight ways to pay.
>
> The method is whatever the owner chose on the Staff tab. This tab
> reads it and applies it: a straight percentage, a percentage of
> service and retail, a percentage after the colour has been paid for,
> a tiered rate that climbs with the month, a flat amount per service,
> hourly only, rent a chair, or nothing at all. No house rate is
> assumed anywhere.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Commission summary under the grid |

**Module constants:** `NAME` ('Commissions').

````python
"""`Commissions` - one row per stylist per month, eight ways to pay.

The method is whatever the owner chose on the Staff tab. This tab
reads it and applies it: a straight percentage, a percentage of
service and retail, a percentage after the colour has been paid for,
a tiered rate that climbs with the month, a flat amount per service,
hourly only, rent a chair, or nothing at all. No house rate is
assumed anywhere.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, lookup_staff, rank_desc,
                     status_cf, yesno)

NAME = "Commissions"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("month", "CommissionPrefix", digits=4, width=11),
        Col("month", "Month", 11, "text", validate=list_from("MonthKeyList"),
            note="The month this covers, as yyyy-mm. Pick it from the "
                 "list so it matches the sales tab exactly."),
        Col("staff", "Stylist ID", 11, "id",
            validate=list_from("StaffList")),
        Col("staff_name", "Stylist", 18, "text",
            formula=lookup_staff("name")),
        Col("method", "Method used", 24, "text",
            formula=lookup_staff("method", missing="None"),
            cf=status_cf(),
            note="Taken from the Staff tab. Change it there and this "
                 "month recalculates."),
        Col("adjust", "Adjustment", 11, "money",
            validate=number_check(-100000, 100000),
            note="A correction, a clawback on a refund, a guarantee "
                 "top-up. Plus or minus."),
        Col("bonus", "Bonus", 11, "money",
            validate=number_check(0, 100000),
            note="Anything on top - a retail competition, a full "
                 "column bonus."),
        Col("approved", "Approved?", 11, "flag", validate=yesno()),
        Col("paid_on", "Paid on", 11, "date", validate=date_check(),
            note="Leave blank until it has actually gone out."),
        Col("notes", "Notes", 24, "wrap"),
        # ------------------------------------------------- what they sold --
        Col("service_net", "Service sales", 12, "money",
            formula=blank(
                "month",
                'ROUND(SUMIFS([[Sales.net]],[[Sales.staff]],{staff}{r},'
                '[[Sales.month]],{month}{r},[[Sales.kind]],'
                '"Service"),2)'),
            cf=[{"type": "data_bar", "bar_color": Palette.primary_mid}]),
        Col("service_gross", "Before discount", 12, "money",
            formula=blank(
                "month",
                'ROUND(SUMIFS([[Sales.gross]],[[Sales.staff]],'
                '{staff}{r},[[Sales.month]],{month}{r},'
                '[[Sales.kind]],"Service"),2)'),
            hidden=True),
        Col("retail_net", "Retail sales", 11, "money",
            formula=blank(
                "month",
                'ROUND(SUMIFS([[Sales.net]],[[Sales.staff]],{staff}{r},'
                '[[Sales.month]],{month}{r},[[Sales.kind]],'
                '"Retail"),2)')),
        Col("retail_gross", "Retail before", 11, "money",
            formula=blank(
                "month",
                'ROUND(SUMIFS([[Sales.gross]],[[Sales.staff]],'
                '{staff}{r},[[Sales.month]],{month}{r},'
                '[[Sales.kind]],"Retail"),2)'),
            hidden=True),
        Col("product_cost", "Product used", 11, "money",
            formula=blank(
                "month",
                'ROUND(SUMIFS([[Sales.cost]],[[Sales.staff]],{staff}{r},'
                '[[Sales.month]],{month}{r},[[Sales.kind]],'
                '"Service"),2)'),
            note="Backbar cost of the services they did - the figure "
                 "the 'percent after product cost' method takes off "
                 "first."),
        Col("services", "Services done", 11, "int",
            formula=blank(
                "month",
                'COUNTIFS([[Sales.staff]],{staff}{r},[[Sales.month]],'
                '{month}{r},[[Sales.kind]],"Service")'),
            hidden=True),
        # ------------------------------------------------- the base and rate
        Col("base_service", "Service base", 12, "money",
            formula=blank(
                "month",
                'IF(CommissionOn="Gross before discounts",'
                '{service_gross}{r},{service_net}{r})'),
            note="Which figure commission is paid on is your choice - "
                 "set it once on Settings, net of discounts or gross "
                 "before them."),
        Col("base_retail", "Retail base", 11, "money",
            formula=blank(
                "month",
                'IF(CommissionOn="Gross before discounts",'
                '{retail_gross}{r},{retail_net}{r})'),
            hidden=True),
        Col("rate_service", "Service rate", 11, "percent",
            formula=blank(
                "month",
                'IF({method}{r}="Tiered on monthly service revenue",'
                'IFERROR(INDEX(TierRate,MATCH({base_service}{r},'
                'TierFrom,1)),0),'
                'IFERROR(IF(INDEX([[Staff.rate_service]],'
                'MATCH({staff}{r},[[Staff.id]],0))="",ServiceCommission,'
                'INDEX([[Staff.rate_service]],MATCH({staff}{r},'
                '[[Staff.id]],0))),ServiceCommission))'),
            note="Their own rate if they have one, otherwise the salon "
                 "default. On the tiered method it is read from the "
                 "bands on Settings instead."),
        Col("rate_retail", "Retail rate", 11, "percent",
            formula=blank(
                "month",
                'IFERROR(IF(INDEX([[Staff.rate_retail]],MATCH({staff}{r},'
                '[[Staff.id]],0))="",RetailCommission,'
                'INDEX([[Staff.rate_retail]],MATCH({staff}{r},'
                '[[Staff.id]],0))),RetailCommission)'),
            hidden=True),
        Col("flat_rate", "Flat per service", 12, "money",
            formula=blank(
                "month",
                'IFERROR(INDEX([[Staff.flat_rate]],MATCH({staff}{r},'
                '[[Staff.id]],0))*1,0)'),
            hidden=True),
        # ----------------------------------------------------- the payment --
        Col("commission", "Commission", 12, "money",
            formula=blank(
                "month",
                'ROUND(IF({method}{r}="Percent of service",'
                '{base_service}{r}*{rate_service}{r},'
                'IF({method}{r}="Percent of service and retail",'
                '{base_service}{r}*{rate_service}{r}+'
                '{base_retail}{r}*{rate_retail}{r},'
                'IF({method}{r}="Percent after product cost",'
                'MAX(0,{base_service}{r}-{product_cost}{r})*'
                '{rate_service}{r},'
                'IF({method}{r}="Tiered on monthly service revenue",'
                '{base_service}{r}*{rate_service}{r},'
                'IF({method}{r}="Flat amount per service",'
                '{services}{r}*{flat_rate}{r},0))))),2)'),
            cf=[{"type": "data_bar", "bar_color": Palette.accent}],
            note="Eligible revenue x the rate, exactly as the method "
                 "says. None, hourly only and rent a chair all come "
                 "out at zero here on purpose - those people are paid "
                 "on the payroll tab or pay you rent instead."),
        Col("tips", "Tips", 10, "money",
            formula=blank(
                "month",
                'ROUND(SUMIFS([[Payments.amount]],[[Payments.staff]],'
                '{staff}{r},[[Payments.month]],{month}{r},'
                '[[Payments.kind]],"Tip"),2)'),
            note="Tips logged to this stylist. They are theirs - shown "
                 "here so the payout is complete, kept out of salon "
                 "revenue everywhere else."),
        Col("payout", "To pay", 12, "money",
            formula=blank(
                "month",
                'ROUND({commission}{r}+IF({adjust}{r}="",0,'
                '{adjust}{r})+IF({bonus}{r}="",0,{bonus}{r}),2)'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}],
            note="Commission plus adjustments and bonus. Tips are not "
                 "added in - they are not yours to pay."),
        Col("share", "Share of their sales", 13, "percent",
            formula=blank(
                "month",
                'IF({base_service}{r}+{base_retail}{r}=0,"",'
                '{payout}{r}/({base_service}{r}+{base_retail}{r}))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0.55,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 0.45,
                 "fmt": "warn"}],
            note="What this month costs you as a percentage of what "
                 "they billed. Over about 45% is worth a conversation."),
        rank_desc("commission", guard='{month}{r}=""',
                  header="Commission rank"),
        Col("alert", "Check", 28, "status",
            formula=blank(
                "month",
                'IF({staff}{r}="","Add - whose commission is this?",'
                'IF(COUNTIFS({month_first}:{month_last},{month}{r},'
                '{staff_first}:{staff_last},{staff}{r})>1,'
                '"Duplicate - two rows for the same month",'
                'IF({method}{r}="None","",'
                'IF({method}{r}="Rent a chair","",'
                'IF({method}{r}="Hourly only","",'
                'IF({payout}{r}=0,"Check - nothing calculated",'
                'IF({approved}{r}<>"Yes","Due - waiting for approval",'
                'IF({paid_on}{r}="","Due - approved but not paid",'
                '"OK"))))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Staff Commissions",
        subtitle="One row per stylist per month. The method is yours; "
                 "the arithmetic is the workbook's.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.deep,
        freeze_cols=2,
        samples=sample(cfg).commissions,
        help_text="Add a row for each stylist each month: month, "
                  "stylist, and any adjustment. The sales, the base, "
                  "the rate and the payout are all read from elsewhere.",
        extra_notes=[
            "Eight methods are supported and none of them is assumed. "
            "Set the method on the Staff tab, per person, and this "
            "tab follows it: percent of service, service and retail, "
            "percent after product cost, tiered, flat per service, "
            "hourly only, rent a chair, or none.",
            "Commission is paid on whichever base you chose on "
            "Settings - net of discounts, or gross before them. One "
            "switch, applied to everybody, so there is never an "
            "argument about which figure was used.",
            "The tiered method reads the bands on the Settings tab. "
            "Edit the bands there to match your own ladder; the "
            "lookup finds the right rung on its own.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Commission summary under the grid."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Commission totals", last)
    sheet.table_head(row, 0, [
        "Rows", "Service sales", "Retail sales", "Commission",
        "Adjustments", "Bonuses", "To pay", "Tips passed on",
        "Waiting for approval", "Approved, not paid"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Commissions.month]],"<>")',
        '=ROUND(SUM([[Commissions.service_net]]),2)',
        '=ROUND(SUM([[Commissions.retail_net]]),2)',
        '=ROUND(SUM([[Commissions.commission]]),2)',
        '=ROUND(SUM([[Commissions.adjust]]),2)',
        '=ROUND(SUM([[Commissions.bonus]]),2)',
        '=ROUND(SUM([[Commissions.payout]]),2)',
        '=ROUND(SUM([[Commissions.tips]]),2)',
        '=COUNTIFS([[Commissions.approved]],"<>Yes",'
        '[[Commissions.month]],"<>")',
        '=COUNTIFS([[Commissions.approved]],"Yes",'
        '[[Commissions.paid_on]],"")',
    ], ["int", "money", "money", "money", "money", "money", "money",
        "money", "int", "int"])
````

### `salon_business_spreadsheet/sheets/payroll.py`

*220 lines · 2 functions*

> `Payroll` - what each pay run actually costs the salon.
>
> Hours at a rate, plus the commission you are releasing this run, plus
> the tips you are passing on, plus the employer costs on top. The last
> column is the one that belongs in the profit and loss: the total cost
> of employing somebody for that period, not the number on their payslip.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Payroll summary under the grid |

**Module constants:** `NAME` ('Payroll').

````python
"""`Payroll` - what each pay run actually costs the salon.

Hours at a rate, plus the commission you are releasing this run, plus
the tips you are passing on, plus the employer costs on top. The last
column is the one that belongs in the profit and loss: the total cost
of employing somebody for that period, not the number on their payslip.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, lookup_staff, month_col,
                     week_col, yesno)

NAME = "Payroll"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("period_end", "PayrollPrefix", digits=4, width=11),
        Col("period_end", "Period ending", 13, "date",
            validate=date_check(),
            note="The last day of the pay period. Everything on this "
                 "tab is grouped by the month this date falls in."),
        Col("period", "Run", 12, "text", validate=list_from("PayPeriodList"),
            note="Weekly, fortnightly, monthly, four weekly or one "
                 "off - label only, so your runs can be whatever "
                 "shape they really are."),
        Col("staff", "Stylist ID", 11, "id",
            validate=list_from("StaffList")),
        Col("staff_name", "Name", 18, "text",
            formula=lookup_staff("name")),
        Col("hours", "Hours", 9, "number1",
            validate=number_check(0, 400),
            note="Hours actually worked in this period."),
        Col("rate", "Hourly rate", 11, "money",
            validate=number_check(0, 1000),
            note="Leave blank to use their rate from the Staff tab, "
                 "or the salon default if they have none."),
        Col("commission", "Commission paid", 13, "money",
            validate=number_check(-100000, 100000),
            note="What you are releasing this run. The column beside "
                 "it shows what the commissions tab calculated for "
                 "the whole month, so you can see what is left."),
        Col("tips", "Tips passed on", 12, "money",
            validate=number_check(0, 100000),
            note="Tips handed over with this run. They are the "
                 "stylist's money passing through, so they are kept "
                 "out of the salon's wage cost."),
        Col("bonus", "Bonus", 10, "money",
            validate=number_check(0, 100000)),
        Col("holiday", "Holiday pay", 11, "money",
            validate=number_check(0, 100000)),
        Col("deductions", "Deductions", 11, "money",
            validate=number_check(0, 100000),
            note="Anything taken off before it is paid - a product "
                 "account, a uniform, an advance."),
        Col("paid", "Paid?", 9, "flag", validate=yesno()),
        Col("paid_on", "Paid on", 11, "date", validate=date_check()),
        Col("method", "Method", 14, "text",
            validate=list_from("PaymentMethodList")),
        Col("notes", "Notes", 22, "wrap"),
        # ---------------------------------------------------- worked out --
        month_col("period_end"),
        week_col("period_end"),
        Col("rate_used", "Rate used", 10, "money",
            formula=blank(
                "period_end",
                'IF({rate}{r}<>"",{rate}{r},IFERROR(IF(INDEX('
                '[[Staff.hourly]],MATCH({staff}{r},[[Staff.id]],0))="",'
                'DefaultHourly,INDEX([[Staff.hourly]],MATCH({staff}{r},'
                '[[Staff.id]],0))),DefaultHourly))'),
            hidden=True),
        Col("base_pay", "Hourly pay", 11, "money",
            formula=blank(
                "period_end",
                'ROUND(IF({hours}{r}="",0,{hours}{r})*'
                '{rate_used}{r},2)')),
        Col("gross_pay", "Gross pay", 11, "money",
            formula=blank(
                "period_end",
                'ROUND({base_pay}{r}+IF({commission}{r}="",0,'
                '{commission}{r})+IF({bonus}{r}="",0,{bonus}{r})+'
                'IF({holiday}{r}="",0,{holiday}{r}),2)'),
            cf=[{"type": "data_bar", "bar_color": Palette.primary_mid}],
            note="Hourly pay plus commission, bonus and holiday pay. "
                 "Tips are not in here - they were never yours."),
        Col("net_pay", "After deductions", 13, "money",
            formula=blank(
                "period_end",
                'ROUND({gross_pay}{r}-IF({deductions}{r}="",0,'
                '{deductions}{r}),2)'),
            note="Before income tax and any statutory deductions - "
                 "this workbook does not calculate those."),
        Col("to_hand_over", "Including tips", 12, "money",
            formula=blank(
                "period_end",
                'ROUND({net_pay}{r}+IF({tips}{r}="",0,{tips}{r}),2)'),
            hidden=True,
            note="What physically leaves you on pay day, tips "
                 "included."),
        Col("employer_cost", "Employer costs", 12, "money",
            formula=blank(
                "period_end",
                'IF(IFERROR(INDEX([[Staff.type]],MATCH({staff}{r},'
                '[[Staff.id]],0))&"","")="Self-employed",0,'
                'IF(IFERROR(INDEX([[Staff.type]],MATCH({staff}{r},'
                '[[Staff.id]],0))&"","")="Chair renter",0,'
                'ROUND({gross_pay}{r}*PayrollBurden,2)))'),
            note="Employer tax, insurance and holiday accrual at the "
                 "percentage on Settings. Self-employed stylists and "
                 "chair renters carry none."),
        Col("total_cost", "Cost to salon", 12, "money",
            formula=blank(
                "period_end",
                'ROUND({gross_pay}{r}+{employer_cost}{r},2)'),
            cf=[{"type": "data_bar", "bar_color": Palette.accent}],
            note="Gross pay plus employer costs. This is the figure "
                 "the profit and loss uses for wages."),
        Col("month_commission", "Month calculated", 13, "money",
            formula=blank(
                "period_end",
                'ROUND(SUMIFS([[Commissions.payout]],'
                '[[Commissions.staff]],{staff}{r},'
                '[[Commissions.month]],{month}{r}),2)'),
            hidden=True),
        Col("month_paid", "Month released", 12, "money",
            formula=blank(
                "period_end",
                'ROUND(SUMIFS({commission_first}:{commission_last},'
                '{staff_first}:{staff_last},{staff}{r},'
                '{month_first}:{month_last},{month}{r}),2)'),
            hidden=True),
        Col("commission_left", "Commission left", 13, "money",
            formula=blank(
                "period_end",
                'ROUND({month_commission}{r}-{month_paid}{r},2)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0.004,
                 "fmt": "warn"},
                {"type": "cell", "criteria": "<", "value": -0.004,
                 "fmt": "bad"}],
            note="This month's calculated commission less everything "
                 "already released in the month's pay runs. Zero means "
                 "you are square; negative means you have overpaid."),
        Col("cost_hour", "Cost an hour", 11, "money",
            formula=blank(
                "period_end",
                'IF(IF({hours}{r}="",0,{hours}{r})=0,"",'
                'ROUND({total_cost}{r}/{hours}{r},2))'),
            hidden=True),
        Col("alert", "Check", 28, "status",
            formula=blank(
                "period_end",
                'IF({staff}{r}="","Add - who is this run for?",'
                'IF(IF({hours}{r}="",{commission}{r}="",FALSE),'
                '"Add - no hours and no commission",'
                'IF({commission_left}{r}<-0.004,'
                '"Check - more commission paid than earned",'
                'IF({paid}{r}<>"Yes","Due - not paid yet",'
                'IF({paid_on}{r}="","Check - marked paid, no date",'
                '"OK")))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Payroll & Tips",
        subtitle="Hours, commission released, tips passed on and the "
                 "employer costs on top - the real cost of a pay run.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.deep,
        freeze_cols=2,
        samples=sample(cfg).payroll,
        help_text="One row per person per pay run. Hours and the "
                  "commission you are releasing are the only numbers "
                  "to type.",
        extra_notes=[
            "Commission is typed here rather than pulled in "
            "automatically, because a monthly commission is often "
            "paid across two or three weekly runs. The 'commission "
            "left' column keeps the score so nobody is paid twice.",
            "Tips pass straight through. They are added to what you "
            "hand over and left out of the wage cost, because they "
            "were never the salon's money.",
            "This is a cost tracker, not a payroll system. It does "
            "not calculate income tax, national insurance or social "
            "security - your payroll provider or accountant does "
            "that, and these totals are what you give them.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Payroll summary under the grid."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "Pay runs, all time", last)
    sheet.table_head(row, 0, [
        "Runs logged", "Hours", "Hourly pay", "Commission released",
        "Bonuses", "Gross pay", "Employer costs", "Cost to salon",
        "Tips passed on", "Runs not paid"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Payroll.period_end]],"<>")',
        '=ROUND(SUM([[Payroll.hours]]),1)',
        '=ROUND(SUM([[Payroll.base_pay]]),2)',
        '=ROUND(SUM([[Payroll.commission]]),2)',
        '=ROUND(SUM([[Payroll.bonus]]),2)',
        '=ROUND(SUM([[Payroll.gross_pay]]),2)',
        '=ROUND(SUM([[Payroll.employer_cost]]),2)',
        '=ROUND(SUM([[Payroll.total_cost]]),2)',
        '=ROUND(SUM([[Payroll.tips]]),2)',
        '=COUNTIFS([[Payroll.paid]],"<>Yes",'
        '[[Payroll.period_end]],"<>")',
    ], ["int", "number1", "money", "money", "money", "money", "money",
        "money", "money", "int"])
````

---

## 6. Costs, stock and suppliers

Every bill with its recurring frequency spread into a monthly overhead, the stock room counted the honest way, and the orders that fill it - including the column that catches a wholesaler's quiet price rise.

### `salon_business_spreadsheet/sheets/expenses.py`

*180 lines · 2 functions*

> `Expenses` - everything the salon spends, in the order a P&L wants it.
>
> Rent, product, wages, card fees, the coffee for the waiting area.
> Recurring costs carry a frequency so the workbook can tell you what
> your fixed monthly overhead really is - the number that decides how
> many heads you need through the door before you have made anything.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Expense summary under the grid |

**Module constants:** `NAME` ('Expenses').

````python
"""`Expenses` - everything the salon spends, in the order a P&L wants it.

Rent, product, wages, card fees, the coffee for the waiting area.
Recurring costs carry a frequency so the workbook can tell you what
your fixed monthly overhead really is - the number that decides how
many heads you need through the door before you have made anything.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, month_col, rank_desc,
                     status_cf, week_col, yesno)

NAME = "Expenses"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "ExpensePrefix", digits=4, width=11),
        Col("date", "Date", 11, "date", validate=date_check(),
            note="The day the cost belongs to, not necessarily the "
                 "day it leaves the bank."),
        Col("category", "Category", 22, "text",
            validate=list_from("ExpenseCategoryList"),
            note="These are the lines of your profit and loss and the "
                 "rows of your budget. Keep to the list."),
        Col("payee", "Paid to", 20, "text",
            note="Landlord, wholesaler, the council, the card "
                 "processor."),
        Col("description", "What for", 28, "text"),
        Col("amount", "Amount", 11, "money",
            validate=number_check(0, 1000000),
            note="Before tax, if you can reclaim it. Including tax if "
                 "you cannot - just be consistent."),
        Col("tax", "Tax", 10, "money", validate=number_check(0, 1000000),
            note="Sales tax or VAT on the bill, if you track it "
                 "separately."),
        Col("method", "Method", 14, "text",
            validate=list_from("PaymentMethodList")),
        Col("recurring", "Recurring?", 11, "flag", validate=yesno(),
            note="Yes for anything that comes round again - rent, "
                 "insurance, software, the music licence."),
        Col("frequency", "How often", 13, "text",
            validate=list_from("FrequencyList"),
            note="Only used when recurring is Yes. It is what turns a "
                 "yearly insurance bill into a monthly overhead."),
        Col("due", "Due", 11, "date", validate=date_check(),
            note="When it has to be paid by. Leave blank for anything "
                 "already settled."),
        Col("paid", "Paid?", 9, "flag", validate=yesno()),
        Col("paid_on", "Paid on", 11, "date", validate=date_check()),
        Col("reference", "Reference", 16, "text",
            note="Invoice or receipt number - the thing your "
                 "accountant will ask for."),
        Col("notes", "Notes", 22, "wrap"),
        # ---------------------------------------------------- worked out --
        month_col("date"),
        week_col("date"),
        Col("total", "Total", 11, "money",
            formula=blank(
                "date",
                'ROUND(IF({amount}{r}="",0,{amount}{r})+'
                'IF({tax}{r}="",0,{tax}{r}),2)'),
            cf=[{"type": "data_bar", "bar_color": Palette.primary_mid}],
            note="Amount plus tax. The profit and loss uses the "
                 "amount; this column is what actually left the bank."),
        Col("monthly", "Monthly equivalent", 14, "money",
            formula=blank(
                "date",
                'IF({recurring}{r}<>"Yes",0,'
                'ROUND(IF({amount}{r}="",0,{amount}{r})*'
                'IF({frequency}{r}="Weekly",52/12,'
                'IF({frequency}{r}="Fortnightly",26/12,'
                'IF({frequency}{r}="Four weekly",13/12,'
                'IF({frequency}{r}="Monthly",1,'
                'IF({frequency}{r}="Quarterly",1/3,'
                'IF({frequency}{r}="Twice a year",1/6,'
                'IF({frequency}{r}="Yearly",1/12,0))))))),2))'),
            note="A yearly bill spread over twelve months, a weekly "
                 "one multiplied up. Add this column and you have "
                 "your fixed monthly overhead."),
        Col("is_fixed", "Fixed cost?", 11, "flag",
            formula=blank(
                "date",
                'IF({recurring}{r}="Yes","Yes","No")'),
            hidden=True),
        Col("days_late", "Days late", 10, "int",
            formula=blank(
                "date",
                'IF({paid}{r}="Yes","",IF({due}{r}="","",'
                'TodayDate-{due}{r}))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "bad"}],
            hidden=True),
        Col("status", "Status", 12, "status",
            formula=blank(
                "date",
                'IF({paid}{r}="Yes","Paid",'
                'IF({due}{r}="","Pending",'
                'IF({due}{r}<TodayDate,"Overdue",'
                'IF({due}{r}<=TodayDate+7,"Due soon","Pending"))))'),
            cf=status_cf()),
        Col("budget", "Budget this month", 13, "money",
            formula=blank(
                "date",
                'IFERROR(INDEX(BudgetAmount,MATCH({category}{r},'
                'BudgetCategory,0))*1,0)'),
            hidden=True,
            note="The monthly budget for this category, from the "
                 "Goals & Budget tab."),
        rank_desc("total", guard='{date}{r}=""', header="Size rank"),
        Col("alert", "Check", 26, "status",
            formula=blank(
                "date",
                'IF({amount}{r}="","Add - no amount",'
                'IF({category}{r}="","Add - no category, it will miss '
                'the P&L",'
                'IF({status}{r}="Overdue","Overdue - pay it",'
                'IF({status}{r}="Due soon","Due - coming up",'
                'IF(IF({recurring}{r}="Yes",{frequency}{r}="",FALSE),'
                '"Add - how often does it recur?",'
                'IF({paid}{r}="Yes","OK","Pending"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Expenses",
        subtitle="Every cost, categorised the way the profit and loss "
                 "needs it, with the recurring ones spread properly.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).expenses,
        help_text="Date, category, amount. Mark the regulars as "
                  "recurring and give them a frequency - that is what "
                  "makes the overhead figure believable.",
        extra_notes=[
            "Wages do not belong here. They come from the payroll tab "
            "so they are counted once and only once - the profit and "
            "loss picks them up from there.",
            "Stock bought for resale is a purchase, not an expense "
            "while it sits on the shelf. Log it on the Suppliers & "
            "Purchases tab; the profit and loss counts the cost of "
            "what you actually sold.",
            "The monthly equivalent column is what turns a pile of "
            "bills into a break-even number. Rent plus insurance plus "
            "software plus the loan, spread evenly, is what you have "
            "to earn before anything is yours.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Expense summary under the grid."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "What it costs to open the doors", last)
    sheet.table_head(row, 0, [
        "Costs logged", "Total", "This month", "Fixed monthly overhead",
        "Unpaid", "Overdue", "Biggest single cost", "Biggest category"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Expenses.date]],"<>")',
        '=ROUND(SUM([[Expenses.amount]]),2)',
        '=ROUND(SUMIFS([[Expenses.amount]],[[Expenses.month]],'
        'ThisMonthKey),2)',
        '=ROUND(SUM([[Expenses.monthly]]),2)',
        '=ROUND(SUMIFS([[Expenses.total]],[[Expenses.paid]],"<>Yes"),2)',
        '=COUNTIFS([[Expenses.status]],"Overdue")',
        '=ROUND(MAX([[Expenses.amount]]),2)',
        '=IFERROR(INDEX([[Expenses.category]],MATCH(MAX('
        '[[Expenses.total]]),[[Expenses.total]],0)),"")',
    ], ["int", "money", "money", "money", "money", "int", "money",
        "text"])
````

### `salon_business_spreadsheet/sheets/inventory.py`

*267 lines · 2 functions*

> `Inventory` - retail stock and backbar, counted the same way.
>
> Stock on hand = opening + received - used in services - sold - waste
> + adjustments. Every one of those columns is visible, so when the
> count is wrong you can see which part of the story is wrong rather
> than typing over the answer.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Stock summary under the grid |

**Module constants:** `NAME` ('Inventory').

````python
"""`Inventory` - retail stock and backbar, counted the same way.

Stock on hand = opening + received - used in services - sold - waste
+ adjustments. Every one of those columns is visible, so when the
count is wrong you can see which part of the story is wrong rather
than typing over the answer.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, rank_desc, round_price,
                     status_cf, yesno)

NAME = "Inventory"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "ProductPrefix", digits=3, width=10),
        Col("name", "Product", 28, "text"),
        Col("brand", "Brand", 16, "text"),
        Col("category", "Category", 24, "text",
            validate=list_from("ProductCategoryList"),
            note="Retail lines and backbar lines are both welcome "
                 "here - the category keeps them apart."),
        Col("size", "Size", 10, "text",
            note="250ml, 100g, a tube, a box of 50."),
        Col("sku", "SKU / barcode", 15, "text"),
        Col("supplier", "Supplier", 20, "text",
            validate=list_from("SupplierList"),
            note="From the supplier list at the bottom of the "
                 "Suppliers & Purchases tab."),
        Col("unit_cost", "Cost each", 11, "money",
            validate=number_check(0, 100000),
            note="What you pay, per unit, delivered."),
        Col("retail_price", "Retail price", 11, "money",
            validate=number_check(0, 100000),
            note="What the client pays. Leave blank for backbar "
                 "stock that is never sold."),
        Col("opening", "Opening stock", 12, "number1",
            validate=number_check(0, 1000000),
            note="What was on the shelf when you started using this "
                 "workbook. Set it once; after that let the movements "
                 "do the work."),
        Col("used", "Used in services", 13, "number1",
            validate=number_check(0, 1000000),
            note="Backbar usage - tubes of colour, lash trays, wax "
                 "pots. Count it monthly rather than per client."),
        Col("waste", "Waste", 9, "number1",
            validate=number_check(0, 1000000),
            note="Spilt, expired, broken, given away."),
        Col("adjust", "Adjustment", 11, "number1",
            validate=number_check(-1000000, 1000000),
            note="The honest column. Plus or minus whatever it takes "
                 "to match the real count, with a note saying why."),
        Col("reorder_level", "Reorder at", 11, "number1",
            validate=number_check(0, 1000000),
            note="The point at which you want to be told. Blank "
                 "means the workbook works one out from how fast it "
                 "sells."),
        Col("reorder_qty", "Order quantity", 12, "number1",
            validate=number_check(0, 1000000)),
        Col("counted_on", "Last counted", 12, "date",
            validate=date_check()),
        Col("expires", "Expires", 11, "date", validate=date_check(),
            note="For developer, lash adhesive, anything with a "
                 "shelf life."),
        Col("retail_item", "Sold to clients?", 13, "flag",
            validate=yesno(),
            note="No for pure backbar stock, so it stays out of the "
                 "retail figures."),
        Col("notes", "Notes", 22, "wrap"),
        # ---------------------------------------------------- worked out --
        Col("received", "Received", 10, "number1",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Purchases.qty_in]],'
                '[[Purchases.product]],{id}{r}),2)'),
            note="Everything booked in on the purchases tab."),
        Col("sold", "Sold", 9, "number1",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.qty_used]],[[Sales.item]],'
                '{id}{r},[[Sales.kind]],"Retail"),2)'),
            note="Units that went out through the till."),
        Col("stock", "Stock on hand", 12, "number1",
            formula=blank(
                "name",
                'ROUND(IF({opening}{r}="",0,{opening}{r})+'
                '{received}{r}-IF({used}{r}="",0,{used}{r})-'
                '{sold}{r}-IF({waste}{r}="",0,{waste}{r})+'
                'IF({adjust}{r}="",0,{adjust}{r}),2)'),
            cf=[{"type": "cell", "criteria": "<=", "value": 0,
                 "fmt": "bad"}],
            note="Opening + received - used - sold - waste + "
                 "adjustments. Nothing hidden."),
        Col("weekly_use", "Used a week", 11, "number1",
            formula=blank(
                "name",
                'ROUND(({sold}{r}+IF({used}{r}="",0,{used}{r}))/'
                'MAX(4,TrackingWeeks),2)'),
            hidden=True,
            note="Average weekly movement since you started tracking."),
        Col("weeks_cover", "Weeks left", 10, "number1",
            formula=blank(
                "name",
                'IF({weekly_use}{r}<=0,"",'
                'ROUND({stock}{r}/{weekly_use}{r},1))'),
            cf=[{"type": "cell", "criteria": "<",
                 "value": "=StockWarningWeeks", "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": 6,
                 "fmt": "warn"}],
            note="At the current rate. Under the warning level on "
                 "Settings and it goes red."),
        Col("reorder_used", "Reorder point", 11, "number1",
            formula=blank(
                "name",
                'IF({reorder_level}{r}<>"",{reorder_level}{r},'
                'ROUND({weekly_use}{r}*StockWarningWeeks,1))'),
            hidden=True),
        Col("status", "Stock status", 12, "status",
            formula=blank(
                "name",
                'IF({stock}{r}<=0,"Out of stock",'
                'IF({stock}{r}<={reorder_used}{r},"Low stock",'
                'IF({retail_item}{r}="No","Backbar","In stock")))'),
            cf=status_cf()),
        Col("order_qty", "Order this many", 12, "number1",
            formula=blank(
                "name",
                'IF({status}{r}="In stock","",IF({status}{r}="Backbar",'
                'IF({stock}{r}<={reorder_used}{r},'
                'MAX(1,IF({reorder_qty}{r}="",'
                'ROUND({weekly_use}{r}*ReorderWeeks,0),'
                '{reorder_qty}{r})),""),'
                'MAX(1,IF({reorder_qty}{r}="",'
                'ROUND({weekly_use}{r}*ReorderWeeks,0),'
                '{reorder_qty}{r}))))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}],
            note="Your order quantity if you set one, otherwise "
                 "enough to cover the reorder window on Settings."),
        Col("value_cost", "Value at cost", 12, "money",
            formula=blank(
                "name",
                'ROUND({stock}{r}*IF({unit_cost}{r}="",0,'
                '{unit_cost}{r}),2)'),
            cf=[{"type": "data_bar", "bar_color": Palette.primary_mid}],
            note="What is sitting on the shelf, in money. Usually a "
                 "bigger number than anybody expects."),
        Col("value_retail", "Value at retail", 12, "money",
            formula=blank(
                "name",
                'ROUND({stock}{r}*IF({retail_price}{r}="",0,'
                '{retail_price}{r}),2)'),
            hidden=True),
        Col("margin", "Margin", 9, "percent",
            formula=blank(
                "name",
                'IF(IF({retail_price}{r}="",0,{retail_price}{r})=0,"",'
                '({retail_price}{r}-IF({unit_cost}{r}="",0,'
                '{unit_cost}{r}))/{retail_price}{r})'),
            cf=[{"type": "cell", "criteria": "<", "value": 0.3,
                 "fmt": "warn"},
                {"type": "cell", "criteria": ">=", "value": 0.4,
                 "fmt": "good"}]),
        Col("suggested_raw", "Suggested raw", 11, "money",
            formula=blank(
                "name",
                'IF({unit_cost}{r}="","",{unit_cost}{r}*RetailMarkup)'),
            hidden=True),
        Col("suggested", "Suggested price", 12, "money",
            formula=blank(
                "name",
                'IF({suggested_raw}{r}="","",ROUND(' +
                round_price("{suggested_raw}{r}") + ',2))'),
            note="Cost x the retail markup on Settings, rounded by "
                 "your rounding rule."),
        Col("revenue", "Retail revenue", 12, "money",
            formula=blank(
                "name",
                'ROUND(SUMIFS([[Sales.net]],[[Sales.item]],{id}{r},'
                '[[Sales.kind]],"Retail"),2)')),
        Col("profit", "Retail profit", 11, "money",
            formula=blank(
                "name",
                'ROUND({revenue}{r}-{sold}{r}*IF({unit_cost}{r}="",0,'
                '{unit_cost}{r}),2)'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"}],
            hidden=True),
        rank_desc("revenue", header="Seller rank"),
        Col("alert", "Action", 28, "status",
            formula=blank(
                "name",
                'IF({unit_cost}{r}="","Add - no cost, no stock value",'
                'IF({status}{r}="Out of stock",'
                '"Out of stock - order today",'
                'IF({status}{r}="Low stock","Order - running low",'
                'IF(IF({expires}{r}="",FALSE,'
                '{expires}{r}<=TodayDate+30),'
                '"Check - expiring within a month",'
                'IF(IF({retail_item}{r}="Yes",'
                '{margin}{r}<0.25,FALSE),'
                '"Below - thin margin for a retail line",'
                'IF(IF({counted_on}{r}="",TRUE,'
                'TodayDate-{counted_on}{r}>90),'
                '"Due - not counted in three months","OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Product Inventory",
        subtitle="Retail and backbar stock, what is left, what it is "
                 "worth and what to order.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).inventory,
        help_text="Set the opening stock once. After that, purchases "
                  "and sales move the numbers; you only type the "
                  "backbar usage, the waste and the odd adjustment.",
        extra_notes=[
            "Backbar stock is tracked the same way as retail, it just "
            "leaves by a different door: 'used in services' instead "
            "of 'sold'. Count it once a month and the weeks-left "
            "column does the rest.",
            "The adjustment column exists because real counts never "
            "match. Use it, write why in the notes, and keep the "
            "movement columns honest.",
            "Order quantity is a suggestion, not an order. Nothing is "
            "sent anywhere - this workbook has no connection to your "
            "supplier.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Stock summary under the grid."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "The stock room", last)
    sheet.table_head(row, 0, [
        "Products listed", "Value at cost", "Value at retail",
        "Out of stock", "Low stock", "To reorder", "Retail revenue",
        "Average retail margin", "Expiring in 30 days"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Inventory.name]],"<>")',
        '=ROUND(SUM([[Inventory.value_cost]]),2)',
        '=ROUND(SUM([[Inventory.value_retail]]),2)',
        '=COUNTIFS([[Inventory.status]],"Out of stock")',
        '=COUNTIFS([[Inventory.status]],"Low stock")',
        '=ROUND(SUM([[Inventory.order_qty]]),0)',
        '=ROUND(SUM([[Inventory.revenue]]),2)',
        '=IFERROR(SUMPRODUCT(IFERROR([[Inventory.margin]]*1,0))/'
        'MAX(1,COUNTIFS([[Inventory.retail_item]],"Yes")),0)',
        '=COUNTIFS([[Inventory.expires]],">0",'
        '[[Inventory.expires]],"<="&TodayDate+30)',
    ], ["int", "money", "money", "int", "int", "number1", "money",
        "percent", "int"])
````

### `salon_business_spreadsheet/sheets/purchases.py`

*240 lines · 2 functions*

> `Purchases` - stock orders, and the supplier book underneath them.
>
> Top half: every order you place, what it cost, whether it turned up.
> Bottom half: an editable supplier directory that feeds the supplier
> drop-down on this tab and on the inventory tab. Add a wholesaler
> there and it appears in the lists straight away.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Order summary, then the editable supplier directory |

**Module constants:** `NAME` ('Purchases'), `SUPPLIER_COLUMNS` (9 items), `SUPPLIER_ROWS` (24).

````python
"""`Purchases` - stock orders, and the supplier book underneath them.

Top half: every order you place, what it cost, whether it turned up.
Bottom half: an editable supplier directory that feeds the supplier
drop-down on this tab and on the inventory tab. Add a wholesaler
there and it appears in the lists straight away.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec, dynamic_name, quote_sheet
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, lookup_product, month_col,
                     status_cf, yesno)

NAME = "Purchases"

#: Supplier directory written under the order grid.
SUPPLIER_COLUMNS = [
    ("Supplier", 24, "text"),
    ("Contact", 18, "text"),
    ("Phone", 16, "text"),
    ("Email", 26, "text"),
    ("Account number", 16, "text"),
    ("Minimum order", 14, "money"),
    ("Lead days", 11, "int"),
    ("Pays by", 14, "text"),
    ("Notes", 30, "text"),
]
SUPPLIER_ROWS = 24


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("date", "PurchasePrefix", digits=4, width=11),
        Col("date", "Ordered", 11, "date", validate=date_check()),
        Col("supplier", "Supplier", 22, "text",
            validate=list_from("SupplierList"),
            note="From the supplier book at the bottom of this tab."),
        Col("product", "Product ID", 11, "id",
            validate=list_from("ProductList"),
            note="The inventory ID. One line per product, as on the "
                 "invoice."),
        Col("product_name", "Product", 24, "text",
            formula=lookup_product("name")),
        Col("qty", "Quantity", 10, "number1",
            validate=number_check(0, 100000)),
        Col("unit_cost", "Cost each", 11, "money",
            validate=number_check(0, 100000),
            note="What this order charged. Leave blank to use the "
                 "cost on the inventory tab."),
        Col("shipping", "Delivery", 10, "money",
            validate=number_check(0, 100000)),
        Col("tax", "Tax", 10, "money", validate=number_check(0, 100000)),
        Col("status", "Status", 14, "text",
            validate=list_from("PurchaseStatusList"), cf=status_cf(),
            note="Nothing counts into stock until this says Received "
                 "or Part received."),
        Col("expected", "Expected", 11, "date", validate=date_check()),
        Col("received_on", "Received", 11, "date", validate=date_check()),
        Col("qty_received", "Quantity in", 11, "number1",
            validate=number_check(0, 100000),
            note="Only needed when a delivery is short. Blank means "
                 "all of it arrived."),
        Col("invoice", "Invoice ref", 14, "text"),
        Col("paid", "Paid?", 9, "flag", validate=yesno()),
        Col("notes", "Notes", 22, "wrap"),
        # ---------------------------------------------------- worked out --
        month_col("date"),
        Col("cost_used", "Cost used", 11, "money",
            formula=blank(
                "date",
                'IF({unit_cost}{r}<>"",{unit_cost}{r},'
                'IFERROR(INDEX([[Inventory.unit_cost]],MATCH('
                '{product}{r},[[Inventory.id]],0))*1,0))'),
            hidden=True),
        Col("goods", "Goods", 11, "money",
            formula=blank(
                "date",
                'ROUND(IF({qty}{r}="",0,{qty}{r})*{cost_used}{r},2)')),
        Col("total", "Order total", 11, "money",
            formula=blank(
                "date",
                'ROUND({goods}{r}+IF({shipping}{r}="",0,'
                '{shipping}{r})+IF({tax}{r}="",0,{tax}{r}),2)'),
            cf=[{"type": "data_bar", "bar_color": Palette.primary_mid}]),
        Col("qty_in", "Into stock", 11, "number1",
            formula=blank(
                "date",
                'IF({status}{r}="Received",'
                'IF({qty_received}{r}<>"",{qty_received}{r},'
                'IF({qty}{r}="",0,{qty}{r})),'
                'IF({status}{r}="Part received",'
                'IF({qty_received}{r}="",0,{qty_received}{r}),0))'),
            note="What the inventory tab counts as received. Draft "
                 "and ordered lines add nothing until they land."),
        Col("short", "Short by", 10, "number1",
            formula=blank(
                "date",
                'IF({status}{r}="Ordered","",'
                'IF({status}{r}="Draft","",'
                'ROUND(IF({qty}{r}="",0,{qty}{r})-{qty_in}{r},2)))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}],
            hidden=True),
        Col("lead_days", "Days to arrive", 12, "int",
            formula=blank(
                "date",
                'IF({received_on}{r}="","",{received_on}{r}-{date}{r})'),
            hidden=True),
        Col("price_change", "Cost moved", 11, "money",
            formula=blank(
                "date",
                'IF({unit_cost}{r}="","",IFERROR(ROUND({unit_cost}{r}-'
                'INDEX([[Inventory.unit_cost]],MATCH({product}{r},'
                '[[Inventory.id]],0)),2),""))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "good"}],
            note="This order against the cost on file. A positive "
                 "number is a price rise you may not have noticed."),
        Col("alert", "Check", 28, "status",
            formula=blank(
                "date",
                'IF({product}{r}="","Add - which product?",'
                'IF({status}{r}="","Add - no status yet",'
                'IF(IF({status}{r}="Ordered",IF({expected}{r}="",FALSE,'
                '{expected}{r}<TodayDate),FALSE),'
                '"Late - expected and not arrived",'
                'IF({short}{r}>0,"Check - delivery came up short",'
                'IF(IF({price_change}{r}="",FALSE,'
                '{price_change}{r}>0),"Check - cost has gone up",'
                'IF({paid}{r}<>"Yes","Due - invoice not paid",'
                '"OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Suppliers & Purchases",
        subtitle="What you ordered, what it cost, whether it turned "
                 "up - and the supplier book underneath.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).purchases,
        help_text="Log the order when you place it and set the status "
                  "to Received when it lands. Only received quantities "
                  "reach the stock figures.",
        extra_notes=[
            "Cost moved is the column that pays for this workbook. "
            "Wholesale prices creep up quietly; this says by how "
            "much, on which line, on which order.",
            "Short deliveries: set the status to Part received and "
            "type what actually arrived in 'quantity in'. The "
            "difference shows as short, and the stock stays honest.",
            "The supplier book at the bottom of this tab is a plain "
            "editable list. Whatever you type in the first column "
            "appears in the supplier drop-downs here and on the "
            "inventory tab.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Order summary, then the editable supplier directory."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index

    row = sheet.section(row, "Ordering, all time", last)
    sheet.table_head(row, 0, [
        "Orders", "Goods", "Delivery", "Order total", "Received",
        "Still on order", "Unpaid invoices", "Late deliveries",
        "Lines where cost rose"])
    row = sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Purchases.date]],"<>")',
        '=ROUND(SUM([[Purchases.goods]]),2)',
        '=ROUND(SUM([[Purchases.shipping]]),2)',
        '=ROUND(SUM([[Purchases.total]]),2)',
        '=ROUND(SUMIFS([[Purchases.total]],[[Purchases.status]],'
        '"Received"),2)',
        '=ROUND(SUMIFS([[Purchases.total]],[[Purchases.status]],'
        '"Ordered"),2)',
        '=ROUND(SUMIFS([[Purchases.total]],[[Purchases.paid]],'
        '"<>Yes"),2)',
        '=COUNTIFS([[Purchases.alert]],"Late*")',
        '=COUNTIFS([[Purchases.price_change]],">0")',
    ], ["int", "money", "money", "money", "money", "money", "money",
        "int", "int"])

    # ------------------------------------------------- supplier directory --
    row += 1
    row = sheet.section(row, "Supplier book", last)
    row = sheet.note(
        row,
        "Type your wholesalers here. The first column feeds the supplier "
        "drop-down on this tab and on the inventory tab, so add a new "
        "one and it is immediately available everywhere. Blank rows are "
        "ignored.",
        last, height=30)
    row += 1

    headers = [h for h, _, _ in SUPPLIER_COLUMNS]
    widths = [w for _, w, _ in SUPPLIER_COLUMNS]
    sheet.table_head(row, 0, headers, widths)
    first = row + 1                      # 0 based first input row
    rows = sample(cfg).suppliers if sample(cfg).suppliers else []
    for offset in range(SUPPLIER_ROWS):
        data = rows[offset] if offset < len(rows) else {}
        for index, (header, _width, style) in enumerate(SUPPLIER_COLUMNS):
            key = header.lower().replace(" ", "_")
            sheet.write(first + offset, index, data.get(key),
                        fmts.input(style))

    sheet_ref = quote_sheet(NAME)
    sheet.book.define_name(
        "SupplierList",
        dynamic_name(NAME, "A", first + 1, first + SUPPLIER_ROWS))
    sheet.book.define_name(
        "SupplierLeadDays",
        f"={sheet_ref}!$G${first + 1}:$G${first + SUPPLIER_ROWS}")

    # spend per supplier, so the book is worth reading
    row = first + SUPPLIER_ROWS + 1
    sheet.table_head(row, 0, ["Suppliers listed", "Spend, all time",
                              "Spend this month", "Biggest supplier"])
    sheet.calc_row(row + 1, 0, [
        f'=COUNTIF({sheet_ref}!$A${first + 1}:$A${first + SUPPLIER_ROWS},'
        f'"?*")',
        '=ROUND(SUM([[Purchases.total]]),2)',
        '=ROUND(SUMIFS([[Purchases.total]],[[Purchases.month]],'
        'ThisMonthKey),2)',
        '=IFERROR(INDEX([[Purchases.supplier]],MATCH(MAX('
        '[[Purchases.total]]),[[Purchases.total]],0)),"")',
    ], ["int", "money", "money", "text"])
````

---

## 7. Growth

Campaigns measured against what they really cost - the money out and the discount given away - and a loyalty scheme whose points come from spend already logged.

### `salon_business_spreadsheet/sheets/marketing.py`

*205 lines · 2 functions*

> `Marketing` - campaigns, offers and what they actually brought in.
>
> A promo code on a sale is the link. Type the code here, use the same
> code on the till, and this tab counts the bookings, the revenue, the
> discount given away and the new clients - then divides by what the
> campaign cost you.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Marketing summary under the grid |

**Module constants:** `NAME` ('Marketing').

````python
"""`Marketing` - campaigns, offers and what they actually brought in.

A promo code on a sale is the link. Type the code here, use the same
code on the till, and this tab counts the bookings, the revenue, the
discount given away and the new clients - then divides by what the
campaign cost you.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, month_col, rank_desc,
                     status_cf)

NAME = "Marketing"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("name", "CampaignPrefix", digits=3, width=10),
        Col("name", "Campaign", 26, "text",
            note="What you would call it out loud - 'January colour "
                 "offer', 'refer a friend', 'Tuesday blow dry'."),
        Col("channel", "Channel", 18, "text",
            validate=list_from("MarketingChannelList")),
        Col("offer", "The offer", 28, "text",
            note="20% off a first colour, free treatment with a cut, "
                 "two for one on Tuesdays."),
        Col("code", "Promo code", 13, "text",
            note="The code you type on the Sales tab. This is what "
                 "links the money back to the campaign - keep it "
                 "short and keep it unique."),
        Col("start", "Starts", 11, "date", validate=date_check()),
        Col("end", "Ends", 11, "date", validate=date_check()),
        Col("spend", "Cost", 11, "money", validate=number_check(0, 1000000),
            note="Ads, print, the influencer's treatment at cost, the "
                 "hour you spent making the reel."),
        Col("target_rev", "Revenue target", 13, "money",
            validate=number_check(0, 1000000)),
        Col("status", "Status", 12, "text",
            validate=list_from("CampaignStatusList"), cf=status_cf()),
        Col("notes", "What happened", 28, "wrap",
            note="Worth a sentence while it is fresh. Next January "
                 "you will want to know whether it was worth "
                 "repeating."),
        # ---------------------------------------------------- worked out --
        month_col("start"),
        Col("uses", "Times used", 11, "int",
            formula=blank(
                "name",
                'IF({code}{r}="",0,COUNTIFS([[Sales.promo]],'
                '{code}{r}))'),
            note="Sales lines carrying this code."),
        Col("revenue", "Revenue", 12, "money",
            formula=blank(
                "name",
                'IF({code}{r}="",0,ROUND(SUMIFS([[Sales.net]],'
                '[[Sales.promo]],{code}{r}),2))'),
            cf=[{"type": "data_bar", "bar_color": Palette.primary_mid}],
            note="Net sales on those lines - after the discount the "
                 "offer gave away."),
        Col("given", "Discount given", 13, "money",
            formula=blank(
                "name",
                'IF({code}{r}="",0,ROUND(SUMIFS([[Sales.discount]],'
                '[[Sales.promo]],{code}{r}),2))'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}]),
        Col("clients", "Clients used it", 12, "int",
            formula=blank(
                "name",
                'IF({code}{r}="",0,SUMPRODUCT(([[Sales.promo]]='
                '{code}{r})*([[Sales.first_of_appt]]="Yes")))'),
            hidden=True),
        Col("new_clients", "New clients", 11, "int",
            formula=blank(
                "name",
                'IF({start}{r}="",0,COUNTIFS([[Clients.source]],'
                '{channel}{r},[[Clients.first_visit]],">="&{start}{r},'
                '[[Clients.first_visit]],"<="&IF({end}{r}="",'
                'TodayDate,{end}{r})))'),
            note="Clients whose 'found you by' matches this channel "
                 "and whose first visit falls inside the dates."),
        Col("bookings", "Bookings", 10, "int",
            formula=blank(
                "name",
                'IF({start}{r}="",0,COUNTIFS([[Appointments.source]],'
                '{channel}{r},[[Appointments.date]],">="&{start}{r},'
                '[[Appointments.date]],"<="&IF({end}{r}="",TodayDate,'
                '{end}{r})))'),
            hidden=True),
        Col("total_cost", "Total cost", 11, "money",
            formula=blank(
                "name",
                'ROUND(IF({spend}{r}="",0,{spend}{r})+{given}{r},2)'),
            note="What you paid out plus what you gave away. Both are "
                 "the cost of the offer."),
        Col("profit", "Left over", 11, "money",
            formula=blank(
                "name",
                'ROUND({revenue}{r}-{total_cost}{r},2)'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "good"}]),
        Col("roi", "Return", 10, "percent",
            formula=blank(
                "name",
                'IF({total_cost}{r}<=0,"",'
                '{profit}{r}/{total_cost}{r})'),
            cf=[{"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "bad"},
                {"type": "cell", "criteria": ">=", "value": 1,
                 "fmt": "good"}],
            note="What you got back for every pound or dollar the "
                 "campaign cost. 100% means it doubled its money."),
        Col("cost_per_client", "Cost per new client", 14, "money",
            formula=blank(
                "name",
                'IF({new_clients}{r}=0,"",'
                'ROUND({total_cost}{r}/{new_clients}{r},2))'),
            cf=[{"type": "cell", "criteria": ">", "value": "=TargetTicket",
                 "fmt": "warn"}],
            note="Over the price of an average visit means the "
                 "campaign only pays if they come back - which is "
                 "exactly what the retention tab is for."),
        Col("vs_target", "Against target", 12, "money",
            formula=blank(
                "name",
                'IF({target_rev}{r}="","",'
                'ROUND({revenue}{r}-{target_rev}{r},2))'),
            cf=[{"type": "cell", "criteria": ">=", "value": 0,
                 "fmt": "good"},
                {"type": "cell", "criteria": "<", "value": 0,
                 "fmt": "warn"}],
            hidden=True),
        rank_desc("revenue", header="Best earner"),
        Col("alert", "Verdict", 28, "status",
            formula=blank(
                "name",
                'IF({code}{r}="","Add - no code, nothing to measure",'
                'IF({status}{r}="Planned","Upcoming",'
                'IF({uses}{r}=0,"Check - code never used",'
                'IF({profit}{r}<0,"Loss - cost more than it made",'
                'IF({roi}{r}>=1,"OK - doubled its money",'
                'IF({roi}{r}>0,"OK - made money","Check - review it"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Marketing & Promotions",
        subtitle="Offers, codes and campaigns, measured against what "
                 "they cost - money out and discount given away.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.accent,
        freeze_cols=2,
        samples=sample(cfg).marketing,
        help_text="Give every offer a code and use that code on the "
                  "Sales tab. Without the code, nothing can be traced "
                  "back here.",
        extra_notes=[
            "The discount you gave away is counted as part of the "
            "cost. A full column of half-price blow dries is not free "
            "marketing, and this tab will not pretend it was.",
            "New clients are matched by channel and date, so keep the "
            "'found you by' column on the Clients tab filled in. It "
            "takes two seconds at the basin and it is the only way "
            "the number means anything.",
            "Cost per new client against your average ticket is the "
            "honest test of an offer. If it costs more to win them "
            "than they spend on the first visit, the campaign is a "
            "bet on them coming back.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Marketing summary under the grid."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "What the marketing did", last)
    sheet.table_head(row, 0, [
        "Campaigns", "Running now", "Spent", "Discount given",
        "Revenue traced", "Left over", "New clients", "Average return",
        "Cost per new client"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Marketing.name]],"<>")',
        '=COUNTIFS([[Marketing.status]],"Running")',
        '=ROUND(SUM([[Marketing.spend]]),2)',
        '=ROUND(SUM([[Marketing.given]]),2)',
        '=ROUND(SUM([[Marketing.revenue]]),2)',
        '=ROUND(SUM([[Marketing.profit]]),2)',
        '=SUM([[Marketing.new_clients]])',
        '=IFERROR(SUM([[Marketing.profit]])/'
        'MAX(0.01,SUM([[Marketing.total_cost]])),0)',
        '=IFERROR(ROUND(SUM([[Marketing.total_cost]])/'
        'MAX(1,SUM([[Marketing.new_clients]])),2),0)',
    ], ["int", "int", "money", "money", "money", "money", "int",
        "percent", "money"])
````

### `salon_business_spreadsheet/sheets/loyalty.py`

*190 lines · 2 functions*

> `Loyalty` - points, rewards and referrals, one row per client.
>
> Points are earned on spend at whatever rate you set, redeemed when
> you say so, and the balance is simply the difference. Referrals are
> counted from the Clients tab, so the credit you owe somebody for
> sending a friend is never a matter of memory.

| Defines | What it does |
|---------|--------------|
| `spec(cfg)` |  |
| `decorate(sheet, fmts, cfg, reg)` | Loyalty summary under the grid |

**Module constants:** `NAME` ('Loyalty').

````python
"""`Loyalty` - points, rewards and referrals, one row per client.

Points are earned on spend at whatever rate you set, redeemed when
you say so, and the balance is simply the difference. Referrals are
counted from the Clients tab, so the credit you owe somebody for
sending a friend is never a matter of memory.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Col, SheetSpec
from ..sample_data import get as sample
from ..writer import date_check, list_from, number_check
from .common import (alert_cf, blank, id_col, lookup_client, rank_desc,
                     status_cf, yesno)

NAME = "Loyalty"


def spec(cfg: Config) -> SheetSpec:
    columns = [
        id_col("client", "LoyaltyPrefix", digits=4, width=11),
        Col("client", "Client ID", 11, "id",
            validate=list_from("ClientList"),
            note="One row per client on the scheme."),
        Col("name", "Client", 22, "text",
            formula=lookup_client("name", key="client")),
        Col("joined", "Joined", 11, "date", validate=date_check()),
        Col("card", "Card number", 14, "text",
            note="Only if you use physical cards. Otherwise leave it "
                 "blank - the client ID is enough."),
        Col("bonus_points", "Bonus points", 12, "int",
            validate=number_check(0, 1000000),
            note="Points you have handed out by hand - a birthday "
                 "gift, an apology, a launch promotion."),
        Col("redeemed", "Points redeemed", 14, "int",
            validate=number_check(0, 1000000),
            note="Total points spent so far. Add to this number each "
                 "time a reward is taken."),
        Col("last_reward", "Last reward", 11, "date",
            validate=date_check()),
        Col("reward_type", "Usual reward", 20, "text",
            validate=list_from("LoyaltyRewardList")),
        Col("referral_paid", "Referral credit paid", 15, "money",
            validate=number_check(0, 100000),
            note="What you have already given them for sending "
                 "people in."),
        Col("active", "On the scheme?", 13, "flag", validate=yesno()),
        Col("notes", "Notes", 24, "wrap"),
        # ---------------------------------------------------- worked out --
        Col("spend", "Spend counted", 12, "money",
            formula=blank(
                "client",
                'ROUND(SUMIFS([[Sales.net]],[[Sales.client]],'
                '{client}{r},[[Sales.date]],">="&IF({joined}{r}="",0,'
                '{joined}{r})),2)'),
            note="Everything they have spent since joining - services "
                 "and retail."),
        Col("visits", "Visits", 8, "int",
            formula=blank(
                "client",
                'IFERROR(INDEX([[Clients.visits]],MATCH({client}{r},'
                '[[Clients.id]],0))*1,0)'),
            hidden=True),
        Col("earned", "Points earned", 12, "int",
            formula=blank(
                "client",
                'ROUND({spend}{r}*PointsPerUnit+IF({bonus_points}{r}='
                '"",0,{bonus_points}{r}),0)'),
            cf=[{"type": "data_bar", "bar_color": Palette.primary_mid}],
            note="Spend x the points rate on Settings, plus any "
                 "bonus points."),
        Col("balance", "Points balance", 12, "int",
            formula=blank(
                "client",
                '{earned}{r}-IF({redeemed}{r}="",0,{redeemed}{r})'),
            cf=[{"type": "cell", "criteria": ">=", "value": "=RewardAt",
                 "fmt": "good"}],
            note="Earned less redeemed. This is the number the client "
                 "cares about."),
        Col("rewards_due", "Rewards due", 11, "int",
            formula=blank(
                "client",
                'IF(RewardAt<=0,0,ROUNDDOWN({balance}{r}/RewardAt,0))'),
            hidden=True),
        Col("reward_value", "Reward owed", 12, "money",
            formula=blank(
                "client",
                'ROUND({rewards_due}{r}*RewardValue,2)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}],
            note="What those points are worth if they cash them in "
                 "today. Worth knowing before you run a double "
                 "points week."),
        Col("to_next", "Points to next", 12, "int",
            formula=blank(
                "client",
                'IF(RewardAt<=0,"",MAX(0,RewardAt*({rewards_due}{r}+1)-'
                '{balance}{r}))'),
            note="The number to tell them at the desk. 'Forty more "
                 "points and your next treatment is on us.'"),
        Col("referrals", "People referred", 13, "int",
            formula=blank(
                "client",
                'COUNTIFS([[Clients.referred_by]],{client}{r})'),
            note="Counted from the 'referred by' column on the "
                 "Clients tab."),
        Col("referral_owed", "Referral credit owed", 15, "money",
            formula=blank(
                "client",
                'ROUND({referrals}{r}*ReferralReward-'
                'IF({referral_paid}{r}="",0,{referral_paid}{r}),2)'),
            cf=[{"type": "cell", "criteria": ">", "value": 0,
                 "fmt": "warn"}]),
        Col("segment", "Segment", 13, "text",
            formula=blank(
                "client",
                'IFERROR(INDEX([[Clients.segment]],MATCH({client}{r},'
                '[[Clients.id]],0))&"","")'),
            cf=status_cf(),
            hidden=True),
        rank_desc("balance", guard='{client}{r}=""', header="Points rank"),
        Col("alert", "Action", 30, "status",
            formula=blank(
                "client",
                'IF({active}{r}="No","",'
                'IF(COUNTIF({client_first}:{client_last},{client}{r})>1,'
                '"Duplicate - this client is on twice",'
                'IF({rewards_due}{r}>0,"Action - reward ready to give",'
                'IF({referral_owed}{r}>0.004,'
                '"Action - referral credit owed",'
                'IF({segment}{r}="Lapsing",'
                '"Chase - points are a reason to call",'
                'IF({to_next}{r}<=RewardAt*0.2,'
                '"Rebook - nearly at a reward","OK"))))))'),
            cf=alert_cf()),
    ]

    return SheetSpec(
        name=NAME,
        title="Client Loyalty",
        subtitle="Points, rewards and referral credit - counted from "
                 "the spend you have already logged.",
        columns=columns,
        rows=cfg.row_count(NAME),
        tab_color=Palette.primary_mid,
        freeze_cols=2,
        samples=sample(cfg).loyalty,
        help_text="Add a row per client on the scheme. Points build "
                  "themselves from the Sales tab; you only record "
                  "what has been redeemed.",
        extra_notes=[
            "Set the points rate, the reward threshold and the reward "
            "value on Settings. One point per unit of spend with a "
            "reward at 200 points is a 10% scheme - worth knowing "
            "before you promise it.",
            "The reward owed column is a liability, not a decoration. "
            "It is what you would have to hand over if every client "
            "cashed in tomorrow.",
            "Referral credit is counted from the 'referred by' column "
            "on the Clients tab, so the person who sent three friends "
            "in gets what they were promised without anybody having "
            "to remember.",
        ],
    )


def decorate(sheet, fmts, cfg: Config, reg) -> None:
    """Loyalty summary under the grid."""
    spec_ = reg[NAME]
    row = spec_.last_data_row + 1
    last = spec_.last_col_index
    row = sheet.section(row, "The scheme", last)
    sheet.table_head(row, 0, [
        "Members", "Active", "Points earned", "Points redeemed",
        "Points outstanding", "Rewards ready", "Reward value owed",
        "Referrals", "Referral credit owed"])
    sheet.calc_row(row + 1, 0, [
        '=COUNTIF([[Loyalty.client]],"<>")',
        '=COUNTIFS([[Loyalty.active]],"Yes")',
        '=SUM([[Loyalty.earned]])',
        '=SUM([[Loyalty.redeemed]])',
        '=SUM([[Loyalty.balance]])',
        '=SUM([[Loyalty.rewards_due]])',
        '=ROUND(SUM([[Loyalty.reward_value]]),2)',
        '=SUM([[Loyalty.referrals]])',
        '=ROUND(SUM([[Loyalty.referral_owed]]),2)',
    ], ["int", "int", "int", "int", "int", "int", "money", "int",
        "money"])
````

---

## 8. The pages

The hand-built tabs: the welcome, the dashboard, the profit and loss, the goals and budget, the analytics, the weekly and monthly reports, and the settings tab that every formula reads.

### `salon_business_spreadsheet/sheets/start_here.py`

*259 lines · 1 function*

> `Start Here` - what this is, what to type, and what it will not do.
>
> The first tab a buyer opens. It sets the file up in six steps, maps
> all twenty tabs, and is honest about the limits: no texts, no
> confirmations, no booking platform, no tax advice.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Start Here'), `LAST_COL` (9), `TABS` (19 items), `STEPS` (6 items), `RULES` (7 items), `LIMITS` (5 items).

````python
"""`Start Here` - what this is, what to type, and what it will not do.

The first tab a buyer opens. It sets the file up in six steps, maps
all twenty tabs, and is honest about the limits: no texts, no
confirmations, no booking platform, no tax advice.
"""

from __future__ import annotations

from ..config import DISCLAIMER, Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter

NAME = "Start Here"
LAST_COL = 9

TABS = [
    ("Dashboard", "The month at a glance: takings against target, "
                  "average ticket, utilisation, rebooking, and the "
                  "jobs that need doing today.", "Read it"),
    ("Appointments", "The diary. Date, time, client, service, stylist "
                     "and status - with a double-booking alarm.",
     "Type here"),
    ("Clients", "The client book: contact details, formula, patch "
                "test, spend, last visit, and who is drifting away.",
     "Type here"),
    ("Services", "Your menu, with the minutes, the product cost and "
                 "what each treatment earns an hour.", "Type here"),
    ("Sales", "The till. One row per thing sold - gross, discount, "
              "refund, net, tax and profit.", "Type here"),
    ("Payments", "Money in and out: deposits, balances, tips, refunds "
                 "and the card fees.", "Type here"),
    ("Staff", "The team, their pay basis and what each chair brings "
              "in an hour.", "Type here"),
    ("Commissions", "One row per stylist per month. Eight methods "
                    "supported - you choose, per person.", "Type here"),
    ("Payroll", "Pay runs: hours, commission released, tips passed "
                "on and the employer costs on top.", "Type here"),
    ("Expenses", "Every cost, categorised the way a profit and loss "
                 "wants it, with the recurring ones spread.",
     "Type here"),
    ("Profit & Loss", "Twelve months of income, cost of sales, staff "
                      "cost, overheads and operating profit.",
     "Read it"),
    ("Inventory", "Retail and backbar stock: what is left, what it is "
                  "worth, what to order.", "Type here"),
    ("Purchases", "Orders placed and received, plus the supplier book "
                  "that feeds the drop-downs.", "Type here"),
    ("Marketing", "Campaigns and offers, measured against what they "
                  "cost - money out and discount given away.",
     "Type here"),
    ("Loyalty", "Points, rewards and referral credit, counted from "
                "the spend already logged.", "Type here"),
    ("Goals & Budget", "This year's targets, the month-by-month plan "
                       "and a budget for every category.",
     "Type here"),
    ("Analytics", "Categories, days, hours, sources and retention - "
                  "the patterns behind the takings.", "Read it"),
    ("Reports", "The week and the month on one printable page.",
     "Read it"),
    ("Settings", "Your salon, your tax rules, your pay model, your "
                 "lists. Set these first.", "Set it up"),
]

STEPS = [
    ("1. Settings first",
     "Open the Settings tab and work down it: salon name, currency, "
     "tax rate and what tax applies to, deposit rule, rounding, the "
     "commission defaults and your targets. Ten minutes here saves "
     "every other tab from guessing."),
    ("2. Put your menu in",
     "On the Services tab, type each treatment with its price, how "
     "many minutes it takes and what the product costs you. The "
     "minutes matter as much as the price - that is what turns a "
     "price list into a profit list."),
    ("3. Add the team",
     "On the Staff tab, add everyone and pick their commission "
     "method. Nothing is assumed: commission only, hourly, a flat "
     "amount per service, tiered, a chair renter who pays you rent, "
     "or none at all."),
    ("4. Bring the clients across",
     "Paste or type your client list - name, phone, email and the "
     "'found you by' column. Everything else fills itself in as you "
     "start logging visits."),
    ("5. Work the diary and the till",
     "Log appointments as you book them and sales as you ring them "
     "up. A deposit goes on the Payments tab against the booking; it "
     "is applied to the bill when the client comes in, so it can "
     "never be counted as revenue twice."),
    ("6. Once a month",
     "Add a commission row per stylist, run the payroll rows, log "
     "the bills on Expenses, count the stock, and read the Profit & "
     "Loss. Twenty minutes, and you know exactly how the month went."),
]

RULES = [
    "Blue cells are yours. Grey cells are formulas - they are locked "
    "so a stray paste cannot break the workbook.",
    "Every sheet is protected with the password 'premium'. Review > "
    "Unprotect Sheet if you want to change the structure.",
    "IDs number themselves from the row. Filter and sort freely, but "
    "do not cut and paste rows around - copy the values instead.",
    "Type dates as real dates. 14/03 or 3/14 depending on where you "
    "are, not '14 march'.",
    "Status is the engine: Completed, Cancelled and No show are the "
    "three final answers an appointment can have. Everything else is "
    "still open and is left out of the completion and cancellation "
    "rates.",
    "Never type over a formula to make a number look right. Find the "
    "row that is wrong instead - the alert columns are there to point "
    "at it.",
    "Add rows by typing in the first empty row of a tab. Every tab "
    "has hundreds already formatted and waiting.",
]

LIMITS = [
    "It does not send confirmations, reminders, texts or emails, and "
    "it does not sync with any booking platform, till or card reader. "
    "Everything in here is typed in or pasted in by you.",
    "It is not accounting, payroll or tax software, and nothing in it "
    "is accounting, payroll or tax advice. It organises your business "
    "so that the conversation with your accountant is a short one.",
    "It does not calculate income tax, national insurance or social "
    "security on wages. It shows what a pay run costs the salon.",
    "It does not take payments or store card details.",
    "Sales tax is whatever you set it to be. There is no country "
    "baked in, and the two switches let you tax retail, services, "
    "both or neither.",
]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary,
                    landscape=False)
    ws = s.ws

    s.banner(cfg.product_name,
             "A complete salon business system in one spreadsheet - "
             "diary, clients, prices, takings, staff pay, stock and "
             "profit.", LAST_COL)
    s.set_col(0, 0, 3)
    s.set_col(1, 1, 26)
    s.set_col(2, LAST_COL, 13)

    row = 4
    row = s.note(
        row,
        "Everything in this workbook is worked out from four tabs you "
        "fill in as you go: the diary, the till, the expenses and the "
        "pay runs. Fill those honestly and the dashboard, the profit "
        "and loss and the reports look after themselves.",
        LAST_COL, height=32)
    row += 1

    # --------------------------------------------------- six steps ----
    row = s.section(row, "Setting up, in six steps", LAST_COL)
    for title, text in STEPS:
        s.write(row, 1, title, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, text, fmts.chrome("note"))
        s.row_height(row, 32)
        row += 1
    row += 1

    # ------------------------------------------------------ the tabs ---
    row = s.section(row, "What is on each tab", LAST_COL)
    row = s.table_head(row, 1, ["Tab", "What it is for", "You"],
                       [22, 60, 12])
    for name, text, kind in TABS:
        if name in reg or name in ("Dashboard", "Profit & Loss",
                                   "Goals & Budget", "Analytics",
                                   "Reports", "Settings"):
            ws.write_url(row, 1, f"internal:{quote_sheet(name)}!A1",
                         fmts.chrome("link"), name)
        else:
            s.write(row, 1, name, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL - 1, text, fmts.chrome("note"))
        s.write(row, LAST_COL, kind, fmts.chrome("label"))
        s.row_height(row, 26)
        row += 1
    row += 1

    # ------------------------------------------------- the house rules --
    row = s.section(row, "Seven rules that keep it working", LAST_COL)
    for rule in RULES:
        s.merge(row, 1, LAST_COL, "-  " + rule, fmts.chrome("note"))
        s.row_height(row, 26)
        row += 1
    row += 1

    # -------------------------------------------------- what it is not --
    row = s.section(row, "What this workbook does not do", LAST_COL)
    for limit in LIMITS:
        s.merge(row, 1, LAST_COL, "-  " + limit, fmts.chrome("note"))
        s.row_height(row, 30)
        row += 1
    row += 1

    # ------------------------------------------------- the numbers used --
    row = s.section(row, "The sums behind the headline numbers",
                    LAST_COL)
    row = s.table_head(row, 1, ["Figure", "How it is worked out"],
                       [26, 70])
    sums = [
        ("Net sales", "Gross, less discounts, less refunds."),
        ("Operating profit",
         "Net sales, less the cost of what was sold, less the full "
         "cost of the pay runs, less the expenses logged."),
        ("Average ticket",
         "Net sales divided by completed appointments that were paid "
         "for."),
        ("Completion rate",
         "Completed appointments divided by the appointments that "
         "reached a final answer - completed, cancelled or no show."),
        ("Cancellation rate",
         "Cancelled appointments over the same final-answer total."),
        ("Client balance",
         "What was billed, less discounts, less every payment "
         "applied - deposits included."),
        ("Commission",
         "Eligible revenue times the rate, by whichever of the eight "
         "methods that stylist is on."),
        ("Service gross profit",
         "Price less the direct costs: the product used and the "
         "commission earned on it."),
        ("Stock on hand",
         "Opening, plus received, less used in services, less sold, "
         "less waste, plus adjustments."),
        ("Budget variance", "Actual spend less the budget. Positive "
                            "means over."),
        ("Repeat client rate",
         "Clients who came back at least twice, over the clients you "
         "have served."),
    ]
    for label, text in sums:
        s.write(row, 1, label, fmts.chrome("label_bold"))
        s.merge(row, 2, LAST_COL, text, fmts.chrome("note"))
        s.row_height(row, 26)
        row += 1
    row += 1

    # ------------------------------------------------------- the small --
    row = s.section(row, "The small print", LAST_COL)
    row = s.note(row, DISCLAIMER, LAST_COL, height=50)
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

### `salon_business_spreadsheet/sheets/dashboard.py`

*438 lines · 2 functions*

> `Dashboard` - the one page that answers "how is the salon doing?"
>
> Sixteen numbers across the top, the jobs that need doing today, who
> is earning what, which services pay best, and twelve months of
> revenue against target. Everything reads from the registers, so it is
> never out of date and there is nothing to refresh.

| Defines | What it does |
|---------|--------------|
| `_by_rank(sheet, rank_key, field, n)` | The nth row of a register by one of its hidden rank columns |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Dashboard'), `LAST_COL` (12), `MONTHS` (12), `TOP_ROWS` (6).

````python
"""`Dashboard` - the one page that answers "how is the salon doing?"

Sixteen numbers across the top, the jobs that need doing today, who
is earning what, which services pay best, and twelve months of
revenue against target. Everything reads from the registers, so it is
never out of date and there is nothing to refresh.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter

NAME = "Dashboard"
LAST_COL = 12
MONTHS = 12
TOP_ROWS = 6


def _by_rank(sheet: str, rank_key: str, field: str, n: int) -> str:
    """The nth row of a register by one of its hidden rank columns."""
    return (f'IFERROR(INDEX([[{sheet}.{field}]],MATCH({n},'
            f'[[{sheet}.{rank_key}]],0)),"")')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws
    money = cfg.currency_symbol

    s.banner("Salon Dashboard",
             "Takings, bookings, clients and chairs - worked out from "
             "the diary, the till and the payroll you have already "
             "filled in.", LAST_COL)
    s.set_col(0, 0, 26)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 13)

    # ----------------------------------------------------------- the header
    row = 4
    s.merge(row, 0, 0, "Salon", fmts.chrome("label_bold"))
    s.merge(row, 1, 3, None, fmts.calc("text"))
    ws.write_formula(row, 1, "=SalonName", fmts.calc("text"))
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

    # ------------------------------------------------------------- the KPIs
    tiles_1 = [
        ("Net sales this month",
         '=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],DashKey),2)',
         "money0", f"target {money}{cfg.revenue_target_month:,.0f}"),
        ("Against target",
         '=IFERROR(SUMIFS([[Sales.net]],[[Sales.month]],DashKey)/'
         'MAX(0.01,RevenueTargetMonth),"")', "percent0",
         "100% means the month is made"),
        ("Operating profit, month",
         "=ROUND(MonthNetSales-MonthCostOfSales-MonthWages-"
         "MonthExpenses,2)", "money0",
         "net sales less stock, wages and costs"),
        ("Profit margin",
         '=IF(MonthNetSales<=0,"",(MonthNetSales-MonthCostOfSales-'
         'MonthWages-MonthExpenses)/MonthNetSales)', "percent",
         "what is left of every dollar"),
    ]
    tiles_2 = [
        ("Clients served, month",
         '=COUNTIFS([[Appointments.month]],DashKey,'
         '[[Appointments.status]],"Completed")', "int",
         "completed appointments"),
        ("Average ticket",
         '=IFERROR(SUMIFS([[Sales.net]],[[Sales.month]],DashKey)/'
         'MAX(1,COUNTIFS([[Appointments.month]],DashKey,'
         '[[Appointments.status]],"Completed")),"")', "money",
         f"target {money}{cfg.target_average_ticket:,.0f}"),
        ("Retail share",
         '=IFERROR(SUMIFS([[Sales.net]],[[Sales.month]],DashKey,'
         '[[Sales.kind]],"Retail")/MAX(0.01,SUMIFS([[Sales.net]],'
         '[[Sales.month]],DashKey)),"")', "percent",
         f"target {cfg.target_retail_percent:.0%}"),
        ("Tips passed on",
         '=ROUND(SUMIFS([[Payments.amount]],[[Payments.month]],DashKey,'
         '[[Payments.kind]],"Tip"),2)', "money0",
         "the team's money, not the salon's"),
    ]
    tiles_3 = [
        ("Completion rate",
         '=IFERROR(COUNTIFS([[Appointments.month]],DashKey,'
         '[[Appointments.status]],"Completed")/'
         'MAX(1,COUNTIFS([[Appointments.month]],DashKey,'
         '[[Appointments.is_final]],"Yes")),"")', "percent",
         "of bookings that reached a final answer"),
        ("Cancellation rate",
         '=IFERROR(COUNTIFS([[Appointments.month]],DashKey,'
         '[[Appointments.status]],"Cancelled")/'
         'MAX(1,COUNTIFS([[Appointments.month]],DashKey,'
         '[[Appointments.is_final]],"Yes")),"")', "percent",
         "cancelled, not no-showed"),
        ("No-show rate",
         '=IFERROR(COUNTIFS([[Appointments.month]],DashKey,'
         '[[Appointments.status]],"No show")/'
         'MAX(1,COUNTIFS([[Appointments.month]],DashKey,'
         '[[Appointments.is_final]],"Yes")),"")', "percent",
         "the most expensive hour in the salon"),
        ("Chair utilisation",
         '=IFERROR(SUMIFS([[Appointments.hours]],'
         '[[Appointments.month]],DashKey,[[Appointments.status]],'
         '"Completed")/MAX(1,ChairHoursWeek*4.33),"")', "percent",
         f"target {cfg.target_utilisation:.0%} of sellable hours"),
    ]
    tiles_4 = [
        ("Rebooking rate",
         '=IFERROR(COUNTIFS([[Appointments.month]],DashKey,'
         '[[Appointments.rebooked]],"Yes")/'
         'MAX(1,COUNTIFS([[Appointments.month]],DashKey,'
         '[[Appointments.status]],"Completed")),"")', "percent",
         f"target {cfg.target_rebook_rate:.0%} leaving with the next "
         f"one booked"),
        ("Repeat client rate",
         '=IFERROR(COUNTIFS([[Clients.retained]],"Yes")/'
         'MAX(1,COUNTIFS([[Clients.visits]],">0")),"")', "percent",
         "clients served who came back at least twice"),
        ("Owed to you",
         '=ROUND(SUMIFS([[Sales.balance]],[[Sales.balance]],">0.004"),2)',
         "money0", "unpaid balances across every client"),
        ("Stock on the shelf",
         '=ROUND(SUM([[Inventory.value_cost]]),2)', "money0",
         "at cost, retail and backbar"),
    ]
    for tiles in (tiles_1, tiles_2, tiles_3, tiles_4):
        for i, (label, formula, style, note) in enumerate(tiles):
            s.kpi(row, i * 3, label, formula, style=style, span=3,
                  note=note)
        row += 4

    # ------------------------------------------------- needs attention ----
    row = s.section(row, "Needs doing today", LAST_COL)
    row = s.table_head(row, 0, ["What", "How many", "What it means",
                                "Where to fix it"],
                       [28, 11, 48, 18])
    alerts = [
        ("Clashes in the diary",
         '=COUNTIF([[Appointments.clash]],"Clash")',
         '="Two bookings for the same stylist at the same time"',
         "Appointments"),
        ("Bookings to confirm",
         '=COUNTIF([[Appointments.alert]],"Confirm*")',
         '="Coming up in the next two days and not confirmed"',
         "Appointments"),
        ("Deposits not taken",
         '=COUNTIF([[Appointments.alert]],"Due - deposit*")',
         '="Bookings over "&TEXT(DepositThreshold,"0")&" with no '
         'deposit against them"', "Appointments"),
        ("Visits with nothing rung up",
         '=COUNTIF([[Appointments.alert]],"Unpaid*")',
         '="Marked completed but no sale was logged"', "Sales"),
        ("Unpaid balances",
         '=COUNTIFS([[Sales.balance]],">0.004")',
         '="Sales where money is still owed"', "Payments"),
        ("Clients lapsing",
         '=COUNTIF([[Clients.segment]],"Lapsing")',
         '="Not in for "&LapsingDays&" days. Call them before they '
         'become inactive"', "Clients"),
        ("Clients due a rebook",
         '=COUNTIF([[Clients.alert]],"Rebook*")',
         '="Past their usual gap with nothing in the book"',
         "Clients"),
        ("Patch tests out of date",
         '=COUNTIF([[Clients.alert]],"Due - patch*")',
         '="Colour clients whose patch test is older than "&'
         'PatchTestDays&" days"', "Clients"),
        ("Products out of stock",
         '=COUNTIF([[Inventory.status]],"Out of stock")',
         '="Nothing left on the shelf to sell or use"', "Inventory"),
        ("Products to reorder",
         '=COUNTIF([[Inventory.status]],"Low stock")',
         '="Below the reorder point or under "&StockWarningWeeks&'
         '" weeks of cover"', "Inventory"),
        ("Deliveries late",
         '=COUNTIF([[Purchases.alert]],"Late*")',
         '="Ordered, expected, still not here"', "Purchases"),
        ("Bills overdue",
         '=COUNTIF([[Expenses.status]],"Overdue")',
         '="Past the due date and not marked paid"', "Expenses"),
        ("Commission to approve",
         '=COUNTIFS([[Commissions.approved]],"<>Yes",'
         '[[Commissions.month]],"<>")',
         '="Calculated and waiting for your sign-off"',
         "Commissions"),
        ("Pay runs not paid",
         '=COUNTIFS([[Payroll.paid]],"<>Yes",'
         '[[Payroll.period_end]],"<>")',
         '="Payroll rows still open"', "Payroll"),
        ("Loyalty rewards ready",
         '=COUNTIF([[Loyalty.alert]],"Action - reward*")',
         '="Clients who have earned something. Tell them"',
         "Loyalty"),
        ("Services below your chair rate",
         '=SUMPRODUCT((IFERROR([[Services.price_hour]]*1,0)>0)*'
         '(IFERROR([[Services.price_hour]]*1,0)<TargetChairRate))',
         '="Earning less than "&TEXT(TargetChairRate,"0")&" an hour '
         'before commission"', "Services"),
    ]
    first_alert = row
    for label, count, message, where in alerts:
        s.write(row, 0, label, fmts.chrome("label_bold"))
        ws.write_formula(row, 1, s.reg.render(count), fmts.calc("int"))
        s.merge(row, 2, LAST_COL - 2, None, fmts.calc("text"))
        ws.write_formula(row, 2, s.reg.render(message), fmts.calc("text"))
        s.merge(row, LAST_COL - 1, LAST_COL, where, fmts.chrome("label"))
        row += 1
    ws.conditional_format(first_alert, 1, row - 1, 1,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("warn")})
    ws.conditional_format(first_alert, 1, row - 1, 1,
                          {"type": "cell", "criteria": "=", "value": 0,
                           "format": fmts.cf("good")})
    row += 1

    # ----------------------------------------------------- the team today --
    row = s.section(row, "Who is earning what", LAST_COL)
    row = s.table_head(row, 0, [
        "Stylist", "Sales this month", "Against target", "Completed",
        "Average ticket", "Retail share", "Utilisation",
        "Rebook rate", "Commission", "Pay this month", "Leaves behind"],
        [20, 14, 13, 11, 13, 12, 12, 12, 12, 13, 13])
    team_first = row
    for n in range(1, 9):
        s.write(row, 0, None, fmts.calc("text"))
        ws.write_formula(
            row, 0,
            s.reg.render('=' + _by_rank("Staff", "rev_total_rank",
                                        "name", n)),
            fmts.calc("text"))
        pairs = [
            (1, "rev_month", "money"),
            (2, "target_gap", "money"),
            (3, "done", "int"),
            (4, "avg_ticket", "money"),
            (5, "retail_share", "percent"),
            (6, "utilisation", "percent"),
            (7, "rebook_rate", "percent"),
            (8, "commission_paid", "money"),
            (9, "cost_month", "money"),
            (10, "margin_month", "money"),
        ]
        for col, key, style in pairs:
            ws.write_formula(
                row, col,
                s.reg.render('=' + _by_rank("Staff", "rev_total_rank",
                                            key, n)),
                fmts.calc(style))
        row += 1
    ws.conditional_format(team_first, 10, row - 1, 10,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    row += 1

    # --------------------------------------------------- best services ----
    row = s.section(row, "What the chairs are selling", LAST_COL)
    row = s.table_head(row, 0, [
        "Most booked", "Times", "Biggest earner", "Revenue",
        "Best an hour", "An hour", "Best clients", "Spent"],
        [24, 9, 24, 12, 24, 11, 22, 12])
    for n in range(1, TOP_ROWS + 1):
        cells = [
            (0, _by_rank("Services", "done_rank", "name", n), "text"),
            (1, _by_rank("Services", "done_rank", "done", n), "int"),
            (2, _by_rank("Services", "revenue_rank", "name", n), "text"),
            (3, _by_rank("Services", "revenue_rank", "revenue", n),
             "money"),
            (4, _by_rank("Services", "profit_hour_rank", "name", n),
             "text"),
            (5, _by_rank("Services", "profit_hour_rank", "price_hour", n),
             "money"),
            (6, _by_rank("Clients", "spend_rank", "name", n), "text"),
            (7, _by_rank("Clients", "spend_rank", "spend", n), "money"),
        ]
        for col, formula, style in cells:
            ws.write_formula(row, col, s.reg.render("=" + formula),
                             fmts.calc(style))
        row += 1
    row += 1

    # ------------------------------------------------- twelve months ------
    row = s.section(row, "The last twelve months", LAST_COL)
    row = s.table_head(row, 0, [
        "Month", "Net sales", "Target", "Against target", "Services",
        "Retail", "Cost of sales", "Wages", "Expenses",
        "Operating profit", "Margin", "Clients served",
        "Average ticket"],
        [12, 13, 12, 13, 12, 11, 13, 12, 12, 14, 10, 12, 13])
    chart_first = row
    for index in range(MONTHS - 1, -1, -1):
        key = (f'TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{index},1),'
               f'"yyyy-mm")')
        label = (f'=TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{index},1),'
                 f'"mmm yy")')
        s.write(row, 0, None, fmts.calc("text"))
        ws.write_formula(row, 0, label, fmts.calc("text"))
        formulas = [
            (1, f'=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{key}),2)',
             "money"),
            (2, "=ROUND(RevenueTargetMonth,2)", "money"),
            (3, f'=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{key})-'
                f'RevenueTargetMonth,2)', "money"),
            (4, f'=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{key},'
                f'[[Sales.kind]],"Service"),2)', "money"),
            (5, f'=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{key},'
                f'[[Sales.kind]],"Retail"),2)', "money"),
            (6, f'=ROUND(SUMIFS([[Sales.cost]],[[Sales.month]],{key}),2)',
             "money"),
            (7, f'=ROUND(SUMIFS([[Payroll.total_cost]],'
                f'[[Payroll.month]],{key}),2)', "money"),
            (8, f'=ROUND(SUMIFS([[Expenses.amount]],[[Expenses.month]],'
                f'{key}),2)', "money"),
            (9, f'=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{key})-'
                f'SUMIFS([[Sales.cost]],[[Sales.month]],{key})-'
                f'SUMIFS([[Payroll.total_cost]],[[Payroll.month]],{key})-'
                f'SUMIFS([[Expenses.amount]],[[Expenses.month]],{key}),2)',
             "money"),
            (10, f'=IF(SUMIFS([[Sales.net]],[[Sales.month]],{key})<=0,"",'
                 f'(SUMIFS([[Sales.net]],[[Sales.month]],{key})-'
                 f'SUMIFS([[Sales.cost]],[[Sales.month]],{key})-'
                 f'SUMIFS([[Payroll.total_cost]],[[Payroll.month]],{key})-'
                 f'SUMIFS([[Expenses.amount]],[[Expenses.month]],{key}))/'
                 f'SUMIFS([[Sales.net]],[[Sales.month]],{key}))',
             "percent"),
            (11, f'=COUNTIFS([[Appointments.month]],{key},'
                 f'[[Appointments.status]],"Completed")', "int"),
            (12, f'=IFERROR(SUMIFS([[Sales.net]],[[Sales.month]],{key})/'
                 f'MAX(1,COUNTIFS([[Appointments.month]],{key},'
                 f'[[Appointments.status]],"Completed")),"")', "money"),
        ]
        for col, formula, style in formulas:
            ws.write_formula(row, col, s.reg.render(formula),
                             fmts.calc(style))
        row += 1
    last_month_row = row - 1
    ws.conditional_format(chart_first, 1, last_month_row, 1,
                          {"type": "data_bar",
                           "bar_color": Palette.primary_mid})
    ws.conditional_format(chart_first, 9, last_month_row, 9,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(chart_first, 3, last_month_row, 3,
                          {"type": "cell", "criteria": ">=", "value": 0,
                           "format": fmts.cf("good")})

    # sparklines beat a chart that Excel 2010 might not draw
    ws.add_sparkline(row, 1, {
        "range": (f"{quote_sheet(NAME)}!$B${chart_first + 1}:"
                  f"$B${last_month_row + 1}"),
        "type": "column", "series_color": Palette.primary_mid})
    s.merge(row, 2, 4, "Net sales, twelve months",
            fmts.chrome("note"))
    ws.add_sparkline(row, 5, {
        "range": (f"{quote_sheet(NAME)}!$J${chart_first + 1}:"
                  f"$J${last_month_row + 1}"),
        "type": "column", "series_color": Palette.accent,
        "negative_points": True})
    s.merge(row, 6, 8, "Operating profit, twelve months",
            fmts.chrome("note"))
    row += 2

    # ------------------------------------------------- month definitions --
    row = s.section(row, "The month in one line", LAST_COL)
    row = s.table_head(row, 0, [
        "Net sales", "Cost of sales", "Gross profit", "Wages",
        "Other costs", "Operating profit", "Margin", "Break-even sales"],
        [14, 14, 14, 13, 13, 15, 10, 15])
    totals_row = row
    s.write(row, 0, None, fmts.calc("money"))
    ws.write_formula(
        row, 0, s.reg.render('=ROUND(SUMIFS([[Sales.net]],'
                             '[[Sales.month]],DashKey),2)'),
        fmts.calc("money"))
    book.define_name("MonthNetSales",
                     f"={quote_sheet(NAME)}!$A${totals_row + 1}")
    ws.write_formula(
        row, 1, s.reg.render('=ROUND(SUMIFS([[Sales.cost]],'
                             '[[Sales.month]],DashKey),2)'),
        fmts.calc("money"))
    book.define_name("MonthCostOfSales",
                     f"={quote_sheet(NAME)}!$B${totals_row + 1}")
    ws.write_formula(row, 2, "=ROUND(MonthNetSales-MonthCostOfSales,2)",
                     fmts.calc("money"))
    ws.write_formula(
        row, 3, s.reg.render('=ROUND(SUMIFS([[Payroll.total_cost]],'
                             '[[Payroll.month]],DashKey),2)'),
        fmts.calc("money"))
    book.define_name("MonthWages",
                     f"={quote_sheet(NAME)}!$D${totals_row + 1}")
    ws.write_formula(
        row, 4, s.reg.render('=ROUND(SUMIFS([[Expenses.amount]],'
                             '[[Expenses.month]],DashKey),2)'),
        fmts.calc("money"))
    book.define_name("MonthExpenses",
                     f"={quote_sheet(NAME)}!$E${totals_row + 1}")
    ws.write_formula(row, 5, "=ROUND(MonthNetSales-MonthCostOfSales-"
                             "MonthWages-MonthExpenses,2)",
                     fmts.calc("money"))
    ws.write_formula(row, 6, '=IF(MonthNetSales<=0,"",'
                             '(MonthNetSales-MonthCostOfSales-'
                             'MonthWages-MonthExpenses)/MonthNetSales)',
                     fmts.calc("percent"))
    ws.write_formula(row, 7, '=IF(MonthNetSales<=0,"",'
                             'ROUND((MonthWages+MonthExpenses)/'
                             'MAX(0.05,(MonthNetSales-MonthCostOfSales)/'
                             'MonthNetSales),2))', fmts.calc("money"))
    row += 2

    row = s.note(
        row,
        "Operating profit here is net sales less the cost of what you "
        "sold, less the full cost of the pay runs, less the expenses "
        "logged for the month. It is not a tax computation and it does "
        "not know about depreciation, drawings or loans - it is the "
        "number that tells you whether the month worked.",
        LAST_COL, height=36)

    s.finish(freeze_row=6, gridlines=False, fit_width=True)
    s.protect()
    return s
````

### `salon_business_spreadsheet/sheets/profit_loss.py`

*316 lines · 2 functions*

> `Profit & Loss` - twelve months, in the order an accountant reads it.
>
> Income, cost of sales, gross profit, staff costs, every expense
> category, operating profit. Each column is a month, worked out from
> the till, the payroll and the expense log - so it agrees with the
> dashboard by construction rather than by luck.

| Defines | What it does |
|---------|--------------|
| `_key(index)` | The yyyy-mm key for `index` months before this one |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Profit & Loss'), `MONTHS` (12), `LAST_COL` (MONTHS + 2).

````python
"""`Profit & Loss` - twelve months, in the order an accountant reads it.

Income, cost of sales, gross profit, staff costs, every expense
category, operating profit. Each column is a month, worked out from
the till, the payroll and the expense log - so it agrees with the
dashboard by construction rather than by luck.
"""

from __future__ import annotations

from ..config import EXPENSE_CATEGORIES, Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter

NAME = "Profit & Loss"
MONTHS = 12
LAST_COL = MONTHS + 2          # label + 12 months + year to date


def _key(index: int) -> str:
    """The yyyy-mm key for `index` months before this one."""
    return (f'TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{index},1),'
            f'"yyyy-mm")')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.primary)
    ws = s.ws

    s.banner("Profit & Loss",
             "Twelve months of income, costs and what was left - built "
             "from the sales, payroll and expenses you have already "
             "logged.", LAST_COL)
    s.set_col(0, 0, 30)
    for c in range(1, LAST_COL + 1):
        s.set_col(c, c, 12)

    row = 4
    row = s.note(
        row,
        "Every column is a calendar month. The year-to-date column adds "
        "the twelve together. Nothing here is a tax return: it is a "
        "management view, so drawings, depreciation, loan capital and "
        "stock still on the shelf are deliberately left out.",
        LAST_COL, height=32)
    row += 1

    # ------------------------------------------------------- month header
    s.write(row, 0, "Line", fmts.chrome("header_left"))
    for index in range(MONTHS - 1, -1, -1):
        col = MONTHS - index
        ws.write_formula(
            row, col,
            f'=TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{index},1),'
            f'"mmm yy")', fmts.chrome("header"))
    s.write(row, MONTHS + 1, "12 months", fmts.chrome("header"))
    s.write(row, MONTHS + 2, "Share", fmts.chrome("header"))
    s.row_height(row, 24)
    header_row = row
    row += 1

    def line(label: str, body: str, style: str = "money",
             mode: str = "calc", bold: bool = False,
             share_of: str = "") -> int:
        """One P&L line: the label, twelve months, the year, the share."""
        nonlocal row
        fmt_label = fmts.chrome("label_bold" if bold else "label")
        s.write(row, 0, label, fmt_label)
        for index in range(MONTHS - 1, -1, -1):
            col = MONTHS - index
            formula = body.format(key=_key(index))
            ws.write_formula(row, col, s.reg.render(formula),
                             fmts.calc(style))
        letters = [f"{chr(66 + i)}" for i in range(MONTHS)]
        total = "+".join(f"{letter}{row + 1}" for letter in letters)
        ws.write_formula(row, MONTHS + 1, f"=ROUND({total},2)",
                         fmts.calc(style))
        if share_of:
            ws.write_formula(
                row, MONTHS + 2,
                f'=IF({share_of}=0,"",{chr(66 + MONTHS)}{row + 1}/'
                f'{share_of})', fmts.calc("percent"))
        else:
            s.write(row, MONTHS + 2, None, fmts.calc("percent"))
        row += 1
        return row - 1

    # ------------------------------------------------------------- income
    row = s.section(row, "Income", LAST_COL)
    service_row = line(
        "Service sales",
        '=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{key},'
        '[[Sales.kind]],"Service"),2)')
    retail_row = line(
        "Retail sales",
        '=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{key},'
        '[[Sales.kind]],"Retail"),2)')
    package_row = line(
        "Packages, courses & gift certificates",
        '=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{key},'
        '[[Sales.kind]],"Package")+SUMIFS([[Sales.net]],'
        '[[Sales.month]],{key},[[Sales.kind]],"Gift certificate")+'
        'SUMIFS([[Sales.net]],[[Sales.month]],{key},[[Sales.kind]],'
        '"Course"),2)')
    other_row = line(
        "Other income (chair rent, fees)",
        '=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{key},'
        '[[Sales.kind]],"Other")+SUMIFS([[Appointments.fee_charged]],'
        '[[Appointments.month]],{key}),2)')
    net_row = line("Net sales", "=0", bold=True)
    # the generic {key} substitution cannot build column references, so
    # the total line is rewritten cell by cell
    for index in range(MONTHS):
        col = index + 1
        letter = chr(66 + index)
        ws.write_formula(
            net_row, col,
            f"=ROUND({letter}{service_row + 1}+{letter}{retail_row + 1}+"
            f"{letter}{package_row + 1}+{letter}{other_row + 1},2)",
            fmts.calc("money"))
    book.define_name("PLNetSalesRow",
                     f"={quote_sheet(NAME)}!$N${net_row + 1}")
    net_total = f"$N${net_row + 1}"
    ws.write_formula(net_row, MONTHS + 2, "=1", fmts.calc("percent"))

    row = s.subsection(row, "What came off the top", LAST_COL)
    line("Discounts given (memo)",
         '=ROUND(SUMIFS([[Sales.discount]],[[Sales.month]],{key}),2)',
         share_of=net_total)
    line("Refunds (memo)",
         '=ROUND(SUMIFS([[Sales.refund]],[[Sales.month]],{key}),2)',
         share_of=net_total)
    line("Sales tax collected (not yours)",
         '=ROUND(SUMIFS([[Sales.tax]],[[Sales.month]],{key}),2)',
         share_of=net_total)
    line("Tips passed to staff (not income)",
         '=ROUND(SUMIFS([[Payments.amount]],[[Payments.month]],{key},'
         '[[Payments.kind]],"Tip"),2)', share_of=net_total)
    row += 1

    # ------------------------------------------------------ cost of sales
    row = s.section(row, "Cost of sales", LAST_COL)
    backbar_row = line(
        "Product used in services",
        '=ROUND(SUMIFS([[Sales.cost]],[[Sales.month]],{key},'
        '[[Sales.kind]],"Service"),2)', share_of=net_total)
    stock_row = line(
        "Retail stock sold",
        '=ROUND(SUMIFS([[Sales.cost]],[[Sales.month]],{key},'
        '[[Sales.kind]],"Retail"),2)', share_of=net_total)
    cos_row = line("Total cost of sales", "=0", bold=True,
                   share_of=net_total)
    for index in range(MONTHS):
        col = index + 1
        letter = chr(66 + index)
        ws.write_formula(
            cos_row, col,
            f"=ROUND({letter}{backbar_row + 1}+{letter}{stock_row + 1},2)",
            fmts.calc("money"))
    gross_row = line("Gross profit", "=0", bold=True, share_of=net_total)
    for index in range(MONTHS):
        col = index + 1
        letter = chr(66 + index)
        ws.write_formula(
            gross_row, col,
            f"=ROUND({letter}{net_row + 1}-{letter}{cos_row + 1},2)",
            fmts.calc("money"))
    margin_row = line("Gross margin", "=0", style="percent", bold=True)
    for index in range(MONTHS):
        col = index + 1
        letter = chr(66 + index)
        ws.write_formula(
            margin_row, col,
            f'=IF({letter}{net_row + 1}=0,"",{letter}{gross_row + 1}/'
            f'{letter}{net_row + 1})', fmts.calc("percent"))
    ws.write_formula(margin_row, MONTHS + 1,
                     f'=IF($N${net_row + 1}=0,"",$N${gross_row + 1}/'
                     f'$N${net_row + 1})', fmts.calc("percent"))
    row += 1

    # --------------------------------------------------------- staff cost
    row = s.section(row, "Staff costs", LAST_COL)
    wages_row = line(
        "Hourly pay",
        '=ROUND(SUMIFS([[Payroll.base_pay]],[[Payroll.month]],{key}),2)',
        share_of=net_total)
    commission_row = line(
        "Commission released",
        '=ROUND(SUMIFS([[Payroll.commission]],[[Payroll.month]],'
        '{key}),2)', share_of=net_total)
    extra_row = line(
        "Bonuses and holiday pay",
        '=ROUND(SUMIFS([[Payroll.bonus]],[[Payroll.month]],{key})+'
        'SUMIFS([[Payroll.holiday]],[[Payroll.month]],{key}),2)',
        share_of=net_total)
    employer_row = line(
        "Employer taxes and cover",
        '=ROUND(SUMIFS([[Payroll.employer_cost]],[[Payroll.month]],'
        '{key}),2)', share_of=net_total)
    staff_row = line("Total staff cost", "=0", bold=True,
                     share_of=net_total)
    for index in range(MONTHS):
        col = index + 1
        letter = chr(66 + index)
        ws.write_formula(
            staff_row, col,
            f"=ROUND({letter}{wages_row + 1}+{letter}{commission_row + 1}+"
            f"{letter}{extra_row + 1}+{letter}{employer_row + 1},2)",
            fmts.calc("money"))
    line("Staff cost as a share of service sales", "=0",
         style="percent")
    share_row = row - 1
    for index in range(MONTHS):
        col = index + 1
        letter = chr(66 + index)
        ws.write_formula(
            share_row, col,
            f'=IF({letter}{service_row + 1}=0,"",'
            f'{letter}{staff_row + 1}/{letter}{service_row + 1})',
            fmts.calc("percent"))
    ws.write_formula(share_row, MONTHS + 1,
                     f'=IF($N${service_row + 1}=0,"",$N${staff_row + 1}/'
                     f'$N${service_row + 1})', fmts.calc("percent"))
    row += 1

    # ----------------------------------------------------- the overheads
    row = s.section(row, "Operating expenses", LAST_COL)
    first_expense = row
    for category in EXPENSE_CATEGORIES:
        if category in ("Salaries & wages", "Commission", "Payroll taxes"):
            continue                    # they live on the payroll tab
        safe = category.replace('"', "")
        line(category,
             '=ROUND(SUMIFS([[Expenses.amount]],[[Expenses.month]],'
             '{key},[[Expenses.category]],"' + safe + '"),2)',
             share_of=net_total)
    last_expense = row - 1
    total_exp_row = line("Total operating expenses", "=0", bold=True,
                         share_of=net_total)
    for index in range(MONTHS):
        col = index + 1
        letter = chr(66 + index)
        ws.write_formula(
            total_exp_row, col,
            f"=ROUND(SUM({letter}{first_expense + 1}:"
            f"{letter}{last_expense + 1}),2)", fmts.calc("money"))
    row += 1

    # ------------------------------------------------------------- profit
    row = s.section(row, "The bottom line", LAST_COL)
    profit_row = line("Operating profit", "=0", bold=True,
                      share_of=net_total)
    for index in range(MONTHS):
        col = index + 1
        letter = chr(66 + index)
        ws.write_formula(
            profit_row, col,
            f"=ROUND({letter}{gross_row + 1}-{letter}{staff_row + 1}-"
            f"{letter}{total_exp_row + 1},2)", fmts.calc("money"))
    profit_margin_row = line("Operating margin", "=0", style="percent",
                             bold=True)
    for index in range(MONTHS):
        col = index + 1
        letter = chr(66 + index)
        ws.write_formula(
            profit_margin_row, col,
            f'=IF({letter}{net_row + 1}=0,"",{letter}{profit_row + 1}/'
            f'{letter}{net_row + 1})', fmts.calc("percent"))
    ws.write_formula(profit_margin_row, MONTHS + 1,
                     f'=IF($N${net_row + 1}=0,"",$N${profit_row + 1}/'
                     f'$N${net_row + 1})', fmts.calc("percent"))
    target_row = line("Profit target", "=ROUND(ProfitTargetMonth,2)")
    variance_row = line("Against target", "=0", bold=True)
    for index in range(MONTHS):
        col = index + 1
        letter = chr(66 + index)
        ws.write_formula(
            variance_row, col,
            f"=ROUND({letter}{profit_row + 1}-{letter}{target_row + 1},2)",
            fmts.calc("money"))
    ws.conditional_format(variance_row, 1, variance_row, MONTHS + 1,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(variance_row, 1, variance_row, MONTHS + 1,
                          {"type": "cell", "criteria": ">=", "value": 0,
                           "format": fmts.cf("good")})
    ws.conditional_format(profit_row, 1, profit_row, MONTHS + 1,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("bad")})
    ws.conditional_format(net_row, 1, net_row, MONTHS,
                          {"type": "data_bar",
                           "bar_color": Palette.primary_mid})
    row += 1

    ws.add_sparkline(row, 1, {
        "range": (f"{quote_sheet(NAME)}!$B${profit_row + 1}:"
                  f"$M${profit_row + 1}"),
        "type": "column", "series_color": Palette.accent,
        "negative_points": True})
    s.merge(row, 2, 5, "Operating profit, month by month",
            fmts.chrome("note"))
    row += 2

    row = s.note(
        row,
        "Wages come from the payroll tab, never from the expense log, so "
        "they are counted once. Stock is counted when it is sold, not "
        "when it is bought, which is why a big delivery does not make a "
        "month look like a disaster. Deposits are money received and "
        "appear as revenue only when the sale is rung up.",
        LAST_COL, height=36)

    s.finish(freeze_row=header_row + 1, freeze_col=1, gridlines=False,
             fit_width=True)
    s.protect()
    return s
````

### `salon_business_spreadsheet/sheets/goals.py`

*364 lines · 2 functions*

> `Goals & Budget` - what you said you would do, against what happened.
>
> Three blocks: the year's targets, a month-by-month plan with the
> actuals beside it, and a budget for every expense category. The
> budget figures here are what the expense tab compares itself to, so
> setting them once makes the whole workbook sharper.

| Defines | What it does |
|---------|--------------|
| `_key(index)` |  |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Goals & Budget'), `LAST_COL` (10), `MONTHS` (12).

````python
"""`Goals & Budget` - what you said you would do, against what happened.

Three blocks: the year's targets, a month-by-month plan with the
actuals beside it, and a budget for every expense category. The
budget figures here are what the expense tab compares itself to, so
setting them once makes the whole workbook sharper.
"""

from __future__ import annotations

from ..config import EXPENSE_CATEGORIES, Config, Palette
from ..layout import Registry, col_letter, quote_sheet
from ..writer import SheetWriter, number_check

NAME = "Goals & Budget"
LAST_COL = 10
MONTHS = 12


def _key(index: int) -> str:
    return (f'TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{index},1),'
            f'"yyyy-mm")')


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.accent)
    ws = s.ws
    sheet_ref = quote_sheet(NAME)

    s.banner("Goals & Budget",
             "The year you planned, the months as they land, and a "
             "budget for every line of cost.", LAST_COL)
    s.set_col(0, 0, 3)
    s.set_col(1, 1, 32)
    s.set_col(2, LAST_COL, 15)

    row = 4
    row = s.note(
        row,
        "Blue cells are yours. The targets you set here are what the "
        "dashboard, the profit and loss and the expense tab measure "
        "against - change them and every red number in the workbook "
        "changes with them.",
        LAST_COL, height=30)
    row += 1

    # ------------------------------------------------------ the year ----
    row = s.section(row, "This year's goals", LAST_COL)
    s.table_head(row, 1, ["Goal", "Target", "Actual so far",
                          "Difference", "On track?"],
                 [32, 15, 15, 15, 15])
    row += 1
    goals = [
        ("Revenue for the year", "money", "=ROUND(RevenueTargetYear,2)",
         '=ROUND(SUMIFS([[Sales.net]],[[Sales.date]],">="&'
         'DATE(YEAR(TodayDate),1,1)),2)'),
        ("Operating profit for the year", "money",
         "=ROUND(ProfitTargetMonth*12,2)",
         '=ROUND(SUMIFS([[Sales.net]],[[Sales.date]],">="&'
         'DATE(YEAR(TodayDate),1,1))-SUMIFS([[Sales.cost]],'
         '[[Sales.date]],">="&DATE(YEAR(TodayDate),1,1))-'
         'SUMIFS([[Payroll.total_cost]],[[Payroll.period_end]],">="&'
         'DATE(YEAR(TodayDate),1,1))-SUMIFS([[Expenses.amount]],'
         '[[Expenses.date]],">="&DATE(YEAR(TodayDate),1,1)),2)'),
        ("Clients served", "int", "=ROUND(ClientTargetYear,0)",
         '=COUNTIFS([[Appointments.status]],"Completed",'
         '[[Appointments.date]],">="&DATE(YEAR(TodayDate),1,1))'),
        ("New clients won", "int", "=ROUND(NewClientTargetYear,0)",
         '=COUNTIFS([[Clients.first_visit]],">="&'
         'DATE(YEAR(TodayDate),1,1))'),
        ("Average ticket", "money", "=ROUND(TargetTicket,2)",
         '=IFERROR(ROUND(SUMIFS([[Sales.net]],[[Sales.date]],">="&'
         'DATE(YEAR(TodayDate),1,1))/MAX(1,COUNTIFS('
         '[[Appointments.status]],"Completed",[[Appointments.date]],'
         '">="&DATE(YEAR(TodayDate),1,1))),2),0)'),
        ("Retail share of sales", "percent",
         "=ROUND(TargetRetailShare,4)",
         '=IFERROR(SUMIFS([[Sales.net]],[[Sales.kind]],"Retail",'
         '[[Sales.date]],">="&DATE(YEAR(TodayDate),1,1))/'
         'MAX(0.01,SUMIFS([[Sales.net]],[[Sales.date]],">="&'
         'DATE(YEAR(TodayDate),1,1))),0)'),
        ("Rebooking rate", "percent", "=ROUND(TargetRebook,4)",
         '=IFERROR(COUNTIFS([[Appointments.rebooked]],"Yes",'
         '[[Appointments.date]],">="&DATE(YEAR(TodayDate),1,1))/'
         'MAX(1,COUNTIFS([[Appointments.status]],"Completed",'
         '[[Appointments.date]],">="&DATE(YEAR(TodayDate),1,1))),0)'),
        ("Repeat client rate", "percent", "=ROUND(TargetRetention,4)",
         '=IFERROR(COUNTIFS([[Clients.retained]],"Yes")/'
         'MAX(1,COUNTIFS([[Clients.visits]],">0")),0)'),
        ("Chair utilisation", "percent", "=ROUND(TargetUtilisation,4)",
         '=IFERROR(SUMIFS([[Appointments.hours]],'
         '[[Appointments.status]],"Completed",[[Appointments.date]],'
         '">="&TodayDate-28)/MAX(1,ChairHoursWeek*4),0)'),
    ]
    first_goal = row
    for label, style, target, actual in goals:
        s.write(row, 1, label, fmts.chrome("label_bold"))
        ws.write_formula(row, 2, s.reg.render(target), fmts.calc(style))
        ws.write_formula(row, 3, s.reg.render(actual), fmts.calc(style))
        ws.write_formula(row, 4, f"=D{row + 1}-C{row + 1}",
                         fmts.calc(style))
        ws.write_formula(
            row, 5,
            f'=IF(C{row + 1}=0,"Set a target",'
            f'IF(D{row + 1}>=C{row + 1},"On track",'
            f'IF(D{row + 1}>=C{row + 1}*0.8,"Close","Behind")))',
            fmts.calc("status"))
        row += 1
    for kind, text in (("good", "On track"), ("warn", "Close"),
                       ("bad", "Behind")):
        ws.conditional_format(first_goal, 5, row - 1, 5,
                              {"type": "text", "criteria": "containing",
                               "value": text, "format": fmts.cf(kind)})
    ws.conditional_format(first_goal, 4, row - 1, 4,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("warn")})
    row += 1

    # the two yearly counts live here rather than on Settings because
    # they are goals, not settings
    row = s.subsection(row, "Two targets only this tab uses", LAST_COL)
    row = s.field(row, "Clients served, target for the year",
                  int(cfg.revenue_target_year /
                      max(1, cfg.target_average_ticket)),
                  style="int", name="ClientTargetYear",
                  validate=number_check(0, 1000000, decimal=False),
                  help_text="Revenue target divided by the average "
                            "ticket is a good first guess.")
    row = s.field(row, "New clients, target for the year",
                  int(cfg.revenue_target_year /
                      max(1, cfg.target_average_ticket) * 0.18),
                  style="int", name="NewClientTargetYear",
                  validate=number_check(0, 1000000, decimal=False))
    row += 1

    # --------------------------------------------------- month by month --
    row = s.section(row, "Month by month", LAST_COL)
    row = s.note(
        row,
        "Type a revenue target for each month - quiet January, busy "
        "December. Leave one blank and the monthly target from Settings "
        "is used instead.",
        LAST_COL, height=28)
    row += 1
    s.table_head(row, 1, [
        "Month", "Revenue target", "Actual", "Difference",
        "Clients served", "Average ticket", "Profit", "Hit it?"],
        [14, 15, 15, 15, 14, 14, 14, 12])
    row += 1
    month_first = row
    for index in range(MONTHS - 1, -1, -1):
        key = _key(index)
        s.write(row, 1, None, fmts.calc("text"))
        ws.write_formula(
            row, 1,
            f'=TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{index},1),'
            f'"mmm yy")', fmts.calc("text"))
        s.write(row, 2, None, fmts.input("money"))
        ws.data_validation(row, 2, row, 2, number_check(0, 10000000))
        ws.write_formula(
            row, 3,
            s.reg.render(f'=ROUND(SUMIFS([[Sales.net]],'
                         f'[[Sales.month]],{key}),2)'),
            fmts.calc("money"))
        ws.write_formula(
            row, 4,
            f"=ROUND(D{row + 1}-IF(C{row + 1}=\"\",RevenueTargetMonth,"
            f"C{row + 1}),2)", fmts.calc("money"))
        ws.write_formula(
            row, 5,
            s.reg.render(f'=COUNTIFS([[Appointments.month]],{key},'
                         f'[[Appointments.status]],"Completed")'),
            fmts.calc("int"))
        ws.write_formula(
            row, 6,
            f'=IFERROR(ROUND(D{row + 1}/MAX(1,F{row + 1}),2),"")',
            fmts.calc("money"))
        ws.write_formula(
            row, 7,
            s.reg.render(
                f'=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{key})-'
                f'SUMIFS([[Sales.cost]],[[Sales.month]],{key})-'
                f'SUMIFS([[Payroll.total_cost]],[[Payroll.month]],{key})-'
                f'SUMIFS([[Expenses.amount]],[[Expenses.month]],{key}),2)'),
            fmts.calc("money"))
        ws.write_formula(
            row, 8,
            f'=IF(D{row + 1}=0,"",IF(E{row + 1}>=0,"Yes","No"))',
            fmts.calc("flag"))
        row += 1
    month_last = row - 1
    ws.conditional_format(month_first, 4, month_last, 4,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("warn")})
    ws.conditional_format(month_first, 4, month_last, 4,
                          {"type": "cell", "criteria": ">=", "value": 0,
                           "format": fmts.cf("good")})
    ws.conditional_format(month_first, 8, month_last, 8,
                          {"type": "text", "criteria": "containing",
                           "value": "No", "format": fmts.cf("bad")})
    ws.conditional_format(month_first, 8, month_last, 8,
                          {"type": "text", "criteria": "containing",
                           "value": "Yes", "format": fmts.cf("good")})
    s.write(row, 1, "Year", fmts.chrome("label_bold"))
    for col in (2, 3, 4, 5, 7):
        letter = col_letter(col)
        ws.write_formula(
            row, col,
            f"=ROUND(SUM({letter}{month_first + 1}:"
            f"{letter}{month_last + 1}),2)",
            fmts.calc("int" if col == 5 else "money"))
    ws.write_formula(row, 6,
                     f'=IF(F{row + 1}=0,"",ROUND(D{row + 1}/'
                     f'F{row + 1},2))', fmts.calc("money"))
    row += 2

    # ---------------------------------------------------- the budget ----
    row = s.section(row, "Budget by category", LAST_COL)
    row = s.note(
        row,
        "Set a monthly budget for each line of cost. The expense tab "
        "reads these figures, so every bill you log knows whether it is "
        "inside the plan. Variance is actual less budget: positive "
        "means you overspent.",
        LAST_COL, height=30)
    row += 1
    s.table_head(row, 1, [
        "Category", "Monthly budget", "This month", "Variance",
        "Last month", "Year to date", "Budget to date",
        "Variance to date"],
        [26, 15, 14, 14, 14, 14, 15, 15])
    row += 1
    budget_first = row
    defaults = {
        "Rent": 2400, "Utilities": 320, "Product & stock": 0,
        "Backbar supplies": 0, "Insurance": 95,
        "Equipment": 120, "Repairs & maintenance": 90,
        "Marketing & advertising": 200,
        "Software & subscriptions": 80, "Card processing fees": 320,
        "Laundry": 165, "Cleaning": 240, "Training & education": 150,
        "Licences & memberships": 25, "Accounting & legal": 120,
        "Telephone & internet": 65, "Travel & parking": 30,
        "Bank charges": 18, "Loan repayment": 420, "Other": 120,
    }
    categories = [c for c in EXPENSE_CATEGORIES
                  if c not in ("Salaries & wages", "Commission",
                               "Payroll taxes")]
    for category in categories:
        s.write(row, 1, category, fmts.input("text"))
        s.write(row, 2, defaults.get(category, 0)
                if cfg.include_sample_data else None,
                fmts.input("money"))
        ws.data_validation(row, 2, row, 2, number_check(0, 1000000))
        ws.write_formula(
            row, 3,
            s.reg.render('=ROUND(SUMIFS([[Expenses.amount]],'
                         '[[Expenses.month]],ThisMonthKey,'
                         f'[[Expenses.category]],$B${row + 1}),2)'),
            fmts.calc("money"))
        ws.write_formula(row, 4,
                         f'=ROUND(D{row + 1}-IF(C{row + 1}="",0,'
                         f'C{row + 1}),2)', fmts.calc("money"))
        ws.write_formula(
            row, 5,
            s.reg.render('=ROUND(SUMIFS([[Expenses.amount]],'
                         '[[Expenses.month]],TEXT(DATE(YEAR(TodayDate),'
                         'MONTH(TodayDate)-1,1),"yyyy-mm"),'
                         f'[[Expenses.category]],$B${row + 1}),2)'),
            fmts.calc("money"))
        ws.write_formula(
            row, 6,
            s.reg.render('=ROUND(SUMIFS([[Expenses.amount]],'
                         '[[Expenses.date]],">="&DATE(YEAR(TodayDate),'
                         '1,1),[[Expenses.category]],'
                         f'$B${row + 1}),2)'),
            fmts.calc("money"))
        ws.write_formula(row, 7,
                         f'=ROUND(IF(C{row + 1}="",0,C{row + 1})*'
                         f'MONTH(TodayDate),2)', fmts.calc("money"))
        ws.write_formula(row, 8, f"=ROUND(G{row + 1}-H{row + 1},2)",
                         fmts.calc("money"))
        row += 1
    budget_last = row - 1
    s.write(row, 1, "Total", fmts.chrome("label_bold"))
    for col in range(2, 9):
        letter = col_letter(col)
        ws.write_formula(
            row, col,
            f"=ROUND(SUM({letter}{budget_first + 1}:"
            f"{letter}{budget_last + 1}),2)", fmts.calc("money"))
    total_row = row
    for column in (4, 8):
        ws.conditional_format(budget_first, column, total_row, column,
                              {"type": "cell", "criteria": ">", "value": 0,
                               "format": fmts.cf("bad")})
        ws.conditional_format(budget_first, column, total_row, column,
                              {"type": "cell", "criteria": "<=",
                               "value": 0, "format": fmts.cf("good")})
    book.define_name(
        "BudgetCategory",
        f"={sheet_ref}!$B${budget_first + 1}:$B${budget_last + 1}")
    book.define_name(
        "BudgetAmount",
        f"={sheet_ref}!$C${budget_first + 1}:$C${budget_last + 1}")
    book.define_name(
        "BudgetMonthlyTotal", f"={sheet_ref}!$C${total_row + 1}")
    row += 2

    # ------------------------------------------------------ break even --
    row = s.section(row, "What you have to take to break even", LAST_COL)
    row = s.table_head(row, 1, [
        "Fixed costs a month", "Staff cost a month", "Gross margin",
        "Break-even sales", "Sales this month", "Still to go",
        "Days left in the month"],
        [18, 18, 14, 16, 16, 14, 16])
    ws.write_formula(
        row, 1,
        s.reg.render('=ROUND(SUM([[Expenses.monthly]]),2)'),
        fmts.calc("money"))
    ws.write_formula(
        row, 2,
        s.reg.render('=ROUND(SUMIFS([[Payroll.total_cost]],'
                     '[[Payroll.month]],ThisMonthKey),2)'),
        fmts.calc("money"))
    ws.write_formula(
        row, 3,
        s.reg.render('=IF(SUMIFS([[Sales.net]],[[Sales.month]],'
                     'ThisMonthKey)<=0,0.85,(SUMIFS([[Sales.net]],'
                     '[[Sales.month]],ThisMonthKey)-'
                     'SUMIFS([[Sales.cost]],[[Sales.month]],'
                     'ThisMonthKey))/SUMIFS([[Sales.net]],'
                     '[[Sales.month]],ThisMonthKey))'),
        fmts.calc("percent"))
    ws.write_formula(row, 4,
                     f'=IFERROR(ROUND((B{row + 1}+C{row + 1})/'
                     f'MAX(0.05,D{row + 1}),2),"")', fmts.calc("money"))
    ws.write_formula(
        row, 5,
        s.reg.render('=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],'
                     'ThisMonthKey),2)'), fmts.calc("money"))
    ws.write_formula(row, 6, f"=ROUND(MAX(0,E{row + 1}-F{row + 1}),2)",
                     fmts.calc("money"))
    ws.write_formula(row, 7,
                     "=DAY(DATE(YEAR(TodayDate),MONTH(TodayDate)+1,0))-"
                     "DAY(TodayDate)", fmts.calc("int"))
    ws.conditional_format(row, 6, row, 6,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("warn")})
    ws.conditional_format(row, 6, row, 6,
                          {"type": "cell", "criteria": "=", "value": 0,
                           "format": fmts.cf("good")})
    row += 2

    row = s.note(
        row,
        "Break-even is the fixed overhead plus the pay run, divided by "
        "the gross margin you are actually achieving. It is the number "
        "to know before you agree to a discount, a new hire or a bigger "
        "unit.",
        LAST_COL, height=30)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    s.protect()
    return s
````

### `salon_business_spreadsheet/sheets/analytics.py`

*351 lines · 2 functions*

> `Analytics` - the patterns you cannot see from behind the desk.
>
> Which categories earn, which days pay, which hours are dead, where
> the clients come from, and whether the ones you win stay. Every
> figure is a formula over the registers, so it moves the moment the
> diary does.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |
| `_hour_label(hour)` |  |

**Module constants:** `NAME` ('Analytics'), `LAST_COL` (10), `DAYS` (7 items).

````python
"""`Analytics` - the patterns you cannot see from behind the desk.

Which categories earn, which days pay, which hours are dead, where
the clients come from, and whether the ones you win stay. Every
figure is a formula over the registers, so it moves the moment the
diary does.
"""

from __future__ import annotations

from ..config import (MARKETING_CHANNELS, SERVICE_CATEGORIES, Config,
                      Palette)
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter

NAME = "Analytics"
LAST_COL = 10
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday",
        "Saturday", "Sunday"]


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.deep)
    ws = s.ws

    s.banner("Business Analytics",
             "Categories, days, hours, sources and retention - the "
             "patterns behind the takings.", LAST_COL)
    s.set_col(0, 0, 3)
    s.set_col(1, 1, 28)
    s.set_col(2, LAST_COL, 14)

    row = 4
    row = s.note(
        row,
        "Nothing on this tab is typed. It is the same data as the rest "
        "of the workbook, cut four different ways, because the answer "
        "to 'why was last month quiet?' is usually in one of these "
        "tables.",
        LAST_COL, height=28)
    row += 1

    # ---------------------------------------------- service categories --
    row = s.section(row, "Where the money comes from, by category",
                    LAST_COL)
    s.table_head(row, 1, [
        "Category", "Services on the menu", "Booked", "Revenue",
        "Share of revenue", "Average price", "Chair hours",
        "Revenue an hour"],
        [26, 16, 11, 14, 14, 13, 13, 14])
    row += 1
    cat_first = row
    for category in SERVICE_CATEGORIES:
        s.write(row, 1, category, fmts.calc("text"))
        formulas = [
            (2, f'=COUNTIFS([[Services.category]],"{category}")', "int"),
            (3, f'=SUMIFS([[Services.done]],[[Services.category]],'
                f'"{category}")', "int"),
            (4, f'=ROUND(SUMIFS([[Services.revenue]],'
                f'[[Services.category]],"{category}"),2)', "money"),
            (5, f'=IFERROR(SUMIFS([[Services.revenue]],'
                f'[[Services.category]],"{category}")/'
                f'MAX(0.01,ServiceRevenueAll),"")', "percent"),
            (6, f'=IFERROR(ROUND(SUMIFS([[Services.revenue]],'
                f'[[Services.category]],"{category}")/'
                f'MAX(1,SUMIFS([[Services.units]],'
                f'[[Services.category]],"{category}")),2),"")', "money"),
            (7, f'=ROUND(SUMIFS([[Services.hours_total]],'
                f'[[Services.category]],"{category}"),1)', "number1"),
            (8, f'=IFERROR(ROUND(SUMIFS([[Services.revenue]],'
                f'[[Services.category]],"{category}")/'
                f'MAX(0.1,SUMIFS([[Services.hours_total]],'
                f'[[Services.category]],"{category}")),2),"")', "money"),
        ]
        for col, formula, style in formulas:
            ws.write_formula(row, col, s.reg.render(formula),
                             fmts.calc(style))
        row += 1
    cat_last = row - 1
    ws.conditional_format(cat_first, 4, cat_last, 4,
                          {"type": "data_bar",
                           "bar_color": Palette.primary_mid})
    ws.conditional_format(cat_first, 8, cat_last, 8,
                          {"type": "cell", "criteria": ">=",
                           "value": "=TargetChairRate",
                           "format": fmts.cf("good")})
    ws.conditional_format(cat_first, 8, cat_last, 8,
                          {"type": "cell", "criteria": "<",
                           "value": "=TargetChairRate",
                           "format": fmts.cf("warn")})
    row += 1

    # ----------------------------------------------------- by weekday ---
    row = s.section(row, "Which days pay", LAST_COL)
    s.table_head(row, 1, [
        "Day", "Appointments", "Completed", "Cancelled", "No shows",
        "Revenue", "Average ticket", "Chair hours", "Revenue an hour"],
        [16, 14, 12, 12, 11, 14, 14, 13, 14])
    row += 1
    day_first = row
    for index, day in enumerate(DAYS):
        short = day[:3]
        s.write(row, 1, day, fmts.calc("text"))
        formulas = [
            (2, f'=COUNTIFS([[Appointments.weekday]],"{short}")', "int"),
            (3, f'=COUNTIFS([[Appointments.weekday]],"{short}",'
                f'[[Appointments.status]],"Completed")', "int"),
            (4, f'=COUNTIFS([[Appointments.weekday]],"{short}",'
                f'[[Appointments.status]],"Cancelled")', "int"),
            (5, f'=COUNTIFS([[Appointments.weekday]],"{short}",'
                f'[[Appointments.status]],"No show")', "int"),
            (6, f'=ROUND(SUMIFS([[Appointments.billed]],'
                f'[[Appointments.weekday]],"{short}",'
                f'[[Appointments.status]],"Completed"),2)', "money"),
            (7, f'=IFERROR(ROUND(SUMIFS([[Appointments.billed]],'
                f'[[Appointments.weekday]],"{short}",'
                f'[[Appointments.status]],"Completed")/'
                f'MAX(1,COUNTIFS([[Appointments.weekday]],"{short}",'
                f'[[Appointments.status]],"Completed")),2),"")',
             "money"),
            (8, f'=ROUND(SUMIFS([[Appointments.hours]],'
                f'[[Appointments.weekday]],"{short}",'
                f'[[Appointments.status]],"Completed"),1)', "number1"),
            (9, f'=IFERROR(ROUND(SUMIFS([[Appointments.billed]],'
                f'[[Appointments.weekday]],"{short}",'
                f'[[Appointments.status]],"Completed")/'
                f'MAX(0.1,SUMIFS([[Appointments.hours]],'
                f'[[Appointments.weekday]],"{short}",'
                f'[[Appointments.status]],"Completed")),2),"")',
             "money"),
        ]
        for col, formula, style in formulas:
            ws.write_formula(row, col, s.reg.render(formula),
                             fmts.calc(style))
        row += 1
    day_last = row - 1
    ws.conditional_format(day_first, 6, day_last, 6,
                          {"type": "data_bar",
                           "bar_color": Palette.primary_mid})
    ws.conditional_format(day_first, 5, day_last, 5,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("warn")})
    row += 1

    # -------------------------------------------------------- by hour ---
    row = s.section(row, "Which hours are dead", LAST_COL)
    row = s.note(
        row,
        "Bookings by start time, across the whole diary. A column of "
        "empty mornings is a column you can fill with an offer, a "
        "training slot, or a later opening time.",
        LAST_COL, height=26)
    row += 1
    s.table_head(row, 1, [
        "Hour", "Appointments", "Completed", "Revenue",
        "Share of bookings", "Cancelled or no show"],
        [12, 14, 12, 14, 16, 20])
    row += 1
    hour_first = row
    for hour in range(7, 21):
        label = _hour_label(hour)
        s.write(row, 1, label, fmts.calc("text"))
        formulas = [
            (2, f'=COUNTIFS([[Appointments.hour]],{hour})', "int"),
            (3, f'=COUNTIFS([[Appointments.hour]],{hour},'
                f'[[Appointments.status]],"Completed")', "int"),
            (4, f'=ROUND(SUMIFS([[Appointments.billed]],'
                f'[[Appointments.hour]],{hour},'
                f'[[Appointments.status]],"Completed"),2)', "money"),
            (5, f'=IFERROR(COUNTIFS([[Appointments.hour]],{hour})/'
                f'MAX(1,COUNTIF([[Appointments.date]],"<>")),"")',
             "percent"),
        ]
        for col, formula, style in formulas:
            ws.write_formula(row, col, s.reg.render(formula),
                             fmts.calc(style))
        ws.write_formula(
            row, 6,
            s.reg.render(
                f'=COUNTIFS([[Appointments.hour]],{hour},'
                f'[[Appointments.status]],"Cancelled")+'
                f'COUNTIFS([[Appointments.hour]],{hour},'
                f'[[Appointments.status]],"No show")'),
            fmts.calc("int"))
        row += 1
    hour_last = row - 1
    ws.conditional_format(hour_first, 2, hour_last, 2,
                          {"type": "data_bar",
                           "bar_color": Palette.accent})
    row += 1

    # ------------------------------------------------- where they come --
    row = s.section(row, "Where the clients come from", LAST_COL)
    s.table_head(row, 1, [
        "Channel", "Clients", "New in 12 months", "Bookings",
        "Revenue from them", "Average spend", "Still active"],
        [24, 12, 16, 12, 16, 14, 13])
    row += 1
    src_first = row
    for channel in MARKETING_CHANNELS[:12] + ["Walk-in", "Referral",
                                              "Repeat client"]:
        s.write(row, 1, channel, fmts.calc("text"))
        formulas = [
            (2, f'=COUNTIFS([[Clients.source]],"{channel}")', "int"),
            (3, f'=COUNTIFS([[Clients.source]],"{channel}",'
                f'[[Clients.first_visit]],">="&TodayDate-365)', "int"),
            (4, f'=COUNTIFS([[Appointments.source]],"{channel}")',
             "int"),
            (5, f'=ROUND(SUMIFS([[Clients.spend]],[[Clients.source]],'
                f'"{channel}"),2)', "money"),
            (6, f'=IFERROR(ROUND(SUMIFS([[Clients.spend]],'
                f'[[Clients.source]],"{channel}")/'
                f'MAX(1,COUNTIFS([[Clients.source]],"{channel}")),2),'
                f'"")', "money"),
            (7, f'=COUNTIFS([[Clients.source]],"{channel}",'
                f'[[Clients.segment]],"Active")+'
                f'COUNTIFS([[Clients.source]],"{channel}",'
                f'[[Clients.segment]],"Regular")+'
                f'COUNTIFS([[Clients.source]],"{channel}",'
                f'[[Clients.segment]],"VIP")', "int"),
        ]
        for col, formula, style in formulas:
            ws.write_formula(row, col, s.reg.render(formula),
                             fmts.calc(style))
        row += 1
    src_last = row - 1
    ws.conditional_format(src_first, 5, src_last, 5,
                          {"type": "data_bar",
                           "bar_color": Palette.primary_mid})
    row += 1

    # ------------------------------------------------------- retention --
    row = s.section(row, "Do they come back?", LAST_COL)
    s.table_head(row, 1, [
        "Month", "Visits", "New clients", "Visits by returning clients",
        "Repeat rate", "Rebooked at the desk", "Rebook rate",
        "Gone quiet since"],
        [14, 12, 13, 18, 13, 18, 13, 15])
    row += 1
    ret_first = row
    for index in range(11, -1, -1):
        key = (f'TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{index},1),'
               f'"yyyy-mm")')
        s.write(row, 1, None, fmts.calc("text"))
        ws.write_formula(
            row, 1,
            f'=TEXT(DATE(YEAR(TodayDate),MONTH(TodayDate)-{index},1),'
            f'"mmm yy")', fmts.calc("text"))
        formulas = [
            (2, f'=COUNTIFS([[Appointments.month]],{key},'
                f'[[Appointments.status]],"Completed")', "int"),
            (3, f'=COUNTIFS([[Clients.first_visit]],">="&'
                f'DATE(YEAR(TodayDate),MONTH(TodayDate)-{index},1),'
                f'[[Clients.first_visit]],"<"&'
                f'DATE(YEAR(TodayDate),MONTH(TodayDate)-{index}+1,1))',
             "int"),
            (6, f'=COUNTIFS([[Appointments.month]],{key},'
                f'[[Appointments.rebooked]],"Yes")', "int"),
        ]
        for col, formula, style in formulas:
            ws.write_formula(row, col, s.reg.render(formula),
                             fmts.calc(style))
        ws.write_formula(row, 4, f"=MAX(0,C{row + 1}-D{row + 1})",
                         fmts.calc("int"))
        ws.write_formula(row, 5,
                         f'=IFERROR(E{row + 1}/MAX(1,C{row + 1}),"")',
                         fmts.calc("percent"))
        ws.write_formula(row, 7,
                         f'=IFERROR(G{row + 1}/MAX(1,C{row + 1}),"")',
                         fmts.calc("percent"))
        ws.write_formula(
            row, 8,
            s.reg.render(
                f'=COUNTIFS([[Clients.last_visit]],">="&'
                f'DATE(YEAR(TodayDate),MONTH(TodayDate)-{index},1),'
                f'[[Clients.last_visit]],"<"&'
                f'DATE(YEAR(TodayDate),MONTH(TodayDate)-{index}+1,1),'
                f'[[Clients.segment]],"Inactive")'),
            fmts.calc("int"))
        row += 1
    ret_last = row - 1
    for column in (5, 7):
        ws.conditional_format(ret_first, column, ret_last, column,
                              {"type": "cell", "criteria": ">=",
                               "value": "=TargetRebook",
                               "format": fmts.cf("good")})
        ws.conditional_format(ret_first, column, ret_last, column,
                              {"type": "cell", "criteria": "<",
                               "value": "=TargetRebook",
                               "format": fmts.cf("warn")})
    ws.add_sparkline(row, 2, {
        "range": (f"{quote_sheet(NAME)}!$C${ret_first + 1}:"
                  f"$C${ret_last + 1}"),
        "type": "column", "series_color": Palette.primary_mid})
    s.merge(row, 3, 5, "Visits, twelve months", fmts.chrome("note"))
    row += 2

    # ---------------------------------------------- the client pyramid --
    row = s.section(row, "The client book, by value", LAST_COL)
    s.table_head(row, 1, [
        "Segment", "Clients", "Share", "Lifetime spend",
        "Average spend", "Average visits", "Average ticket"],
        [18, 12, 12, 16, 14, 14, 14])
    row += 1
    for segment in ("VIP", "Regular", "Active", "New", "Lapsing",
                    "Inactive", "Not been yet"):
        s.write(row, 1, segment, fmts.calc("text"))
        formulas = [
            (2, f'=COUNTIFS([[Clients.segment]],"{segment}")', "int"),
            (3, f'=IFERROR(COUNTIFS([[Clients.segment]],"{segment}")/'
                f'MAX(1,ClientCount),"")', "percent"),
            (4, f'=ROUND(SUMIFS([[Clients.spend]],[[Clients.segment]],'
                f'"{segment}"),2)', "money"),
            (5, f'=IFERROR(ROUND(SUMIFS([[Clients.spend]],'
                f'[[Clients.segment]],"{segment}")/'
                f'MAX(1,COUNTIFS([[Clients.segment]],"{segment}")),2),'
                f'"")', "money"),
            (6, f'=IFERROR(ROUND(SUMIFS([[Clients.visits]],'
                f'[[Clients.segment]],"{segment}")/'
                f'MAX(1,COUNTIFS([[Clients.segment]],"{segment}")),1),'
                f'"")', "number1"),
            (7, f'=IFERROR(ROUND(SUMIFS([[Clients.spend]],'
                f'[[Clients.segment]],"{segment}")/'
                f'MAX(1,SUMIFS([[Clients.visits]],'
                f'[[Clients.segment]],"{segment}")),2),"")', "money"),
        ]
        for col, formula, style in formulas:
            ws.write_formula(row, col, s.reg.render(formula),
                             fmts.calc(style))
        row += 1
    row += 1

    row = s.note(
        row,
        "Two tables worth acting on: the hours table tells you when to "
        "run an offer, and the category table tells you which service "
        "to run it on. Discounting the thing that already earns the "
        "most an hour is the most common mistake in the trade.",
        LAST_COL, height=30)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    s.protect()
    return s


def _hour_label(hour: int) -> str:
    suffix = "am" if hour < 12 else "pm"
    display = hour if hour <= 12 else hour - 12
    if hour == 0:
        display = 12
    return f"{display}{suffix}"
````

### `salon_business_spreadsheet/sheets/reports.py`

*401 lines · 1 function*

> `Reports` - the week and the month, on one printable page.
>
> Pick a week ending date and a month at the top; everything below
> re-reads itself. Print it for the Monday meeting, the accountant or
> the folder you keep for the bank.

| Defines | What it does |
|---------|--------------|
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Reports'), `LAST_COL` (9), `TOP_ROWS` (6).

````python
"""`Reports` - the week and the month, on one printable page.

Pick a week ending date and a month at the top; everything below
re-reads itself. Print it for the Monday meeting, the accountant or
the folder you keep for the bank.
"""

from __future__ import annotations

from ..config import Config, Palette
from ..layout import Registry, quote_sheet
from ..writer import SheetWriter, date_check, list_from

NAME = "Reports"
LAST_COL = 9
TOP_ROWS = 6


def build(book, fmts, cfg: Config, reg: Registry) -> SheetWriter:
    s = SheetWriter(book, fmts, cfg, reg, NAME, tab_color=Palette.deep)
    ws = s.ws
    sheet_ref = quote_sheet(NAME)

    s.banner("Weekly & Monthly Reports",
             "Two pickers at the top, two reports underneath. Nothing "
             "to rebuild and nothing to refresh.", LAST_COL)
    s.set_col(0, 0, 3)
    s.set_col(1, 1, 30)
    s.set_col(2, LAST_COL, 15)

    row = 4
    row = s.field(row, "Week ending (a Sunday works best)",
                  cfg.today, style="date", name="ReportWeekEnd",
                  validate=date_check(),
                  help_text="The report covers the seven days up to "
                            "and including this date.")
    row = s.field(row, "Month", f"{cfg.today.year:04d}-"
                                f"{cfg.today.month:02d}",
                  style="text", name="ReportMonth",
                  validate=list_from("MonthKeyList"),
                  help_text="Pick any month in yyyy-mm form.")
    week_start = "(ReportWeekEnd-6)"
    row += 1

    # ------------------------------------------------------- the week ---
    row = s.section(row, "The week", LAST_COL)
    s.merge(row, 1, 4, None, fmts.calc("text"))
    ws.write_formula(row, 1,
                     '="Week of "&TEXT(ReportWeekEnd-6,"d mmm")&'
                     '" to "&TEXT(ReportWeekEnd,"d mmm yyyy")',
                     fmts.calc("text"))
    row += 1
    s.table_head(row, 1, [
        "Measure", "This week", "Week before", "Change",
        "Change %", "Four week average"],
        [30, 15, 15, 14, 12, 18])
    row += 1
    week_first = row
    measures = [
        ("Net sales", "money",
         '=ROUND(SUMIFS([[Sales.net]],[[Sales.date]],">="&{from},'
         '[[Sales.date]],"<="&{to}),2)'),
        ("Service sales", "money",
         '=ROUND(SUMIFS([[Sales.net]],[[Sales.kind]],"Service",'
         '[[Sales.date]],">="&{from},[[Sales.date]],"<="&{to}),2)'),
        ("Retail sales", "money",
         '=ROUND(SUMIFS([[Sales.net]],[[Sales.kind]],"Retail",'
         '[[Sales.date]],">="&{from},[[Sales.date]],"<="&{to}),2)'),
        ("Clients seen", "int",
         '=COUNTIFS([[Appointments.status]],"Completed",'
         '[[Appointments.date]],">="&{from},[[Appointments.date]],'
         '"<="&{to})'),
        ("New clients", "int",
         '=COUNTIFS([[Clients.first_visit]],">="&{from},'
         '[[Clients.first_visit]],"<="&{to})'),
        ("Average ticket", "money",
         '=IFERROR(ROUND(SUMIFS([[Sales.net]],[[Sales.date]],">="&'
         '{from},[[Sales.date]],"<="&{to})/MAX(1,COUNTIFS('
         '[[Appointments.status]],"Completed",[[Appointments.date]],'
         '">="&{from},[[Appointments.date]],"<="&{to})),2),"")'),
        ("Cancellations", "int",
         '=COUNTIFS([[Appointments.status]],"Cancelled",'
         '[[Appointments.date]],">="&{from},[[Appointments.date]],'
         '"<="&{to})'),
        ("No shows", "int",
         '=COUNTIFS([[Appointments.status]],"No show",'
         '[[Appointments.date]],">="&{from},[[Appointments.date]],'
         '"<="&{to})'),
        ("Chair hours worked", "number1",
         '=ROUND(SUMIFS([[Appointments.hours]],'
         '[[Appointments.status]],"Completed",[[Appointments.date]],'
         '">="&{from},[[Appointments.date]],"<="&{to}),1)'),
        ("Tips taken", "money",
         '=ROUND(SUMIFS([[Payments.amount]],[[Payments.kind]],"Tip",'
         '[[Payments.date]],">="&{from},[[Payments.date]],"<="&'
         '{to}),2)'),
        ("Money into the bank", "money",
         '=ROUND(SUMIFS([[Payments.net_cash]],[[Payments.date]],">="&'
         '{from},[[Payments.date]],"<="&{to}),2)'),
        ("Costs logged", "money",
         '=ROUND(SUMIFS([[Expenses.amount]],[[Expenses.date]],">="&'
         '{from},[[Expenses.date]],"<="&{to}),2)'),
    ]
    for label, style, body in measures:
        s.write(row, 1, label, fmts.chrome("label_bold"))
        this_week = body.format(**{"from": week_start,
                                   "to": "ReportWeekEnd"})
        last_week = body.format(**{"from": "(ReportWeekEnd-13)",
                                   "to": "(ReportWeekEnd-7)"})
        four_week = body.format(**{"from": "(ReportWeekEnd-27)",
                                   "to": "ReportWeekEnd"})
        ws.write_formula(row, 2, s.reg.render(this_week),
                         fmts.calc(style))
        ws.write_formula(row, 3, s.reg.render(last_week),
                         fmts.calc(style))
        ws.write_formula(row, 4, f"=ROUND(C{row + 1}-D{row + 1},2)",
                         fmts.calc(style))
        ws.write_formula(row, 5,
                         f'=IFERROR((C{row + 1}-D{row + 1})/'
                         f'ABS(D{row + 1}),"")', fmts.calc("percent"))
        if style == "int":
            ws.write_formula(
                row, 6, f"=ROUND(({s.reg.render(four_week)[1:]})/4,1)",
                fmts.calc("number1"))
        else:
            ws.write_formula(
                row, 6, f"=ROUND(({s.reg.render(four_week)[1:]})/4,2)",
                fmts.calc(style))
        row += 1
    week_last = row - 1
    ws.conditional_format(week_first, 4, week_last, 4,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("good")})
    ws.conditional_format(week_first, 4, week_last, 4,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("warn")})
    row += 1

    # --------------------------------------------------- week, by day ---
    row = s.subsection(row, "The week, day by day", LAST_COL)
    s.table_head(row, 1, [
        "Day", "Date", "Appointments", "Completed", "Net sales",
        "Chair hours", "No shows"],
        [14, 13, 14, 12, 14, 13, 12])
    row += 1
    day_first = row
    for offset in range(6, -1, -1):
        day_ref = f"(ReportWeekEnd-{offset})"
        s.write(row, 1, None, fmts.calc("text"))
        ws.write_formula(row, 1, f'=TEXT({day_ref},"dddd")',
                         fmts.calc("text"))
        ws.write_formula(row, 2, f"={day_ref}", fmts.calc("date"))
        ws.write_formula(
            row, 3,
            s.reg.render(f'=COUNTIFS([[Appointments.date]],{day_ref})'),
            fmts.calc("int"))
        ws.write_formula(
            row, 4,
            s.reg.render(f'=COUNTIFS([[Appointments.date]],{day_ref},'
                         f'[[Appointments.status]],"Completed")'),
            fmts.calc("int"))
        ws.write_formula(
            row, 5,
            s.reg.render(f'=ROUND(SUMIFS([[Sales.net]],[[Sales.date]],'
                         f'{day_ref}),2)'), fmts.calc("money"))
        ws.write_formula(
            row, 6,
            s.reg.render(f'=ROUND(SUMIFS([[Appointments.hours]],'
                         f'[[Appointments.date]],{day_ref},'
                         f'[[Appointments.status]],"Completed"),1)'),
            fmts.calc("number1"))
        ws.write_formula(
            row, 7,
            s.reg.render(f'=COUNTIFS([[Appointments.date]],{day_ref},'
                         f'[[Appointments.status]],"No show")'),
            fmts.calc("int"))
        row += 1
    day_last = row - 1
    ws.conditional_format(day_first, 5, day_last, 5,
                          {"type": "data_bar",
                           "bar_color": Palette.primary_mid})
    ws.add_sparkline(row, 5, {
        "range": (f"{sheet_ref}!$F${day_first + 1}:$F${day_last + 1}"),
        "type": "column", "series_color": Palette.primary_mid})
    s.merge(row, 1, 4, "Net sales through the week",
            fmts.chrome("note"))
    row += 2

    # ------------------------------------------------------ the month ---
    row = s.section(row, "The month", LAST_COL)
    s.merge(row, 1, 4, None, fmts.calc("text"))
    ws.write_formula(
        row, 1,
        '="Month of "&TEXT(DATE(VALUE(LEFT(ReportMonth,4)),'
        'VALUE(RIGHT(ReportMonth,2)),1),"mmmm yyyy")',
        fmts.calc("text"))
    row += 1
    s.table_head(row, 1, [
        "Measure", "This month", "Month before", "Change",
        "Change %", "Share of sales"],
        [30, 15, 15, 14, 12, 15])
    row += 1
    prev = ('TEXT(DATE(VALUE(LEFT(ReportMonth,4)),'
            'VALUE(RIGHT(ReportMonth,2))-1,1),"yyyy-mm")')
    month_first = row
    month_measures = [
        ("Net sales", "money", "yes",
         '=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{m}),2)'),
        ("Service sales", "money", "yes",
         '=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{m},'
         '[[Sales.kind]],"Service"),2)'),
        ("Retail sales", "money", "yes",
         '=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{m},'
         '[[Sales.kind]],"Retail"),2)'),
        ("Packages and vouchers", "money", "yes",
         '=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{m},'
         '[[Sales.kind]],"Package")+SUMIFS([[Sales.net]],'
         '[[Sales.month]],{m},[[Sales.kind]],"Gift certificate"),2)'),
        ("Discounts given", "money", "yes",
         '=ROUND(SUMIFS([[Sales.discount]],[[Sales.month]],{m}),2)'),
        ("Refunds", "money", "yes",
         '=ROUND(SUMIFS([[Sales.refund]],[[Sales.month]],{m}),2)'),
        ("Cost of sales", "money", "yes",
         '=ROUND(SUMIFS([[Sales.cost]],[[Sales.month]],{m}),2)'),
        ("Staff cost", "money", "yes",
         '=ROUND(SUMIFS([[Payroll.total_cost]],[[Payroll.month]],'
         '{m}),2)'),
        ("Other expenses", "money", "yes",
         '=ROUND(SUMIFS([[Expenses.amount]],[[Expenses.month]],'
         '{m}),2)'),
        ("Operating profit", "money", "yes",
         '=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],{m})-'
         'SUMIFS([[Sales.cost]],[[Sales.month]],{m})-'
         'SUMIFS([[Payroll.total_cost]],[[Payroll.month]],{m})-'
         'SUMIFS([[Expenses.amount]],[[Expenses.month]],{m}),2)'),
        ("Clients seen", "int", "",
         '=COUNTIFS([[Appointments.month]],{m},'
         '[[Appointments.status]],"Completed")'),
        ("Cancellations", "int", "",
         '=COUNTIFS([[Appointments.month]],{m},'
         '[[Appointments.status]],"Cancelled")'),
        ("No shows", "int", "",
         '=COUNTIFS([[Appointments.month]],{m},'
         '[[Appointments.status]],"No show")'),
        ("Tips", "money", "",
         '=ROUND(SUMIFS([[Payments.amount]],[[Payments.month]],{m},'
         '[[Payments.kind]],"Tip"),2)'),
        ("Sales tax collected", "money", "",
         '=ROUND(SUMIFS([[Sales.tax]],[[Sales.month]],{m}),2)'),
        ("Deposits taken", "money", "",
         '=ROUND(SUMIFS([[Payments.amount]],[[Payments.month]],{m},'
         '[[Payments.kind]],"Deposit"),2)'),
    ]
    sales_row = month_first                     # net sales row, for shares
    for label, style, share, body in month_measures:
        s.write(row, 1, label, fmts.chrome("label_bold"))
        ws.write_formula(row, 2, s.reg.render(body.format(m="ReportMonth")),
                         fmts.calc(style))
        ws.write_formula(row, 3, s.reg.render(body.format(m=prev)),
                         fmts.calc(style))
        ws.write_formula(row, 4, f"=ROUND(C{row + 1}-D{row + 1},2)",
                         fmts.calc(style))
        ws.write_formula(row, 5,
                         f'=IFERROR((C{row + 1}-D{row + 1})/'
                         f'ABS(D{row + 1}),"")', fmts.calc("percent"))
        if share:
            ws.write_formula(
                row, 6,
                f'=IF(C{sales_row + 1}=0,"",C{row + 1}/'
                f'C{sales_row + 1})', fmts.calc("percent"))
        else:
            s.write(row, 6, None, fmts.calc("percent"))
        row += 1
    month_last = row - 1
    ws.conditional_format(month_first, 4, month_last, 4,
                          {"type": "cell", "criteria": ">", "value": 0,
                           "format": fmts.cf("good")})
    ws.conditional_format(month_first, 4, month_last, 4,
                          {"type": "cell", "criteria": "<", "value": 0,
                           "format": fmts.cf("warn")})
    row += 1

    # ------------------------------------------------- month, by person --
    row = s.subsection(row, "The month, by stylist", LAST_COL)
    s.table_head(row, 1, [
        "Stylist", "Service sales", "Retail sales", "Commission",
        "Tips", "Clients seen", "Average ticket"],
        [20, 15, 14, 14, 12, 13, 14])
    row += 1
    for n in range(1, 9):
        name = (f'IFERROR(INDEX([[Staff.name]],MATCH({n},'
                f'[[Staff.rev_total_rank]],0)),"")')
        who = (f'IFERROR(INDEX([[Staff.id]],MATCH({n},'
               f'[[Staff.rev_total_rank]],0)),"")')
        s.write(row, 1, None, fmts.calc("text"))
        ws.write_formula(row, 1, s.reg.render("=" + name),
                         fmts.calc("text"))
        cells = [
            (2, f'=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],'
                f'ReportMonth,[[Sales.kind]],"Service",'
                f'[[Sales.staff]],{who}),2)', "money"),
            (3, f'=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],'
                f'ReportMonth,[[Sales.kind]],"Retail",'
                f'[[Sales.staff]],{who}),2)', "money"),
            (4, f'=ROUND(SUMIFS([[Commissions.commission]],'
                f'[[Commissions.month]],ReportMonth,'
                f'[[Commissions.staff]],{who}),2)', "money"),
            (5, f'=ROUND(SUMIFS([[Payments.amount]],'
                f'[[Payments.month]],ReportMonth,[[Payments.kind]],'
                f'"Tip",[[Payments.staff]],{who}),2)', "money"),
            (6, f'=COUNTIFS([[Appointments.month]],ReportMonth,'
                f'[[Appointments.status]],"Completed",'
                f'[[Appointments.staff]],{who})', "int"),
        ]
        for col, formula, style in cells:
            ws.write_formula(row, col, s.reg.render(formula),
                             fmts.calc(style))
        ws.write_formula(row, 7,
                         f'=IF(G{row + 1}=0,"",ROUND((C{row + 1}+'
                         f'D{row + 1})/G{row + 1},2))',
                         fmts.calc("money"))
        row += 1
    row += 1

    # ---------------------------------------------- month, top services --
    row = s.subsection(row, "The month's best sellers", LAST_COL)
    s.table_head(row, 1, [
        "Service", "Times", "Revenue", "Client", "Visits", "Spend"],
        [26, 10, 14, 24, 10, 14])
    row += 1
    for n in range(1, TOP_ROWS + 1):
        service = (f'IFERROR(INDEX([[Services.name]],MATCH({n},'
                   f'[[Services.revenue_rank]],0)),"")')
        service_id = (f'IFERROR(INDEX([[Services.id]],MATCH({n},'
                      f'[[Services.revenue_rank]],0)),"")')
        client = (f'IFERROR(INDEX([[Clients.name]],MATCH({n},'
                  f'[[Clients.spend_rank]],0)),"")')
        client_id = (f'IFERROR(INDEX([[Clients.id]],MATCH({n},'
                     f'[[Clients.spend_rank]],0)),"")')
        cells = [
            (1, "=" + service, "text"),
            (2, f'=COUNTIFS([[Sales.month]],ReportMonth,[[Sales.item]],'
                f'{service_id})', "int"),
            (3, f'=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],'
                f'ReportMonth,[[Sales.item]],{service_id}),2)', "money"),
            (4, "=" + client, "text"),
            (5, f'=COUNTIFS([[Appointments.month]],ReportMonth,'
                f'[[Appointments.client]],{client_id},'
                f'[[Appointments.status]],"Completed")', "int"),
            (6, f'=ROUND(SUMIFS([[Sales.net]],[[Sales.month]],'
                f'ReportMonth,[[Sales.client]],{client_id}),2)',
             "money"),
        ]
        for col, formula, style in cells:
            ws.write_formula(row, col, s.reg.render(formula),
                             fmts.calc(style))
        row += 1
    row += 1

    # --------------------------------------------------- year to date ---
    row = s.section(row, "Year to date", LAST_COL)
    s.table_head(row, 1, [
        "Net sales", "Cost of sales", "Staff cost", "Other expenses",
        "Operating profit", "Margin", "Clients seen", "New clients"],
        [16, 15, 14, 15, 16, 10, 13, 13])
    row += 1
    ytd = '">="&DATE(YEAR(TodayDate),1,1)'
    cells = [
        (1, f'=ROUND(SUMIFS([[Sales.net]],[[Sales.date]],{ytd}),2)',
         "money"),
        (2, f'=ROUND(SUMIFS([[Sales.cost]],[[Sales.date]],{ytd}),2)',
         "money"),
        (3, f'=ROUND(SUMIFS([[Payroll.total_cost]],'
            f'[[Payroll.period_end]],{ytd}),2)', "money"),
        (4, f'=ROUND(SUMIFS([[Expenses.amount]],[[Expenses.date]],'
            f'{ytd}),2)', "money"),
        (7, f'=COUNTIFS([[Appointments.status]],"Completed",'
            f'[[Appointments.date]],{ytd})', "int"),
        (8, f'=COUNTIFS([[Clients.first_visit]],{ytd})', "int"),
    ]
    for col, formula, style in cells:
        ws.write_formula(row, col, s.reg.render(formula),
                         fmts.calc(style))
    ws.write_formula(row, 5,
                     f"=ROUND(B{row + 1}-C{row + 1}-D{row + 1}-"
                     f"E{row + 1},2)", fmts.calc("money"))
    ws.write_formula(row, 6,
                     f'=IF(B{row + 1}=0,"",F{row + 1}/B{row + 1})',
                     fmts.calc("percent"))
    row += 2

    row = s.note(
        row,
        "Change the week ending date or the month at the top and every "
        "number on this tab moves with it. Print range is set, so "
        "Ctrl+P gives you something you can hand over.",
        LAST_COL, height=28)

    s.finish(freeze_row=6, gridlines=False, fit_width=True)
    s.protect()
    return s
````

### `salon_business_spreadsheet/sheets/settings.py`

*495 lines · 2 functions*

> `Settings` - every number and every list the workbook runs on.
>
> Nothing in this workbook is hard coded to one country, one pay model
> or one price list. The tax switches, the commission basis, the
> rounding rule, the loyalty scheme and all twenty-odd drop-downs are
> set here, once, and every other tab follows.

| Defines | What it does |
|---------|--------------|
| `_month_keys(today, back, ahead)` | `yyyy-mm` keys around today, oldest first |
| `build(book, fmts, cfg, reg)` |  |

**Module constants:** `NAME` ('Settings'), `LAST_COL` (11), `LISTS` (24 items).

````python
"""`Settings` - every number and every list the workbook runs on.

Nothing in this workbook is hard coded to one country, one pay model
or one price list. The tax switches, the commission basis, the
rounding rule, the loyalty scheme and all twenty-odd drop-downs are
set here, once, and every other tab follows.
"""

from __future__ import annotations

import datetime as _dt
from typing import List, Tuple

from ..config import (APPOINTMENT_STATUSES, BOOKING_SOURCES,
                      CAMPAIGN_STATUSES, CANCEL_REASONS, CLIENT_STATUSES,
                      COMMISSION_METHODS, COMMISSION_TIERS,
                      CONTACT_METHODS, EXPENSE_CATEGORIES, FREQUENCIES,
                      ID_PREFIXES, LOYALTY_REWARDS, MARKETING_CHANNELS,
                      PAY_PERIODS, PAYMENT_KINDS, PAYMENT_METHODS,
                      PRODUCT_CATEGORIES, PURCHASE_STATUSES,
                      DISCLAIMER, ROUNDING_RULES, SALE_KINDS,
                      SERVICE_CATEGORIES,
                      SERVICE_STATUSES, STAFF_STATUSES, STAFF_TYPES,
                      STYLIST_LEVELS, Config, Palette)
from ..layout import Registry, col_letter, dynamic_name, quote_sheet
from ..writer import SheetWriter, list_from, number_check

NAME = "Settings"
LAST_COL = 11

#: name -> values. Everything the workbook offers you in a drop-down.
LISTS: List[Tuple[str, List[str]]] = [
    ("YesNoList", ["Yes", "No"]),
    ("ServiceCategoryList", list(SERVICE_CATEGORIES)),
    ("ServiceStatusList", list(SERVICE_STATUSES)),
    ("StylistLevelList", list(STYLIST_LEVELS)),
    ("AppointmentStatusList", list(APPOINTMENT_STATUSES)),
    ("SourceList", list(BOOKING_SOURCES)),
    ("CancelReasonList", list(CANCEL_REASONS)),
    ("ContactList", list(CONTACT_METHODS)),
    ("ClientStatusList", list(CLIENT_STATUSES)),
    ("SaleKindList", list(SALE_KINDS)),
    ("PaymentKindList", list(PAYMENT_KINDS)),
    ("PaymentMethodList", list(PAYMENT_METHODS)),
    ("StaffTypeList", list(STAFF_TYPES)),
    ("StaffStatusList", list(STAFF_STATUSES)),
    ("CommissionMethodList", list(COMMISSION_METHODS)),
    ("PayPeriodList", list(PAY_PERIODS)),
    ("ExpenseCategoryList", list(EXPENSE_CATEGORIES)),
    ("FrequencyList", list(FREQUENCIES)),
    ("ProductCategoryList", list(PRODUCT_CATEGORIES)),
    ("PurchaseStatusList", list(PURCHASE_STATUSES)),
    ("MarketingChannelList", list(MARKETING_CHANNELS)),
    ("CampaignStatusList", list(CAMPAIGN_STATUSES)),
    ("LoyaltyRewardList", list(LOYALTY_REWARDS)),
    ("RoundingRuleList", list(ROUNDING_RULES)),
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
             "Your salon, your prices, your pay model, your lists. Set "
             "these once and every tab in the workbook follows.",
             LAST_COL)
    s.set_col(0, 0, 3)
    s.set_col(1, 2, 22)
    s.set_col(3, 3, 20)
    s.set_col(4, LAST_COL, 17)

    row = 4
    row = s.note(
        row,
        "Blue cells are yours to change. Everything else is locked and "
        "calculated. Nothing here is specific to one country: set your "
        "own tax rate and switches, your own commission basis, your own "
        "currency, and the workbook does the rest.",
        LAST_COL, height=32)
    row += 1

    # ------------------------------------------------------- the salon ---
    row = s.section(row, "The salon", LAST_COL)
    row = s.field(row, "Salon name", cfg.salon.name, name="SalonName",
                  value_end=5, help_col=7,
                  help_text="Printed at the top of every report.")
    row = s.field(row, "Owner", cfg.salon.owner, name="OwnerName",
                  value_end=5, help_col=7)
    row = s.field(row, "Address", cfg.salon.address, value_end=5,
                  help_col=7)
    row = s.field(row, "Phone", cfg.salon.phone, value_end=5, help_col=7)
    row = s.field(row, "Email", cfg.salon.email, value_end=5, help_col=7)
    row = s.field(row, "Website", cfg.salon.website, value_end=5,
                  help_col=7)
    row = s.field(row, "Booking link", cfg.salon.booking_link,
                  value_end=5, help_col=7,
                  help_text="Where clients book. This workbook does "
                            "not connect to it.")
    row = s.field(row, "Social", cfg.salon.instagram, value_end=5,
                  help_col=7, help_text="Handle used on the marketing "
                                        "tab.")
    row = s.field(row, "Strapline", cfg.salon.tagline, value_end=5,
                  help_col=7)
    row = s.field(row, "Opening hours", cfg.salon.opening_hours,
                  value_end=5, help_col=7)
    row = s.field(row, "Tax / business number", cfg.salon.tax_id,
                  value_end=5, help_col=7)
    row = s.field(row, "Currency symbol", cfg.currency_symbol,
                  name="CurrencySign",
                  help_text="Used on screen. The cell formats follow the "
                            "symbol you build the file with.")
    row = s.field(row, "Currency code", cfg.currency_code)
    row = s.field(row, "Tracking started", cfg.today - _dt.timedelta(days=70),
                  style="date", name="TrackingStart",
                  help_text="The day you began logging in this file. "
                            "Stock cover and averages measure from here.")
    row = s.field(row, "Today", None, style="date", name="TodayDate",
                  mode="calc", formula="=TODAY()",
                  help_text="Live. Everything that says 'overdue', "
                            "'lapsing' or 'this month' reads this cell.")
    row = s.field(row, "This month", None, style="text",
                  name="ThisMonthKey", mode="calc",
                  formula='=TEXT(TODAY(),"yyyy-mm")',
                  help_text="The month key the dashboard totals on.")
    row = s.field(row, "Weeks tracked", None, style="number1",
                  name="TrackingWeeks", mode="calc",
                  formula="=MAX(4,ROUND((TodayDate-TrackingStart)/7,0))",
                  help_text="Used for 'used a week' and the stock cover "
                            "figures.")
    row += 1

    # ------------------------------------------------- opening & capacity --
    row = s.section(row, "Opening hours and capacity", LAST_COL)
    row = s.field(row, "Opens at", cfg.open_hour, style="number1",
                  name="OpenHour", validate=number_check(0, 24),
                  help_text="As a 24 hour number - 9 is 9am, 19 is 7pm.")
    row = s.field(row, "Closes at", cfg.close_hour, style="number1",
                  name="CloseHour", validate=number_check(0, 24))
    row = s.field(row, "Days open a week", cfg.open_days, style="int",
                  name="OpenDays", validate=number_check(1, 7,
                                                         decimal=False))
    row = s.field(row, "Chairs / stations", cfg.chairs, style="int",
                  name="Chairs", validate=number_check(1, 60,
                                                       decimal=False))
    row = s.field(row, "Sellable chair hours a week", None,
                  style="number1", name="ChairHoursWeek", mode="calc",
                  formula="=MAX(1,(CloseHour-OpenHour)*OpenDays*Chairs)",
                  help_text="Hours x days x chairs. The utilisation "
                            "figure on the dashboard divides by this.")
    row += 1

    # --------------------------------------------------------- money & tax
    row = s.section(row, "Money, tax and deposits", LAST_COL)
    row = s.note(
        row,
        "Sales tax works differently everywhere, so nothing is assumed. "
        "Set the rate and say what it applies to - retail only, services "
        "only, both, or neither. Set both switches to No if tax is not "
        "something you charge.",
        LAST_COL, height=30)
    row += 1
    row = s.field(row, "Sales tax rate", cfg.tax_rate, style="percent",
                  name="TaxRate", validate=number_check(0, 1),
                  help_text="8.25% is 0.0825 or 8.25% - both work.")
    row = s.field(row, "Tax on services?", cfg.tax_on_services,
                  name="TaxOnServices", validate=list_from("YesNoList"))
    row = s.field(row, "Tax on retail?", cfg.tax_on_retail,
                  name="TaxOnRetail", validate=list_from("YesNoList"))
    row = s.field(row, "Tax registered?", cfg.tax_registered,
                  validate=list_from("YesNoList"),
                  help_text="For your own reference - it changes "
                            "nothing in the formulas.")
    row = s.field(row, "Card fee, percent", cfg.card_fee_percent,
                  style="percent", name="CardFeePercent",
                  validate=number_check(0, 1),
                  help_text="What the card machine takes on each "
                            "payment.")
    row = s.field(row, "Card fee, fixed", cfg.card_fee_fixed,
                  style="money", name="CardFeeFixed",
                  validate=number_check(0, 100))
    row = s.field(row, "Deposit, percent of price", cfg.deposit_percent,
                  style="percent", name="DepositPercent",
                  validate=number_check(0, 1))
    row = s.field(row, "Ask for a deposit above", cfg.deposit_threshold,
                  style="money", name="DepositThreshold",
                  validate=number_check(0, 100000),
                  help_text="Bookings worth more than this suggest a "
                            "deposit on the appointment tab.")
    row = s.field(row, "Cancellation notice, hours",
                  cfg.cancellation_hours, style="int",
                  name="CancellationHours",
                  validate=number_check(0, 336, decimal=False))
    row = s.field(row, "No-show fee", cfg.no_show_fee, style="money",
                  name="NoShowFee", validate=number_check(0, 10000))
    row = s.field(row, "Price rounding", cfg.rounding_rule,
                  name="RoundingRule",
                  validate=list_from("RoundingRuleList"),
                  help_text="How calculated prices are rounded. Every "
                            "rule rounds up, never down.")
    row = s.field(row, "Target chair rate an hour", cfg.target_chair_rate,
                  style="money", name="TargetChairRate",
                  validate=number_check(0, 10000),
                  help_text="What an hour of chair time has to earn "
                            "before commission. The suggested prices "
                            "on the service menu are built from it.")
    row = s.field(row, "Review prices after, days", cfg.price_review_days,
                  style="int", name="PriceReviewDays",
                  validate=number_check(0, 3650, decimal=False))
    row += 1

    # ----------------------------------------------------------- staff pay
    row = s.section(row, "How the team is paid", LAST_COL)
    row = s.note(
        row,
        "These are defaults, not rules. Each person carries their own "
        "method and their own rate on the Staff tab; these fill the gaps "
        "when a rate is left blank.",
        LAST_COL, height=28)
    row += 1
    row = s.field(row, "Default service commission",
                  cfg.service_commission, style="percent",
                  name="ServiceCommission", validate=number_check(0, 1))
    row = s.field(row, "Default retail commission", cfg.retail_commission,
                  style="percent", name="RetailCommission",
                  validate=number_check(0, 1))
    row = s.field(row, "Apprentice commission", cfg.apprentice_commission,
                  style="percent", name="ApprenticeCommission",
                  validate=number_check(0, 1))
    row = s.field(row, "Default hourly rate", cfg.hourly_rate,
                  style="money", name="DefaultHourly",
                  validate=number_check(0, 1000))
    row = s.field(row, "Employer costs on wages", cfg.payroll_burden,
                  style="percent", name="PayrollBurden",
                  validate=number_check(0, 1),
                  help_text="Employer tax, insurance and holiday "
                            "accrual as a percentage of gross pay.")
    row = s.field(row, "Commission is paid on", cfg.commission_on,
                  value_end=5, name="CommissionOn",
                  validate={"validate": "list",
                            "source": ["Net of discounts",
                                       "Gross before discounts"]},
                  help_col=7,
                  help_text="One switch, applied to everybody.")
    row = s.field(row, "Chair rent a week", cfg.chair_rent_week,
                  style="money", name="ChairRentWeek",
                  validate=number_check(0, 10000))
    row += 1

    row = s.subsection(row, "Tiered commission ladder", LAST_COL)
    row = s.note(
        row,
        "Used by the 'tiered on monthly service revenue' method. The "
        "rate that applies is the highest band the stylist's month "
        "reaches. Edit the bands to match your own ladder - keep the "
        "'from' column in ascending order.",
        LAST_COL, height=30)
    row += 1
    s.table_head(row, 1, ["Monthly service revenue from", "Rate"])
    first_tier = row + 1
    tier_rows = len(COMMISSION_TIERS) + 3
    for index in range(tier_rows):
        if index < len(COMMISSION_TIERS):
            floor, rate = COMMISSION_TIERS[index]
            s.write(first_tier + index, 1, floor, fmts.input("money"))
            s.write(first_tier + index, 2, rate, fmts.input("percent"))
        else:
            s.write(first_tier + index, 1, None, fmts.input("money"))
            s.write(first_tier + index, 2, None, fmts.input("percent"))
    book.define_name(
        "TierFrom",
        f"={sheet_ref}!$B${first_tier + 1}:$B${first_tier + tier_rows}")
    book.define_name(
        "TierRate",
        f"={sheet_ref}!$C${first_tier + 1}:$C${first_tier + tier_rows}")
    row = first_tier + tier_rows + 1

    # ------------------------------------------------------------- targets
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
    row = s.field(row, "Average ticket target", cfg.target_average_ticket,
                  style="money", name="TargetTicket",
                  validate=number_check(0, 100000))
    row = s.field(row, "Retail share target", cfg.target_retail_percent,
                  style="percent", name="TargetRetailShare",
                  validate=number_check(0, 1),
                  help_text="Retail as a share of everything you sell. "
                            "Ten to fifteen percent is a good salon.")
    row = s.field(row, "Rebooking target", cfg.target_rebook_rate,
                  style="percent", name="TargetRebook",
                  validate=number_check(0, 1))
    row = s.field(row, "Retention target", cfg.target_retention,
                  style="percent", name="TargetRetention",
                  validate=number_check(0, 1))
    row = s.field(row, "Utilisation target", cfg.target_utilisation,
                  style="percent", name="TargetUtilisation",
                  validate=number_check(0, 1))
    row = s.field(row, "VIP spend, 12 months", cfg.vip_spend_year,
                  style="money", name="VipSpend",
                  validate=number_check(0, 1000000),
                  help_text="Spend in a year that marks a client as "
                            "VIP on the CRM.")
    row += 1

    # --------------------------------------------------------------- stock
    row = s.section(row, "Stock and ordering", LAST_COL)
    row = s.field(row, "Retail markup on cost", cfg.retail_markup,
                  style="number1", name="RetailMarkup",
                  validate=number_check(1, 20),
                  help_text="2.0 means a $10 product retails at $20 "
                            "before rounding.")
    row = s.field(row, "Reorder cover, weeks",
                  round(cfg.default_reorder_days / 7, 1),
                  style="number1", name="ReorderWeeks",
                  validate=number_check(0.5, 52),
                  help_text="How many weeks of stock an order should "
                            "bring in.")
    row = s.field(row, "Warn when stock falls below, weeks",
                  cfg.stock_warning_weeks, style="number1",
                  name="StockWarningWeeks", validate=number_check(0, 52))
    row = s.field(row, "Backbar cost as a share of a service",
                  cfg.product_cost_percent, style="percent",
                  name="ProductCostPercent", validate=number_check(0, 1),
                  help_text="Only used when a service has no product "
                            "cost of its own.")
    row += 1

    # -------------------------------------------------------- client care
    row = s.section(row, "Client care, loyalty and referrals", LAST_COL)
    row = s.field(row, "Usual gap between visits, weeks",
                  cfg.rebook_weeks, style="int", name="RebookWeeks",
                  validate=number_check(1, 104, decimal=False))
    row = s.field(row, "Lapsing after, days", cfg.lapsing_days,
                  style="int", name="LapsingDays",
                  validate=number_check(1, 1000, decimal=False))
    row = s.field(row, "Inactive after, days", cfg.inactive_days,
                  style="int", name="InactiveDays",
                  validate=number_check(1, 2000, decimal=False))
    row = s.field(row, "Patch test valid for, days", cfg.patch_test_days,
                  style="int", name="PatchTestDays",
                  validate=number_check(1, 2000, decimal=False))
    row = s.field(row, "Loyalty points per unit spent",
                  cfg.loyalty_points_per, style="number1",
                  name="PointsPerUnit", validate=number_check(0, 100))
    row = s.field(row, "Reward at, points", cfg.loyalty_reward_at,
                  style="int", name="RewardAt",
                  validate=number_check(1, 100000, decimal=False))
    row = s.field(row, "Reward is worth", cfg.loyalty_reward_value,
                  style="money", name="RewardValue",
                  validate=number_check(0, 10000),
                  help_text="One point per dollar with a $20 reward at "
                            "200 points is a 10% scheme. Worth knowing "
                            "before you promise it.")
    row = s.field(row, "Referral reward", cfg.referral_reward,
                  style="money", name="ReferralReward",
                  validate=number_check(0, 10000))
    row += 1

    # ------------------------------------------------------------ prefixes
    row = s.section(row, "ID prefixes", LAST_COL)
    row = s.note(
        row,
        "Every record numbers itself from its row. Change a prefix here "
        "and the IDs on that tab change with it - but do it before you "
        "have typed a hundred appointments, not after.",
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
            name,
            f"={sheet_ref}!${col_letter(column + 2)}${line + 1}")
    row = top + half + 1

    # -------------------------------------------------- calculated totals
    row = s.section(row, "Workbook totals", LAST_COL)
    row = s.note(
        row,
        "Calculated once here so that several tabs can share them. "
        "Nothing to change.",
        LAST_COL, height=24)
    row += 1
    row = s.field(row, "Service revenue, all time", None, style="money",
                  name="ServiceRevenueAll", mode="calc",
                  formula='=ROUND(SUMIFS([[Sales.net]],[[Sales.kind]],'
                          '"Service"),2)')
    row = s.field(row, "Retail revenue, all time", None, style="money",
                  name="RetailRevenueAll", mode="calc",
                  formula='=ROUND(SUMIFS([[Sales.net]],[[Sales.kind]],'
                          '"Retail"),2)')
    row = s.field(row, "Net sales, all time", None, style="money",
                  name="NetSalesAll", mode="calc",
                  formula="=ROUND(SUM([[Sales.net]]),2)")
    row = s.field(row, "Clients on file", None, style="int",
                  name="ClientCount", mode="calc",
                  formula='=COUNTIF([[Clients.name]],"<>")')
    row += 1

    # ---------------------------------------------------------- the lists
    row = s.section(row, "Drop-down lists", LAST_COL)
    row = s.note(
        row,
        "Add new entries directly under the last one in a column - every "
        "list has room to grow and blank rows are ignored. Do not delete "
        "a whole column: a drop-down pointing at a missing list stops "
        "working.",
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

    # ------------------------------------------------------- month keys --
    row = s.section(row, "Month keys", LAST_COL)
    row = s.note(
        row,
        "The months the commission, payroll and budget tabs pick from. "
        "They are plain text in yyyy-mm form so that every monthly total "
        "in the workbook groups on exactly the same thing.",
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

    row = s.note(row, DISCLAIMER, LAST_COL, height=44)

    s.finish(freeze_row=4, gridlines=False, fit_width=True)
    s.protect()
    return s
````

---

## 9. The quality tools

Not shipped with the workbook, but the reason it can be trusted: a static validator, a bracket and placeholder checker, a demo data checker that rebuilds the arithmetic in Python, and a recalculation harness that really evaluates every formula.

### `tools/validate_salon_workbook.py`

*654 lines · 11 functions*

> Static checks on a generated Salon Business workbook.
>
> Run it with any Python that has openpyxl installed::
>
>     python tools/validate_salon_workbook.py output/salon.xlsx
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
| `_name_target(wb, name)` | The one cell a defined name points at, or None |
| `check_pages(wb, problems, notes)` | The calculator pages, checked through their defined names |
| `_specs()` | Every register specification, keyed by sheet name |
| `check_chain(wb, problems, notes)` | The quote-to-job-to-invoice chain only holds while these do |
| `check(path)` |  |

**Module constants:** `FUNCTION_RE` (re.compile(...)), `NAME_RE` (re.compile(...)), `BARE_COL_RE` (re.compile(...)), `KNOWN_FUNCTIONS` ({'IF', 'IFERROR', 'SUM', 'SUMIF', 'SUMIFS', 'COUNT', 'COU...), `MODERN_ONLY` ({'XLOOKUP', 'IFS', 'MAXIFS', 'MINIFS', 'TEXTJOIN', 'CONCA...), `PRECEDENCE_OK` ({'D#+U#/#', 'D#+U#/#-TodayDate'}), `SEEDED_INPUTS` (0 entries), `CHAIN_CHECKS` (57 items), `PAGE_FORMULAS` (14 items), `PAGE_INPUTS` (15 items).

````python
"""Static checks on a generated Salon Business workbook.

Run it with any Python that has openpyxl installed::

    python tools/validate_salon_workbook.py output/salon.xlsx

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
from salon_business_spreadsheet.build import SHEET_ORDER as EXPECTED_SHEETS  # noqa: E402

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
    "TIMEVALUE", "WORKDAY", "EOMONTH", "FREQUENCY", "PERCENTILE",
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
    # an appointment's finish time: the start time plus the minutes,
    # and only the minutes are turned into a fraction of a day
    "D#+U#/#",
    "D#+U#/#-TodayDate",
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
        shape = re.sub(r"\d+", "#", argument)
        if shape in PRECEDENCE_OK:
            continue
        sums = _top_level(argument, "+-")
        if not sums:
            continue
        if any(divide > sums[0] for divide in _top_level(argument, "/")):
            return argument
    return ""


#: input cells that ship with a formula in them as a helpful default -
#: the buyer is meant to be able to type a fixed value straight over it
SEEDED_INPUTS = {
    # nothing on the salon workbook ships a formula in a blue cell
}

#: the automation chain, link by link: (sheet, column key, what the
#: formula in that column has to still be pointing at)
CHAIN_CHECKS = [
    # ---- the menu prices and what they earn ------------------------------
    ("Services", "hours", ("column", "Services", "duration")),
    ("Services", "cost_total", ("column", "Services", "cost_product")),
    ("Services", "profit", ("column", "Services", "cost_total")),
    ("Services", "price_hour", ("column", "Services", "hours")),
    ("Services", "revenue", ("sheet", "Sales")),
    # ---- the diary -------------------------------------------------------
    ("Appointments", "price_used", ("sheet", "Services")),
    ("Appointments", "minutes_used", ("sheet", "Services")),
    ("Appointments", "finish", ("column", "Appointments", "minutes_used")),
    ("Appointments", "clash", ("column", "Appointments", "finish")),
    ("Appointments", "billed", ("sheet", "Sales")),
    ("Appointments", "deposit_paid", ("sheet", "Payments")),
    ("Appointments", "is_final", ("text", "Cancelled")),
    # ---- the client book -------------------------------------------------
    ("Clients", "visits", ("sheet", "Appointments")),
    ("Clients", "spend", ("sheet", "Sales")),
    ("Clients", "balance", ("sheet", "Sales")),
    ("Clients", "segment", ("text", "VipSpend")),
    ("Clients", "retained", ("column", "Clients", "visits")),
    # ---- the till --------------------------------------------------------
    ("Sales", "gross", ("column", "Sales", "price_used")),
    ("Sales", "net", ("column", "Sales", "discount")),
    ("Sales", "tax", ("text", "TaxOnRetail")),
    ("Sales", "cost", ("sheet", "Services")),
    ("Sales", "paid_direct", ("sheet", "Payments")),
    ("Sales", "paid_deposit", ("sheet", "Payments")),
    ("Sales", "balance", ("column", "Sales", "paid")),
    # ---- the money -------------------------------------------------------
    ("Payments", "counts", ("text", "Tip")),
    ("Payments", "fee", ("text", "CardFeePercent")),
    ("Payments", "net_cash", ("column", "Payments", "fee")),
    # ---- the team --------------------------------------------------------
    ("Staff", "rev_service", ("sheet", "Sales")),
    ("Staff", "utilisation", ("column", "Staff", "hours_booked")),
    ("Staff", "commission_paid", ("sheet", "Commissions")),
    ("Staff", "cost_month", ("sheet", "Payroll")),
    ("Staff", "margin_month", ("column", "Staff", "cost_month")),
    # ---- commission, every method ---------------------------------------
    ("Commissions", "service_net", ("sheet", "Sales")),
    ("Commissions", "base_service", ("text", "CommissionOn")),
    ("Commissions", "rate_service", ("text", "TierRate")),
    ("Commissions", "commission", ("text", "Percent after product cost")),
    ("Commissions", "payout", ("column", "Commissions", "commission")),
    # ---- payroll ---------------------------------------------------------
    ("Payroll", "base_pay", ("column", "Payroll", "rate_used")),
    ("Payroll", "employer_cost", ("text", "PayrollBurden")),
    ("Payroll", "total_cost", ("column", "Payroll", "employer_cost")),
    ("Payroll", "commission_left", ("column", "Payroll",
                                    "month_commission")),
    # ---- costs and stock -------------------------------------------------
    ("Expenses", "monthly", ("text", "Quarterly")),
    ("Expenses", "budget", ("text", "BudgetAmount")),
    ("Inventory", "received", ("sheet", "Purchases")),
    ("Inventory", "sold", ("sheet", "Sales")),
    ("Inventory", "stock", ("column", "Inventory", "received")),
    ("Inventory", "status", ("column", "Inventory", "reorder_used")),
    ("Inventory", "suggested_raw", ("text", "RetailMarkup")),
    ("Inventory", "suggested", ("column", "Inventory",
                                "suggested_raw")),
    ("Purchases", "qty_in", ("text", "Part received")),
    ("Purchases", "price_change", ("sheet", "Inventory")),
    # ---- marketing and loyalty ------------------------------------------
    ("Marketing", "revenue", ("sheet", "Sales")),
    ("Marketing", "total_cost", ("column", "Marketing", "given")),
    ("Marketing", "roi", ("column", "Marketing", "profit")),
    ("Loyalty", "earned", ("text", "PointsPerUnit")),
    ("Loyalty", "balance", ("column", "Loyalty", "earned")),
    ("Loyalty", "referrals", ("sheet", "Clients")),
]

#: cells reached by a defined name that have to hold a live formula -
#: the calculator pages are where the value of the file lives, and a
#: name pointing at an empty cell means a page quietly stopped working.
PAGE_FORMULAS = [
    ("TodayDate", "settings: the live date every alert reads"),
    ("ThisMonthKey", "settings: the month key the totals group on"),
    ("TrackingWeeks", "settings: weeks of history, for the averages"),
    ("ChairHoursWeek", "settings: sellable chair hours a week"),
    ("ServiceRevenueAll", "settings: service revenue, all time"),
    ("RetailRevenueAll", "settings: retail revenue, all time"),
    ("NetSalesAll", "settings: net sales, all time"),
    ("ClientCount", "settings: how many clients are on file"),
    ("DashKey", "dashboard: the month the tiles report on"),
    ("DashPrevKey", "dashboard: last month, for the comparisons"),
    ("MonthNetSales", "dashboard: net sales this month"),
    ("MonthCostOfSales", "dashboard: cost of what was sold"),
    ("MonthWages", "dashboard: the pay runs this month"),
    ("MonthExpenses", "dashboard: the bills this month"),
]

#: cells a buyer has to be able to type in
PAGE_INPUTS = [
    ("SalonName", "settings: the salon's name"),
    ("TaxRate", "settings: the sales tax rate"),
    ("TaxOnServices", "settings: does tax apply to services"),
    ("TaxOnRetail", "settings: does tax apply to retail"),
    ("CommissionOn", "settings: net or gross commission basis"),
    ("ServiceCommission", "settings: default service commission"),
    ("PayrollBurden", "settings: employer costs on wages"),
    ("DepositPercent", "settings: deposit as a share of the price"),
    ("RoundingRule", "settings: how prices are rounded"),
    ("TargetChairRate", "settings: what an hour of chair time must earn"),
    ("RevenueTargetMonth", "settings: the monthly revenue target"),
    ("PointsPerUnit", "settings: loyalty points per unit spent"),
    ("TrackingStart", "settings: the day tracking began"),
    ("ReportWeekEnd", "reports: the week being reported on"),
    ("ReportMonth", "reports: the month being reported on"),
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
    """The calculator pages, checked through their defined names."""
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
    notes.append(f"calculator pages: {live} live formulas, "
                 f"{typed} typeable inputs verified")


def _specs():
    """Every register specification, keyed by sheet name."""
    from salon_business_spreadsheet.build import REGISTER_MODULES
    from salon_business_spreadsheet.config import Config
    cfg = Config()
    return {module.NAME: module.spec(cfg) for module in REGISTER_MODULES}


def check_chain(wb, problems: list, notes: list) -> None:
    """The quote-to-job-to-invoice chain only holds while these do.

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


def check(path: str) -> int:
    problems: list[str] = list(check_raw_xml(path))
    notes: list[str] = []
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
              else "output/Salon Business Management Spreadsheet.xlsx")
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

### `tools/check_salon_sample.py`

*451 lines · 6 functions*

> Does the demo salon actually hang together?
>
> The validator checks the formulas; this checks the *story* they are
> told. It reads only the literal cells a user would have typed - never
> a formula column - rebuilds the arithmetic in Python, and complains if
> the example data would make the workbook look wrong on the first open:
> double bookings, visits with nothing rung up, payments that do not
> match the till, commission paid that was never earned, stock that goes
> negative, IDs that point at nothing.
>
>     python tools/check_salon_sample.py output/salon.xlsx

| Defines | What it does |
|---------|--------------|
| `_specs(cfg)` |  |
| `_rows(ws, spec, key_column)` | Every filled data row as {column key: literal value} |
| `_month(value)` |  |
| `_day(value)` |  |
| `_minutes(value)` |  |
| `check(path)` |  |

**Module constants:** `MONEY` (0.015).

````python
#!/usr/bin/env python3
"""Does the demo salon actually hang together?

The validator checks the formulas; this checks the *story* they are
told. It reads only the literal cells a user would have typed - never
a formula column - rebuilds the arithmetic in Python, and complains if
the example data would make the workbook look wrong on the first open:
double bookings, visits with nothing rung up, payments that do not
match the till, commission paid that was never earned, stock that goes
negative, IDs that point at nothing.

    python tools/check_salon_sample.py output/salon.xlsx
"""

from __future__ import annotations

import datetime as _dt
import os
import re
import sys
from collections import Counter, defaultdict

try:
    import openpyxl
except ImportError:                                   # pragma: no cover
    print("openpyxl is required: pip install openpyxl")
    raise SystemExit(2)

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from salon_business_spreadsheet.build import REGISTER_MODULES  # noqa: E402
from salon_business_spreadsheet.config import (COMMISSION_TIERS,  # noqa: E402
                                               Config, rid)

MONEY = 0.015           # a cent and a half of rounding slack


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


def check(path: str) -> int:                          # noqa: C901
    cfg = Config()
    specs = _specs(cfg)
    wb = openpyxl.load_workbook(path, data_only=False)
    problems: list = []
    notes: list = []

    data = {}
    anchors = {
        "Services": "name", "Clients": "name", "Staff": "name",
        "Appointments": "date", "Sales": "date", "Payments": "date",
        "Commissions": "month", "Payroll": "period_end",
        "Expenses": "date", "Inventory": "name", "Purchases": "date",
        "Marketing": "name", "Loyalty": "client",
    }
    for name, anchor in anchors.items():
        if name not in wb.sheetnames:
            problems.append(f"missing sheet {name}")
            continue
        data[name] = _rows(wb[name], specs[name], anchor)
    if problems:
        for line in problems:
            print("    - " + line)
        return 1

    counts = ", ".join(f"{name} {len(rows)}" for name, rows in data.items())
    notes.append("demo rows: " + counts)
    if not data["Appointments"]:
        print(f"--- {path}\n  . no example data in this build (blank "
              f"version) - nothing to check")
        return 0

    # ---------------------------------------------------- the ID spine --
    service_by_id = {rid("Service", row["_row"]): row
                     for row in data["Services"]}
    client_by_id = {rid("Client", row["_row"]): row
                    for row in data["Clients"]}
    staff_by_id = {rid("Staff", row["_row"]): row for row in data["Staff"]}
    product_by_id = {rid("Product", row["_row"]): row
                     for row in data["Inventory"]}
    appt_by_id = {rid("Appointment", row["_row"]): row
                  for row in data["Appointments"]}
    sale_by_id = {rid("Sale", row["_row"]): row for row in data["Sales"]}

    def known(table, value, where, kind):
        if value in (None, ""):
            return False
        if value not in table:
            problems.append(f"{where}: {kind} '{value}' does not exist")
            return False
        return True

    for row in data["Appointments"]:
        where = f"Appointments row {row['_row']}"
        known(client_by_id, row.get("client"), where, "client")
        known(service_by_id, row.get("service"), where, "service")
        known(staff_by_id, row.get("staff"), where, "stylist")
    for row in data["Sales"]:
        where = f"Sales row {row['_row']}"
        if row.get("appointment"):
            known(appt_by_id, row["appointment"], where, "appointment")
        if row.get("client"):
            known(client_by_id, row["client"], where, "client")
        if row.get("staff"):
            known(staff_by_id, row["staff"], where, "stylist")
        if row.get("kind") == "Service":
            known(service_by_id, row.get("item"), where, "service")
        if row.get("kind") == "Retail":
            known(product_by_id, row.get("item"), where, "product")
    for row in data["Payments"]:
        where = f"Payments row {row['_row']}"
        if row.get("sale"):
            known(sale_by_id, row["sale"], where, "sale")
        if row.get("appointment"):
            known(appt_by_id, row["appointment"], where, "appointment")

    # ------------------------------------------------- double bookings --
    diary = defaultdict(list)
    for row in data["Appointments"]:
        if row.get("status") == "Cancelled":
            continue
        service = service_by_id.get(row.get("service"), {})
        minutes = row.get("minutes") or service.get("duration") or 0
        buffer_ = service.get("buffer") or 0
        start = _minutes(row.get("start"))
        finish = start + int(minutes) + int(buffer_)
        diary[(row.get("staff"), _day(row.get("date")))].append(
            (start, finish, row["_row"]))
    clashes = 0
    for (who, day), slots in diary.items():
        slots.sort()
        for first, second in zip(slots, slots[1:]):
            if second[0] < first[1]:
                clashes += 1
                if clashes < 6:
                    problems.append(
                        f"double booking: {who} on {day}, rows "
                        f"{first[2]} and {second[2]}")
    notes.append(f"diary: {len(data['Appointments'])} bookings, "
                 f"{clashes} clashes")

    # --------------------------------------------- every visit rung up --
    sales_by_appt = defaultdict(list)
    for row in data["Sales"]:
        if row.get("appointment"):
            sales_by_appt[row["appointment"]].append(row)
    missing_sale = 0
    for ref, row in appt_by_id.items():
        if row.get("status") != "Completed":
            continue
        if row.get("price") == 0:
            continue                                  # chair renter
        if not sales_by_appt.get(ref):
            missing_sale += 1
    if missing_sale:
        problems.append(f"{missing_sale} completed visits have no sale "
                        f"logged against them")

    # ----------------------------------------------- money against till --
    def sale_total(row):
        kind = row.get("kind")
        qty = row.get("qty") or 1
        if row.get("unit_price") not in (None, ""):
            price = float(row["unit_price"]) * qty
        elif kind == "Service":
            price = float(service_by_id.get(row.get("item"), {})
                          .get("price") or 0)
        elif kind == "Retail":
            price = float(product_by_id.get(row.get("item"), {})
                          .get("retail_price") or 0) * qty
        else:
            price = 0.0
        discount = round(price * float(row.get("discount_pct") or 0), 2) + \
            float(row.get("discount_amt") or 0)
        net = round(price - discount - float(row.get("refund") or 0), 2)
        taxed = ((kind == "Retail" and cfg.tax_on_retail == "Yes") or
                 (kind != "Retail" and cfg.tax_on_services == "Yes"))
        tax = round(net * cfg.tax_rate, 2) if taxed else 0.0
        return round(net + tax, 2), net

    paid_by_sale = Counter()
    deposits_by_appt = Counter()
    tips = 0.0
    for row in data["Payments"]:
        amount = float(row.get("amount") or 0)
        kind = row.get("kind")
        if kind == "Tip":
            tips += amount
            continue
        if kind == "Refund":
            amount = -amount
        if row.get("sale"):
            paid_by_sale[row["sale"]] += amount
        elif row.get("appointment"):
            deposits_by_appt[row["appointment"]] += amount
    first_of_appt = {}
    for ref, row in sale_by_id.items():
        appt = row.get("appointment")
        if appt and appt not in first_of_appt:
            first_of_appt[appt] = ref
    unpaid = 0
    overpaid = 0
    net_sales = 0.0
    gross_by_month = Counter()
    for ref, row in sale_by_id.items():
        total, net = sale_total(row)
        net_sales += net
        gross_by_month[_month(row.get("date"))] += net
        paid = paid_by_sale.get(ref, 0.0)
        if first_of_appt.get(row.get("appointment")) == ref:
            paid += deposits_by_appt.get(row["appointment"], 0.0)
        balance = round(total - paid, 2)
        if balance > MONEY:
            unpaid += 1
        elif balance < -MONEY:
            overpaid += 1
            if overpaid < 6:
                problems.append(
                    f"Sales row {row['_row']}: paid {paid:,.2f} against "
                    f"a bill of {total:,.2f} - money applied twice?")
    notes.append(f"till: {len(sale_by_id)} sales, net {net_sales:,.0f}, "
                 f"{unpaid} left unpaid on purpose, {overpaid} overpaid")
    if unpaid > len(sale_by_id) * 0.1:
        problems.append(f"{unpaid} unpaid sales is too many for a demo "
                        f"({len(sale_by_id)} sales)")

    # ------------------------------------------------------- commission --
    by_staff_month = defaultdict(lambda: dict(service=0.0, retail=0.0,
                                              cost=0.0, count=0))
    for ref, row in sale_by_id.items():
        who = row.get("staff")
        if not who:
            continue
        bucket = by_staff_month[(who, _month(row.get("date")))]
        _total, net = sale_total(row)
        kind = row.get("kind")
        if kind == "Service":
            service = service_by_id.get(row.get("item"), {})
            bucket["service"] += net
            bucket["cost"] += float(service.get("product_cost") or 0)
            bucket["count"] += 1
        elif kind == "Retail":
            bucket["retail"] += net

    def tier_rate(base: float) -> float:
        rate = COMMISSION_TIERS[0][1]
        for floor, value in COMMISSION_TIERS:
            if base >= floor:
                rate = value
        return rate

    def commission_for(who: str, month: str) -> float:
        person = staff_by_id.get(who, {})
        bucket = by_staff_month.get((who, month))
        if not bucket:
            return 0.0
        method = person.get("method") or "None"
        service = round(bucket["service"], 2)
        retail = round(bucket["retail"], 2)
        cost = round(bucket["cost"], 2)
        rate = float(person.get("rate_service") or cfg.service_commission)
        retail_rate = float(person.get("rate_retail") or
                            cfg.retail_commission)
        if method == "Percent of service":
            return round(service * rate, 2)
        if method == "Percent of service and retail":
            return round(round(service * rate, 2) +
                         round(retail * retail_rate, 2), 2)
        if method == "Percent after product cost":
            return round(max(0.0, service - cost) * rate, 2)
        if method == "Tiered on monthly service revenue":
            return round(service * tier_rate(service), 2)
        if method == "Flat amount per service":
            return round(bucket["count"] *
                         float(person.get("flat_rate") or 0), 2)
        return 0.0

    seen = Counter()
    for row in data["Commissions"]:
        key = (row.get("staff"), row.get("month"))
        seen[key] += 1
        if seen[key] > 1:
            problems.append(f"Commissions row {row['_row']}: a second row "
                            f"for {key[0]} in {key[1]}")
        if row.get("staff") not in staff_by_id:
            problems.append(f"Commissions row {row['_row']}: unknown "
                            f"stylist {row.get('staff')}")
        if not re.match(r"^\d{4}-\d{2}$", str(row.get("month") or "")):
            problems.append(f"Commissions row {row['_row']}: month "
                            f"'{row.get('month')}' is not yyyy-mm")
    methods = {staff_by_id[k[0]].get("method")
               for k in seen if k[0] in staff_by_id}
    notes.append(f"commissions: {len(data['Commissions'])} rows covering "
                 f"{len(methods)} different pay methods")

    released = Counter()
    for row in data["Payroll"]:
        month = _month(row.get("period_end"))
        if row.get("commission"):
            released[(row.get("staff"), month)] += float(row["commission"])
        if row.get("staff") not in staff_by_id:
            problems.append(f"Payroll row {row['_row']}: unknown stylist")
        if not row.get("hours") and not row.get("commission"):
            problems.append(f"Payroll row {row['_row']}: neither hours "
                            f"nor commission")
    this_month = f"{cfg.today.year:04d}-{cfg.today.month:02d}"
    for (who, month), amount in released.items():
        if month == this_month:
            continue
        earned = commission_for(who, month)
        adjust = sum(float(row.get("adjust") or 0) + float(
            row.get("bonus") or 0)
            for row in data["Commissions"]
            if row.get("staff") == who and row.get("month") == month)
        if amount - (earned + adjust) > 1.0:
            problems.append(
                f"payroll releases {amount:,.2f} of commission to {who} "
                f"in {month} but only {earned + adjust:,.2f} was earned")
    notes.append(f"payroll: {len(data['Payroll'])} runs, "
                 f"{len(released)} stylist-months with commission")

    # ------------------------------------------------------------ stock --
    received = Counter()
    for row in data["Purchases"]:
        status = row.get("status")
        qty = float(row.get("qty") or 0)
        got = row.get("qty_received")
        if status == "Received":
            received[row.get("product")] += float(got) if got else qty
        elif status == "Part received":
            received[row.get("product")] += float(got or 0)
        if row.get("product") not in product_by_id:
            problems.append(f"Purchases row {row['_row']}: unknown product")
    sold = Counter()
    for row in data["Sales"]:
        if row.get("kind") == "Retail":
            sold[row.get("item")] += float(row.get("qty") or 1)
    negative = 0
    for ref, row in product_by_id.items():
        stock = (float(row.get("opening") or 0) + received[ref] -
                 float(row.get("used") or 0) - sold[ref] -
                 float(row.get("waste") or 0) +
                 float(row.get("adjust") or 0))
        if stock < 0:
            negative += 1
            if negative < 6:
                problems.append(
                    f"Inventory row {row['_row']} ({row.get('name')}): "
                    f"stock works out at {stock:,.1f}")
    notes.append(f"stock: {len(product_by_id)} products, {negative} would "
                 f"go negative")

    # -------------------------------------------------------- marketing --
    codes = [row.get("code") for row in data["Marketing"] if row.get("code")]
    duplicates = [code for code, n in Counter(codes).items() if n > 1]
    if duplicates:
        problems.append(f"marketing codes used twice: {duplicates}")
    used_codes = {row.get("promo") for row in data["Sales"]
                  if row.get("promo")}
    unknown = used_codes - set(codes)
    if unknown:
        problems.append(f"sales use promo codes that are not on the "
                        f"marketing tab: {sorted(unknown)}")
    notes.append(f"marketing: {len(codes)} codes, {len(used_codes)} of "
                 f"them actually used on sales")

    members = [row.get("client") for row in data["Loyalty"]]
    repeats = [c for c, n in Counter(members).items() if n > 1]
    if repeats:
        problems.append(f"loyalty: {len(repeats)} clients listed twice")

    # ------------------------------------------------------- the story --
    months = sorted(gross_by_month)
    recent = [m for m in months if m < this_month][-3:]
    if recent:
        average = sum(gross_by_month[m] for m in recent) / len(recent)
        notes.append("net sales by month: " + ", ".join(
            f"{m} {gross_by_month[m]:,.0f}" for m in months[-4:]))
        target = cfg.revenue_target_month
        if average < target * 0.6 or average > target * 1.6:
            problems.append(
                f"the demo averages {average:,.0f} a month against a "
                f"target of {target:,.0f} - the dashboard will look odd")
    completed = sum(1 for row in data["Appointments"]
                    if row.get("status") == "Completed")
    final = sum(1 for row in data["Appointments"]
                if row.get("status") in ("Completed", "Cancelled",
                                         "No show"))
    if final:
        notes.append(f"completion rate {completed / final:.0%}, "
                     f"average ticket "
                     f"{net_sales / max(1, completed):,.2f}, "
                     f"tips {tips:,.0f}")

    print(f"--- {path}")
    for note in notes:
        print(f"  . {note}")
    if problems:
        print(f"  ! {len(problems)} problem(s):")
        for line in problems[:30]:
            print("    - " + line)
        return 1
    print("  OK - the example salon adds up")
    return 0


if __name__ == "__main__":
    target = (sys.argv[1] if len(sys.argv) > 1
              else "output/Salon Business Management Spreadsheet.xlsx")
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

#### Appointments

The diary: who, with whom, how long, how much - and a double-booking alarm that checks every row.

34 columns, 1200 rows of capacity, 14 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Day | `weekday` | text | formula |
| D | Start | `start` | time | you type it (checked) |
| E | Client ID | `client` | id | drop down `ClientList` |
| F | Client | `client_name` | text | formula |
| G | Service ID | `service` | id | drop down `ServiceList` |
| H | Service | `service_name` | text | formula |
| I | Stylist ID | `staff` | id | drop down `StaffList` |
| J | Stylist | `staff_name` | text | formula |
| K | Minutes | `minutes` | int | you type it (checked) |
| L | Price | `price` | money | you type it (checked) |
| M | Status | `status` | text | drop down `AppointmentStatusList` |
| N | Booked via | `source` | text | drop down `SourceList` |
| O | Deposit asked | `deposit_due` | money | you type it (checked) |
| P | Notes | `notes` | wrap | you type it |
| Q | If cancelled, why | `cancel_reason` | text | drop down `CancelReasonList` |
| R | Late fee charged | `fee_charged` | money | you type it (checked) |
| S | Month | `month` | text | formula, hidden |
| T | Week | `week` | int | formula, hidden |
| U | Minutes used | `minutes_used` | int | formula, hidden |
| V | Finishes | `finish` | time | formula |
| W | Hour | `hour` | int | formula, hidden |
| X | Chair hours | `hours` | number1 | formula, hidden |
| Y | Price charged | `price_used` | money | formula |
| Z | Deposit wanted | `deposit_wanted` | money | formula, hidden |
| AA | Deposit paid | `deposit_paid` | money | formula |
| AB | Billed | `billed` | money | formula |
| AC | Final? | `is_final` | flag | formula, hidden |
| AD | Booked ahead | `lead_days` | int | formula, hidden |
| AE | Booked on | `logged` | date | you type it (checked) |
| AF | Rebooked? | `rebooked` | flag | formula |
| AG | Clash? | `clash` | flag | formula |
| AH | Action | `alert` | status | formula |

#### Clients

Who they are, what they have spent, when they were last in - and which of them you are about to lose.

43 columns, 500 rows of capacity, 17 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Client | `name` | text | you type it |
| C | Phone | `phone` | text | you type it |
| D | Email | `email` | text | you type it |
| E | Contact by | `contact` | text | drop down `ContactList` |
| F | First seen | `joined` | date | you type it (checked) |
| G | Found you by | `source` | text | drop down `SourceList` |
| H | Referred by | `referred_by` | text | you type it |
| I | Usual stylist | `staff` | id | drop down `StaffList` |
| J | Stylist | `staff_name` | text | formula, hidden |
| K | Birthday | `birthday` | date | you type it (checked) |
| L | Formula & preferences | `preferences` | wrap | you type it |
| M | Allergies & notes | `allergies` | wrap | you type it |
| N | Patch test | `patch_test` | date | you type it (checked) |
| O | Marketing OK? | `marketing_ok` | flag | you type it (checked) |
| P | Photos OK? | `photos_ok` | flag | you type it (checked) |
| Q | Tags | `tags` | text | you type it |
| R | Status | `set_status` | text | drop down `ClientStatusList` |
| S | Notes | `notes` | wrap | you type it |
| T | Visits | `visits` | int | formula |
| U | Visits, 12m | `visits_year` | int | formula, hidden |
| V | Cancelled | `cancels` | int | formula, hidden |
| W | No shows | `no_shows` | int | formula |
| X | First visit | `first_visit` | date | formula |
| Y | Last visit | `last_visit` | date | formula |
| Z | Days since | `days_since` | int | formula |
| AA | Next raw | `next_raw` | number | formula, hidden |
| AB | Next booked | `next_visit` | date | formula |
| AC | Usual gap | `gap_days` | int | formula |
| AD | Rebook by | `due_date` | date | formula |
| AE | Lifetime spend | `spend` | money | formula |
| AF | On services | `spend_service` | money | formula, hidden |
| AG | On retail | `spend_retail` | money | formula, hidden |
| AH | Spend, 12m | `spend_year` | money | formula |
| AI | Average ticket | `avg_ticket` | money | formula |
| AJ | A year is worth | `yearly_value` | money | formula |
| AK | Balance owed | `balance` | money | formula |
| AL | Loyalty points | `points` | int | formula |
| AM | Segment | `segment` | status | formula |
| AN | Came back? | `retained` | flag | formula, hidden |
| AO | Spend rank | `spend_rank` | int | formula, hidden |
| AP | Visit rank | `visits_rank` | int | formula, hidden |
| AQ | Action | `alert` | status | formula |

#### Services

The menu, the chair time, the product cost and what each treatment actually earns an hour.

39 columns, 200 rows of capacity, 14 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Service | `name` | text | you type it |
| C | Category | `category` | text | drop down `ServiceCategoryList` |
| D | Description | `description` | wrap | you type it |
| E | Minutes | `duration` | int | you type it (checked) |
| F | Turnaround | `buffer` | int | you type it (checked) |
| G | Price | `price` | money | you type it (checked) |
| H | Product cost | `product_cost` | money | you type it (checked) |
| I | Add-on? | `addon` | flag | you type it (checked) |
| J | Pays commission? | `commissionable` | flag | you type it (checked) |
| K | Bookable online? | `online` | flag | you type it (checked) |
| L | Stylist level | `level` | text | drop down `StylistLevelList` |
| M | Status | `status` | text | drop down `ServiceStatusList` |
| N | Price set on | `reviewed` | date | you type it (checked) |
| O | Notes | `notes` | wrap | you type it |
| P | Chair hours | `hours` | number | formula, hidden |
| Q | Product cost used | `cost_product` | money | formula |
| R | Commission | `cost_commission` | money | formula |
| S | Direct cost | `cost_total` | money | formula |
| T | Gross profit | `profit` | money | formula |
| U | Margin | `margin` | percent | formula |
| V | Price an hour | `price_hour` | money | formula |
| W | Profit an hour | `profit_hour` | money | formula |
| X | Suggested, raw | `suggested_raw` | money | formula, hidden |
| Y | Suggested price | `suggested` | money | formula |
| Z | Price gap | `gap` | money | formula |
| AA | Booked | `bookings` | int | formula |
| AB | Completed | `done` | int | formula |
| AC | Sold | `units` | int | formula, hidden |
| AD | Revenue | `revenue` | money | formula |
| AE | Share of sales | `revenue_share` | percent | formula, hidden |
| AF | Profit earned | `profit_total` | money | formula, hidden |
| AG | Chair hours used | `hours_total` | number1 | formula, hidden |
| AH | Retail attached | `retail_attach` | percent | formula |
| AI | Last done | `last_done` | date | formula |
| AJ | Revenue rank | `revenue_rank` | int | formula, hidden |
| AK | Rate rank | `profit_hour_rank` | int | formula, hidden |
| AL | Popularity | `done_rank` | int | formula, hidden |
| AM | Action | `alert` | status | formula |

#### Sales

Every service, product and voucher sold - gross, discount, refund, net, tax and the profit left over.

38 columns, 1200 rows of capacity, 14 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Appointment | `appointment` | id | you type it |
| D | Client ID | `client` | id | drop down `ClientList` |
| E | Client | `client_name` | text | formula, hidden |
| F | Stylist ID | `staff` | id | drop down `StaffList` |
| G | Stylist | `staff_name` | text | formula, hidden |
| H | Kind | `kind` | text | drop down `SaleKindList` |
| I | Service / product | `item` | id | you type it |
| J | Description | `description` | text | you type it |
| K | Qty | `qty` | number1 | you type it (checked) |
| L | Unit price | `unit_price` | money | you type it (checked) |
| M | Discount % | `discount_pct` | percent | you type it (checked) |
| N | Discount $ | `discount_amt` | money | you type it (checked) |
| O | Promo code | `promo` | text | you type it |
| P | Refunded | `refund` | money | you type it (checked) |
| Q | Notes | `notes` | wrap | you type it |
| R | Month | `month` | text | formula, hidden |
| S | Week | `week` | int | formula, hidden |
| T | Qty used | `qty_used` | number1 | formula, hidden |
| U | Item | `item_name` | text | formula |
| V | Price used | `price_used` | money | formula, hidden |
| W | Gross | `gross` | money | formula |
| X | Discount | `discount` | money | formula |
| Y | Net sale | `net` | money | formula |
| Z | Tax | `tax` | money | formula |
| AA | Total due | `total` | money | formula |
| AB | Cost | `cost` | money | formula |
| AC | Gross profit | `profit` | money | formula |
| AD | First line? | `first_of_appt` | flag | formula, hidden |
| AE | Paid to this sale | `paid_direct` | money | formula, hidden |
| AF | Deposit applied | `paid_deposit` | money | formula, hidden |
| AG | Paid | `paid` | money | formula |
| AH | Balance | `balance` | money | formula |
| AI | Paid? | `status` | status | formula |
| AJ | Retail too? | `retail_with` | flag | formula, hidden |
| AK | Sale rank | `net_rank` | int | formula, hidden |
| AL | Check | `alert` | status | formula |

#### Payments

Every payment, deposit, tip and refund - what it settles, what the card machine took, what reached the bank.

23 columns, 1200 rows of capacity, 11 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Client ID | `client` | id | drop down `ClientList` |
| D | Client | `client_name` | text | formula, hidden |
| E | What for | `kind` | text | drop down `PaymentKindList` |
| F | Method | `method` | text | drop down `PaymentMethodList` |
| G | Amount | `amount` | money | you type it (checked) |
| H | Sale ID | `sale` | id | you type it |
| I | Appointment | `appointment` | id | you type it |
| J | Stylist ID | `staff` | id | drop down `StaffList` |
| K | Stylist | `staff_name` | text | formula, hidden |
| L | Reference | `reference` | text | you type it |
| M | Reconciled? | `banked` | flag | you type it (checked) |
| N | Notes | `notes` | wrap | you type it |
| O | Month | `month` | text | formula, hidden |
| P | Week | `week` | int | formula, hidden |
| Q | Settles a bill? | `counts` | flag | formula, hidden |
| R | In or out | `direction` | int | formula, hidden |
| S | Processing fee | `fee` | money | formula |
| T | Into the bank | `net_cash` | money | formula |
| U | Tip? | `is_tip` | flag | formula, hidden |
| V | Applied to | `applied_to` | text | formula |
| W | Check | `alert` | status | formula |

#### Staff

The team, their pay basis, and what each chair brings in an hour.

41 columns, 40 rows of capacity, 18 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Name | `name` | text | you type it |
| C | Role | `role` | text | you type it |
| D | Level | `level` | text | drop down `StylistLevelList` |
| E | Employment | `type` | text | drop down `StaffTypeList` |
| F | Started | `start_date` | date | you type it (checked) |
| G | Status | `status` | text | drop down `StaffStatusList` |
| H | Phone | `phone` | text | you type it |
| I | Email | `email` | text | you type it |
| J | Hours a week | `hours_week` | number1 | you type it (checked) |
| K | Days worked | `days` | text | you type it |
| L | Commission method | `method` | text | drop down `CommissionMethodList` |
| M | Service rate | `rate_service` | percent | you type it (checked) |
| N | Retail rate | `rate_retail` | percent | you type it (checked) |
| O | Flat per service | `flat_rate` | money | you type it (checked) |
| P | Hourly rate | `hourly` | money | you type it (checked) |
| Q | Chair rent | `rent_week` | money | you type it (checked) |
| R | Monthly target | `target_month` | money | you type it (checked) |
| S | Notes | `notes` | wrap | you type it |
| T | Appointments | `appts` | int | formula, hidden |
| U | Completed | `done` | int | formula |
| V | No shows | `no_shows` | int | formula, hidden |
| W | Clients seen | `clients` | int | formula, hidden |
| X | Hours booked | `hours_booked` | number1 | formula |
| Y | Hours available | `hours_avail` | number1 | formula, hidden |
| Z | Utilisation | `utilisation` | percent | formula |
| AA | Service sales | `rev_service` | money | formula |
| AB | Retail sales | `rev_retail` | money | formula |
| AC | Total sales | `rev_total` | money | formula |
| AD | Sales this month | `rev_month` | money | formula |
| AE | Against target | `target_gap` | money | formula |
| AF | Average ticket | `avg_ticket` | money | formula |
| AG | Retail share | `retail_share` | percent | formula |
| AH | Earned an hour | `rate_hour` | money | formula |
| AI | Rebook rate | `rebook_rate` | percent | formula |
| AJ | Tips | `tips` | money | formula, hidden |
| AK | Commission | `commission_paid` | money | formula, hidden |
| AL | Pay this month | `cost_month` | money | formula |
| AM | Leaves behind | `margin_month` | money | formula |
| AN | Sales rank | `rev_total_rank` | int | formula, hidden |
| AO | Action | `alert` | status | formula |

#### Commissions

One row per stylist per month. The method is yours; the arithmetic is the workbook's.

27 columns, 240 rows of capacity, 7 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Month | `month` | text | drop down `MonthKeyList` |
| C | Stylist ID | `staff` | id | drop down `StaffList` |
| D | Stylist | `staff_name` | text | formula |
| E | Method used | `method` | text | formula |
| F | Adjustment | `adjust` | money | you type it (checked) |
| G | Bonus | `bonus` | money | you type it (checked) |
| H | Approved? | `approved` | flag | you type it (checked) |
| I | Paid on | `paid_on` | date | you type it (checked) |
| J | Notes | `notes` | wrap | you type it |
| K | Service sales | `service_net` | money | formula |
| L | Before discount | `service_gross` | money | formula, hidden |
| M | Retail sales | `retail_net` | money | formula |
| N | Retail before | `retail_gross` | money | formula, hidden |
| O | Product used | `product_cost` | money | formula |
| P | Services done | `services` | int | formula, hidden |
| Q | Service base | `base_service` | money | formula |
| R | Retail base | `base_retail` | money | formula, hidden |
| S | Service rate | `rate_service` | percent | formula |
| T | Retail rate | `rate_retail` | percent | formula, hidden |
| U | Flat per service | `flat_rate` | money | formula, hidden |
| V | Commission | `commission` | money | formula |
| W | Tips | `tips` | money | formula |
| X | To pay | `payout` | money | formula |
| Y | Share of their sales | `share` | percent | formula |
| Z | Commission rank | `commission_rank` | int | formula, hidden |
| AA | Check | `alert` | status | formula |

#### Payroll

Hours, commission released, tips passed on and the employer costs on top - the real cost of a pay run.

30 columns, 360 rows of capacity, 14 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Period ending | `period_end` | date | you type it (checked) |
| C | Run | `period` | text | drop down `PayPeriodList` |
| D | Stylist ID | `staff` | id | drop down `StaffList` |
| E | Name | `staff_name` | text | formula |
| F | Hours | `hours` | number1 | you type it (checked) |
| G | Hourly rate | `rate` | money | you type it (checked) |
| H | Commission paid | `commission` | money | you type it (checked) |
| I | Tips passed on | `tips` | money | you type it (checked) |
| J | Bonus | `bonus` | money | you type it (checked) |
| K | Holiday pay | `holiday` | money | you type it (checked) |
| L | Deductions | `deductions` | money | you type it (checked) |
| M | Paid? | `paid` | flag | you type it (checked) |
| N | Paid on | `paid_on` | date | you type it (checked) |
| O | Method | `method` | text | drop down `PaymentMethodList` |
| P | Notes | `notes` | wrap | you type it |
| Q | Month | `month` | text | formula, hidden |
| R | Week | `week` | int | formula, hidden |
| S | Rate used | `rate_used` | money | formula, hidden |
| T | Hourly pay | `base_pay` | money | formula |
| U | Gross pay | `gross_pay` | money | formula |
| V | After deductions | `net_pay` | money | formula |
| W | Including tips | `to_hand_over` | money | formula, hidden |
| X | Employer costs | `employer_cost` | money | formula |
| Y | Cost to salon | `total_cost` | money | formula |
| Z | Month calculated | `month_commission` | money | formula, hidden |
| AA | Month released | `month_paid` | money | formula, hidden |
| AB | Commission left | `commission_left` | money | formula |
| AC | Cost an hour | `cost_hour` | money | formula, hidden |
| AD | Check | `alert` | status | formula |

#### Expenses

Every cost, categorised the way the profit and loss needs it, with the recurring ones spread properly.

25 columns, 600 rows of capacity, 14 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Date | `date` | date | you type it (checked) |
| C | Category | `category` | text | drop down `ExpenseCategoryList` |
| D | Paid to | `payee` | text | you type it |
| E | What for | `description` | text | you type it |
| F | Amount | `amount` | money | you type it (checked) |
| G | Tax | `tax` | money | you type it (checked) |
| H | Method | `method` | text | drop down `PaymentMethodList` |
| I | Recurring? | `recurring` | flag | you type it (checked) |
| J | How often | `frequency` | text | drop down `FrequencyList` |
| K | Due | `due` | date | you type it (checked) |
| L | Paid? | `paid` | flag | you type it (checked) |
| M | Paid on | `paid_on` | date | you type it (checked) |
| N | Reference | `reference` | text | you type it |
| O | Notes | `notes` | wrap | you type it |
| P | Month | `month` | text | formula, hidden |
| Q | Week | `week` | int | formula, hidden |
| R | Total | `total` | money | formula |
| S | Monthly equivalent | `monthly` | money | formula |
| T | Fixed cost? | `is_fixed` | flag | formula, hidden |
| U | Days late | `days_late` | int | formula, hidden |
| V | Status | `status` | status | formula |
| W | Budget this month | `budget` | money | formula, hidden |
| X | Size rank | `total_rank` | int | formula, hidden |
| Y | Check | `alert` | status | formula |

#### Inventory

Retail and backbar stock, what is left, what it is worth and what to order.

36 columns, 300 rows of capacity, 18 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Product | `name` | text | you type it |
| C | Brand | `brand` | text | you type it |
| D | Category | `category` | text | drop down `ProductCategoryList` |
| E | Size | `size` | text | you type it |
| F | SKU / barcode | `sku` | text | you type it |
| G | Supplier | `supplier` | text | drop down `SupplierList` |
| H | Cost each | `unit_cost` | money | you type it (checked) |
| I | Retail price | `retail_price` | money | you type it (checked) |
| J | Opening stock | `opening` | number1 | you type it (checked) |
| K | Used in services | `used` | number1 | you type it (checked) |
| L | Waste | `waste` | number1 | you type it (checked) |
| M | Adjustment | `adjust` | number1 | you type it (checked) |
| N | Reorder at | `reorder_level` | number1 | you type it (checked) |
| O | Order quantity | `reorder_qty` | number1 | you type it (checked) |
| P | Last counted | `counted_on` | date | you type it (checked) |
| Q | Expires | `expires` | date | you type it (checked) |
| R | Sold to clients? | `retail_item` | flag | you type it (checked) |
| S | Notes | `notes` | wrap | you type it |
| T | Received | `received` | number1 | formula |
| U | Sold | `sold` | number1 | formula |
| V | Stock on hand | `stock` | number1 | formula |
| W | Used a week | `weekly_use` | number1 | formula, hidden |
| X | Weeks left | `weeks_cover` | number1 | formula |
| Y | Reorder point | `reorder_used` | number1 | formula, hidden |
| Z | Stock status | `status` | status | formula |
| AA | Order this many | `order_qty` | number1 | formula |
| AB | Value at cost | `value_cost` | money | formula |
| AC | Value at retail | `value_retail` | money | formula, hidden |
| AD | Margin | `margin` | percent | formula |
| AE | Suggested raw | `suggested_raw` | money | formula, hidden |
| AF | Suggested price | `suggested` | money | formula |
| AG | Retail revenue | `revenue` | money | formula |
| AH | Retail profit | `profit` | money | formula, hidden |
| AI | Seller rank | `revenue_rank` | int | formula, hidden |
| AJ | Action | `alert` | status | formula |

#### Purchases

What you ordered, what it cost, whether it turned up - and the supplier book underneath.

25 columns, 300 rows of capacity, 14 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Ordered | `date` | date | you type it (checked) |
| C | Supplier | `supplier` | text | drop down `SupplierList` |
| D | Product ID | `product` | id | drop down `ProductList` |
| E | Product | `product_name` | text | formula |
| F | Quantity | `qty` | number1 | you type it (checked) |
| G | Cost each | `unit_cost` | money | you type it (checked) |
| H | Delivery | `shipping` | money | you type it (checked) |
| I | Tax | `tax` | money | you type it (checked) |
| J | Status | `status` | text | drop down `PurchaseStatusList` |
| K | Expected | `expected` | date | you type it (checked) |
| L | Received | `received_on` | date | you type it (checked) |
| M | Quantity in | `qty_received` | number1 | you type it (checked) |
| N | Invoice ref | `invoice` | text | you type it |
| O | Paid? | `paid` | flag | you type it (checked) |
| P | Notes | `notes` | wrap | you type it |
| Q | Month | `month` | text | formula, hidden |
| R | Cost used | `cost_used` | money | formula, hidden |
| S | Goods | `goods` | money | formula |
| T | Order total | `total` | money | formula |
| U | Into stock | `qty_in` | number1 | formula |
| V | Short by | `short` | number1 | formula, hidden |
| W | Days to arrive | `lead_days` | int | formula, hidden |
| X | Cost moved | `price_change` | money | formula |
| Y | Check | `alert` | status | formula |

#### Marketing

Offers, codes and campaigns, measured against what they cost - money out and discount given away.

25 columns, 100 rows of capacity, 10 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Campaign | `name` | text | you type it |
| C | Channel | `channel` | text | drop down `MarketingChannelList` |
| D | The offer | `offer` | text | you type it |
| E | Promo code | `code` | text | you type it |
| F | Starts | `start` | date | you type it (checked) |
| G | Ends | `end` | date | you type it (checked) |
| H | Cost | `spend` | money | you type it (checked) |
| I | Revenue target | `target_rev` | money | you type it (checked) |
| J | Status | `status` | text | drop down `CampaignStatusList` |
| K | What happened | `notes` | wrap | you type it |
| L | Month | `month` | text | formula, hidden |
| M | Times used | `uses` | int | formula |
| N | Revenue | `revenue` | money | formula |
| O | Discount given | `given` | money | formula |
| P | Clients used it | `clients` | int | formula, hidden |
| Q | New clients | `new_clients` | int | formula |
| R | Bookings | `bookings` | int | formula, hidden |
| S | Total cost | `total_cost` | money | formula |
| T | Left over | `profit` | money | formula |
| U | Return | `roi` | percent | formula |
| V | Cost per new client | `cost_per_client` | money | formula |
| W | Against target | `vs_target` | money | formula, hidden |
| X | Best earner | `revenue_rank` | int | formula, hidden |
| Y | Verdict | `alert` | status | formula |

#### Loyalty

Points, rewards and referral credit - counted from the spend you have already logged.

24 columns, 500 rows of capacity, 10 of them typed in by hand.

| Col | Header | Key | Format | Filled by |
|-----|--------|-----|--------|-----------|
| A | ID | `id` | id | formula |
| B | Client ID | `client` | id | drop down `ClientList` |
| C | Client | `name` | text | formula |
| D | Joined | `joined` | date | you type it (checked) |
| E | Card number | `card` | text | you type it |
| F | Bonus points | `bonus_points` | int | you type it (checked) |
| G | Points redeemed | `redeemed` | int | you type it (checked) |
| H | Last reward | `last_reward` | date | you type it (checked) |
| I | Usual reward | `reward_type` | text | drop down `LoyaltyRewardList` |
| J | Referral credit paid | `referral_paid` | money | you type it (checked) |
| K | On the scheme? | `active` | flag | you type it (checked) |
| L | Notes | `notes` | wrap | you type it |
| M | Spend counted | `spend` | money | formula |
| N | Visits | `visits` | int | formula, hidden |
| O | Points earned | `earned` | int | formula |
| P | Points balance | `balance` | int | formula |
| Q | Rewards due | `rewards_due` | int | formula, hidden |
| R | Reward owed | `reward_value` | money | formula |
| S | Points to next | `to_next` | int | formula |
| T | People referred | `referrals` | int | formula |
| U | Referral credit owed | `referral_owed` | money | formula |
| V | Segment | `segment` | text | formula, hidden |
| W | Points rank | `balance_rank` | int | formula, hidden |
| X | Action | `alert` | status | formula |

---

*10,677 lines of Python in 37 files.*
